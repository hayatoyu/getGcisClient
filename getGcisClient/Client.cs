using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Net;
using System.Net.Sockets;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System.Windows.Forms;
using getGcisClient.Output;

namespace getGcisClient
{
    class Client : QueryAction
    {
        

        /*
         * IP 和 Port 有預設值，存放在 AppConfig檔案中
         */
        public Client(string IPAddr, string port, string FilePath, string OutputFolder, Form1 form)
        {
            this.IPAddr = IPAddr;
            try
            {
                this.port = int.Parse(port);
                this.TimeOut = int.Parse(ConfigurationManager.AppSettings["timeOut"]);
            }
            catch (Exception e)
            {
                PrintErrMsgToConsole(e);
                this.port = 1357;
                this.TimeOut = 60000;
            }
            this.FilePath = FilePath;
            this.SaveFolder = OutputFolder;
            this.comList = new List<CompanyInfo>();
            this.myForm = form;
        }

        public override void StartQuery()        
        {

            if (client != null)
                client.Close();
            client = new TcpClient();
            try
            {
                client.Connect(IPAddr, port);
                client.ReceiveBufferSize = 1024;
                CommunicateWithServer();
            }
            catch (Exception e)
            {
                PrintErrMsgToConsole(e);
            }
            finally
            {
                // 在 CommunicateWithServer() 結束後，才將連線按鈕開啟。因為該按鈕是由主執行緒產生的，必須委託給主執行緒去處理。
                myForm.Invoke(new UpdateBtn(() => { myForm.btn_Connect.Enabled = true; }));
            }
            
            
        }



        /*
* 當收到 Server 傳來的訊息時
*  1. added 表示已經收到連線請求，並排入等待。
*  2. ready 表示 Server 已經準備好接收查詢列表。
*  3. 收到 result: 開頭的訊息，表示此為 Server 回傳的公司基本資料。
*  4. finish 表示查詢作業已經完成。
*  Server 回傳的資料因為是一筆一筆回傳，基本上不會超過 buffer 大小。
*  本來有想過全部查完再一次回傳，但風險太高，途中如果發生錯誤，那 Client 可能經過漫長等待還無法收到回應。
*  一筆一筆傳先儲存在 Client 端，即使 Server 發生錯誤，Client 也能輸出部分資料，下次再從還沒查的地方繼續。
*/
        private void CommunicateWithServer()
        {
            int reclength = 0;
            NetworkStream netStream = client.GetStream();
            byte[] buffer = new byte[1024];
            List<CompanyInfoResult> result = new List<CompanyInfoResult>();
            string savepath = SaveFolder + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss");
            // 這裡要讀取使用者選擇的輸出檔案類型，同樣要委託主執行緒去取得。
            savepath += myForm.Invoke(new returnFileType(() => { return myForm.cb_OutputType.SelectedItem.ToString(); }));
            OutputFile writer = null;
            if (savepath.EndsWith(".csv"))
                writer = new OutputCsv(result);
            else
                writer = new OutputExcel(result);
            while (client.Connected)
            {
                try
                {
                    if ((reclength = netStream.Read(buffer, 0, buffer.Length)) != 0)
                    {
                        string recMessage = Encoding.UTF8.GetString(buffer, 0, reclength);
                        Console.WriteLine(recMessage);

                        if (recMessage.Equals("finish"))
                        {
                            writer.Output(savepath);
                            
                            result.Clear();
                            client.Close();
                            MessageBox.Show("匯出完成！");

                        }
                        else if (recMessage.Equals("added"))
                        {
                            Console.WriteLine("已連線，等待接受服務...");
                        }
                        else if (recMessage.Equals("ready"))
                        {
                            Console.WriteLine("開始讀取 Excel 公司列表...");
                            comList.Clear();
                            ReadFromExcel();
                            ComRequest request = new ComRequest { comList = comList.ToArray() };
                            SendToServer(netStream, JsonConvert.SerializeObject(request));
                            Console.WriteLine("送出查詢清單...");
                            client.ReceiveTimeout = TimeOut;
                        }
                        else if (recMessage.Equals("JsonErr"))
                        {
                            Console.WriteLine("商業司 API 回傳錯誤，請洽程式開發人員");
                            writer.Output(savepath);
                            result.Clear();
                            client.Close();
                            MessageBox.Show("商業司 API 錯誤，若有部份資料已導出，請洽程式開發人員");
                        }
                        else if (recMessage.StartsWith("result:"))
                        {
                            // 公司資料
                            try
                            {
                                recMessage = recMessage.Replace("result:", string.Empty);
                                var info = JsonConvert.DeserializeObject<CompanyInfoResult>(recMessage);
                                result.Add(info);
                            }
                            catch (Exception e)
                            {
                                PrintErrMsgToConsole(e);
                            }
                        }
                    }
                }
                catch (IOException e)
                {
                    // Server斷線
                    MessageBox.Show(e.Message);
                    netStream.Close();
                    client.Close();
                }


            }
            if (result.Count > 0)
            {
                writer.Output(savepath);
                MessageBox.Show("連線中斷，先導出部分資料");
            }
                
            

        }        
        

        private void SendToServer(NetworkStream ns, string content)
        {
            byte[] sendByte = Encoding.UTF8.GetBytes(content);
            try
            {
                ns.Write(sendByte, 0, sendByte.Length);
                ns.Flush();
            }
            catch(Exception e)
            {
                PrintErrMsgToConsole(e);
            }
        }

        
    }
}
