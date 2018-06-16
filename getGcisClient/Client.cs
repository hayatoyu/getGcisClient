using System;
using System.Collections.Generic;
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
    class Client
    {
        public string IPAddr { private set; get; }
        public int port { private set; get; }
        public string FilePath { private set; get; }
        public string OutputFolder { private set; get; }
        private TcpClient client;
        private List<string> comList;
        private Form1 myForm;
        delegate void UpdateBtn();
        delegate string returnFileType();

        public Client(string IPAddr, string port, string FilePath, string OutputFolder, Form1 form)
        {
            this.IPAddr = IPAddr;
            try
            {
                this.port = int.Parse(port);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            this.FilePath = FilePath;
            this.OutputFolder = OutputFolder;
            this.comList = new List<string>();
            this.myForm = form;
        }

        public void Connect()
        {

            if (client != null)
                client.Close();
            client = new TcpClient();
            try
            {
                client.Connect(IPAddr, port);
                client.ReceiveBufferSize = 1024;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            CommunicateWithServer();
            myForm.Invoke(new UpdateBtn(() => { myForm.btn_Connect.Enabled = true; }));
        }

        private void CommunicateWithServer()
        {
            int reclength = 0;
            NetworkStream netStream = client.GetStream();
            byte[] buffer = new byte[1024];
            List<CompanyInfoResult> result = new List<CompanyInfoResult>();
            string savepath = OutputFolder + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss");
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
                            MessageBox.Show("匯出完成！");
                            result.Clear();
                            client.Close();

                        }
                        else if (recMessage.Equals("added"))
                        {
                            Console.WriteLine("已連線，等待接受服務...");
                        }
                        else if (recMessage.Equals("ready"))
                        {
                            comList.Clear();
                            ReadFromExcel();
                            ComRequest request = new ComRequest { comList = comList.ToArray() };
                            SendToServer(netStream, JsonConvert.SerializeObject(request));
                        }
                        else if (recMessage.StartsWith("result:"))
                        {
                            // 一般資料
                            try
                            {
                                recMessage = recMessage.Replace("result:", string.Empty);
                                var info = JsonConvert.DeserializeObject<CompanyInfoResult>(recMessage);
                                result.Add(info);
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine(e.Message);
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

        private void ReadFromExcel()
        {
            IWorkbook wb = null;
            ISheet ws = null;
            int index = 1;
            using (FileStream fs = new FileStream(FilePath, FileMode.Open, FileAccess.Read))
            {
                if (FilePath.EndsWith(".xlsx"))
                    wb = new XSSFWorkbook(fs);
                else
                    wb = new HSSFWorkbook(fs);

                ws = wb.GetSheetAt(0);

                // 先取得要處理的最初行
                while (ws.GetRow(index) != null)
                {
                    if (ws.GetRow(index).GetCell(1) == null
                        || string.IsNullOrEmpty(ws.GetRow(index).GetCell(1).StringCellValue))
                        break;
                    index++;
                }

                // 執行到最末行
                while (ws.GetRow(index) != null
                    && !string.IsNullOrEmpty(ws.GetRow(index).GetCell(0).StringCellValue))
                {
                    comList.Add(ws.GetRow(index).GetCell(0).StringCellValue);
                    index++;
                }
            }
        }

        private void SendToServer(NetworkStream ns, string content)
        {
            byte[] sendByte = Encoding.UTF8.GetBytes(content);
            ns.Write(sendByte, 0, sendByte.Length);
            ns.Flush();
        }
    }
}
