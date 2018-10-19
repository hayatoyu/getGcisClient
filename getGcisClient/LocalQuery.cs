using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Net;
using System.IO;
using Newtonsoft.Json;
using System.Windows.Forms;
using getGcisClient.Output;

namespace getGcisClient
{
    class LocalQuery : QueryAction
    {
        private List<CompanyInfoResult> comResults;
        
        public LocalQuery(string FilePath,string SaveFolder,Form1 form1)
        {
            this.FilePath = FilePath;
            this.SaveFolder = SaveFolder;
            this.myForm = form1;
            comResults = new List<CompanyInfoResult>();
            comList = new List<CompanyInfo>();
        }

        public override void StartQuery()
        {
            string savepath = SaveFolder + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss");            
            savepath += myForm.cb_OutputType.SelectedItem.ToString();
            OutputFile writer = null;
            if (savepath.EndsWith(".csv"))
                writer = new OutputCsv(comResults);
            else
                writer = new OutputExcel(comResults);
            
            StringBuilder stbr = new StringBuilder();
            string paramCom = "Company_Name like comName and Company_Status eq 01";
            string paramID = "Business_Accounting_NO eq comID";
            int errCount = 0, index = 0;
            string comName = string.Empty,comID = string.Empty;            
            Console.WriteLine("程式將從本地直接查詢商業司 API...");            
            ReadFromExcel();
            Console.WriteLine("讀取公司列表完成...準備開始查詢 API...");
            Console.WriteLine("共有 {0} 條資料待查詢...", comList.Count);
            
            while(index < comList.Count)
            {
                if(index % 100 == 0 && index > 0)
                {
                    Console.WriteLine("已連續查詢 100 條，將等待 10 秒繼續...");
                    Thread.Sleep(10000);
                }
                var currentCom = comList[index];
                if(!string.IsNullOrEmpty(currentCom.Company_Name))
                    comName = currentCom.Company_Name.Trim();
                if(!string.IsNullOrEmpty(currentCom.Business_Accounting_NO))
                    comID = currentCom.Business_Accounting_NO.Trim();
                stbr.Clear();

                // 這邊分成用公司名稱和統編兩種
                // 統編優先
                if(!string.IsNullOrEmpty(comID))
                {
                    stbr.Append("http://").Append("data.gcis.nat.gov.tw")
                        .Append("/od/data/api/5F64D864-61CB-4D0D-8AD9-492047CC1EA6")
                        .Append("?$format=json&$filter=")
                        .Append(paramID.Replace("comID", comID))
                        .Append("&$skip=0&$top=50");
                }
                else
                {
                    stbr.Append("http://").Append("data.gcis.nat.gov.tw")
                                .Append("/od/data/api/6BBA2268-1367-4B42-9CCA-BC17499EBE8C")
                                .Append("?$format=json&$filter=")
                                .Append(paramCom.Replace("comName", comName));
                }
                
                Console.WriteLine("開始查詢第 {0} / {1} 條資料： {2} ", index + 1, comList.Count, comName);
                HttpWebResponse response = null;

                try
                {
                    WebRequest request = WebRequest.Create(stbr.ToString());
                    response = request.GetResponse() as HttpWebResponse;
                }
                catch(WebException e)
                {
                    Console.WriteLine(e.Message);
                }

                if(response != null && response.StatusCode == HttpStatusCode.OK)
                {
                    try
                    {
                        using (Stream stream = response.GetResponseStream())
                        {
                            using (StreamReader reader = new StreamReader(stream))
                            {
                                string resFromAPI = reader.ReadToEnd();
                                CompanyInfoResult comResult = null;
                                if(!string.IsNullOrEmpty(resFromAPI))
                                {
                                    CompanyInfo[] comInfos = null;
                                    try
                                    {
                                        comInfos = JsonConvert.DeserializeObject<CompanyInfo[]>(resFromAPI);                                    
                                    }
                                    catch(JsonReaderException e)
                                    {
                                        Console.WriteLine(resFromAPI);
                                        PrintErrMsgToConsole(e);
                                        Console.WriteLine("{0} 查詢過程中發生錯誤，程式已中止", comName);
                                        Thread.Sleep(3000);
                                        writer.Output(savepath);
                                        MessageBox.Show("商業司 API 錯誤，部份資料已導出，請洽程式開發人員");                                        
                                        return;
                                    }
                                    CompanyInfo cInfo = null;
                                    bool NameMatch = false;
                                    if(comInfos.Length > 1)
                                    {
                                        cInfo = comInfos.Where(c => c.Company_Name.Equals(comName)).FirstOrDefault();
                                        NameMatch = (cInfo != default(CompanyInfo));
                                    }
                                    if(cInfo == null || cInfo == default(CompanyInfo))
                                    {
                                        cInfo = comInfos[0];
                                    }
                                    comResult = new CompanyInfoResult
                                    {
                                        Business_Accounting_NO = cInfo.Business_Accounting_NO,
                                        Company_Status_Desc = cInfo.Company_Status_Desc,
                                        Company_Name = cInfo.Company_Name,
                                        Capital_Stock_Amount = cInfo.Capital_Stock_Amount,
                                        Paid_In_Capital_Amount = cInfo.Paid_In_Capital_Amount,
                                        Responsible_Name = cInfo.Responsible_Name,
                                        Company_Location = cInfo.Company_Location,
                                        Register_Organization_Desc = cInfo.Register_Organization_Desc,
                                        Company_Setup_Date = cInfo.Company_Setup_Date,
                                        Change_Of_Approval_Data = cInfo.Change_Of_Approval_Data,
                                        Duplicate = (comInfos.Length > 1),
                                        NameMatch = NameMatch,
                                        ErrNotice = false
                                    };
                                    
                                }
                                else
                                {
                                    comResult = new CompanyInfoResult
                                    {
                                        Company_Name = comName,
                                        NoData = true
                                    };
                                }
                                //Console.WriteLine(JsonConvert.SerializeObject(comResult));
                                comResults.Add(comResult);
                                Thread.Sleep(2000);
                            }
                        }
                        Console.WriteLine("查詢 {0} 完成", comName);
                        index++;
                    }
                    catch(IOException e)
                    {
                        PrintErrMsgToConsole(e);
                        errCount++;
                        Console.WriteLine("查詢 {0} 時出現連線錯誤，將等候 10 秒重試...", comName);
                        Thread.Sleep(10000);
                        continue;
                    }
                    catch(JsonSerializationException e)
                    {
                        PrintErrMsgToConsole(e);
                        errCount++;
                        Console.WriteLine("查詢 {0} 時回應資料無法解析，將等候 5 秒 重試...", comName);
                        Thread.Sleep(5000);
                        continue;
                    }
                    catch(Exception e)
                    {
                        Console.WriteLine(e.Message);
                        errCount++;
                        Console.WriteLine("查詢 {0} 時發生不明原因錯誤，將等候 10 秒重試...", comName);
                        Thread.Sleep(10000);
                        continue;
                    }
                    finally
                    {
                        if(errCount >= 3)
                        {
                            index++;
                            errCount = 0;
                            CompanyInfoResult err = new CompanyInfoResult
                            {
                                Company_Name = comName,
                                ErrNotice = true
                            };
                            //Console.WriteLine(JsonConvert.SerializeObject(err));
                            comResults.Add(err);
                            Console.WriteLine("查詢 {0} 時發生錯誤已達 3 次，錯誤代碼 {1}，將暫時跳過", comName, response.StatusCode.ToString());
                            
                        }
                    }
                }
                else
                {
                    if(errCount >= 3)
                    {
                        index++;
                        errCount = 0;

                        CompanyInfoResult err = new CompanyInfoResult
                        {
                            Company_Name = comName,
                            ErrNotice = true
                        };
                        comResults.Add(err);
                        Console.WriteLine("查詢 {0} 時發生錯誤已達 3 次，將暫時跳過...", comName);
                        continue;
                    }
                    errCount++;
                    Console.WriteLine("查詢 {0} 時出現連線錯誤，將等候 10 秒重試...", comName);
                    Thread.Sleep(10000);
                    continue;
                }
            }

            Console.WriteLine("批次查詢作業完畢");
            writer.Output(savepath);
            MessageBox.Show("批次查詢作業完畢！");
        }
    }
}
