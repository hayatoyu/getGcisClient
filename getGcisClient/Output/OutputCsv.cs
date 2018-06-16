using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace getGcisClient.Output
{
    /*
     * 輸出成 CSV
     */
    class OutputCsv : OutputFile
    {
        public OutputCsv(List<CompanyInfoResult> result) : base(result) { }

        public override void Output(string FilePath)
        {
            StringBuilder stbr = new StringBuilder();

            using (StreamWriter writer = new StreamWriter(FilePath,false,Encoding.UTF8))
            {
                stbr.Append("公司名稱,").Append("統一編號,").Append("公司狀況描述,")
                    .Append("資本總額(元),").Append("實收資本額(元),").Append("代表人姓名,")
                    .Append("公司地址,").Append("登記機關名稱,").Append("核准設立日期,")
                    .Append("最後核准變更日期,").AppendLine("備註");

                writer.Write(stbr.ToString());
                stbr.Clear();
                for(int i = 0;i < result.Count;i++)
                {
                    string Note = string.Empty;
                    stbr.Append(result[i].Company_Name + ",").Append(result[i].Business_Accounting_NO + ",")
                        .Append(result[i].Company_Status_Desc + ",");
                    if (result[i].Capital_Stock_Amount != null)
                        stbr.Append(result[i].Capital_Stock_Amount + ",");
                    else
                        stbr.Append("0,");
                    if (result[i].Paid_In_Capital_Amount != null)
                        stbr.Append(result[i].Paid_In_Capital_Amount + ",");
                    else
                        stbr.Append("0,");
                    stbr.Append(result[i].Responsible_Name + ",").Append(result[i].Company_Location + ",")
                        .Append(result[i].Register_Organization_Desc + ",").Append(result[i].Company_Setup_Date + ",")
                        .Append(result[i].Change_Of_Approval_Data);
                    if (result[i].Duplicate)
                        Note += "此名稱有多筆公司資料";
                    if (result[i].ErrNotice)
                        Note += "查詢此筆資料時連線錯誤，請人工確認是否正確";
                    if (result[i].NoData)
                        Note += "查無資料！";
                    if (!string.IsNullOrEmpty(Note))
                        stbr.Append("," + Note);

                    writer.WriteLine(stbr.ToString());
                    stbr.Clear();
                }


            }
        }
    }
}
