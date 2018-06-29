using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using System.Threading;

namespace getGcisClient.Output
{
    /*
     * 輸出成 Excel
     * 能有 xls 和 xlsx 兩種格式
     */
    class OutputExcel : OutputFile
    {
        public OutputExcel(List<CompanyInfoResult> result) : base(result) { }
        
        public override void Output(string FilePath)
        {
            IWorkbook wb = null;
            if (FilePath.EndsWith(".xlsx"))
                wb = new XSSFWorkbook();
            else
                wb = new HSSFWorkbook();
            ISheet ws = wb.CreateSheet();
            IRow TitleRow = ws.CreateRow(0);
            IRow DataRow;
            TitleRow.CreateCell(0).SetCellValue("公司名稱");
            TitleRow.CreateCell(1).SetCellValue("統一編號");
            TitleRow.CreateCell(2).SetCellValue("公司狀況描述");
            TitleRow.CreateCell(3).SetCellValue("資本總額(元)");
            TitleRow.CreateCell(4).SetCellValue("實收資本額(元)");
            TitleRow.CreateCell(5).SetCellValue("代表人姓名");
            TitleRow.CreateCell(6).SetCellValue("公司地址");
            TitleRow.CreateCell(7).SetCellValue("登記機關名稱");
            TitleRow.CreateCell(8).SetCellValue("核准設立日期");
            TitleRow.CreateCell(9).SetCellValue("最後核准變更日期");
            TitleRow.CreateCell(10).SetCellValue("備註");

            result = CompanySort(result);
            for (int i = 0; i < result.Count; i++)
            {
                string Note = string.Empty;
                DataRow = ws.CreateRow(i + 1);
                DataRow.CreateCell(0).SetCellValue(result[i].Company_Name);
                DataRow.CreateCell(1).SetCellValue(result[i].Business_Accounting_NO);
                DataRow.CreateCell(2).SetCellValue(result[i].Company_Status_Desc);
                if (result[i].Capital_Stock_Amount != null)
                    DataRow.CreateCell(3).SetCellValue(result[i].Capital_Stock_Amount.ToString());
                else
                    DataRow.CreateCell(3).SetCellValue("0");
                if (result[i].Paid_In_Capital_Amount != null)
                    DataRow.CreateCell(4).SetCellValue(result[i].Paid_In_Capital_Amount.ToString());
                else
                    DataRow.CreateCell(4).SetCellValue("0");
                DataRow.CreateCell(5).SetCellValue(result[i].Responsible_Name);
                DataRow.CreateCell(6).SetCellValue(result[i].Company_Location);
                DataRow.CreateCell(7).SetCellValue(result[i].Register_Organization_Desc);
                DataRow.CreateCell(8).SetCellValue(result[i].Company_Setup_Date);
                DataRow.CreateCell(9).SetCellValue(result[i].Change_Of_Approval_Data);
                if (result[i].Duplicate)
                {
                    Note += "此名稱有多筆公司資料，";
                    Note += result[i].NameMatch ? "系統已搜尋到公司名稱完全吻合之公司資料。" : "系統未搜尋到公司名稱完全吻合之公司資料，建議人工檢查。";
                }                                    
                if (result[i].ErrNotice)
                    Note += "查詢此筆資料時連線錯誤，請人工確認是否正確";
                if (result[i].NoData)
                    Note += "查無資料！";
                if (!string.IsNullOrEmpty(Note))
                    DataRow.CreateCell(10).SetCellValue(Note);

            }
            using (FileStream fs = new FileStream(FilePath, FileMode.Create))
            {
                wb.Write(fs);
                Thread.Sleep(3000);
                wb.Close();
            }
        }

    }
}
