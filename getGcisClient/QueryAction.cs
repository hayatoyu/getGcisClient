using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Sockets;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;


namespace getGcisClient
{
    abstract class QueryAction
    {
        public string IPAddr { get; protected set; }
        public int port { get; protected set; }
        public string FilePath { get; protected set; }
        public string SaveFolder { get; protected set; }
        public int TimeOut { get; protected set; }
        protected TcpClient client;
        protected List<string> comList;
        protected Form1 myForm;
        protected delegate void UpdateBtn();
        protected delegate string returnFileType();

        public abstract void StartQuery();
        

        protected void ReadFromExcel()
        {
            IWorkbook wb = null;
            ISheet ws = null;
            IRow row = null;
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
                    row = ws.GetRow(index);

                    /* NPOI 對 Excel 的操作，沒有像 C# Cell.Value2(dynamic) 的用法
                     * 應該是 Java 本身就不支援 dynamic 型別所致
                     * 所以要讀取 Value 前，需要先將單元格轉換為對應的型別
                     * */

                    if (row.GetCell(1) == null)
                    {
                        break;
                    }
                    else
                    {
                        row.GetCell(1).SetCellType(CellType.String);
                        if (string.IsNullOrEmpty(row.GetCell(1).StringCellValue))
                        {
                            break;
                        }
                    }
                    index++;
                }
                row.GetCell(0).SetCellType(CellType.String);
                // 執行到最末行
                while (row != null
                    && !string.IsNullOrEmpty(row.GetCell(0).StringCellValue))
                {
                    comList.Add(row.GetCell(0).StringCellValue);
                    index++;
                    row = ws.GetRow(index);
                    //Console.WriteLine(index);
                    if (row != null && row.GetCell(0) != null)
                        row.GetCell(0).SetCellType(CellType.String);
                }
            }
        }

        protected void PrintErrMsgToConsole(Exception e)
        {
            Console.WriteLine("錯誤類型：{0}\n錯誤訊息：{1}\n堆疊：{2}", e.GetType(), e.Message, e.StackTrace);
        }
    }
}
