using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace getGcisClient.Output
{
    /*
     * 輸出檔案的抽象類
     * 建構子則是先加入查詢好的結果集
     */
    public abstract class OutputFile
    {
        protected List<CompanyInfoResult> result { get; set; }

        public OutputFile(List<CompanyInfoResult> result)
        {
            this.result = result;
        }

        public abstract void Output(string FilePath);

        protected List<CompanyInfoResult> CompanySort(List<CompanyInfoResult> result)
        {
            return (from c in result orderby c.Paid_In_Capital_Amount,c.Capital_Stock_Amount,c.Business_Accounting_NO, c.Company_Name ascending select c).ToList<CompanyInfoResult>();
        }
    }
}
