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

        protected OutputFile(List<CompanyInfoResult> result)
        {
            this.result = result;
        }

        public abstract void Output(string FilePath);
    }
}
