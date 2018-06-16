using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace getGcisClient.Output
{
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
