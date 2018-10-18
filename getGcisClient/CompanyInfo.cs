using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace getGcisClient
{
    /*
     * 詳見 Server 端
     */
    public class CompanyInfo
    {
        [JsonProperty("Business_Accounting_NO")]
        public string Business_Accounting_NO { get; set; }

        [JsonProperty("Company_Status_Desc")]
        public string Company_Status_Desc { get; set; }

        [JsonProperty("Company_Name")]
        public string Company_Name { get; set; }

        [JsonProperty("Capital_Stock_Amount")]
        public long? Capital_Stock_Amount { get; set; }

        [JsonProperty("Paid_In_Capital_Amount")]
        public long? Paid_In_Capital_Amount { get; set; }

        [JsonProperty("Responsible_Name")]
        public string Responsible_Name { get; set; }

        [JsonProperty("Company_Location")]
        public string Company_Location { get; set; }

        [JsonProperty("Register_Organization_Desc")]
        public string Register_Organization_Desc { get; set; }

        [JsonProperty("Company_Setup_Date")]
        public string Company_Setup_Date { get; set; }

        [JsonProperty("Change_Of_Approval_Data")]
        public string Change_Of_Approval_Data { get; set; }

        [JsonProperty(Required = Required.Default)]
        public string Revoke_App_Date { get; set; }

        [JsonProperty(Required = Required.Default)]
        public string Case_Status { get; set; }

        [JsonProperty(Required = Required.Default)]
        public string Case_Status_Desc { get; set; }

        [JsonProperty(Required = Required.Default)]
        public string Sus_App_Date { get; set; }

        [JsonProperty(Required = Required.Default)]
        public string Sus_Beg_Date { get; set; }

        [JsonProperty(Required = Required.Default)]
        public string Sus_End_Date { get; set; }
    }

    public class CompanyInfoResult : CompanyInfo
    {
        [JsonProperty("Duplicate")]
        public bool Duplicate { get; set; }

        [JsonProperty("NameMatch")]
        public bool NameMatch { get; set; }

        [JsonProperty("ErrNotice")]
        public bool ErrNotice { get; set; }

        [JsonProperty("NoData")]
        public bool NoData { get; set; }

        
    }
}
