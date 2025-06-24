using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GH_LINK_ATM_ROU_INDIVIDUAL_AND_BULK_RECONS
{
    public class GhLinkReport
    {
        public string NARRATIVE { get; set; }

        public string LCY_AMT { get; set; }

        public string FCY_AMT { get; set; }

        public string AC_CCY { get; set; }

        public string SOURCE_REF_NO { get; set; }

        public string AUTH_ID { get; set; }

        public string AC_BRANCH { get; set; }

        public string TRN_REF_NO { get; set; }

        public string ADDL_TEXT { get; set; }

        public string TRN_DT { get; set; }

        public string VALUE_DT { get; set; }

        public string TRN_DESC { get; set; }

        public string DRCR { get; set; }

        public string USER_ID { get; set; }

        public string AC_ENTRY_SR_NO { get; set; }

        public string AC_NO { get; set; }

        public string ATM_RRN { get; set; }

    }

    public class GhLinkInd
    {
        public string VALUE_DT { get; set; }

        public string ADDL_TEXT { get; set; }

        public string ATM_RRN { get; set; }

        public string LCY_AMT { get; set; }

        public string DC { get; set; }

        public string USER_ID { get; set; }



    }

    public class OpenItems
    {
        public string ATM_RRN { get; set; }

        public string C { get; set; }

        public string D { get; set; }

        [JsonProperty("Grand Total")]
        public string GrandTotal { get; set; }

        public string VALUE_DT { get; set; }

        public string ADDL_TEXT { get; set; }

        public string USER_ID { get; set; }

    }
}
