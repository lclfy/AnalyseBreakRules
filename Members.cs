using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AnalyseBreakRules
{
    public class Members
    {
        /*
         * 人员类
            Members
            姓名
            name
           行车组
            team
            职名
            jobName
            政治面貌
            politicalOutlook
         * */
        public string name { get; set; }
        public string team { get; set; }
        public string jobName { get; set; }
        public string politicalOutlook { get; set; }

        public Members()
        {
            name = "";
            team = "";
            jobName = "";
            politicalOutlook = "";
        }
    }
}
