using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AnalyseBreakRules
{
    public class BreakRules
    {
        /*
         * 违标内容
        breakRuleContents
        时间time datetime
        责任人peopleLiable
        班组team
        职名jobName
        政治面貌politicalOutlook 0群众 1党员
        类别breakRuleClass 0问题 1轻微 2一般 3严重
        违标关键字
        breakRuleKeyWord
        分析会主持人（处理人）
        analyseHost
        参会人员
        analyseTeam

        分析内容
        analyseContent
        措施
        analyseMeasures
        */

        public string breakRuleContents { get; set; }
        public string time { get; set; }
        public string peopleLiable { get; set; }
        public string team { get; set; }
        public string jobName { get; set; }
        //政治面貌
        public string politicalOutlook { get; set; }
        //0问题 ，1轻微， 2一般， 3严重
        public int breakRuleClass { get; set; }
        public string breakRuleKeyWord { get; set; }
        public string analyseHost { get; set; }
        public string analyseTeam { get; set; }
        public string analyseContent { get; set; }
        public string analyseSolutions { get; set; }
        public string treatWay { get; set; }
        //测试用，被找出的关键字是什么
        public string keyWord { get; set; }

        public BreakRules()
        {
            breakRuleContents = "";
            time = "";
            peopleLiable = "";
            team = "";
            jobName = "";
            politicalOutlook = "";
            breakRuleClass = -1;
            breakRuleKeyWord = "";
            analyseHost = "";
            analyseTeam = "";
            analyseContent = "";
            analyseSolutions = "";
            treatWay = "";
            keyWord = "";
        }
    }
}
