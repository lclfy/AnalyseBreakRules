using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AnalyseBreakRules
{
    class BreakRules
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

        string breakRuleContents { get; set; }
        string time { get; set; }
        string peopleLiable { get; set; }
        string team { get; set; }
        string jobName { get; set; }
        //政治面貌，0群众，1党员
        int politicalOutlook { get; set; }
        //0问题 ，1轻微， 2一般， 3严重
        int breakRuleClass { get; set; }
        string breakRuleKeyWord { get; set; }
        string analyseHost { get; set; }
        string analyseTeam { get; set; }
        string analyseContent { get; set; }
        string analyseSolutions { get; set; }
    }
}
