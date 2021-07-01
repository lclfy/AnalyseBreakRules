using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AnalyseBreakRules
{
    public class ClassificationOfRuleBreaks
    {//违标分类
     //违标等级，0最高，优先找0的
        public int rankOfRuleBreaks { get; set; }
        public string[] keyWords { get; set; }
        //解决方案，从模板里获取，随机抽取一个，分析内容与解决方案
        public List<string> analyseText { get; set; }
        public List<string> solutions { get; set; }
        //双重预防机制文件名称与所在位置
        public string fileName { get; set; }
        public string fileContent { get; set; }

        public ClassificationOfRuleBreaks()
        {
            rankOfRuleBreaks = -1;
            analyseText = new List<string>();
            solutions = new List<string>() ;
            fileContent = "";
            fileName = "";
        }
    }
}
