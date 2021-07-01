using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AnalyseBreakRules
{
    class ClassificationOfRuleBreaks
    {//违标分类
         List<string> keyWords { get; set; }
        //解决方案，从模板里获取，随机抽取一个
        List<string> solutions { get; set; }
        //双重预防机制文件名称与所在位置
        string fileName { get; set; }
        string fileContent { get; set; }

    }
}
