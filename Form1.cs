using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using CCWin;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace AnalyseBreakRules
{
    public partial class Form1 : Skin_Mac
    {
        List<string> breakRulefileNames;
        //示例文件
        List<string> exampleFilesName;
        //样板（空违标分析表）
        string modelFileName;
        List<string> problemsInfo;
        //成员
        List<Members> allMembers;
        //违标分析列表
        List<BreakRules> allBreakRules;
        //违标分类列表
        List<ClassificationOfRuleBreaks> classificationOfRules;
        //会议主持人与参会人员
        string host;
        string teamMembers;
        string[] randomTeamMembers;
        string southStationMenber;
        string southEMUGarageMenber;
        string eastEMUGarageMember;
        string innercityMember;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            refreshObjects();
            getProblemsText();
            ImportExcelFiles(0);
            importWordFiles();
            getExamples();
        }

        private void refreshObjects()
        {
            host = "姚英";
            teamMembers = "姚英，赵明，罗思聪";
            if (!checkBox1.Checked)
            {
                label6.Visible = false;
                label7.Visible = false;
                label8.Visible = false;
                label9.Visible = false;
                textBox2.Visible = false;
                textBox3.Visible = false;
                textBox4.Visible = false;
                textBox5.Visible = false;
            }
            else
            {
                label6.Visible = true;
                label7.Visible = true;
                label8.Visible = true;
                label9.Visible = true;
                textBox2.Visible = true;
                textBox3.Visible = true;
                textBox4.Visible = true;
                textBox5.Visible = true;
            }
            textBox1.Text = teamMembers;
            southStationMenber = "岳云峰";
            textBox2.Text = southStationMenber;
            southEMUGarageMenber = "杨焱鑫";
            textBox3.Text = southEMUGarageMenber;
            innercityMember = "张亚辉";
            textBox4.Text = innercityMember;
            eastEMUGarageMember = "张志";
            textBox5.Text = eastEMUGarageMember;
            randomTeamMembers = teamMembers.Split('，');
            breakRulefileNames = new List<string>();
            exampleFilesName = new List<string>();
            modelFileName = "";
            problemsInfo = new List<string>();
            allMembers = new List<Members>();
            allBreakRules = new List<BreakRules>();
            classificationOfRules = new List<ClassificationOfRuleBreaks>();
        }

        private void getProblemsText()
        {
            problemsInfo = new List<string>();
            //添加问题分析名称，0问题 1轻微 2一般 3严重
            string problemText0 = "列问题一件。";
            problemsInfo.Add(problemText0);
            string problemText1 = "按轻微违标，扣款50元，纳入星级职工考核。";
            problemsInfo.Add(problemText1);
            string problemText2 = "按一般违标，扣款100元，纳入星级职工考核。";
            problemsInfo.Add(problemText2);
            string problemText3 = "按严重违标，扣款200元，纳入星级职工考核。";
            problemsInfo.Add(problemText3);

        }

        private bool ImportExcelFiles(int type)
        {
            //读问题查询表
            if(type == 1)
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();   //显示选择文件对话框 
                openFileDialog1.Multiselect = false;
                openFileDialog1.Filter = "Excel 文件 |*.xlsx;*.xls";
                //openFileDialog1.InitialDirectory = Application.StartupPath + "\\时刻表\\";
                openFileDialog1.FilterIndex = 2;
                openFileDialog1.RestoreDirectory = true;
                IWorkbook workBook;
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    String fileNames = "已选择：";
                    string fileName = openFileDialog1.FileName;
                    {
                        //try
                        {
                            FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
                            if (fileName.IndexOf(".xlsx") > 0) // 2007版本  
                            {
                                //try
                                {
                                    workBook = new XSSFWorkbook(fileStream);  //xlsx数据读入workbook  
                                    allBreakRules = getProblems(workBook);
                                }
                                //catch (Exception e)
                                {
                                    //MessageBox.Show("读取问题查询表时出现错误\n" + fileName + "\n错误内容：" + e.ToString().Split('在')[0], "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                    //return false;
                                }

                            }
                            else if (fileName.IndexOf(".xls") > 0) // 2003版本  
                            {
                               // try
                                {
                                    workBook = new HSSFWorkbook(fileStream);  //xls数据读入workbook  
                                    allBreakRules = getProblems(workBook);

                                }
                               // catch (Exception e)
                                {
                                  //  MessageBox.Show("读取问题查询表时出现错误\n" + fileName + "\n错误内容：" + e.ToString().Split('在')[0], "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                   // return false;
                                }
                            }
                            fileStream.Close();
                        }
                        //catch (IOException)
                        {
                           // MessageBox.Show("读取问题查询表时出现错误\n" + fileName, "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                          //  return false;
                        }
                    }
                    fileNames = fileNames + fileName;
                    label1.Text = fileNames;
                }
            }
            //读成员名单
            else if(type == 0)
            {
                //加载成员名单
                string fileName = Application.StartupPath + "\\Members.xls";
                IWorkbook workBook;
                FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
                try
                {
                    workBook = new HSSFWorkbook(fileStream);  //xlsx数据读入workbook  
                    allMembers = getMembers(workBook);
                }
                catch (Exception e)
                {
                    MessageBox.Show("读取成员名单(Member.xls)出现错误,将无法获取职名,政治面貌等\n" + fileName + "\n错误内容：" + e.ToString().Split('在')[0], "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return false;
                }
                fileStream.Close();
                //label2.Text = "已加载成员名单";
                //加载分类
                string fileNameClass = Application.StartupPath + "\\Classifications.xls";
                IWorkbook workBookClass;
                FileStream fileStreamClass = new FileStream(fileNameClass, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
                try
                {
                    workBookClass = new HSSFWorkbook(fileStreamClass);  //xlsx数据读入workbook  
                    classificationOfRules= getClassifications(workBookClass);
                }
                catch (Exception e)
                {
                    MessageBox.Show("读取违标分类列表(Classifications.xls)出现错误\n" + fileNameClass + "\n错误内容：" + e.ToString().Split('在')[0], "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return false;
                }
                fileStreamClass.Close();
                label2.Text = "已加载违标分类";

            }
            return true;
        }

        private void importWordFiles()
        {
            try
            {//找示例
                //验证“路径”文本框中的路径是否存在
                string path = Application.StartupPath + "\\Examples\\";
                if (Directory.Exists(path))
                {
                    //搜索“路径”文件夹下面的文件
                    string[] files = Directory.GetFiles(path, "*.doc", SearchOption.AllDirectories);    //从“路径”中搜索所有的文件
                    SortedList<string, FileInfo> fileList = new SortedList<string, FileInfo>();     //声明一个字典，用于存储文件信息
                                                                                                    //验证是否在“路径”中搜索到了文件
                    if (files.Length > 0)
                    {
                        //搜索到了文件，继续执行

                        foreach (string f in files)    //把文件夹中搜索到文件信息全部存储到刚才声明的字典中
                        {
                            FileInfo fi = new FileInfo(f);  //根据路径获取文件信息
                            fileList.Add(fi.Name, fi);  //存储文件信息到字典中
                        }
                        //把存储到文件信息字典中的数据显示到listview中
                        int counter = 0;
                        foreach (FileInfo item in fileList.Values)
                        {
                            if (!item.ToString().Contains("$"))
                            {
                                counter++;
                                exampleFilesName.Add(item.ToString());
                            }
                            /*
                            //验证一下里面有没有表格
                            FileFormat fileFormat = FileFormat.Docx;
                            Document doc = new Document();
                            doc.LoadFromFile(item.ToString(), fileFormat);
                            if (doc.Sections[0] != null)
                            {

                            }
                            */

                        }
                        label4.Text = "已找到" + counter + "个分析示例";
                    }
                }
                else
                {
                    MessageBox.Show("没有找到Examples文件夹，无法搜索示例。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("没有找到Examples文件夹，无法搜索示例。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            try
            {//找模板（空表）
                string fileNameModel = Application.StartupPath + "\\Empty.docx";
                try
                {
                    FileFormat fileFormat = FileFormat.Docx;
                    Document doc = new Document();
                    doc.LoadFromFile(fileNameModel, fileFormat);
                    modelFileName = fileNameModel;
                }
                catch (Exception e)
                {
                    MessageBox.Show("读取违标模板(Empty.docx)出现错误\n" + fileNameModel + "\n错误内容：" + e.ToString().Split('在')[0], "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("读取违标模板(Empty.docx)出现错误", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        //获取分类
        private List<ClassificationOfRuleBreaks> getClassifications(IWorkbook ClassificationsWorkbook)
        {
            List<ClassificationOfRuleBreaks> _cofRuleBreaks = new List<ClassificationOfRuleBreaks>();
            IWorkbook _cWorkBook = ClassificationsWorkbook;
            ISheet _sheet;
            if(_cWorkBook == null)
            {
                return _cofRuleBreaks;
            }
            _sheet = _cWorkBook.GetSheetAt(0);
            if(_sheet == null)
            {
                return _cofRuleBreaks;
            }
            for(int  i = 2; i <= _sheet.LastRowNum; i++)
            {
                IRow row = _sheet.GetRow(i);
                if(row == null)
                {
                    continue;
                }
                ClassificationOfRuleBreaks _corb = new ClassificationOfRuleBreaks();
                if (row.GetCell(0) == null)
                {
                    continue;
                }
                if(row.GetCell(1) == null)
                {
                    continue;
                }
                //填上
                if(row.GetCell(0).ToString().Trim().Length != 0 &&
                    row.GetCell(1).ToString().Trim().Length != 0)
                {
                    int rank = -1;
                    int.TryParse(row.GetCell(0).ToString().Trim(), out rank);
                    _corb.rankOfRuleBreaks = rank;
                    _corb.keyWords = row.GetCell(1).ToString().Trim().Split(',');
                    _cofRuleBreaks.Add(_corb);
                }
            }
            return _cofRuleBreaks;
        }

        //获取成员
        private List<Members> getMembers(IWorkbook MemberWorkbook)
        {
            List<Members> members = new List<Members>();
            IWorkbook _mWorkBook = MemberWorkbook;
            ISheet _sheet;
            if (_mWorkBook == null)
            {
                return members;
            }
            _sheet = _mWorkBook.GetSheetAt(0);
            if (_sheet == null)
            {
                return members;
            }
            for (int i = 2; i <= _sheet.LastRowNum; i++)
            {
                IRow row = _sheet.GetRow(i);
                if (row == null)
                {
                    continue;
                }
                Members _mem = new Members();
                if (row.GetCell(2) == null)
                {
                    continue;
                }
                //填上
                if (row.GetCell(0).ToString().Trim().Length != 0)
                {
                    _mem.name = row.GetCell(0).ToString().Trim();
                    if (row.GetCell(1) != null &&
                        row.GetCell(1).ToString().Length != 0)
                    {
                        _mem.team = row.GetCell(1).ToString().Trim();
                    }
                    if (row.GetCell(2) != null &&
    row.GetCell(2).ToString().Length != 0)
                    {
                        _mem.jobName = row.GetCell(2).ToString().Trim();
                    }
                    if (row.GetCell(3) != null &&
    row.GetCell(3).ToString().Length != 0)
                    {
                        _mem.politicalOutlook = row.GetCell(3).ToString().Trim();
                    }
                    members.Add(_mem);
                }
            }
            return members;
        }

        //获取违标内容
        private List<BreakRules> getProblems(IWorkbook problemWorkbook)
        {
            IWorkbook _problemWorkbook = problemWorkbook;
            List<BreakRules> _breakRules = new List<BreakRules>();
            int titleRowCount = -1;
            int problemColumn = -1;
            int timeColumn = -1;
            int rankOfBreakRulesColumn = -1;
            int hostColumn = -1;
            //责任人
            int peopleColumn = -1;
            //整改措施，大概是不需要的
            int solutionsColumn = -1;
            ISheet sheet = _problemWorkbook.GetSheetAt(0);
            if(sheet == null)
            {
                return _breakRules;
            }
            //找标题，不会超过5行
            for(int i = 0; i < 5; i++)
            {
                IRow row = sheet.GetRow(i);
                if(row == null)
                {
                    continue;
                }
                for(int j = 0; j <= row.LastCellNum; j++)
                {
                    if (row.GetCell(j) != null)
                    {
                        string text = row.GetCell(j).ToString().Trim();
                        if (text.Contains("问题内容"))
                        {
                            problemColumn = j;
                            titleRowCount = i;
                        }
                        else if (text.Trim().Equals("责任人"))
                        {
                            peopleColumn = j;
                        }
                        else if (text.Contains("发现时间"))
                        {
                            timeColumn = j;
                        }
                        else if (text.Contains("严重程度"))
                        {
                            rankOfBreakRulesColumn = j;
                        }
                        else if (text.Contains("处理人"))
                        {
                            hostColumn = j;
                        }
                        else if (text.Contains("整改措施"))
                        {
                            solutionsColumn = j;
                        }

                    }
                }
            }

            if(titleRowCount == -1)
            {
                return _breakRules;
            }
            //再一遍，找出内容

            for (int ij = titleRowCount+1; ij <= sheet.LastRowNum; ij++)
            {
                BreakRules _br = new BreakRules();
                IRow row = sheet.GetRow(ij);
                if(row == null)
                {
                    continue;
                }
                //问题内容
                if(row.GetCell(problemColumn) != null &&
                    row.GetCell(problemColumn).ToString().Trim().Length != 0)
                {
                    _br.breakRuleContents = row.GetCell(problemColumn).ToString().Trim();
                }
                //责任人
                if (row.GetCell(peopleColumn) != null &&
                    row.GetCell(peopleColumn).ToString().Trim().Length != 0)
                {
                    _br.peopleLiable = row.GetCell(peopleColumn).ToString().Trim().Replace(" ","");
                }
                //处理人
                if (row.GetCell(hostColumn) != null &&
                    row.GetCell(hostColumn).ToString().Trim().Length != 0)
                {
                    _br.analyseHost = row.GetCell(hostColumn).ToString().Trim().Replace(" ", "");
                }
                //整改措施
                if (row.GetCell(solutionsColumn) != null &&
                    row.GetCell(solutionsColumn).ToString().Trim().Length != 0)
                {
                    _br.analyseSolutions = row.GetCell(solutionsColumn).ToString().Trim();
                }
                //发现时间
                if (row.GetCell(timeColumn) != null &&
                 row.GetCell(timeColumn).ToString().Trim().Length != 0)
                {
                    _br.time = row.GetCell(timeColumn).ToString().Split(' ')[0];
                }
                //违标等级
                if (row.GetCell(rankOfBreakRulesColumn) != null &&
                row.GetCell(rankOfBreakRulesColumn).ToString().Trim().Length != 0)
                {
                    if (row.GetCell(rankOfBreakRulesColumn).ToString().Trim().Contains("其他"))
                    {
                        _br.breakRuleClass = 0;
                        _br.treatWay = problemsInfo[0];
                    }
                    else  if (row.GetCell(rankOfBreakRulesColumn).ToString().Trim().Contains("轻微"))
                    {
                        _br.breakRuleClass = 1;
                        _br.treatWay = problemsInfo[1];
                    }
                    else if (row.GetCell(rankOfBreakRulesColumn).ToString().Trim().Contains("一般"))
                    {
                        _br.breakRuleClass = 2;
                        _br.treatWay = problemsInfo[2];
                    }
                    else if (row.GetCell(rankOfBreakRulesColumn).ToString().Trim().Contains("严重"))
                    {
                        _br.breakRuleClass = 3;
                        _br.treatWay = problemsInfo[3];
                    }
                }
                //把人名和职名，政治面貌，行车组匹配
                foreach(Members _m in allMembers)
                {
                    if (_m.name.Equals(_br.peopleLiable))
                    {
                        _br.jobName = _m.jobName;
                        _br.politicalOutlook = _m.politicalOutlook;
                        _br.team = _m.team;
                    }
                }
                //把同班组人员找出来
                //加上主持人，班组长，随机干部与责任人
                Random _rd = new Random(Guid.NewGuid().GetHashCode());
                string randomName = "";
                randomName = randomTeamMembers[_rd.Next(0, randomTeamMembers.Length)].Trim();
                _br.analyseHost = randomName;
                _br.analyseTeam = _br.analyseHost ;
                //严重违标所有管理人员都来
                if(_br.breakRuleClass == 3)
                {
                    _br.analyseTeam = _br.analyseTeam + teamMembers;
                }

                if(_br.team.Contains("线路所")||
                    _br.team.Contains("城际")||
                    _br.team.Contains("南动车所") ||
                    _br.team.Contains("南站")||
                   _br.team.Contains("备班"))
                {
                    if(_br.team.Contains("线路所")||
                        _br.team.Contains("城际站")||
                        _br.team.Contains("备班"))
                    {
                        _br.analyseTeam = _br.analyseTeam + "，"+innercityMember+"，" + _br.peopleLiable;
                    }
                    else if( _br.team.Contains("南站"))
                    {
                        _br.analyseTeam = _br.analyseTeam + "，"+southStationMenber+"，" + _br.peopleLiable;
                    }
                    else if (_br.team.Contains("南动车所"))
                    {
                        _br.analyseTeam = _br.analyseTeam + "，" + southEMUGarageMenber + "，" + _br.peopleLiable;
                    }
                    else if (_br.team.Contains("东动车所"))
                    {
                        _br.analyseTeam = _br.analyseTeam + "，" + eastEMUGarageMember + "，" + _br.peopleLiable;
                    }
                }
                else
                {
                    foreach (Members _m in allMembers)
                    {
                        //城际站线路所只来自己和辉哥
                        if (_m.team.Equals(_br.team) && checkBox1.Checked)
                        {
                            _br.analyseTeam = _br.analyseTeam + "，" + _m.name;
                        }
                        else if(_m.team.Equals(_br.team) && _m.jobName.Contains("客运值班员") && !checkBox1.Checked)
                        {
                            _br.analyseTeam = _br.analyseTeam + "，" + _m.name;
                        }
                        else if (_m.name.Equals(_br.peopleLiable) && !checkBox1.Checked)
                        {
                            _br.analyseTeam = _br.analyseTeam + "，" + _m.name;
                        }
                    }
                }

                //参会人员
                //_br.analyseTeam = teamMembers;

                _br = getBreakRuleSolutions(_br);
                _breakRules.Add(_br);
            }



            return _breakRules;
        }

        private BreakRules getBreakRuleSolutions(BreakRules _br)
        {
            //在分类中根据keyword找出违标的对应方案
            string originalText = _br.breakRuleContents;
            bool hasGetKeyWord = false;
            //随机获取一种分析与解决方案
            Random rd = new Random();
            //先找等级0的
            foreach (ClassificationOfRuleBreaks _corb in classificationOfRules)
            {
                if(_corb.rankOfRuleBreaks == 0)
                {
                    foreach(string _keyWord in _corb.keyWords)
                    {
                        //找到了
                        if (originalText.Contains(_keyWord))
                        {
                            //随机抽取一种方法
                            hasGetKeyWord = true;
                            _br.keyWord = _keyWord;
                            if(_corb.solutions.Count != 0)
                            {
                                int randomNumber = rd.Next(0, _corb.solutions.Count - 1);
                                if (_corb.analyseText.Count > randomNumber)
                                {
                                    _br.analyseContent = _corb.analyseText[randomNumber];
                                }
                                if (_corb.solutions[randomNumber] != null)
                                {
                                    _br.analyseSolutions = _corb.solutions[randomNumber];
                                }
                            }
                        }
                    }
                }
            }
            //没找到，再从等级1的里面找
            if (!hasGetKeyWord)
            {
                foreach (ClassificationOfRuleBreaks _corb in classificationOfRules)
                {
                    if (_corb.rankOfRuleBreaks == 1)
                    {
                        foreach (string _keyWord in _corb.keyWords)
                        {
                            //找到了
                            if (originalText.Contains(_keyWord))
                            {
                                hasGetKeyWord = true;
                                _br.keyWord = _keyWord;
                                if(_corb.solutions.Count > 0)
                                {
                                    int randomNumber = rd.Next(0, _corb.solutions.Count);
                                    if (_corb.analyseText.Count > randomNumber)
                                    {
                                        _br.analyseContent = _corb.analyseText[randomNumber];
                                    }
                                    if (_corb.solutions[randomNumber] != null)
                                    {
                                        _br.analyseSolutions = _corb.solutions[randomNumber];
                                    }
                                }

                            }
                        }
                    }
                }
            }

            return _br;
        } 

        //获得示例中的分析与措施
        private void getExamples()
        {
            List<ClassificationOfRuleBreaks> _allCORB = classificationOfRules;
            foreach(string fileName in exampleFilesName)
            {
                FileFormat fileFormat = FileFormat.Doc;
                Document doc = new Document();
                doc.LoadFromFile(fileName, fileFormat);
                if (doc.Sections[0] == null)
                {
                    continue;
                }
                if (doc.Sections[0].Tables[0] == null)
                {
                    continue;
                }
                else
                {
                    Table table = doc.Sections[0].Tables[0] as Table;
                    bool hasGotIt = true;
                    bool hasGotLevel0 = false;
                    int targetLv0Place = -1;
                    int targetLv1Place = -1;
                    string content = "";
                    //遍历表格中的段落，找到需要的位置添加(先找lv0的)
                    /*
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        TableRow row = table.Rows[i];
                        for (int j = 0; j < row.Cells.Count; j++)
                        {
                            TableCell cell = row.Cells[j];
                            foreach (Paragraph paragraph in cell.Paragraphs)
                            {
                                //判断它符合哪个keyword,0和1都找一遍
                                string Text = paragraph.Text.ToString().Trim();
                                if (Text.Contains("概况："))
                                {
                                    for(int ij = 0; ij < _allCORB.Count; ij++)
                                    {
                                        for(int k = 0; k < _allCORB[ij].keyWords.Length; k++)
                                        {
                                            //命中
                                            if (Text.Contains(_allCORB[ij].keyWords[k]))
                                            {
                                                content = Text;
                                                if(_allCORB[ij].rankOfRuleBreaks == 0 && !hasGotLevel0)
                                                {
                                                    hasGotIt = true;
                                                    hasGotLevel0 = true;
                                                    targetLv0Place = ij;
                                                }

                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    //再找lv1的
                    if (!hasGotLevel0)
                    {
                        for (int i = 0; i < table.Rows.Count; i++)
                        {
                            TableRow row = table.Rows[i];
                            for (int j = 0; j < row.Cells.Count; j++)
                            {
                                TableCell cell = row.Cells[j];
                                foreach (Paragraph paragraph in cell.Paragraphs)
                                {
                                    //判断它符合哪个keyword,0和1都找一遍
                                    string Text = paragraph.Text.ToString().Trim();
                                    if (Text.Contains("概况："))
                                    {
                                        for (int ij = 0; ij < _allCORB.Count; ij++)
                                        {
                                            for (int k = 0; k < _allCORB[ij].keyWords.Length; k++)
                                            {
                                                //命中
                                                if (Text.Contains(_allCORB[ij].keyWords[k]))
                                                {
                                                    if (_allCORB[ij].rankOfRuleBreaks == 1)
                                                    {
                                                        hasGotIt = true;
                                                        targetLv1Place = ij;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    */
                    string test = content;
                    if (hasGotIt)
                    {
                        //遍历表格中的段落，找到需要的位置添加
                        for (int i = 0; i < table.Rows.Count; i++)
                        {
                            TableRow row = table.Rows[i];
                            for (int j = 0; j < row.Cells.Count; j++)
                            {
                                TableCell cell = row.Cells[j];
                                foreach (Paragraph paragraph in cell.Paragraphs)
                                {
                                    //判断它符合哪个keyword,0和1都找一遍
                                    string Text = paragraph.Text.ToString().Trim();
                                    //把相应的分析与措施添加进去
                                    if(targetLv0Place != -1)
                                    {
                                        if (Text.Contains("分析："))
                                        {
                                            foreach(Paragraph temp_paragraph in cell.Paragraphs)
                                            {
                                                string alltext = temp_paragraph.Text.ToString().Trim();
                                                _allCORB[targetLv0Place].analyseText.Add(alltext.Replace("分析：", ""));
                                            }
                                          
                                        }
                                        if (Text.Contains("措施："))
                                        {
                                            foreach (Paragraph temp_paragraph in cell.Paragraphs)
                                            {
                                                string alltext = temp_paragraph.Text.ToString().Trim();
                                                _allCORB[targetLv0Place].solutions.Add(alltext.Replace("分析：", ""));
                                            }
                                        }
                                    }
                                    else if(targetLv1Place != -1)
                                    {
                                        if (Text.Contains("分析："))
                                        {
                                            foreach (Paragraph temp_paragraph in cell.Paragraphs)
                                            {
                                                string alltext = temp_paragraph.Text.ToString().Trim();
                                                _allCORB[targetLv1Place].analyseText.Add(alltext.Replace("分析：", ""));
                                            }
                                        }
                                        if (Text.Contains("措施："))
                                        {
                                            foreach (Paragraph temp_paragraph in cell.Paragraphs)
                                            {
                                                string alltext = temp_paragraph.Text.ToString().Trim();
                                                _allCORB[targetLv1Place].solutions.Add(alltext.Replace("分析：", ""));
                                            }
                                        }
                                    }

                                }
                            }
                        }
                    }
                }
            }
        }

        //填写
        private void fillBreakRules()
        {
            List<BreakRules> _allBr = allBreakRules;
            int count = 1;
            foreach(BreakRules _br in _allBr)
            {
                FileFormat fileFormat = FileFormat.Docx;
                Document doc = new Document();
                doc.LoadFromFile(modelFileName, fileFormat);
                if (doc.Sections[0] == null)
                {
                    MessageBox.Show("模板违标分析(Model.docx)出现问题：无表格。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                if (doc.Sections[0].Tables[0] == null)
                {
                    MessageBox.Show("模板违标分析(Model.docx)出现问题：无表格。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                else
                {
                    Table table = doc.Sections[0].Tables[0] as Table;
                    //遍历表格中的段落，找到后在相应位置填写内容
                    for (int i =0;i<table.Rows.Count;i++)
                    {
                        TableRow row = table.Rows[i];
                        for (int j = 0; j < row.Cells.Count; j++)
                        {
                            TableCell cell = row.Cells[j];
                            foreach (Paragraph paragraph in cell.Paragraphs)
                            {
                                //时间,在下一个格子里填上
                                if (paragraph.Text.ToString().Trim().Equals("时间"))
                                {
                                    DateTime dt = new DateTime();
                                    Random rd = new Random();
                                    bool hasGot = false;
                                    if (_br.time.Contains("-"))
                                    {
                                        try
                                        {
                                            dt = DateTime.Parse(_br.time);
                                            hasGot = true;
                                        }
                                        catch (Exception e)
                                        {

                                        }
                                    }
                                    else if (_br.time.Contains("/"))
                                    {

                                        try
                                        {
                                            dt = DateTime.Parse(_br.time);
                                            hasGot = true;
                                        }
                                        catch(Exception e)
                                        {
                                            
                                        }

                                    }
                                    TextRange tableRange;
                                    if (hasGot)
                                    {

                                        tableRange = table[i, j + 1].Paragraphs[0].AppendText(dt.AddDays(rd.Next(1, 2)).ToString("MM月dd日"));
                                    }
                                    else
                                    {
                                        tableRange = table[i, j + 1].Paragraphs[0].AppendText(_br.time);
                                    }
                                    tableRange.CharacterFormat.FontName = "宋体";
                                    tableRange.CharacterFormat.FontSize = 12;
                                }
                                if (paragraph.Text.ToString().Trim().Equals("责任人"))
                                {
                                    TextRange tableRange = table[i, j + 1].Paragraphs[0].AppendText(_br.peopleLiable);
                                    tableRange.CharacterFormat.FontName = "宋体";
                                    tableRange.CharacterFormat.FontSize = 12;
                                }
                                if (paragraph.Text.ToString().Trim().Equals("班组"))
                                {
                                    TextRange tableRange = table[i, j + 1].Paragraphs[0].AppendText(_br.team);
                                    tableRange.CharacterFormat.FontName = "宋体";
                                    tableRange.CharacterFormat.FontSize = 12;
                                }
                                if (paragraph.Text.ToString().Trim().Equals("职名"))
                                {
                                    TextRange tableRange = table[i, j + 1].Paragraphs[0].AppendText(_br.jobName);
                                    tableRange.CharacterFormat.FontName = "宋体";
                                    tableRange.CharacterFormat.FontSize = 12;
                                }
                                if (paragraph.Text.ToString().Trim().Equals("政治"))
                                {
                                    TextRange tableRange = table[i, j + 1].Paragraphs[0].AppendText(_br.politicalOutlook);
                                    tableRange.CharacterFormat.FontName = "宋体";
                                    tableRange.CharacterFormat.FontSize = 12;
                                }
                                if (paragraph.Text.ToString().Trim().Equals("类别"))
                                {
                                    string classifications = "";
                                    if(_br.breakRuleClass == 0)
                                    {
                                        classifications = "其他问题";
                                    }
                                    else if(_br.breakRuleClass == 1)
                                    {
                                        classifications = "轻微违标";
                                    }
                                    else if (_br.breakRuleClass == 2)
                                    {
                                        classifications = "一般违标";
                                    }
                                    else if (_br.breakRuleClass == 3)
                                    {
                                        classifications = "严重违标";
                                    }
                                    TextRange tableRange = table[i, j + 1].Paragraphs[0].AppendText(classifications);
                                    tableRange.CharacterFormat.FontName = "宋体";
                                    tableRange.CharacterFormat.FontSize = 12;
                                }
                                if (paragraph.Text.ToString().Trim().Equals("主持人"))
                                {
                                    TextRange tableRange = table[i, j + 1].Paragraphs[0].AppendText(_br.analyseHost);
                                    tableRange.CharacterFormat.FontName = "宋体";
                                    tableRange.CharacterFormat.FontSize = 12;
                                }
                                if (paragraph.Text.ToString().Trim().Equals("参加"))
                                {
                                    TextRange tableRange = table[i, j + 1].Paragraphs[0].AppendText(_br.analyseTeam);
                                    tableRange.CharacterFormat.FontName = "宋体";
                                    tableRange.CharacterFormat.FontSize = 12;
                                }
                                if (paragraph.Text.ToString().Trim().Equals("概况："))
                                {
                                    TextRange tableRange = table[i, j].Paragraphs[0].AppendText(_br.breakRuleContents);
                                    tableRange.CharacterFormat.FontName = "宋体";
                                    tableRange.CharacterFormat.FontSize = 12;
                                }
                                if (paragraph.Text.ToString().Trim().Equals("分析："))
                                {
                                    TextRange tableRange = table[i, j].Paragraphs[0].AppendText( _br.analyseContent);
                                    tableRange.CharacterFormat.FontName = "宋体";
                                    tableRange.CharacterFormat.FontSize = 12;
                                }
                                if (paragraph.Text.ToString().Trim().Equals("措施："))
                                {
                                    TextRange tableRange = table[i, j].Paragraphs[0].AppendText(_br.analyseSolutions);
                                    tableRange.CharacterFormat.FontName = "宋体";
                                    tableRange.CharacterFormat.FontSize = 12;
                                }
                                if (paragraph.Text.ToString().Trim().Equals("处理："))
                                {
                                    TextRange tableRange = table[i, j].Paragraphs[0].AppendText( _br.treatWay);
                                    tableRange.CharacterFormat.FontName = "宋体";
                                    tableRange.CharacterFormat.FontSize = 12;
                                }
                            }
                        }
                    }
                }
                //另存为文件，存在“Outputs”文件夹
                doc.SaveToFile(Application.StartupPath + "\\Outputs\\"+count.ToString()+"违标分析-"+_br.peopleLiable+".docx", FileFormat.Docx2013);
                count++;
            }
            MessageBox.Show("完成", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            System.Diagnostics.ProcessStartInfo info = new System.Diagnostics.ProcessStartInfo();
            //info.WorkingDirectory = Application.StartupPath;
            info.FileName = Application.StartupPath + "\\Outputs\\";
            info.Arguments = "";
            try
            {
                System.Diagnostics.Process.Start(info);
            }
            catch (System.ComponentModel.Win32Exception we)
            {
                MessageBox.Show(this, we.Message);
                return;
            }
        }

        private void importSearchedProblems_btn_Click(object sender, EventArgs e)
        {
            ImportExcelFiles(1);
        }

        //开始按钮
        private void button1_Click(object sender, EventArgs e)
        {
            fillBreakRules();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if(textBox1.Text.Length != 0)
            {
                teamMembers = textBox1.Text;
                teamMembers.Replace(",", "，");
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (textBox2.Text.Length != 0)
            {
                southStationMenber = textBox2.Text;
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            if (textBox3.Text.Length != 0)
            {
                southEMUGarageMenber = textBox3.Text;
            }
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            if (textBox4.Text.Length != 0)
            {
                innercityMember = textBox4.Text;
            }
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            if (textBox5.Text.Length != 0)
            {
                eastEMUGarageMember = textBox5.Text;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }
    }
}
