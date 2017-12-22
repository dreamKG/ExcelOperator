using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ZYCJ.Utility;
using System.Configuration;
using ZYCJ.Model;
using System.Globalization;

namespace ZYCJ
{
    public class Program
    {
        public static List<String> professorNames = getProfessorNameFromTemplate(ExcelUtility.GetWorkBook(GetConfig("targetTemplate")).GetSheetAt(0));


        public static void Main(string[] args)
        {
            IWorkbook crosswise = ExcelUtility.GetWorkBook(GetConfig("crosswise"));
            IWorkbook lengthways = ExcelUtility.GetWorkBook(GetConfig("lengthways"));
            IWorkbook paper = ExcelUtility.GetWorkBook(GetConfig("paper"));
            IWorkbook targetTemplate = ExcelUtility.GetWorkBook(GetConfig("targetTemplate"));

            Dictionary<string, AcademicAchievements> resultDictionary = new Dictionary<string, AcademicAchievements>();

            //ISheet templateSheet = targetTemplate.GetSheetAt(0);
            //professorNames = getProfessorNameFromTemplate(templateSheet);

            foreach (var item in professorNames)
            {
                resultDictionary.Add(item, new AcademicAchievements(item));
            }

            ISheet paperSheet = paper.GetSheetAt(0);
            getPaperInfo(paperSheet, resultDictionary);
            ISheet crosswiseSheet = crosswise.GetSheetAt(0);
            getProjectInfo(crosswiseSheet, resultDictionary);
            ISheet resultSheet = targetTemplate.GetSheetAt(0);
            setResultValue(resultSheet,resultDictionary);

            List<ISheet> sheets = new List<ISheet>() {resultSheet };
            IWorkbook workbook = ExcelUtility.CloneSheets(sheets);
            ExcelUtility.WriteWorkbookToFile(workbook, GetConfig("result"));


        }

        private static string GetConfig(string configName)
        {
            string result = ConfigurationManager.AppSettings.Get(configName);
            return result;
        }

        private static List<string> getProfessorNameFromTemplate(ISheet sheet)
        {
            List<string> result = new List<string>();
            for (int i = 3; i < 55; i++)
            {
                IRow row = sheet.GetRow(i);
                ICell cell = row.GetCell(2);
                String name;
                name = cell.StringCellValue;
                result.Add(name);
            }
            return result;
        }

        private static void getPaperInfo(ISheet sheet, Dictionary<string, AcademicAchievements> dic)
        {
            foreach (KeyValuePair<string, AcademicAchievements> pair in dic)
            {
                int infoIndex = 1;
                int maxRowCount = sheet.LastRowNum;
                for (int i = 1; i <= maxRowCount; i++)
                {
                    IRow row = sheet.GetRow(i);
                    ICell namesCell = row.GetCell(22);
                    string names = namesCell.StringCellValue;
                    
                    if (names.Contains(pair.Key))
                    {
                        ICell titleCell = row.GetCell(0);
                        string title = titleCell.StringCellValue;
                        ICell dateCell = row.GetCell(1);
                        string date = dateCell.StringCellValue;
                        DateTime dateTime = Convert.ToDateTime(date + " 00:00:00");
                        date = dateTime.ToString("yyyy.MM");
                        ICell periodicalCell = row.GetCell(2);
                        string periodical = periodicalCell.StringCellValue;
                        dic[pair.Key].paperInfo.Append(infoIndex);
                        dic[pair.Key].paperInfo.Append(".");
                        dic[pair.Key].paperInfo.Append(names);
                        dic[pair.Key].paperInfo.Append(".");
                        dic[pair.Key].paperInfo.Append(title);
                        dic[pair.Key].paperInfo.Append(".");
                        dic[pair.Key].paperInfo.Append(periodical);
                        dic[pair.Key].paperInfo.Append(".");
                        dic[pair.Key].paperInfo.Append(date);
                        dic[pair.Key].paperInfo.Append("；");
                        infoIndex++;
                    }
                }
            }
        }

        private static void getProjectInfo(ISheet sheet, Dictionary<string, AcademicAchievements> dic) {
            foreach (KeyValuePair<string, AcademicAchievements> pair in dic)
            {
                int infoIndex = 1;
                int maxRowCount = sheet.LastRowNum;
                for (int i = 1; i <= maxRowCount; i++)
                {
                    IRow row = sheet.GetRow(i);
                    ICell namesCell = row.GetCell(24);
                    string names = namesCell.StringCellValue;

                    if (names.StartsWith(pair.Key))
                    {
                        ICell projectPartyCell = row.GetCell(7);
                        string projectParty = projectPartyCell.StringCellValue +"课题";
                        ICell titleCell = row.GetCell(1);
                        string title = titleCell.StringCellValue;
                        ICell sdCell = row.GetCell(19);
                        string sd = sdCell.StringCellValue;
                        ICell edCell = row.GetCell(20);
                        string ed = edCell.StringCellValue;

                        DateTime sdTime = Convert.ToDateTime(sd + " 00:00:00");
                        sd = sdTime.ToString("yyyy/MM");
                        DateTime edTime = Convert.ToDateTime(ed + " 00:00:00");
                        ed = edTime.ToString("yyyy/MM");

                        dic[pair.Key].crosswiseProject.Append(infoIndex);
                        dic[pair.Key].crosswiseProject.Append(".");
                        dic[pair.Key].crosswiseProject.Append(projectParty);
                        dic[pair.Key].crosswiseProject.Append(",");
                        dic[pair.Key].crosswiseProject.Append(title);
                        dic[pair.Key].crosswiseProject.Append(",");
                        dic[pair.Key].crosswiseProject.Append(sd);
                        dic[pair.Key].crosswiseProject.Append("-");
                        dic[pair.Key].crosswiseProject.Append(ed);
                        dic[pair.Key].crosswiseProject.Append("；");
                        infoIndex++;
                    }
                    
                }
            }
        }

        private static void setResultValue(ISheet sheet, Dictionary<string, AcademicAchievements> dic)
        {
            
            for (int i = 3; i < 55; i++)
            {
                IRow row = sheet.GetRow(i);
                ICell nameCell = row.GetCell(2);
                string name = nameCell.StringCellValue;

                string paperInfo = dic[name].paperInfo.ToString();
                if (paperInfo.EndsWith("；")) {
                    paperInfo = paperInfo.Remove(paperInfo.Length-1,1);
                }

                string crosswiseProject = dic[name].crosswiseProject.ToString();
                if (crosswiseProject.EndsWith("；"))
                {
                    crosswiseProject = crosswiseProject.Remove(crosswiseProject.Length - 1, 1);
                }

                ICell targetCell = row.GetCell(9);
                if (targetCell == null)
                {
                    row.CreateCell(9);
                    targetCell = row.GetCell(9);
                }
                targetCell.SetCellValue(paperInfo + crosswiseProject);
            }
        }
    }
}
