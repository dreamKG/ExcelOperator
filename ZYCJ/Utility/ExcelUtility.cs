using System;
using System.Collections.Generic;
using System.Linq;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using ZYCJ.Model;

namespace ZYCJ.Utility
{
    public static class ExcelUtility
    {
        /// <summary>
        /// 获取workbook
        /// </summary>
        /// <param name="fileName">文件名</param>
        /// <returns></returns>
        public static IWorkbook GetWorkBook(string fileName)
        {
            using (FileStream fs = File.OpenRead(fileName))
            {
                IWorkbook workbook = WorkbookFactory.Create(fs);
                return workbook;
            }
        }

        /// <summary>
        /// 生成一个克隆多个sheet的workbook
        /// </summary>
        /// <param name="sheets"></param>
        /// <returns></returns>
        public static IWorkbook CloneSheets(List<ISheet> sheets)
        {
            IWorkbook workbook = new NPOI.XSSF.UserModel.XSSFWorkbook();
            foreach (var item in sheets)
            {
                ISheet sheet = workbook.CreateSheet(item.SheetName);
                for (int i = 0; i <= item.LastRowNum; i++)
                {
                    IRow row = sheet.CreateRow(i);
                    int lastCellNumber = GetLastCellNum(item.GetRow(i)) - 1;
                    for (int j = 0; j <= lastCellNumber; j++)
                    {
                        ICell cell = item.GetRow(i).GetCell(j);
                        if (cell.CellType == CellType.Numeric)
                        {
                            row.CreateCell(j).SetCellValue(cell.NumericCellValue);
                        }
                        if (cell.CellType == CellType.String)
                        {
                            row.CreateCell(j).SetCellValue(cell.StringCellValue);
                        }
                        if (cell.CellType == CellType.Boolean)
                        {
                            row.CreateCell(j).SetCellValue(cell.BooleanCellValue);
                        }
                        if (cell.CellType == CellType.Formula)
                        {
                            row.CreateCell(j).SetCellValue(cell.CellFormula);
                        }
                    }
                }
            }
            return workbook;
        }

        /// <summary>
        /// 在特定目录中匹配日志并返回其workbook
        /// </summary>
        /// <param name="sourceDirectory">日志所在目录</param>
        /// <param name="logName">日志匹配字符串</param>
        /// <returns></returns>
        public static IWorkbook GetWorkbook(string sourceDirectory, string logName)
        {
            DirectoryInfo di = new DirectoryInfo(sourceDirectory);
            var files = di.GetFiles($"*{logName}*");
            if (files.Length == 0)
                throw new Exception($"未在{sourceDirectory}下找到匹配特征字符串为*{logName}*的文件，");
            var workbook = GetWorkBook(files[0].FullName);
            return workbook;
        }

        /// <summary>
        /// 将workbook生成新文件
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="targetPath">目标路径</param>
        public static void WriteWorkbookToFile(IWorkbook workbook, string targetPath)
        {
            MemoryStream ms = new MemoryStream();
            workbook.Write(ms);

            using (FileStream fs = new FileStream(targetPath, FileMode.Create, FileAccess.Write))
            {
                byte[] bArr = ms.ToArray();
                fs.Write(bArr, 0, bArr.Length);
                fs.Flush();
            }
        }

        /// <summary>
        /// 取出特定单元格数值
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="positionInfo"></param>
        /// <returns></returns>
        public static double GetCellNumericValue(ISheet sheet, PositionInfo positionInfo, string sourceDirectory = "", int sheetIndex = 0, string fileName = "")
        {
            int x = positionInfo.X;
            int y = positionInfo.Y;

            double result = 0;

            try
            {
                IRow row = sheet.GetRow(y);
                ICell cell = row.GetCell(x);
                result = cell.NumericCellValue;
            }
            catch (Exception ex)
            {
                if (string.IsNullOrEmpty(sourceDirectory) || string.IsNullOrEmpty(fileName))
                {
                    throw ex;
                }
                else
                {
                    string errorInfo = string.Format(
                        "单元格格式非法:路径:{0},文件名标识:{1},sheet:{2},位置信息 X:{3} Y:{4}",
                        sourceDirectory,
                        fileName,
                        sheetIndex + 1,
                        (positionInfo.X + 1).ToString(),
                        (positionInfo.Y + 1).ToString()
                        );
                    throw new Exception(errorInfo);
                }
            }

            return result;
        }

        /// <summary>
        /// 取出特定单元格下所有单元格的数值（包括特定）
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="positionInfo"></param>
        /// <param name="sourceDirectory"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static List<double> GetColNumbers(ISheet sheet, PositionInfo positionInfo, string sourceDirectory = "", int sheetIndex = 0, string fileName = "")
        {
            int x = positionInfo.X;
            int y = positionInfo.Y;

            List<double> result = new List<double>();

            for (int count = y; count <= sheet.LastRowNum; count++)
            {
                IRow row = sheet.GetRow(count);
                ICell cell = row.GetCell(x);

                try
                {
                    double cellValue = cell.NumericCellValue;
                    result.Add(cellValue);
                }
                catch (Exception ex)
                {
                    if (string.IsNullOrEmpty(sourceDirectory) || string.IsNullOrEmpty(fileName))
                    {
                        throw ex;
                    }
                    else
                    {
                        string errorInfo = string.Format(
                            "单元格格式非法:路径:{0},文件名标识:{1},sheet:{2},位置信息 X:{3} Y:{4}",
                            sourceDirectory,
                            fileName,
                            sheetIndex + 1,
                            (positionInfo.X + 1).ToString(),
                            (count + 1).ToString()
                            );
                        throw new Exception(errorInfo);
                    }
                }
            }

            if (result.Count == 0)
            {
                string errorInfo = string.Format(
                            "未取到竖列数据:路径:{0},文件名标识:{1},sheet:{2},位置信息 X:{3} Y:{4}",
                            sourceDirectory,
                            fileName,
                            sheetIndex + 1,
                            (positionInfo.X + 1).ToString(),
                            (positionInfo.Y + 1).ToString()
                        );
                throw new Exception(errorInfo);
            }
            else
            {
                return result;
            }
        }

        /// <summary>
        /// 取出带表头名称的竖列数据
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="positionInfo"></param>
        /// <param name="sourceDirectory"></param>
        /// <param name="sheetIndex"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static KeyValuePair<string, List<double>> GetColNumbersWithHeader(ISheet sheet, PositionInfo positionInfo, string sourceDirectory = "", int sheetIndex = 0, string fileName = "")
        {
            string headerName = "";
            List<double> data = new List<double>();

            int x = positionInfo.X;
            int y = positionInfo.Y;

            for (int count = y; count <= sheet.LastRowNum; count++)
            {
                IRow row = sheet.GetRow(count);
                ICell cell = row.GetCell(x);

                if (count == y)
                {
                    headerName = cell.StringCellValue;
                }
                else
                {
                    double dataValue = cell.NumericCellValue;
                    data.Add(dataValue);
                }
            }

            if (string.IsNullOrEmpty(headerName) && data.Count == 0)
            {
                string errorInfo = string.Format(
                    "未取到竖列数据:路径:{0},文件名标识:{1},sheet:{2},位置信息 X:{3} Y:{4}",
                            sourceDirectory,
                            fileName,
                            sheetIndex + 1,
                            (positionInfo.X + 1).ToString(),
                            (positionInfo.Y + 1).ToString()
                    );
                throw new Exception(errorInfo);
            }

            KeyValuePair<string, List<double>> result = new KeyValuePair<string, List<double>>(headerName, data);
            return result;
        }

        /// <summary>
        /// 覆盖竖列数据
        /// </summary>
        /// <param name="data"></param>
        /// <param name="sheet"></param>
        /// <param name="positionInfo"></param>
        /// <param name="sourceDirectory"></param>
        /// <param name="sheetIndex"></param>
        /// <param name="fileName"></param>
        public static void UpdateColNumbers(List<double> data, ref ISheet sheet, string rtgNumTargetChannelName, string sourceDirectory = "", int sheetIndex = 0, string fileName = "")
        {


            IRow headersRow = sheet.GetRow(2);
            int dragonTvIndex = -1;
            for (int i = 0; i < headersRow.LastCellNum; i++)
            {
                ICell headerCell = headersRow.GetCell(i);
                if (headerCell.StringCellValue.Trim().Equals(rtgNumTargetChannelName.Trim()))
                {
                    dragonTvIndex = i;
                    break;
                }
            }

            if (dragonTvIndex == -1)
            {
                string errorInfo = string.Format(
                    "未找到频道“上海东方卫视”:路径:{0},文件名标识:{1},sheet:{2}",
                            sourceDirectory,
                            fileName,
                            sheetIndex + 1
                    );
                throw new Exception(errorInfo);
            }

            int x = dragonTvIndex;
            int y = 3;

            int toUpdateCount = sheet.LastRowNum - y + 1;
            if (toUpdateCount != data.Count)
            {
                string errorInfo = string.Format(
                            "更新数值数量不符合预期:路径:{0},文件名标识:{1},sheet:{2}更新起始位置信息 X:{3} Y:{4}",
                            sourceDirectory,
                            fileName,
                            sheetIndex + 1,
                            (x + 1).ToString(),
                            (3 + 1).ToString()
                        );
                throw new Exception(errorInfo);
            }

            for (int rowCount = y; rowCount <= sheet.LastRowNum; rowCount++)
            {
                IRow row = sheet.GetRow(rowCount);
                ICell cell = row.GetCell(dragonTvIndex);
                double dataValue = data[rowCount - y];
                cell.SetCellValue(dataValue);
            }

        }

        /// <summary>
        /// 增添带表头名称的竖列数据
        /// </summary>
        /// <param name="data"></param>
        /// <param name=""></param>
        /// <param name="positionInfo"></param>
        /// <param name="sourceDirectory"></param>
        /// <param name="fileName"></param>
        /// <param name="sheetIndex"></param>
        public static void InsertColNumbersWithHeader(Dictionary<string, List<double>> data, ref ISheet sheet, PositionInfo positionInfo, string sourceDirectory = "")
        {
            int x = positionInfo.X;
            int y = positionInfo.Y;

            int sheetDataRowsWithOutHeader = sheet.LastRowNum - y;

            //验证数据数目是否一致
            foreach (var item in data)
            {
                int dataCount = item.Value.Count;
                if (dataCount != sheetDataRowsWithOutHeader)
                {
                    string errorInfo = string.Format(
                        "目标数据数目不符合预期:路径{0},目标名称{1}",
                        sourceDirectory,
                        item.Key
                        );
                    throw new Exception(errorInfo);
                }
            }

            for (int count = y; count <= sheet.LastRowNum; count++)
            {

                IRow row = sheet.GetRow(count);

                foreach (var dataItem in data)
                {
                    int lastCellNum = row.LastCellNum - 1;
                    ICell cell = row.CreateCell(lastCellNum + 1);

                    int colNumCount = count - y - 1;
                    if (colNumCount != -1)
                    {
                        //set data
                        cell.SetCellValue(dataItem.Value[colNumCount]);
                    }
                    else
                    {
                        //set header
                        cell.SetCellValue(dataItem.Key);

                        //set empty metric name
                        IRow rowOverhead = sheet.GetRow(count - 1);
                        int lastCellNumOverhead = rowOverhead.LastCellNum - 1;
                        ICell cellOverhead = rowOverhead.CreateCell(lastCellNumOverhead + 1);
                    }

                }


            }

        }

        /// <summary>
        /// 判断workbook中sheet的某row的有效cell数目
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        public static int GetLastCellNum(IRow row)
        {
            //for (int i = 1; i<=row.LastCellNum;i++)
            //{
            //    if (row.GetCell(i - 1) == null||string.IsNullOrEmpty(row.GetCell(i-1).ToString()))
            //    {
            //        return i-1;
            //    }
            //}
            return row.LastCellNum;
        }
    }
}
