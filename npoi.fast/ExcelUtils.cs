/*
 * @Author: HUA 
 * @Date: 2022-10-07 11:13:39 
 * @Last Modified by:   HUA 
 * @Last Modified time: 2022-10-07 11:13:39 
 */
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data;
using System.IO;
using System.Reflection;

namespace H.Npoi.Fast
{
    public class ExcelUtils
    {
        /// <summary>
        /// 导入excel，写入list
        /// </summary>
        /// <param name="filePath">excel文件地址</param>
        /// <param name="columns">列数</param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static List<string[]> ToList(string filePath, int columns)
        {
            if (!File.Exists(filePath))
            {
                throw new Exception("文件不存在");
            }
            List<string[]> list = new List<string[]>();
            string suffix = Path.GetExtension(filePath).ToLower();//后缀
            using FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            IWorkbook workbook=CreateWorkbook(fs,suffix);
            ISheet sheet = workbook.GetSheetAt(0);
            for (int i = 1; i < sheet.PhysicalNumberOfRows; i++)//跳过第一行表头
            {
                string[] rowArray = new string[columns];
                IRow row = sheet.GetRow(i);
                if (row != null && !string.IsNullOrEmpty(row.Cells[0].StringCellValue))
                {
                    for (int j = 0; j < columns; j++)
                    {
                        if (row.GetCell(j) != null)
                        {
                            rowArray[j] = row.GetCell(j).StringCellValue;
                        }
                    }
                    list.Add(rowArray);
                }
            }
            return list;
        }

        /// <summary>
        /// 创建工作簿
        /// </summary>
        /// <param name="stream">excel文件流</param>
        /// <param name="suffix">后缀</param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static IWorkbook CreateWorkbook(Stream stream,string suffix)
        {
            if (suffix == ".xls")
            {
                return new HSSFWorkbook(stream);
            }
            else if (suffix == ".xlsx")
            {
                return new XSSFWorkbook(stream);
            }
            else
            {
                throw new Exception("文件格式错误");
            }
        }

        /// <summary>
        /// 创建工作簿
        /// </summary>
        /// <param name="suffix">后缀</param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static IWorkbook CreateWorkbook(string suffix)
        {
            if (suffix == ".xls")
            {
                return new HSSFWorkbook();
            }
            else if (suffix == ".xlsx")
            {
                return new XSSFWorkbook();
            }
            else
            {
                throw new Exception("文件格式错误");
            }
        }

        /// <summary>
        /// 创建工作簿
        /// </summary>
        /// <param name="index">0 .xls,1 .xlsx</param>
        /// <returns></returns>
        public static IWorkbook CreateWorkbook(int index=0)
        {
            return index == 0 ? new HSSFWorkbook() : new XSSFWorkbook();
        }

        /// <summary>
        /// 获取实体导出属性
        /// </summary>
        /// <typeparam name="T">实体类</typeparam>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static List<ExportAttribute> EntityAttrs<T>()
        {
            List<ExportAttribute> list = new List<ExportAttribute>();
            Type type = typeof(T);
            PropertyInfo[] propertyInfos = type.GetProperties();
            if (propertyInfos.Length == 0)
            {
                throw new Exception("失败：实体属性为空");
            }
            foreach (PropertyInfo propInfo in propertyInfos)
            {
                ExportAttribute? attr = propInfo.Parse();
                if (attr != null)
                {
                    list.Add(attr);
                }
                else
                {
                    ExportAttribute? classAttr = type.Parse(propInfo);
                    if (classAttr != null)
                    {
                        list.Add(classAttr);
                    }
                }
            }
            return list;
        }

    }
}