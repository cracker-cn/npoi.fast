using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;

namespace npoi.helper
{
    public class ExportUtils
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
            if (suffix == ".xlsx")
            {
                return new XSSFWorkbook(stream);
            }
            else if (suffix == ".xls")
            {
                return new HSSFWorkbook(stream);
            }
            else
            {
                throw new Exception("文件格式错误");
            }
        }
    }
}