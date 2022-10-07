/*
 * @Author: HUA 
 * @Date: 2022-10-07 11:14:58 
 * @Last Modified by:   HUA 
 * @Last Modified time: 2022-10-07 11:14:58 
 */
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace H.Npoi.Fast
{
    public class BaseExport<T>where T:class
    {
        public void Do(T entity)
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet=wb.CreateSheet();
            List<ExportAttribute> attributes = ExcelUtils.EntityAttrs<T>();
            IRow row = sheet.CreateRow(1);
            for (int i = 0; i < attributes.Count; i++)
            {
                ExportAttribute attr =attributes[i];
                ICellStyle headStyle = wb.CreateCellStyle();
                headStyle.Alignment = attr.Xalign;
                headStyle.VerticalAlignment = attr.Yalign;
                ICell cell = row.CreateCell(i+1);
            }
            foreach(PropertyInfo item in typeof(T).GetProperties())
            {
                Console.WriteLine(item.GetValue(entity));
            }
        }
    }
}
