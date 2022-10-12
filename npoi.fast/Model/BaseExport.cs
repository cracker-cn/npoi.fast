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
    public sealed class BaseExport<T>where T:class
    {
        public void Do(T entity)
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet=wb.CreateSheet();
            List<ExportAttribute> attributes = ExcelUtils.EntityAttrs<T>();
            IRow row = sheet.CreateRow(0);
            for (int i = 0; i < attributes.Count; i++)
            {
                ExportAttribute attr =attributes[i];
                ICellStyle headStyle = wb.CreateCellStyle();
                headStyle.Alignment = attr.Xalign;
                headStyle.VerticalAlignment = attr.Yalign;
                headStyle.FillForegroundColor =attr.BgColor;
                headStyle.FillPattern = FillPattern.SolidForeground;
                IFont font = wb.CreateFont();
                font.Color = attr.Color;
                font.IsBold = attr.Isbold;
                font.FontHeightInPoints = attr.FontSize;
                headStyle.SetFont(font);
                ICell cell = row.CreateCell(i);
                cell.CellStyle = headStyle;
                cell.SetCellValue(attr.GetName());
                sheet.SetColumnWidth(i,256*(attr.GetName()?.Length??0)*6);
            }
            IRow content = sheet.CreateRow(1);
            var props = typeof(T).GetProperties();
            for (int i = 0; i < props.Length; i++)
            {
                ICellStyle cellStyle = wb.CreateCellStyle();
                cellStyle.VerticalAlignment=VerticalAlignment.Center;
                cellStyle.Alignment = HorizontalAlignment.Center;
                ICell cell = content.CreateCell(i);
                cell.CellStyle = cellStyle;
                cell.SetCellValue(props[i].GetValue(entity)?.ToString());
            }
            using FileStream fs = new FileStream(@"D:\Upload\1.xls", FileMode.OpenOrCreate, FileAccess.Write);
            wb.Write(fs);
        }
    }
}
