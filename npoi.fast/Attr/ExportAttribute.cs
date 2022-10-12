/*
 * @Author: HUA 
 * @Date: 2022-10-07 11:13:22 
 * @Last Modified by:   HUA 
 * @Last Modified time: 2022-10-07 11:13:22 
 */
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System.ComponentModel.DataAnnotations;

namespace H.Npoi.Fast
{
    /// <summary>
    /// 导出设置表头
    /// </summary>
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Property)]
    public class ExportAttribute : System.Attribute
    {
        /// <summary>
        /// 表头文字
        /// </summary>
        private string? Name;
        /// <summary>
        /// 尺寸
        /// </summary>
        [Range(6,128)]
        public int FontSize = 14;
        /// <summary>
        /// 是否加粗
        /// </summary>
        public bool Isbold = true;
        /// <summary>
        /// 文字颜色
        /// </summary>
        public short Color = HSSFColor.White.Index;
        /// <summary>
        /// 背景色
        /// </summary>
        public short BgColor = HSSFColor.SkyBlue.Index;
        /// <summary>
        /// 单元格水平对齐方式
        /// </summary>
        public HorizontalAlignment Xalign = HorizontalAlignment.Center;
        /// <summary>
        /// 单元格垂直对齐方式
        /// </summary>
        public VerticalAlignment Yalign = VerticalAlignment.Center;

        public ExportAttribute()
        {
           
        }

        public ExportAttribute(string name)
        {
            Name = name;
        }

        public string? GetName()
        {
            return Name;
        }
        public void SetName(string name)
        {
            Name=name;
        }
    }
}