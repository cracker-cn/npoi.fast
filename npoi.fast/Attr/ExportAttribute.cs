/*
 * @Author: HUA 
 * @Date: 2022-10-07 11:13:22 
 * @Last Modified by:   HUA 
 * @Last Modified time: 2022-10-07 11:13:22 
 */
using NPOI.SS.UserModel;

namespace H.Npoi.Fast
{
    /// <summary>
    /// 导出设置，Xalign,Yalign作用于所有行，其余只对表头有效
    /// </summary>
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Property)]
    public class ExportAttribute : System.Attribute
    {
        private string? Name;
        public int FontSize = 14;
        public bool Isbold = false;
        public string Color = "#000000";
        public string? BgColor;
        public HorizontalAlignment Xalign = HorizontalAlignment.Center;
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