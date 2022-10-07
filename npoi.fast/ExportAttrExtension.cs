/*
 * @Author: HUA
 * @Date: 2022-10-07 11:13:43
 * @Last Modified by:   HUA
 * @Last Modified time: 2022-10-07 11:13:43
 */

using NPOI.SS.UserModel;
using System.Reflection;

namespace H.Npoi.Fast
{
    public static class ExportAttrExtension
    {
        public static ExportAttribute? Parse(this PropertyInfo property)
        {
            object? obj = property.GetCustomAttribute(typeof(ExportAttribute), true);
            if (obj == null)
            {
                return null;
            }
            ExportAttribute attr = obj as ExportAttribute ?? new ExportAttribute();
            if (string.IsNullOrWhiteSpace(attr.GetName()))
            {
                attr.SetName(property.Name);
            }
            return attr;
        }

        public static ExportAttribute? Parse(this Type type, PropertyInfo property)
        {
            ExportAttribute? classAttr = type.GetCustomAttribute(typeof(ExportAttribute), true) as ExportAttribute;
            if (classAttr != null)
            {
                classAttr.SetName(property.Name);
            }
            return classAttr;
        }
    }
}