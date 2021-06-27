using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Reflection;

namespace spreadsheet_helper.Model
{
    /// <summary>
    /// Helper class for manipulating and managing spreadsheet columns
    /// </summary>
    public class HelperColumns
    {
        public HelperColumns(Type type, string classProperty, ICellStyle style, string title, bool? includeOnSheet)
        {
            Style = new XSSFWorkbook().CreateCellStyle();
            Style.CloneStyleFrom(style);
            ClassProperty = type.GetProperty(classProperty);
            Title = !String.IsNullOrWhiteSpace(title) ? title : ClassProperty.Name;
            IncludeOnSheet = includeOnSheet != null ? includeOnSheet : true;
        }

        public HelperColumns(PropertyInfo classProperty, ICellStyle style, string title, bool? includeOnSheet)
        {
            ClassProperty = classProperty;
            Style = new XSSFWorkbook().CreateCellStyle();
            Style.CloneStyleFrom(style);
            Title = !String.IsNullOrWhiteSpace(title) ? title : classProperty.Name;
            IncludeOnSheet = includeOnSheet != null ? includeOnSheet : true;
        }

        public PropertyInfo ClassProperty { get; set; }
        public ICellStyle Style { get; set; }
        public string Title { get; set; }
        public bool? IncludeOnSheet { get; set; }
    }
}
