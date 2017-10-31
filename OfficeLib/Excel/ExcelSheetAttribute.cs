using System;

namespace OfficeLib.XLS
{
    /// <summary>Excel sheet attribute</summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false)]
    public class ExcelSheetAttribute : PageAttribute
    {
        /// <summary>Defeult : Row max value</summary>
        public static readonly UInt32 DEF_ROW_MAX = 1000;
        /// <summary>Defeult : Column max value</summary>
        public static readonly UInt32 DEF_COL_MAX = 1000;

        /// <summary>Row max value</summary>
        public UInt32 RowMax { get; set; }
        /// <summary>Column max value</summary>
        public UInt32 ColMax { get; set; }

        /// <summary>Constructor</summary>
        public ExcelSheetAttribute(EnumSheetPermission SheetParm) : base(SheetParm)
        {
            this.RowMax = DEF_ROW_MAX;
            this.ColMax = DEF_COL_MAX;
        }
    }
}
