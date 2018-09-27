using System;

namespace OfficeLib.XLS
{
    /// <summary>
    /// Commands used in Excel
    /// </summary>
    public class ExcelCommands
    {
        /// <summary>Cells</summary>
        public const String OBJECT_CELL = "Cells";
        /// <summary>Range</summary>
        public const String OBJECT_RANGE = "Range";
        /// <summary>Sheets object ID</summary>
        public const String OBJECT_SHEET = "Sheets";
        /// <summary>WorkBooks object ID</summary>
        public const String OBJECT_WORKBOOKS = "Workbooks";


        /// <summary>Column</summary>
        public const String PROP_COL = "Column";
        /// <summary>Fomula</summary>
        public const String PROP_FOMULA = "Formula";
        /// <summary>Interior</summary>
        public const String PROP_INTERIOR = "Interior";
        /// <summary>Row</summary>
        public const String PROP_ROW = "Row";
        /// <summary>SheetsInNewWorkbook</summary>
        public const String PROP_SHEET_IN_NEW_WORKBOOK = "SheetsInNewWorkbook";
        /// <summary>Text</summary>
        public const String PROP_TEXT = "Text";
        /// <summary>Value</summary>
        public const String PROP_VALUE = "Value";
        /// <summary>Value2</summary>
        /// <remarks>
        /// The only difference between this property 
        /// and the Value property is that the Value2 property
        /// doesn’t use the Currency and Date data types.
        /// You can return values formatted with
        /// these data types as floating-point numbers
        /// by using the Double data type.
        /// </remarks>
        public const String PROP_VALUE2 = "Value2";

        /// <summary>SpecialCells</summary>
        public const String PROP_SPECIAL_CALLS = "SpecialCells";
    }
}
