using System;
using System.Diagnostics;

namespace OfficeLib.XLS
{
    /// <summary>
    /// Excel Class
    /// </summary>
    /// <remarks>
    /// Excel xxx Object Library wrapper.
    /// "xxx" is version.
    /// This library works with arbitrary version.
    /// "*.xls","*.xlsx","*.xlsm" ... etc
    /// </remarks>
    public class Excel : OfficeCore
    {
        #region --- Constant ---
        /// <summary>Application object ID</summary>
        protected const String PROG_ID = "Excel.Application";
        /// <summary>WorkBooks object ID</summary>
        protected const String OBJECT_WORKBOOKS = "Workbooks";
        /// <summary>Sheets object ID</summary>
        protected const String OBJECT_SHEET = "Sheets";
        /// <summary>Cells</summary>
        protected const String OBJECT_CELL = "Cells";
        /// <summary>Range</summary>
        protected const String OBJECT_RANGE = "Range";

        /// <summary>Value</summary>
        protected const String PROP_VALUE = "Value";
        /// <summary>Value2</summary>
        /// <remarks>
        /// The only difference between this property 
        /// and the Value property is that the Value2 property
        /// doesnÅft use the Currency and Date data types.
        /// You can return values formatted with
        /// these data types as floating-point numbers
        /// by using the Double data type.
        /// </remarks>
        protected const String PROP_VALUE2 = "Value2";
        /// <summary>Text</summary>
        protected const String PROP_TEXT = "Text";
        /// <summary>Fomula</summary>
        protected const String PROP_FOMULA = "Formula";

        /// <summary>Row</summary>
        protected const String PROP_ROW = "Row";
        /// <summary>Column</summary>
        protected const String PROP_COL = "Column";
        /// <summary>SheetsInNewWorkbook</summary>
        protected const String PROP_SHEET_IN_NEW_WORKBOOK = "SheetsInNewWorkbook";

        /// <summary>argument count of "Open" method</summary>
        protected static readonly Int32 ARGS_OPEN = 15;

        /// <summary>Row</summary>
        public static readonly Int32 ROW = 0;
        /// <summary>Columun</summary>
        public static readonly Int32 COL = 1;

        /// <summary>XlCorruptLoad Enumeration</summary>
        /// <remarks>Specifies the processing for a file when it is opened.</remarks>
        public enum XlCorruptLoad : Int32
        {
            /// <summary>Workbook is opened normally.</summary>
            xlNormalLoad = 0,
            /// <summary>Workbook is opened in repair mode.</summary>
            xlRepairFile = 1,
            /// <summary>Workbook is opened in extract data mode.</summary>
            xlExtractData = 2,
        }

        /// <summary>XlPlatform Enumeration</summary>
        /// <remarks>Specifies the platform on which a text file originated.</remarks>
        public enum XlPlatform : Int32
        {
            /// <summary>Macintosh</summary>
            xlMacintosh = 1,
            /// <summary>MS-DOS</summary>
            xlMSDOS = 3,
            /// <summary>Microsoft Windows</summary>
            xlWindows = 2,
        }

        /// <summary>XlGetValueType Enumeration</summary>
        /// <remarks>Specifies the format for retrieving values from a cell.</remarks>
        public enum XlGetValueFormat : Int32
        {
            /// <summary>Value</summary>
            xlValue = 0,
            /// <summary>Floating-point numbers</summary>
            xlValue2 = 1,
            /// <summary>Text</summary>
            xlText = 2,
            /// <summary>Fomula</summary>
            xlFormula = 3,
        }
    
        #endregion

        #region --- Properties ---
        /// <summary>Book object</summary>
        public Object Book { get; private set; }

        /// <summary>Sheet object</summary>
        public Object Sheet { get; private set; }

        /// <summary>Current sheet name</summary>
        public String CurrentSheetName
        {
            get { return this.Sheet.GetProperty(PROP_NAME) as String; }
        }

        /// <summary>Sheet names (internal variable)</summary>
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        protected String[] _sheetNames;
        
        /// <summary>Sheet names in the Book</summary>
        public String[] SheetNames
        {   // When not yet acquired : When acquired
            get { return this._sheetNames ?? GetWorkBookSheetNames(); }
        }
        #endregion

        #region --- Constructor ---
        /// <summary>
        /// Constructor
        /// </summary>
        public Excel() : base(PROG_ID) { this._sheetNames = null; }
        #endregion

        #region --- Open ---
        /// <summary>
        /// Excel file open
        /// </summary>
        /// <param name="file">File path</param>
        /// <returns></returns>
        public override Boolean Open(String file)
        {
            // Argument creation for Open
            // Items other than File path are set to Type.Missing
            return Open(new Object[] { this.Path });
        }

        /// <summary>
        /// Excel file open (details)
        /// </summary>
        /// <param name="options">file open arguments</param>
        /// <remarks>
        /// Argument parameter details
        /// No, Argument,                    Optional,   Type
        /// 01, FileName,                    true,       String
        /// 02, UpdateLinks,                 true,       AutomationSecurity
        /// 03, ReadOnly,                    true,       Boolean
        /// 04, Format,                      true,       Int32(1[Tab] 2[,] 3[Speace] 4[;] 5[none] 6[custom])
        /// 05, Password,                    true,       String
        /// 06, WriteResPassword,            true,       String
        /// 07, IgnoreReadOnlyRecommended,   true,       Boolean
        /// 08, Origin,                      true,       XlPlatform(enum)
        /// 09, Delimiter,                   true,       Char (If the format is of 6. Specify custom Delimiter)
        /// 10, Editable,                    true,       Boolean
        /// 11, Notify,                      true,       Boolean
        /// 12, Converter,                   true,       FileConverters
        /// 13, AddToMru,                    true,       Boolean
        /// 14, Local,                       true,       Boolean
        /// 15, CorruptLoad,                 true,       XlCorruptLoad(enum)
        /// 
        /// See MSDN for further details.
        /// </remarks>
        /// <returns></returns>
        public Boolean Open(params Object[] options)
        {
            try
            {
                this._sheetNames = null;
                this.Path = System.IO.Path.GetFullPath(options[0] as String ?? "");

                if (!System.IO.File.Exists(this.Path)) { return false; }

                base.CreateApplication();
                // Disable warning indication. Prevent program from stopping with warning
                this.Application.SetProperty(PROP_DISP_ALERT,
                                             new Object[] { MsoTriState.msoFalse });

                // Get the Excel book object
                this.WorkArea = this.Application.GetProperty(OBJECT_WORKBOOKS);

                // Todo: It needs to correspond to a file with a password
                // Open the Excel book
                this.Book = this.WorkArea.Method(METHOD_OPEN, SetOpenArguments(options));

                if (this.Book == null) { return false; }
                // By setting it to "saved", the save dialog is not displayed at the end.
                this.Book.SetProperty(PROP_SAVED, new Object[] { MsoTriState.msoTrue });
            }
            catch (Exception) { throw; }
            return true;
        }

        /// <summary>
        /// Set Open Arguments
        /// </summary>
        /// <param name="args">Arguments</param>
        /// <returns></returns>
        private Object[] SetOpenArguments(Object[] args)
        {
            var result = new Object[ARGS_OPEN];

            for (var i = 0; i < result.Length; i++)
            {   // For items for which there is no setting, set [Type.Missing]
                if (i < args.Length)
                {
                    result[i] = args[i] ?? Type.Missing;
                }
                else
                {
                    result[i] = Type.Missing;
                }
            }
            return result;
        }

        /// <summary>
        /// Create a new Excel file
        /// </summary>
        /// <param name="file">File path</param>
        /// <returns></returns>
        public Boolean New(String file)
        {
            try
            {
                this.Path = System.IO.Path.GetFullPath(file);
                base.CreateApplication();

                // Create a sheet. The argument is the number of sheets to be created
                this.Application.SetProperty(PROP_SHEET_IN_NEW_WORKBOOK, new Object[] { 1 });
                // Disable warning indication. Prevent program from stopping with warning
                this.Application.SetProperty(PROP_DISP_ALERT,
                                             new Object[] { MsoTriState.msoFalse });
                // Get the Excel Workbooks object
                this.WorkArea = this.Application.GetProperty(OBJECT_WORKBOOKS);
                // Create new Workbook object
                this.WorkArea.GetProperty(METHOD_ADD);
                // Get the Excel Workbooks object
                this.Book = this.WorkArea.GetProperty(PROP_ITEM, new Object[] { 1 });

                if (this.Book == null) { return false; }
                // By setting it to "saved", the save dialog is not displayed at the end.
                this.Book.SetProperty(PROP_SAVED, new Object[] { MsoTriState.msoTrue });
            }
            catch (Exception) { throw; }
            return true;
        }
        #endregion

        #region --- Close ---
        /// <summary>
        /// Close Excel
        /// </summary>
        /// <remarks>
        /// It must be done at the end.
        /// If not called, the process will keep capturing the file.
        /// </remarks>
        public override void Close()
        {
            try
            {   // Sheet list clear
                this._sheetNames = null;

                if (this.Book != null)
                {   // Close the Book
                    this.Book.Method(METHOD_CLOSE);
                }
                // Quit the Application
                QuitAplication();
            }
            catch (Exception) { throw; }
            finally
            {   // free the sheet, book and work area
                ReleaseObjects(this.Sheet, this.Book, this.WorkArea);
            }
        }
        #endregion

        #region --- Sheet ---
        /// <summary>
        /// Select the Sheet
        /// </summary>
        /// <param name="sheetName">Sheet name</param>
        /// <returns>Success(true), Failure(false)</returns>
        public Boolean SelectSheet(String sheetName)
        {
            try
            {
                this.Sheet = this.Book.GetProperty(OBJECT_SHEET, new Object[] { sheetName });
            }
            catch (Exception)
            {
                return false;
            }
            return this.Sheet != null;
        }

        /// <summary>
        /// Get a list of sheet names
        /// </summary>
        /// <returns>Sheet names</returns>
        public String[] GetWorkBookSheetNames()
        {
            String[] result = null;
            Object sheets = null;
            try
            {   // Get number of sheets
                sheets = this.Book?.GetProperty(OBJECT_SHEET);
                Object countObject = sheets?.GetProperty(PROP_COUNT); 
                var count = Convert.ToInt32(countObject ?? 0);
                Object sheet = null;
                result = new String[count];

                for (var i = 0; i < count; i++)
                {   // Get a Sheet on the basis of the number
                    sheet = sheets.GetProperty(PROP_ITEM, new Object[] { i + 1 });
                    // Get Sheet name
                    result[i] = sheet?.GetProperty(PROP_NAME) as String;
                }
            }
            catch (Exception) { throw; }
            finally { ReleaseObject(sheets); }
            return result;
        }

        /// <summary>
        /// Add the sheet
        /// </summary>
        /// <param name="sheetName">Sheet name</param>
        public Boolean AddSheet(String sheetName)
        {
            // Todo: Add sheet method
            throw new Exception("work in progress");
        }

        /// <summary>
        /// Remove the sheet
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public Boolean RemoveSheet(String sheetName)
        {
            // Todo: Remove sheet method
            throw new Exception("work in progress");
        }
        #endregion

        #region --- Cell ---
        #region --- Get Value ---
        /// <summary>
        /// Get value from sheet
        /// </summary>
        /// <param name="startCol">Start column</param>
        /// <param name="startRow">Start row</param>
        /// <param name="endCol">End column</param>
        /// <param name="endRow">End row</param>
        /// <param name="referenceFormat">Cell value reference format</param>
        /// <remarks>
        /// Cell value options. [Value, Value2, Text, Fomula]
        /// 
        /// example : Input "=DATE(2017,3,27)"
        /// value = 2017/03/27
        /// value2 = 42821
        /// Text = 27 March 2007
        /// Fomula = DATE(2017,3,27)
        /// </remarks>
        public String[,] GetCellValue(UInt32 startRow, UInt32 startCol,
                                      UInt32 endRow, UInt32 endCol,
                                      XlGetValueFormat referenceFormat)
        {
            if (this.Sheet == null ){ return null; }
            Object range = null;
            Object values = null;
            try
            {   // Reference range acquisition
                range = GetRange(startRow, startCol, endRow, endCol);
                values = GetValue(range, referenceFormat);

                // The value was an array type
                if (values is Object[,])    
                {   
                    var temp = values as Object[,];
                    var result = new String[temp.GetLength(0), temp.GetLength(1)];

                    for (var r = 0; r < result.GetLength(0); r++)
                    {
                        for (var c = 0; c < result.GetLength(1); c++)
                        {   // convert object to string
                            result[r, c] = temp[r + 1, c + 1]?.ToString() ?? String.Empty;
                        }
                    }
                    return result;
                }
                // The value was Object type
                else if (values is Object)
                {
                    return new String[,] { { values?.ToString() ?? String.Empty } };
                }
                return null;
            }
            catch (Exception) { throw; }
            finally
            {
                values = null;
                ReleaseObjects(range);
            }
        }

        /// <summary>
        /// Get the value by reference format specification
        /// </summary>
        /// <param name="range">Object of the range to get the value</param>
        /// <param name="referenceFormat">Cell value reference format</param>
        /// <returns></returns>
        private Object GetValue(Object range, XlGetValueFormat referenceFormat)
        {
            switch (referenceFormat)
            {
                case XlGetValueFormat.xlValue:
                    return range.GetProperty(PROP_VALUE);
                case XlGetValueFormat.xlValue2:
                    return range.GetProperty(PROP_VALUE2);
                case XlGetValueFormat.xlText:
                    return range.GetProperty(PROP_TEXT);
                case XlGetValueFormat.xlFormula:
                    return range.GetProperty(PROP_FOMULA);
                default:
                    return null;
            }
        }

        /// <summary>
        /// Get value from sheet
        /// </summary>
        /// <param name="startAdress">Start adress</param>
        /// <param name="endAdress">End adress</param>
        /// <param name="referenceFormat">Cell value reference type</param>
        public String[,] GetCellValue(String startAdress, String endAdress,
                                      XlGetValueFormat referenceFormat)
        {
            var start = startAdress.ToAddress();
            var end = endAdress.ToAddress();
            return GetCellValue(start.Row, start.Column,
                                end.Row, end.Column, referenceFormat);
        }
        #endregion

        #region --- Set Value ---
        /// <summary> 
        /// Set value to sheet
        /// </summary>
        /// <param name="values">Setting values</param>
        /// <param name="startCell">Start Address</param>
        /// <param name="referenceFormat">Cell value reference format</param>
        public Boolean SetCellValue(Object[,] values, String startCell,
                                    XlGetValueFormat referenceFormat)
        {
            var startAddress = startCell.ToAddress();
            return SetCellValue(values, startAddress.Row, startAddress.Column, referenceFormat);
        }

        /// <summary> 
        /// Set value to sheet
        /// </summary>
        /// <param name="values">Setting values</param>
        /// <param name="startRow">Start row</param>
        /// <param name="startCol">Start column</param>
        /// <param name="referenceFormat">Cell value reference format</param>
        public Boolean SetCellValue(Object[,] values, UInt32 startRow, UInt32 startCol,
                                    XlGetValueFormat referenceFormat)
        {
            UInt32 endRow = startRow + (UInt32)values.GetLength(0) - 1;
            UInt32 endCol = startCol + (UInt32)values.GetLength(1) - 1;

            return SetCellValue(values, startRow, startCol, endRow, endCol, referenceFormat);
        }

        /// <summary>
        /// Set value to sheet
        /// </summary>
        /// <param name="values">Setting values</param>
        /// <param name="startAddressString">Start Address</param>
        /// <param name="endAddressString">End Address</param>
        /// <param name="referenceFormat">Cell value reference format</param>
        public Boolean SetCellValue(Object[,] values, String startAddressString,
                                    String endAddressString, XlGetValueFormat referenceFormat)
        {
            var startAddress = startAddressString.ToAddress();
            var endAddress = endAddressString.ToAddress();
            return SetCellValue(values, startAddress.Row, startAddress.Column,
                                        endAddress.Row, endAddress.Column, referenceFormat);
        }

        /// <summary>
        /// Set value to sheet
        /// </summary>
        /// <param name="values">Setting values</param>
        /// <param name="startRow">Start row</param>
        /// <param name="startCol">Start column</param>
        /// <param name="endRow">End row</param>
        /// <param name="endCol">End column</param>
        /// <param name="referenceFormat">Cell value reference format</param>
        public Boolean SetCellValue(Object[,] values, UInt32 startRow, UInt32 startCol,
                                    UInt32 endRow, UInt32 endCol, XlGetValueFormat referenceFormat)
        {
            Object range = null;
            try
            {
                range = GetRange(startRow, startCol, endRow, endCol);
                Object setValue = ReplaceNullValue(values, startRow, startCol,
                                                           endRow, endCol);
                range.SetProperty(GetSetFormat(referenceFormat), new Object[] { setValue });
            }
            catch (Exception) { throw; }
            finally
            {
                ReleaseObject(range);
            }
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="referenceFormat"></param>
        /// <returns></returns>
        private String GetSetFormat(XlGetValueFormat referenceFormat)
        {
            switch (referenceFormat)
            {
                case XlGetValueFormat.xlValue:
                    return PROP_VALUE;
                case XlGetValueFormat.xlValue2:
                    return PROP_VALUE2;
                case XlGetValueFormat.xlText:
                    return PROP_TEXT;
                case XlGetValueFormat.xlFormula:
                    return PROP_FOMULA;
            }
            return PROP_VALUE2;
        }
        #endregion
        #endregion

        #region --- Save ---
        /// <summary>
        /// Save
        /// </summary>
        public void Save()
        {
            try
            {   // Get this Book path
                Object path = this.Book.GetProperty(PROP_PATH);
                if ((path as String).IsNullOrEmpty())
                {
                    if (this.Path.IsNullOrEmpty())
                    {
                        throw new NullReferenceException();
                    }
                    SaveAs(this.Path);
                }
                else
                {
                    this.Book.Method(METHOD_SAVE);
                }
            }
            catch (Exception) { throw; }
        }

        /// <summary>
        /// Save As
        /// </summary>
        /// <param name="file">Save file path</param>
        public void SaveAs(String file)
        {
            try
            {
                if(file.IsNullOrEmpty()) { throw new NullReferenceException(); }
                this.Book.Method(METHOD_SAVE_AS,
                                 new Object[] { System.IO.Path.GetFullPath(file) });
            }
            catch (Exception) { throw; }
        }
        #endregion

        #region --- Function ---
        /// <summary>Convert Jag Array to Rectangular Array</summary>
        /// <param name="src">Jag Array</param>
        public static Object[,] ConvertSetValue<T>(T[][] src)
        {
            Int32 columns = src.ColumnsMax();
            Int32 r = 0;
            Int32 c = 0;
            var result = new Object[src.Length, columns];
            
            for (; r < src.Length; r++)
            {
                for (c = 0; c < columns; c++)
                {
                    result[r, c] = c < src[r].Length ? (Object)src[r][c] : null;
                }
            }
            return result;
        }

        /// <summary>Get Cell range</summary>
        /// <param name="startRow">Start row</param>
        /// <param name="startCol">Start column</param>
        /// <param name="endRow">End row</param>
        /// <param name="endCol">End colmun</param>
        public Object GetRange(UInt32 startRow, UInt32 startCol, UInt32 endRow, UInt32 endCol)
        {
            try
            {   // Get cell address
                Object stCell = this.Sheet.GetProperty(OBJECT_CELL,
                                            new Object[] { startRow, startCol });
                Object edCell = this.Sheet.GetProperty(OBJECT_CELL,
                                            new Object[] { endRow, endCol });
                // return cell range
                return this.Sheet.GetProperty(OBJECT_RANGE, new Object[] { stCell, edCell });
            }
            catch (Exception) { throw; }
        }

        /// <summary>
        /// Display blanks instead of "N / A".
        /// </summary>
        /// <param name="values">source values</param>
        /// <param name="startRow">Start row</param>
        /// <param name="startCol">Start column</param>
        /// <param name="endRow">End row</param>
        /// <param name="endCol">End colmun</param>
        public Object[,] ReplaceNullValue(Object[,] values, UInt32 startRow, UInt32 startCol,
                                                            UInt32 endRow, UInt32 endCol)
        {
            // Adjust the size of the variable according to the range of the cell to be written.
            var result = new Object[(endRow - startRow) + 1, (endCol - startCol) + 1];

            Int32 row = result.GetLength(0);
            Int32 col = result.GetLength(1);
            Int32 valRow = values.GetLength(0);
            Int32 valCol = values.GetLength(1);

            for (var i = 0; i < row; i++)
            {
                for (var j = 0; j < col; j++)
                {   // Is it not within the range?
                    result[i, j] = (valRow > i && valCol > j) ? values[i, j] ?? "" : "";
                }
            }
            return result;
        }
        #endregion

        /// <summary>
        /// ToString
        /// </summary>
        public override String ToString() 
            => String.Format("Book:{0}", System.IO.Path.GetFileNameWithoutExtension(this.Path));
    }
}