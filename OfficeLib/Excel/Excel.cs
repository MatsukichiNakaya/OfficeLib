using System;
using System.Diagnostics;
using System.Reflection.Emit;
using System.Runtime.Remoting.Messaging;
using static OfficeLib.Commands;
using static OfficeLib.XLS.ExcelCommands;

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

        /// <summary>argument count of "Open" method</summary>
        protected static readonly Int32 ARGS_OPEN = 15;

        /// <summary>Row</summary>
        public static readonly Int32 ROW = 0;
        /// <summary>Columun</summary>
        public static readonly Int32 COL = 1;
    
        #endregion

        #region --- Properties ---
        /// <summary>Book object</summary>
        internal Object Book { get; private set; }

        /// <summary>Sheet object</summary>
        internal Object Sheet { get; private set; }

        /// <summary>Active sheet name</summary>
        public String ActiveSheetName
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

        /// <summary>Auto Save</summary>
        public Boolean IsAutoSave { get; set; } = false;
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
        /// Excel file open (optional auto save)
        /// </summary>
        /// <param name="file">File path</param>
        /// <param name="isAutoSave">true : Auto, false : Manual</param>
        /// <returns></returns>
        public Boolean Open(String file, Boolean isAutoSave)
        {
            this.IsAutoSave = isAutoSave;
            return Open(file);
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
                result[i] = i < args.Length ? args[i] ?? Type.Missing : Type.Missing;
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
            try {
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
            try {
                // Option auto save
                if (this.IsAutoSave) {
                    Save();
                }

                // Sheet list clear
                this._sheetNames = null;

                if (this.Book != null) {
                    // Close the Book
                    this.Book.Method(METHOD_CLOSE);
                }
                // Quit the Application
                QuitAplication();
            }
            catch (Exception) { throw; }
            finally {
                // free the sheet, book and work area
                ReleaseObjects(this.Sheet, this.Book, this.WorkArea);
            }
        }
        #endregion

        #region --- Sheet ---
        /// <summary>
        /// Select the Sheet(Sheet name)
        /// </summary>
        /// <param name="sheetName">Sheet name</param>
        /// <returns>Success(true), Failure(false)</returns>
        public Boolean SelectSheet(String sheetName)
        {
            try {
                this.Sheet = this.Book.GetProperty(OBJECT_SHEET, new Object[] { sheetName });
            }
            catch (Exception) {
                return false;
            }
            return this.Sheet != null;
        }

        /// <summary>
        /// Select the Sheet(Sheet index)
        /// </summary>
        /// <param name="sheetIndex"></param>
        /// <returns></returns>
        public Boolean SelectSheet(Int32 sheetIndex)
        {
            try {
                this.Sheet = this.Book.GetProperty(OBJECT_SHEET, new Object[] { this.SheetNames[sheetIndex] });
            }
            catch (Exception) {
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
            try {
                // Get number of sheets
                sheets = this.Book?.GetProperty(OBJECT_SHEET);
                Object countObject = sheets?.GetProperty(PROP_COUNT);
                var count = Convert.ToInt32(countObject ?? 0);
                Object sheet = null;
                result = new String[count];

                for (var i = 0; i < count; i++) {
                    // Get a Sheet on the basis of the number
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
            if (this.SheetNames.Contains(sheetName)) { return false; }
            Object sheets = null;
            Object newSheet = null;
            Boolean result = false;
            try {
                sheets = this.Book.GetProperty(OBJECT_SHEET);
                newSheet = sheets.Method(METHOD_ADD, new Object[] { Type.Missing, Type.Missing, 1, XlSheetType.xlWorksheet });
                newSheet.SetProperty(PROP_NAME, new Object[] { sheetName });
                result = true;
            }
            catch (Exception) {
                result = false;
            }
            finally
            {
                ReleaseObjects(sheets, newSheet);
            }

            return result;
        }

        /// <summary>
        /// Remove the sheet
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public Boolean RemoveSheet(String sheetName)
        {
            // Sheet not found.
            if (!SelectSheet(sheetName)) { return false; }

            this.Sheet.Method(METHOD_DELETE);

            // Set Ather sheet
            if (!SelectSheet(0)) {
                // Workbook Item is none.
                ReleaseObject(this.Sheet);
            }
            return true;
        }

        /// <summary>
        /// Copy the sheet
        /// </summary>
        /// <param name="sourceSheetName"></param>
        /// <param name="destSheetName"></param>
        /// <returns></returns>
        public Boolean CopySheet(String sourceSheetName, String destSheetName)
        {
            Object sheets = null;
            Object target = null;
            Object firstSheet = null;
            Object destSheet = null;

            Boolean result = false;
            try {
                sheets = this.Book?.GetProperty(OBJECT_SHEET);
                if(sheets  == null) { return false; }

                target = sheets.GetProperty(PROP_ITEM, new Object[] { sourceSheetName });
                firstSheet = sheets.GetProperty(PROP_ITEM, new Object[] { 1 });
                if (target == null || firstSheet == null) { return false; }

                // Copy sheets to the first position.
                target.Method(METHOD_COPY, new Object[] { firstSheet, Type.Missing });

                // Change name.
                destSheet = sheets.GetProperty(PROP_ITEM, new Object[] { 1 });
                destSheet.SetProperty(PROP_NAME, new Object[] { destSheetName });

                result = true;
            }
            finally {
                ReleaseObjects(target, firstSheet, destSheet, sheets);
            }
            return result;
        }

        /// <summary>
        /// Move the sheet
        /// </summary>
        /// <param name="sourceSheetName"></param>
        /// <param name="beforeSheetName"></param>
        /// <param name="afterSheetName"></param>
        /// <returns></returns>
        public Boolean MoveSheet(String sourceSheetName, String beforeSheetName = null, String afterSheetName = null)
        {
            Object sheets = null;
            Object target = null;
            Object followSheet = null;
            Boolean result = false;

            try {
                sheets = this.Book.GetProperty(OBJECT_SHEET);
                if (sheets == null) { return false; }

                target = sheets.GetProperty(PROP_ITEM, new Object[] { sourceSheetName });
                if (target == null) { return false; }

                if (String.IsNullOrWhiteSpace(beforeSheetName)) {
                    if (String.IsNullOrWhiteSpace(afterSheetName)) { return false; }
                    followSheet = sheets.GetProperty(PROP_ITEM, new Object[] { afterSheetName });
                    target.Method(METHOD_MOVE, new Object[] { Type.Missing, followSheet });
                }
                else {
                    followSheet = sheets.GetProperty(PROP_ITEM, new Object[] { beforeSheetName });
                    target.Method(METHOD_MOVE, new Object[] { followSheet, Type.Missing });
                }
                result = true;
            }
            finally {
                ReleaseObjects(target, followSheet, sheets);
            }
            return result;
        }

        /// <summary>
        /// Set sheet properties
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="propertyName"></param>
        /// <param name="values"></param>
        public void SetSheetProperty(String sheetName, String propertyName, params Object[] values)
        {
            Object sheet = null;
            try {
                sheet = this.Book.GetProperty(OBJECT_SHEET, new Object[] { sheetName });
                sheet.SetProperty(propertyName, values);
            }
            finally {
                ReleaseObject(sheet);
            }
        }

        /// <summary>
        /// Get sheet properties
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="propertyName"></param>
        /// <param name="values"></param>
        /// <returns></returns>
        public Object GetSheetProperty(String sheetName, String propertyName, params Object[] values)
        {
            Object sheet = null;
            Object result = null;
            try {
                sheet = this.Book.GetProperty(OBJECT_SHEET, new Object[] { sheetName });
                result = sheet.GetProperty(propertyName, values);
            }
            finally {
                ReleaseObject(sheet);
            }
            return result;
        }

        /// <summary>
        /// Get Sheet Object class
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public SheetObject GetSheet(String sheetName)
        {
            return new SheetObject(
                this.Book.GetProperty(OBJECT_SHEET, new Object[] { sheetName }));
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
            try {
                // Reference range acquisition
                range = GetRange(startRow, startCol, endRow, endCol);
                values = GetValue(range, referenceFormat);

                // The value was an array type
                if (values is Object[,]) {   
                    var temp = values as Object[,];
                    var result = new String[temp.GetLength(0), temp.GetLength(1)];

                    for (var r = 0; r < result.GetLength(0); r++) {
                        for (var c = 0; c < result.GetLength(1); c++) {
                            // convert object to string
                            result[r, c] = temp[r + 1, c + 1]?.ToString() ?? String.Empty;
                        }
                    }
                    return result;
                }
                // The value was Object type
                else if (values is Object) {
                    return new String[,] { { values?.ToString() ?? String.Empty } };
                }
                return null;
            }
            catch (Exception) { throw; }
            finally {
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
            switch (referenceFormat) {
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
        /// Get cell refarence format string
        /// </summary>
        /// <param name="referenceFormat">Value format type</param>
        /// <returns>Format String</returns>
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

        #region --- Specia Cell ---
        /// <summary>
        /// Get the largest cell of the edited range
        /// </summary>
        /// <returns>largest cell string</returns>
        public String GetLastCell()
        {
            Object cells = null;
            Object range = null;
            String result = String.Empty;
            try
            {
                cells = this.Sheet?.GetProperty(OBJECT_CELL);
                range = cells?.GetProperty(PROP_SPECIAL_CALLS, new Object[] { XlCellType.xlCellTypeLastCell });
                Int32 row = range?.GetProperty(PROP_ROW).To<Int32>() ?? 1;
                Int32 col = range?.GetProperty(PROP_COL).To<Int32>() ?? 1;
                
                result = $"{Address.ToExcelColumnString((UInt32)col)}{row}";
            }
            finally
            {
                ReleaseObjects(cells, range);
            }
            return result;
        }
        #endregion

        #region --- Copy ---
        /// <summary>
        /// Cell Copy
        /// </summary>
        /// <param name="sheetName">Sheet name</param>
        /// <param name="startAdress">Start adress</param>
        /// <param name="endAdress">End adress</param>
        public void CellCopy(String sheetName, String startAdress, String endAdress)
        {
            this.SelectSheet(sheetName);
            var start = startAdress.ToAddress();
            var end = endAdress.ToAddress();

            Object range = null;
            try {
                range = GetRange(start.Row, start.Column, end.Row, end.Column);
                range.Method(METHOD_COPY);
            }
            catch (Exception) { throw; }
            finally {
                ReleaseObject(range);
            }
        }
        #endregion

        #region --- Paste ---
        /// <summary>
        /// Cell Paste
        /// </summary>
        /// <param name="sheetName">Sheet Name</param>
        /// <param name="startAdress">Start adress</param>
        /// <param name="endAdress">End adress</param>
        public void CellPaste(String sheetName, String startAdress, String endAdress)
        {
            this.SelectSheet(sheetName);
            var start = startAdress.ToAddress();
            var end = endAdress.ToAddress();

            Object range = null;
            try {
                range = GetRange(start.Row, start.Column, end.Row, end.Column);
                range.Method(METHOD_SPECIAL_PASTE, new Object[] { XlPasteType.xlPasteAll });

                this.Application.SetProperty(PROP_CUTCOPY_MODE, new Object[] { MsoTriState.msoFalse });
            }
            catch (Exception) { throw; }
            finally {
                ReleaseObject(range);
            }
        }

        /// <summary>
        /// Paste to another book
        /// </summary>
        /// <param name="bookName">Book path</param>
        /// <param name="sheetName">Sheet name</param>
        /// <param name="startAdress">Start adress</param>
        /// <param name="endAdress">End adress</param>
        public Boolean AtherBookCellPaste(String bookName, String sheetName, String startAdress, String endAdress)
        {
            return AtherBookCellPaste(bookName, sheetName, startAdress, endAdress, XlPasteType.xlPasteAll);
        }

        /// <summary>
        /// Paste to another book
        /// </summary>
        /// <param name="bookName">Book path</param>
        /// <param name="sheetName">Sheet name</param>
        /// <param name="startAdress">Start adress</param>
        /// <param name="endAdress">End adress</param>
        /// <param name="pasteType"></param>
        public Boolean AtherBookCellPaste(String bookName, String sheetName, String startAdress, String endAdress, XlPasteType pasteType)
        {
            // Open the ather Workbook.
            if (!(this.WorkArea.Method(METHOD_OPEN,
                            SetOpenArguments(new Object[] {
                                System.IO.Path.GetFullPath(bookName)
                            })) is Object atherBook)) {
                return false;
            }
            atherBook.SetProperty(PROP_SAVED, new Object[] { MsoTriState.msoTrue });

            if (!(atherBook.GetProperty(OBJECT_SHEET, new Object[] { sheetName }) is Object sheet)) {
                ReleaseObjects(atherBook);
                return false;
            }
            var start = startAdress.ToAddress();
            var end = endAdress.ToAddress();
            Object range = null;
            try {
                range = sheet.GetProperty(OBJECT_RANGE,
                    new Object[] {
                            sheet.GetProperty(OBJECT_CELL, new Object[] { start.Row, start.Column }),
                            sheet.GetProperty(OBJECT_CELL, new Object[] { end.Row, end.Column })
                    });
                // Clipboard data paste
                range?.Method(METHOD_SPECIAL_PASTE, new Object[] { pasteType });
                this.Application.SetProperty(PROP_CUTCOPY_MODE, new Object[] { MsoTriState.msoFalse });

                // Get this Book path
                atherBook.Method(METHOD_SAVE);
                atherBook.Method(METHOD_CLOSE);
            }
            catch (Exception) { return false; }
            finally {
                ReleaseObjects(range, sheet, atherBook);
            }
            return true;
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

            for (; r < src.Length; r++) {
                for (c = 0; c < columns; c++) {
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
        internal Object GetRange(UInt32 startRow, UInt32 startCol, UInt32 endRow, UInt32 endCol)
        {
            try {
                // Get cell address
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

            for (var i = 0; i < row; i++) {
                for (var j = 0; j < col; j++) {
                    // Is it not within the range?
                    result[i, j] = (valRow > i && valCol > j) ? values[i, j] ?? "" : "";
                }
            }
            return result;
        }
#if false
        /// <summary>
        /// 
        /// </summary>
        /// <param name="target"></param>
        /// <param name="command"></param>
        /// <param name="args"></param>
        /// <returns></returns>
        public static Object Macro(Object target, String command, params Object[] args)
        {
            return target.Method(command, args);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="command"></param>
        /// <param name="args"></param>
        /// <returns></returns>
        public Object BookMacro(String command, params Object[] args)
        {
            return this.Book.Method(command, args);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="command"></param>
        /// <param name="args"></param>
        /// <returns></returns>
        public Object SheetMacro(String sheetName, String command, params Object[] args)
        {
            SelectSheet(sheetName);
            return this.Sheet.Method(command, args);
        }
#endif
        #endregion

        #region --- Object ---
        ///// <summary>
        ///// Get ActiveSheet in Chartobjects
        ///// </summary>
        ///// <returns>ChartObjects</returns>
        //public Object[] GetCurrentSheetCharts()
        //{
        //    var count = this.Sheet.Method("ChartObjects").GetProperty(PROP_COUNT).To<Int32>();
        //    var result = new Object[count];
        //    for (int i = 0; i < result.Length; i++) {
        //        result[i] = this.Sheet.Method("ChartObjects", new Object[] { i + 1 }).GetProperty("Chart");
        //    }
        //    return result;
        //}

        /// <summary>
        /// Set ChartSeries Name
        /// </summary>
        /// <param name="chartIndex">chart object index</param>
        /// <param name="names">series names</param>
        public void SetChartSeriesName(Int32 chartIndex, String[] names)
        {
            Object chart = null;
            try {
                chart = this.Sheet.Method(METHOD_CHART_OBJECTS, new Object[] { chartIndex }).GetProperty(PROP_CHART);

                for (int i = 0; i < names.Length; i++) {
                    chart.Method(METHOD_SERIES_COLLECTION, new Object[] { i + 1 })
                         .SetProperty(PROP_NAME, new Object[] { $"=\"{names[i]}\"" });
                }
            }
            finally {
                ReleaseObject(chart);
            }
        }

        /// <summary>
        /// Active sheet add chart
        /// </summary>
        public void AddChart()
        {
            Double.TryParse(this.Version, out var ver);

            if ((Int32)XlVersion.Excel2007 <= ver) {
                this.Sheet.GetProperty(OBJECT_SHAPES).Method("AddChart");
            }
            else {
                throw new Exception("Functions not supported in this version");
            }
        }

        /// <summary>
        /// Set chart type
        /// </summary>
        /// <param name="chartIndex"></param>
        /// <param name="xlType"></param>
        public void SetChartType(Int32 chartIndex, XlChartType xlType)
        {
            Object chart = null;
            try {
                Double.TryParse(this.Version, out var ver);
                if ((Int32)XlVersion.Excel2007 <= ver) {
                    // 2007
                    chart = this.Sheet.Method(METHOD_CHART_OBJECTS, new Object[] { chartIndex }).GetProperty(PROP_CHART);
                    chart.SetProperty(PROP_CHART_TYPE, new Object[] { xlType });
                }
                else {
                    // 2003 VBA code sample
                    // ActiveSheet.ChartObjects(1).Activate
                    // ActiveChart.ChartArea.Select
                    // ActiveChart.ChartType = xlLine
                    throw new Exception("Functions not supported in this version");
                }
            }
            finally {
                ReleaseObject(chart);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="chartIndex"></param>
        /// <param name="xlTypes"></param>
        public void SetChartTypeSeries(Int32 chartIndex, XlChartType[] xlTypes)
        {
            for (var i = 0; i < xlTypes.Length; i++) {
                SetChartTypeSeries(chartIndex, i + 1, xlTypes[i]);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="chartIndex"></param>
        /// <param name="seriesIndex"></param>
        /// <param name="xlType"></param>
        public void SetChartTypeSeries(Int32 chartIndex, Int32 seriesIndex, XlChartType xlType)
        {
            Object chart = null;
            try {
                Double.TryParse(this.Version, out var ver);
                if ((Int32)XlVersion.Excel2007 <= ver) {
                    // 2007
                    chart = this.Sheet.Method(METHOD_CHART_OBJECTS, new Object[] { chartIndex })
                                      .GetProperty(PROP_CHART).Method(METHOD_SERIES_COLLECTION, new Object[] { seriesIndex });
                    chart.SetProperty(PROP_CHART_TYPE, new Object[] { xlType });
                }
                else {
                    // 2003 VBA code sample
                    // ActiveSheet.ChartObjects(1).Activate
                    // ActiveChart.ChartArea.Select
                    // ActiveChart.ChartType = xlLine
                    throw new Exception("Functions not supported in this version");
                }
            }
            finally {
                ReleaseObject(chart);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="chartIndex"></param>
        /// <returns></returns>
        public Int32[] GetChartPosition(Int32 chartIndex)
        {
            var result = new Int32[4];
            Object chart = null;
            try {
                Double.TryParse(this.Version, out var ver);
                if ((Int32)XlVersion.Excel2007 <= ver) {
                    // 2007
                    chart = this.Sheet.Method(METHOD_CHART_OBJECTS, new Object[] { chartIndex });
                    result[0] = chart.GetProperty(PROP_TOP).To<Int32>();
                    result[1] = chart.GetProperty(PROP_LEFT).To<Int32>();
                    result[2] = chart.GetProperty(PROP_HEIGHT).To<Int32>();
                    result[3] = chart.GetProperty(PROP_WIDTH).To<Int32>();
                }
                else {
                    // 2003 VBA code sample
                    // ActiveSheet.ChartObjects(1).Activate
                    // ActiveChart.ChartArea.Select
                    // ActiveChart.ChartType = xlLine
                    throw new Exception("Functions not supported in this version");
                }
            }
            finally {
                ReleaseObject(chart);
            }
            return result;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="chartIndex"></param>
        /// <param name="location"></param>
        public void SetChartPosition(Int32 chartIndex, Int32[] location)
        {
            Object chart = null;
            try {
                Double.TryParse(this.Version, out var ver);
                if ((Int32)XlVersion.Excel2007 <= ver) {
                    // 2007
                    chart = this.Sheet.Method(METHOD_CHART_OBJECTS, new Object[] { chartIndex });
                    chart.SetProperty(PROP_TOP, new Object[] { location[0] });
                    chart.SetProperty(PROP_LEFT, new Object[] { location[1] });
                    chart.SetProperty(PROP_HEIGHT, new Object[] { location[2] });
                    chart.SetProperty(PROP_WIDTH, new Object[] { location[3] });
                }
                else {
                    // 2003 VBA code sample
                    // ActiveSheet.ChartObjects(1).Activate
                    // ActiveChart.ChartArea.Select
                    // ActiveChart.ChartType = xlLine
                    throw new Exception("Functions not supported in this version");
                }
            }
            finally {
                ReleaseObject(chart);
            }
        }
        #endregion

        /// <summary>
        /// ToString
        /// </summary>
        public override String ToString()
            => String.Format("Book:{0}", System.IO.Path.GetFileNameWithoutExtension(this.Path));
    }
}