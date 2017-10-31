using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace OfficeLib.XLS
{
    /// <summary>Excel Workbook class</summary>
    public class WorkBook
    {
        #region --- Property ---
        /// <summary>Book name</summary>
        public String Name { get; private set; }
        /// <summary>Book paht</summary>
        public String Path { get; set; }
        /// <summary>Book has Sheets List</summary>
        public Dictionary<String, WorkSheet> WorkSheets { get; private set; }

        /// <summary>(private variable) sheet names</summary>
        private HashSet<String> _sheetNames;

        /// <summary>Book has sheet names</summary>
        public String[] SheetNames
        {
            get { return GetKeys().ToArray(); }
        }
        #endregion

        #region --- Constructor ---
        /// <summary>Constructor</summary>
        public WorkBook()
        {
            this.WorkSheets = new Dictionary<String, WorkSheet>();
            this.Name = String.Empty;
            this._sheetNames = new HashSet<String>();
        }

        /// <summary>Constructor</summary>
        /// <param name="filePath">Fike path</param>
        public WorkBook(String filePath)
        {
            this.Path = filePath;
            this.Name = System.IO.Path.GetFileNameWithoutExtension(filePath);
            this.WorkSheets = new Dictionary<String, WorkSheet>();
            this._sheetNames = new HashSet<String>();
        }
        #endregion

        #region --- Indexer ---
        /// <summary>Retrieve sheet object from sheet name</summary>
        public WorkSheet this[String sheetName]
        {
            get { return this.WorkSheets[sheetName]; }
            set { this.WorkSheets[sheetName] = value; }
        }
        #endregion

        #region --- Read ---
        /// <summary>
        /// Read all the sheets in the book
        /// </summary>
        public virtual void Read()
        {
            if (!File.Exists(this.Path))
            { throw new FileNotFoundException("File not found.", this.Path); }
            // Setting the Wrokebook name
            this.Name = System.IO.Path.GetFileNameWithoutExtension(this.Path);

            // Create the Excel instance
            using (var excel = new Excel())
            {
                if (!excel.Open(System.IO.Path.GetFullPath(this.Path)))
                {
                    throw new ArgumentException();
                }
                foreach (var name in excel.SheetNames)
                {
                    if (!this.WorkSheets.ContainsKey(name))
                    {   // Add unconfigured sheet
                        this.WorkSheets.Add(name, new WorkSheet(name));
                    }
                }
                // Read data on each sheet
                foreach (var sheet in this.WorkSheets)
                {
                    if (CanRead(excel, sheet.Value))
                    {
                        this[sheet.Key].Read(excel);
                    }
                }
            }
        }

        /// <summary>
        /// Read sheet specification
        /// </summary>
        /// <param name="sheetName">Sheet name</param>
        public virtual void Read(String sheetName)
        {
            if (!File.Exists(this.Path))
            { throw new FileNotFoundException("File not found.", this.Path); }
            
            // Setting the Wrokebook name
            this.Name = System.IO.Path.GetFileNameWithoutExtension(this.Path);

            using (var excel = new Excel())
            {
                if(!excel.Open(System.IO.Path.GetFullPath(this.Path)))
                {
                    throw new ArgumentException();
                }
                if (CanRead(excel, this[sheetName]))
                {
                    this[sheetName].Read(excel);
                }
            }
        }

        /// <summary>
        /// Read preset sheets
        /// </summary>
        public virtual void ReadPreset()
        {
            if (!File.Exists(this.Path))
            { throw new FileNotFoundException("File not found.", this.Path); }
            // Setting the Wrokebook name
            this.Name = System.IO.Path.GetFileNameWithoutExtension(this.Path);

            // Create the Excel instance
            using (var excel = new Excel())
            {
                if (!excel.Open(System.IO.Path.GetFullPath(this.Path)))
                {
                    throw new ArgumentException();
                }
                // Read data on each sheet
                foreach (var sheet in this.WorkSheets)
                {
                    if (CanRead(excel, sheet.Value))
                    {
                        this[sheet.Key].Read(excel);
                    }
                }
            }
        }

        /// <summary>
        /// Confirm whether it is readable
        /// </summary>
        /// <param name="excel">Excel instance</param>
        /// <param name="sheet">Sheet instance</param>
        /// <returns></returns>
        protected virtual Boolean CanRead(Excel excel, WorkSheet sheet)
        {
            // Does the workbook have a seat?
            Boolean hasSheet = excel.SheetNames.Contains(sheet.Name);
            Boolean hasAttribute = false;
            var attr = sheet.GetType().GetCustomAttribute(
                                        typeof(ExcelSheetAttribute))
                                        as ExcelSheetAttribute;
            if (attr != null)
            {   // Does the sheet have read permission?
                hasAttribute = ((UInt32)attr.Permission)
                                .ContainsBitFlag((Int32)EnumSheetPermission.Read);
            }
            // Both true?
            return hasSheet && hasAttribute;
        }
        #endregion

        #region --- Write ---
        /// <summary>
        /// Batch Writing of Book
        /// </summary>
        public virtual void WriteBook()
        {
            if (!File.Exists(this.Path))
            { throw new FileNotFoundException("File not found.", this.Path); }
            // If there is no sheet, the process is terminated
            if (this.WorkSheets.Count < 1) { return; }

            using (var excel = new Excel())
            {
                excel.Open(System.IO.Path.GetFullPath(this.Path));
                // Write data on each sheet
                foreach (KeyValuePair<String, WorkSheet> sheet in this.WorkSheets)
                {
                    if (this.CanWrite(excel, sheet.Value))
                    {
                        this[sheet.Key].Write(excel);
                    }
                }
                excel.Save();
            }
        }

        /// <summary>
        /// Write sheet specification
        /// </summary>
        public virtual void WriteSheet(String sheetName)
        {
            if (!File.Exists(this.Path))
            { throw new FileNotFoundException("File not found.", this.Path); }
            // If there is no sheet, the process is terminated
            if (this.WorkSheets.Count < 1) { return; }

            using (var excel = new Excel())
            {
                excel.Open(System.IO.Path.GetFullPath(this.Path));
                if (this.CanWrite(excel, this[sheetName]))
                {
                    this[sheetName].Write(excel);
                    excel.Save();
                }
            }
        }

        /// <summary>
        /// Confirm whether it is writable
        /// </summary>
        /// <param name="excel">Excel instance</param>
        /// <param name="sheet">Sheet instance</param>
        /// <returns></returns>
        protected virtual Boolean CanWrite(Excel excel, WorkSheet sheet)
        {
            // Does the workbook have a seat?
            Boolean hasSheet = excel.SheetNames.Contains(sheet.Name);
            Boolean hasAttribute = false;
            var attr = sheet.GetType().GetCustomAttribute(
                                        typeof(ExcelSheetAttribute))
                                        as ExcelSheetAttribute;
            if (attr != null)
            {   // Does the sheet have write permission?
                hasAttribute = ((UInt32)attr.Permission)
                                .ContainsBitFlag((Int32)EnumSheetPermission.Write);
            }
            // Both true?
            return hasSheet && hasAttribute;
        }
        #endregion

        #region --- Add Sheet ---
        /// <summary>
        /// Add sheet
        /// </summary>
        /// <param name="sheet">Sheet instance</param>
        public void AddSheet<T>(T sheet) where T : WorkSheet
        {
            this.WorkSheets.Add(sheet.Name, sheet);
        }

        /// <summary>
        /// Add multiple sheets
        /// </summary>
        /// <param name="sheets">Sheet array</param>
        public void AddRangeSheet<T>(T[] sheets) where T : WorkSheet
        {
            foreach (T sheet in sheets)
            {
                this.AddSheet(sheet);
            }
        }
        #endregion

        /// <summary>
        /// Get this worksheet names
        /// </summary>
        private IEnumerable<String> GetKeys()
        {
            foreach (var key in this.WorkSheets.Keys)
            {
                yield return key;
            }
        }
    }
}
