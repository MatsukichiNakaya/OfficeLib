using System;
using System.Collections.Generic;
using System.Reflection;

namespace OfficeLib.XLS
{
    /// <summary>Worksheet class</summary>
    [ExcelSheet(EnumSheetPermission.ReadWrite)]
    public class WorkSheet
    {
        #region --- Property ---
        /// <summary>Sheet name</summary>
        public String Name { get; private set; }
        /// <summary>Tables defined in the sheet</summary>
        public Dictionary<String, Field<Object>> Tables { get; set; }
        /// <summary>The entire loaded field</summary>
        public Field<Object> EntireField { get; set; }

        /// <summary>Max row</summary>
        public Int32 MaxRow { get { return this.EntireField.Row; } }
        /// <summary>Max column</summary>
        public Int32 MaxColumn { get { return this.EntireField.Column; } }

        /// <summary>Name list from tables defined</summary>
        public String[] TableNames { get { return System.Linq.Enumerable.ToArray(this.Tables.Keys); } }
        #endregion

        #region --- Constructor ---
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="name"></param>
        /// <param name="endAddress"></param>
        public WorkSheet(String name, Address endAddress)
            : this(name, endAddress.Column, endAddress.Row) { }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="name"></param>
        /// <param name="endAddress"></param>
        public WorkSheet(String name, String endAddress) 
            : this (name, endAddress.ToAddress()){ }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="name"></param>
        /// <param name="maxColumns"></param>
        /// <param name="maxRows"></param>
        public WorkSheet(String name, String maxColumns, UInt32 maxRows)
            : this(name, maxColumns.ToColumnNumber(), maxRows) { } 

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="name">Sheet name</param>
        /// <param name="maxColumns"></param>
        /// <param name="maxRows"></param>
        public WorkSheet(String name, UInt32 maxColumns = 100, UInt32 maxRows = 100)
        {
            this.Name = name;

            // Initialize tables
            this.Tables = new Dictionary<String, Field<Object>>();

            // Get the sheet information
            var att = this.GetType()
                      .GetCustomAttribute(typeof(ExcelSheetAttribute)) as ExcelSheetAttribute;
            // If it can not get is, use the default value
            UInt32 row = maxRows;
            UInt32 col = maxColumns;

            // Initialize field
            Object[][] tableRange = null;
            tableRange = new Object[row][];
            for (var r = 0; r < tableRange.Length; r++)
            {
                tableRange[r] = new Object[col];
            }
            this.EntireField = new Field<Object>(tableRange);
        }
        #endregion

        #region --- Indexer ---
        /// <summary>Get row</summary>
        /// <param name="row">Row number of the sheet</param>
        public Object[] this[Int32 row]
        {
            get { return EntireField.Data[row]; }
            set { EntireField.Data[row] = value; }
        }

        /// <summary>Get cell value</summary>
        /// <param name="address">Sheet address</param>
        public Object this[Address address]
        {
            get { return EntireField.Data[address.Row - 1][address.Column - 1]; }
            set { EntireField.Data[address.Row - 1][address.Column - 1] = value; }
        }

        /// <summary>Get Values in the specified range</summary>
        /// <param name="startAddress">Start cell</param>
        /// <param name="endAddress">End cell</param>
        public Object[][] this[Address startAddress, Address endAddress]
        {
            get
            {
                Int32 width = Math.Abs((Int32)(startAddress.Column - endAddress.Column));
                return EntireField.Data.RangeTake
                        ((Int32)startAddress.Row - 1, (Int32)endAddress.Row,
                        (Int32)startAddress.Column - 1, width + 1).ToJagArray();
            }
        }

        /// <summary>Get Values in the specified range</summary>
        /// <param name="startAddrStr">Start cell</param>
        /// <param name="endAddrStr">End cell</param>
        public Object[][] this[String startAddrStr, String endAddrStr]
        {
            get { return this[startAddrStr.ToAddress(), endAddrStr.ToAddress()]; }
        }

        /// <summary>Get table</summary>
        /// <param name="tableName">Table name</param>
        public Field<Object> this[String tableName]
        {
            get { return this.Tables[tableName]; }
            set { this.Tables[tableName] = value; }
        }
        #endregion

        #region --- Read ---
        /// <summary>
        /// Read the sheet
        /// </summary>
        /// <param name="excel">Excel instance</param>
        /// <remarks>
        /// Field size Default(1000 x 1000)
        /// </remarks>
        public virtual void Read(Excel excel)
        {
            if (!excel.SelectSheet(this.Name))
            {
                throw new Exception(String.Format("{0} is Nothing. Can not Read.", this, Name));
            }
            // Start "A1" from there Get the maximum amount of data
            this.EntireField = new Field<Object>(
                                        excel.GetCellValue(1, 1,
                                                           (UInt32)this.MaxRow,
                                                           (UInt32)this.MaxColumn,
                                                           Excel.XlGetValueFormat.xlValue));
            // Initialize tables
            this.Tables = new Dictionary<String, Field<Object>>();
        }

        /// <summary>
        /// Get table from its own field
        /// </summary>
        /// <param name="excel">Excel instance</param>
        /// <param name="startAddress">Satart Position</param>
        /// <param name="endAddress">End Position</param>
        /// <returns>Field Data</returns>
        /// <remarks>
        /// Get value from excel instance
        /// Since it is acquired directly from Excel, 
        /// it can be obtained freely if it is within the allowable range of Excel
        /// </remarks>
        public virtual Field<Object> GetTable(Excel excel, 
                                              Address startAddress, Address endAddress)
            => new Field<Object>(
                excel.GetCellValue(startAddress.ReferenceString, 
                                   endAddress.ReferenceString,
                                   Excel.XlGetValueFormat.xlValue), startAddress);

        /// <summary>
        /// Get table from its own field
        /// </summary>
        /// <param name="startAddress">Satart Position</param>
        /// <param name="endAddress">End Position</param>
        /// <returns>Field Data</returns>
        /// <remarks>
        /// * Points to note when specifying the range.
        /// It will never be acquired beyond the scope of the Field variable that is acquired in advance
        /// </remarks>
        public virtual Field<Object> GetTable(Address startAddress, Address endAddress)
            => new Field<Object>(this[startAddress, endAddress], startAddress);

        /// <summary>
        /// Get table from its own field
        /// </summary>
        /// <param name="startAddrStr">Satart Position</param>
        /// <param name="endAddrStr">End Position</param>
        /// <returns>表データ</returns>
        /// <remarks>
        /// * Points to note when specifying the range.
        /// It will never be acquired beyond the scope of the Field variable that is acquired in advance
        /// </remarks>
        public virtual Field<Object> GetTable(String startAddrStr, String endAddrStr)
            => new Field<Object>(this[startAddrStr.ToAddress(), endAddrStr.ToAddress()],
                                 startAddrStr.ToAddress());
        #endregion

        #region --- Update ---
        /// <summary>
        /// Update the Table from the Field
        /// </summary>
        public virtual void UpdateTableFromField()
        {
            if(this.Tables == null) { return; }

            foreach (var key in this.Tables.Keys)
            {   // update table
                this.Tables[key] = GetTable(this.Tables[key].StartAddress.ReferenceString,
                                            this.Tables[key].EndAddress.ReferenceString);
            }
        }

        /// <summary>
        /// Update the Field from the Table
        /// </summary>
        public virtual void UpdateFieldFromTable()
        {
            if(this.Tables == null) { return; }

            UInt32 row = 0;
            UInt32 col = 0;
            foreach (var table in this.Tables)
            {   // Address to index
                row = table.Value.StartAddress.Row;
                col = table.Value.StartAddress.Column;
                // Set Field from table
                for (var r = row; r < table.Value.Row; r++)
                {
                    for (var c = col; c < table.Value.Column; c++)
                    {
                        this.EntireField[(Int32)r][c] = table.Value[(Int32)(r - row)][c - col];
                    }
                }
            }
        }
        #endregion

        #region --- Write ---
        /// <summary>
        /// Write to Excel
        /// </summary>
        /// <param name="excel">Excel instance</param>
        public virtual void Write(Excel excel)
        {
            if (!excel.SelectSheet(this.Name))
            {
                throw new Exception(String.Format("{0} is Nothing. Can not write.", this, Name));
            }
            if(this.EntireField == null) { return; }

            Write(excel, EntireField.Data, "A1".ToAddress());
        }

        /// <summary>
        /// Write to Excel
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="excel">Excel instance</param>
        /// <param name="value">Setting value</param>
        /// <param name="startAddress">Satart Position</param>
        public virtual void Write<T>(Excel excel, T[][] value, Address startAddress)
        {
            SetTable(excel, value, startAddress);
        }

        /// <summary>
        /// 表データ設定処理
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="excel">Excel instance</param>
        /// <param name="value">設定値</param>
        /// <param name="startAddress">開始位置</param>
        protected virtual void SetTable<T>(Excel excel, T[][] value, Address startAddress)
        {
            excel.SetCellValue(Excel.ConvertSetValue(value), startAddress.ReferenceString);
        }

        /// <summary>
        /// 表データ設定処理
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="excel">Excel instance</param>
        /// <param name="value">設定値</param>
        /// <param name="startAddress">開始位置</param>
        /// <param name="endAddress">終了位置</param>
        protected virtual void SetTable<T>(Excel excel, T[][] value, 
                                            Address startAddress, Address endAddress)
        {
            excel.SetCellValue(Excel.ConvertSetValue(value),
                               startAddress.Row, startAddress.Column, 
                               endAddress.Row, startAddress.Column);
        }
        #endregion

        #region シートにテーブルの定義追加
        /// <summary>テーブルの追加</summary>
        protected virtual void AddTable(String key, Field<Object> table)
        {
            this.Tables.Add(key, table);
        }

        /// <summary>テーブルの追加</summary>
        public virtual void AddTable(String key, Address startAddress, Address endAddress)
        {
            this.AddTable(key, startAddress.ReferenceString, endAddress.ReferenceString);
        }

        /// <summary>テーブルの追加</summary>
        public virtual void AddTable(String key, String startAddress, String endAddress)
        {
            this.Tables.Add(key, GetTable(startAddress, endAddress));
        }
        #endregion
    }
}
