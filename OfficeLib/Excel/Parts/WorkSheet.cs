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
            var tableRange = new Object[row][];
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
            get { return this.EntireField.Data[row]; }
            set { this.EntireField.Data[row] = value; }
        }

        /// <summary>Get cell value</summary>
        /// <param name="address">Sheet address</param>
        public Object this[Address address]
        {
            get { return this.EntireField.Data[address.Row - 1][address.Column - 1]; }
            set { this.EntireField.Data[address.Row - 1][address.Column - 1] = value; }
        }

        /// <summary>Get Values in the specified range</summary>
        /// <param name="startAddress">Start cell</param>
        /// <param name="endAddress">End cell</param>
        public Object[][] this[Address startAddress, Address endAddress]
        {
            get { return this.EntireField[startAddress, endAddress]; }
            set { this.EntireField[startAddress, endAddress] = value; }
        }

        /// <summary>Get Values in the specified range</summary>
        /// <param name="startAddrStr">Start cell</param>
        /// <param name="endAddrStr">End cell</param>
        public Object[][] this[String startAddrStr, String endAddrStr]
        {
            get { return this[startAddrStr.ToAddress(), endAddrStr.ToAddress()]; }
            set { this[startAddrStr.ToAddress(), endAddrStr.ToAddress()] = value; }
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
                throw new Exception(String.Format("{0} is Nothing. Can not Read.", this, this.Name));
            }
            // Start "A1" from there Get the maximum amount of data
            this.EntireField = new Field<Object>(
                                        excel.GetCellValue(1, 1,
                                                           (UInt32)this.MaxRow,
                                                           (UInt32)this.MaxColumn,
                                                           XlGetValueFormat.xlValue));
            // Initialize tables
            this.Tables = new Dictionary<String, Field<Object>>();
        }

        /// <summary>
        /// Read the sheet
        /// </summary>
        /// <param name="excel">Excel instance</param>
        /// <param name="format">Get value format type</param>
        public virtual void Read(Excel excel, XlGetValueFormat format)
        {
            if (!excel.SelectSheet(this.Name))
            {
                throw new Exception(String.Format("{0} is Nothing. Can not Read.", this, this.Name));
            }
            // Start "A1" from there Get the maximum amount of data
            this.EntireField = new Field<Object>(
                                        excel.GetCellValue(1, 1,
                                                           (UInt32)this.MaxRow,
                                                           (UInt32)this.MaxColumn,
                                                           format));
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
                                   XlGetValueFormat.xlValue), startAddress);

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
        /// <returns>Table data</returns>
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
                throw new Exception(String.Format("{0} is Nothing. Can not write.", this, this.Name));
            }
            if(this.EntireField == null) { return; }

            Write(excel, this.EntireField.Data, "A1".ToAddress());
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
        /// Table data setting
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="excel">Excel instance</param>
        /// <param name="value">Setting value</param>
        /// <param name="startAddress">Satart Position</param>
        protected virtual void SetTable<T>(Excel excel, T[][] value, Address startAddress)
        {
            excel.SetCellValue(Excel.ConvertSetValue(value),
                               startAddress.ReferenceString,
                               XlGetValueFormat.xlValue2);
        }

        /// <summary>
        /// Set value to sheet
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="excel"></param>
        /// <param name="value"></param>
        /// <param name="address"></param>
        /// <param name="format"></param>
        protected virtual void SetValue<T>(Excel excel, T value, Address address, XlGetValueFormat format)
        {
            var values = new Object[1, 1] { { value } };
            excel.SetCellValue(values, address.ReferenceString, format);
        }

        /// <summary>
        /// Table data setting
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="excel">Excel instance</param>
        /// <param name="value">Setting value</param>
        /// <param name="startAddress">Satart Position</param>
        /// <param name="endAddress">End Position</param>
        protected virtual void SetTable<T>(Excel excel, T[][] value, 
                                            Address startAddress, Address endAddress)
        {
            excel.SetCellValue(Excel.ConvertSetValue(value),
                               startAddress.Row, startAddress.Column, 
                               endAddress.Row, startAddress.Column,
                               XlGetValueFormat.xlValue2);
        }
        #endregion

        #region --- Added table definition ---
        /// <summary>Add table</summary>
        /// <param name="key">Table name</param>
        /// <param name="table">Table data</param>
        protected virtual void AddTable(String key, Field<Object> table)
        {
            this.Tables.Add(key, table);
        }

        /// <summary>Add table</summary>
        /// <param name="key">Table name</param>
        /// <param name="startAddress">Start Position</param>
        /// <param name="endAddress">End Position</param>
        public virtual void AddTable(String key, Address startAddress, Address endAddress)
        {
            this.AddTable(key, startAddress.ReferenceString, endAddress.ReferenceString);
        }

        /// <summary>Add table</summary>
        /// <param name="key">Table name</param>
        /// <param name="startAddress">Start Position(String)</param>
        /// <param name="endAddress">End Position(String)</param>
        public virtual void AddTable(String key, String startAddress, String endAddress)
        {
            this.Tables.Add(key, GetTable(startAddress, endAddress));
        }
        #endregion
    }
}
