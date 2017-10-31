using System;
using System.Linq;
using System.Collections.Generic;

namespace OfficeLib.XLS
{
    /// <summary>
    /// Table using jagged array
    /// </summary>
    public class Field<T>
    {
        /// <summary>Table Column length</summary>
        public Int32 Column { get; private set; }

        /// <summary>Table Row length</summary>
        public Int32 Row { get { return this.Data.Length; } }

        /// <summary>Table data</summary>
        /// <remarks>
        /// Like a rectangular array, Each element has the same length
        /// </remarks>
        public T[][] Data { get; private set; }

        /// <summary>Starting position of Location(Left Top)</summary>
        public Address StartAddress { get; set; }

        /// <summary>Ending position of Location(Right Bottom)</summary>
        public Address EndAddress { get; private set; }

        #region --- Constructor ---
        /// <summary>
        /// Constructor
        /// </summary>
        public Field()
        {
            SetProperties(new T[][] { new T[] { default(T) } }, "A1".ToAddress());
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="value">Data handled as a table</param>
        public Field(T[][] value)
        {
            SetProperties(value?.ToRectLikeJagArray() ??
                                new T[][] { new T[] { default(T) } },
                          "A1".ToAddress());
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="value">Data handled as a table</param>
        /// <param name="startAddress">Start address</param>
        public Field(T[][] value, Address startAddress)
        {
            SetProperties(value?.ToRectLikeJagArray() ??
                                new T[][] { new T[] { default(T) } },
                          startAddress);
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="value">Data handled as a table</param>
        public Field(T[,] value)
        {
            SetProperties(RectToJag(value) ?? new T[][] { new T[] { default(T) } },
                          "A1".ToAddress());
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="value">Data handled as a table</param>
        /// <param name="startAddress">Start address</param>
        public Field(T[,] value, Address startAddress)
        {
            SetProperties(RectToJag(value) ?? new T[][] { new T[] { default(T) } },
                          startAddress);
        }

        /// <summary>
        /// Rect array to jag array
        /// </summary>
        private T[][] RectToJag(T[,] value)
        {
            if(value == null) { return null; }

            T[][] result = new T[value.GetLength(0)][];
            Int32 col = value.GetLength(1);
            for (var r = 0; r < result.Length; r++)
            {
                result[r] = new T[col];
                for (var c = 0; c < col; c++)
                {
                    result[r][c] = value[r, c];
                }
            }
            return result;
        }

        /// <summary>
        /// Set values of Constructor
        /// </summary>
        /// <param name="value"></param>
        /// <param name="address"></param>
        private void SetProperties(T[][] value, Address address)
        {
            this.Data = value;
            this.Column = this.Data.Length <= 0 ? 0 : this.Data[0].Length;

            this.StartAddress = address;
            this.EndAddress = Address.Shift(this.StartAddress, this.Column, this.Row);
        }
        #endregion

        #region --- Indexer ---
        /// <summary>
        /// Get by specifying the row of the table
        /// </summary>
        public T[] this[Int32 row] { get { return this.Data[row]; } }

        /// <summary>
        /// Get the cell Value
        /// </summary>
        public Object this[Address range] { get { return GetCellValue(range); } }

        /// <summary>
        /// Get the cell Value
        /// </summary>
        public Object this[String range] { get { return GetCellValue(range.ToAddress()); } }
        #endregion

        #region --- operator ---
            /// <summary>
            /// Joining tables(Horizontally)
            /// </summary>
            /// <remarks>Based on the less row</remarks>
        public static Field<T> operator &(Field<T> leftTable, Field<T> rightTable)
         => new Field<T>(leftTable.Data.InnerJoin(rightTable.Data))
            { StartAddress = leftTable.StartAddress };

        /// <summary>
        /// Joining tables(Horizontally)
        /// </summary>
        /// <remarks>Based on the one with more rows</remarks>
        public static Field<T> operator |(Field<T> leftTable, Field<T> rightTable)
         => new Field<T>(leftTable.Data.OuterJoin(rightTable.Data))
            { StartAddress = leftTable.StartAddress };

        /// <summary>
        /// Joining tables(vertically)
        /// </summary>
        /// <remarks>Based on the one with more Columns</remarks>
        public static Field<T> operator +(Field<T> topTable, Field<T> bottomTable)
         => new Field<T>(topTable.Data.Union(bottomTable.Data))
            { StartAddress = topTable.StartAddress };

        /// <summary>
        /// Delete columns of the table from the right
        /// </summary>
        public static Field<T> operator <<(Field<T> baseTable, Int32 minus)
         => new Field<T>(baseTable.Data.ColumnRemoveEnd(minus))
            { StartAddress = baseTable.StartAddress };

        /// <summary>
        /// Delete columns of the table from the left
        /// </summary>
        public static Field<T> operator >>(Field<T> baseTable, Int32 minus)
         => new Field<T>(baseTable.Data.ColumnRemoveStart(minus))
            { StartAddress = Address.Shift(baseTable.StartAddress, minus, 0) };

        /// <summary>
        /// Delete columns of the table from the top
        /// </summary>
        public static Field<T> operator -(Field<T> baseTable, Int32 minus)
        {
            Int32 len = baseTable.Data.Length - minus;
            len = len < 0 ? 0 : len;

            T[][] newTable = new T[len][];
            for (var i = minus; i < len + minus; i++)
            {
                newTable[i - minus] = new T[baseTable.Column];
                baseTable.Data[i].CopyTo(newTable[i - minus], 0);
            }
            return new Field<T>(newTable)
            { StartAddress = Address.Shift(baseTable.StartAddress, 0, minus) };
        }

        /// <summary>
        /// Delete columns of the table from the bottom
        /// </summary>
        public static Field<T> operator ^(Field<T> baseTable, Int32 minus)
        {
            Int32 len = baseTable.Data.Length - minus;
            len = len < 0 ? 0 : len;

            T[][] newTable = new T[len][];
            for (var i = 0; i < newTable.Length; i++)
            {
                newTable[i] = new T[baseTable.Column];
                baseTable.Data[i].CopyTo(newTable[i], 0);
            }
            return new Field<T>(newTable) { StartAddress = baseTable.StartAddress };
        }

        /// <summary>
        /// Vertical segmentation of Tables
        /// </summary>
        public static Field<T>[] operator /(Field<T> baseTable, Int32 divid)
        {
            Field<T>[] result = new Field<T>[divid];
            try
            {
                T[][][] val = baseTable.Data.VDividingArray(divid);
                result = new Field<T>[val.Length];
                for (var i = 0; i < result.Length; i++)
                {
                    result[i] = new Field<T>(val[i]);
                }
            }
            catch (Exception) { throw; }
            return result;
        }

        /// <summary>
        /// Horizontal segmentation of Tables
        /// </summary>
        public static Field<T>[] operator %(Field<T> baseTable, Int32 divid)
        {
            Field<T>[] result = new Field<T>[divid];
            try
            {
                T[][][] val = baseTable.Data.HDividingArray(divid);
                result = new Field<T>[val.Length];
                for (var i = 0; i < result.Length; i++)
                {
                    result[i] = new Field<T>(val[i]);
                }
            }
            catch (Exception) { throw; }
            return result;
        }
        
        /// <summary>
        /// Cast
        /// </summary>
        public static explicit operator Field<T>(T[][] values)
            => new Field<T>(values);
        #endregion

        /// <summary>
        /// Vertical segmentation of Table, and get the table
        /// </summary>
        public Field<T> TakeVerticalField(Int32 start, Int32 length)
        {
            Field<T> result = this >> start;
            Int32 takelength = result.Data.Max(row => row.Length) - length;

            return result << takelength;
        }

        /// <summary>
        /// Extract values vertically from the table
        /// </summary>
        public IEnumerable<T> TakeVertical(Int32 column)
            => this.Data.TakeVertical(column);

        /// <summary>
        /// Convert table contents to specified type
        /// </summary>
        /// <typeparam name="TOutput"></typeparam>
        public Field<TOutput> Convert<TOutput>()
            => new Field<TOutput>(this.Data.ConvertAll<T, TOutput>());

        /// <summary>
        /// Get the cell Value
        /// </summary>
        private Object GetCellValue(Address range)
        {
            UInt32 col = range.Column - this.StartAddress.Column;
            UInt32 row = range.Row - this.StartAddress.Row;

            // out of range
            if (col < 0 || this.Column <= col) { return null; }
            if (row < 0 || this.Row <= row) { return null; }

            return this.Data[row][col];
        }
    }
}