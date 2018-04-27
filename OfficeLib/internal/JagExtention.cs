using System;
using System.Collections.Generic;

namespace OfficeLib
{
    /// <summary>Extended function class of Jag array</summary>
    internal static class JagExtention
    {
        #region Combine
        /// <summary>Outer Join(Include all)</summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="srcTable">Table on the right</param>
        /// <param name="joinTable">Table on the left</param>
        /// <remarks>
        /// Combine with reference to the more array elements.
        /// Elements without short arrays are created with default values.
        /// </remarks>
        public static T[][] OuterJoin<T>(this T[][] srcTable, T[][] joinTable)
        {
            T[][] result = null;
            if (srcTable == null) { throw new NullReferenceException(); }
            if (joinTable == null) { return srcTable; }
            try
            {   // Get longer length by comparing the length of the array.
                Int32 rows = srcTable.Length > joinTable.Length ?
                                srcTable.Length : joinTable.Length;
                result = Join(srcTable, joinTable, rows);
            }
            catch (Exception) { throw; }
            return result;
        }

        /// <summary>Outer Join(Common row only)</summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="srcTable">Table on the right</param>
        /// <param name="joinTable">Table on the left</param>
        /// <remarks>
        /// Combine with the smaller number of elements of the array as the reference.
        /// The longer line is truncated.
        /// </remarks>
        public static T[][] InnerJoin<T>(this T[][] srcTable, T[][] joinTable)
        {
            T[][] result = null;
            if (srcTable == null) { throw new NullReferenceException(); }
            if (joinTable == null) { return srcTable; }
            try
            {   // Get shorter length of array vertical length.
                Int32 rows = srcTable.Length < joinTable.Length ?
                                srcTable.Length : joinTable.Length;
                result = Join(srcTable, joinTable, rows);
            }
            catch (Exception) { throw; }
            return result;
        }

        /// <summary>Outer join with the specified number of rows</summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="srcTable">Table on the right</param>
        /// <param name="joinTable">Table on the left</param>
        /// <param name="rows">specific number of rows</param>
        /// <remarks>
        /// Combine  two tables with the specified number of rows.
        /// If it exceeds specified line, it creates with default value.
        /// Trimmed rows are truncated.
        /// </remarks>
        private static T[][] Join<T>(T[][] srcTable, T[][] joinTable, Int32 rows)
        {
            T[][] result = null;
            Int32 colSrc = srcTable.ColumnsMax();
            Int32 colJoin = joinTable.ColumnsMax();
            result = new T[rows][];
            for (var row = 0; row < result.Length; row++)
            {
                result[row] = new T[colSrc + colJoin];

                if (row < srcTable.Length)
                {
                    srcTable[row].CopyTo(result[row], 0);
                }
                if (row < joinTable.Length) { joinTable[row].CopyTo(result[row], colSrc); }
            }
            return result;
        }

        /// <summary>Union</summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="srcTable">Table on the top</param>
        /// <param name="unionTable">Table on the bottom</param>
        /// <remarks>
        /// Combine the table to the underside
        /// Each column is Combined as it is.
        /// </remarks>
        public static T[][] Union<T>(this T[][] srcTable, params T[][][] unionTable)
        {
            var result = new List<T[]>();
            try
            {
                var list = new List<T[][]> { srcTable };
                list.AddRange(unionTable); 
                foreach (T[][] rows in list)
                {
                    foreach (T[] row in rows) { result.Add(row); }
                }
            }
            catch (Exception) { throw; }
            return result.ToArray();
        }
        #endregion

        #region Split
        /// <summary>Split array vertically</summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="src">Table</param>
        /// <param name="dividCount">Division number</param>
        public static T[][][] VDividingArray<T>(this T[][] src, Int32 dividCount)
        {
            var result = new List<T[][]>();
            try
            {
                Int32 colLength = src.ColumnsMax();
                Int32 dividCol = colLength / dividCount;
                Int32 pos = 0;
                // End if not divisible
                if (colLength % dividCount > 0) { return new T[][][] { src }; }
                // End because it is an infinite loop
                if (dividCol <= 0) { return new T[][][] { src }; }

                while (pos < colLength)
                {
                    var tables = new List<T[]>();
                    for (var row = 0; row < src.Length; row++)
                    {
                        try { tables.Add(src[row].SkipTake(pos, dividCol)); }
                        catch (Exception) { break; }
                    }
                    result.Add(tables.ToArray());
                    pos += dividCol;
                }
            }
            catch (Exception) { throw; }
            return result.ToArray();
        }

        /// <summary>Split array horizontally</summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="src">Table</param>
        /// <param name="divid">Division number</param>
        public static T[][][] HDividingArray<T>(this T[][] src, Int32 divid)
        {
            var result = new List<T[][]>();

            Int32 dividRow = src.Length / divid;
            // End if not divisible
            if (src.Length % divid > 0) { return new T[][][] { src }; }
            // End because it is an infinite loop
            if (dividRow <= 0) { return new T[][][] { src }; }

            var rowQueue = new Queue<T[]>(src);
            while (rowQueue.Count > 0)
            {
                var row = new List<T[]>();
                for (var i = 0; i < dividRow; i++) 
                {
                    try { row.Add(rowQueue.Dequeue()); }
                    catch (Exception) { break; }
                }
                result.Add(row.ToArray());
            }
            return result.ToArray();
        }
        #endregion

        #region Delete
        /// <summary>Remove the specified number of columns from the first column</summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="src">Table</param>
        /// <param name="trimCount">Number of columns to delete</param>
        public static T[][] ColumnRemoveStart<T>(this T[][] src, Int32 trimCount)
        {
            var result = new T[src.Length][];
            try
            {
                Int32 colLength = src.ColumnsMax();
                // If all Trim is done, null is returned because there are no elements.
                if (colLength - trimCount <= 0) { return null; }

                for (var row = 0; row < result.Length; row++)
                {
                    result[row] = new T[colLength - trimCount];
                    for (var col = 0; col < result[row].Length; col++)
                    {
                        try { result[row][col] = src[row][col + trimCount]; }
                        // Resize and end when there are no more elements.
                        catch (Exception) { Array.Resize(ref result[row], col); break; }
                    }
                }
            }
            catch (Exception) { throw; }
            return result;
        }

        /// <summary>Remove the specified number of columns from the last column</summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="src">Table</param>
        /// <param name="trimCount">Number of columns to delete</param>
        public static T[][] ColumnRemoveEnd<T>(this T[][] src, Int32 trimCount)
        {
            var result = new T[src.Length][];
            try
            {
                Int32 colLength = src.ColumnsMax();
                // If all Trim is done, null is returned because there are no elements.
                if (colLength - trimCount <= 0) { return null; }

                for (var row = 0; row < result.Length; row++)
                {
                    result[row] = new T[colLength - trimCount];
                    for (var col = 0; col < result[row].Length; col++)
                    {
                        try { result[row][col] = src[row][col]; }
                        // Resize and end when there are no more elements.
                        catch (Exception) { Array.Resize(ref result[row], col); break; }
                    }
                }
            }
            catch (Exception) { throw; }
            return result;
        }
        #endregion

        #region Convert
        /// <summary>
        /// Convert to rectangular array
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="src">Jag Array</param>
        public static T[,] ToRectArray<T>(this T[][] src)
        {
            T[,] result = new T[src.Length, src.ColumnsMax()];
            for (var row = 0; row < result.GetLength(0); row++)
            {
                for (var col = 0; col < result.GetLength(1); col++)
                {
                    // If there is an element, assign the value as it is
                    try { result[row, col] = src[row][col]; }
                    // If there is no value, set a default value.
                    catch (Exception) { result[row, col] = default(T); }
                }
            }
            return result;
        }

        /// <summary>
        /// Align horizontal width of jag array like rectangular array
        /// </summary>
        /// <param name="src">Jag Array</param>
        public static T[][] ToRectLikeJagArray<T>(this T[][] src)
        {
            var result = new T[src.Length][];
            Int32 colMax = src.ColumnsMax();
            for (var i = 0; i < result.Length; i++)
            {
                result[i] = new T[colMax];
                for (var j = 0; j < result[i].Length; j++)
                {
                    result[i][j] = src[i].Length > j ?
                        src[i][j] : default(T);
                }
            }
            return result;
        }

        /// <summary>
        /// Convert a one-dimensional array to a two-dimensional array
        /// </summary>
        /// <param name="src">Array</param>
        /// <param name="divCount">Element count(Row count)</param>
        /// <remarks>
        /// Specify a divisible number of one-dimensional elements.
        /// Truncated decimal places are truncated.
        /// </remarks>
        public static T[][] ToJagArray<T>(this IEnumerable<T> src, Int32 divCount)
        {
            var result = new T[divCount][];
            Int32 columnLength = System.Linq.Enumerable.Count(src) / divCount;

            for (var count = 0; count < result.Length; count++)
            {
                result[count] = new T[columnLength];
                src.SkipTake(count * columnLength, columnLength).CopyTo(result[count], 0);
            }
            return result;
        }

        /// <summary>
        /// Convert an enumeration enumeration to a two-dimensional array
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="src">Enumeration enumeration</param>
        /// <returns>Jag Array</returns>
        public static T[][] ToJagArray<T>(this IEnumerable<IEnumerable<T>> src)
        {
            var result = new T[System.Linq.Enumerable.Count(src)][];
            foreach(var row in System.Linq.Enumerable.Select(src, (val, idx) => new { val, idx }))
            {   // äeçsÇîzóÒÇ…ïœä∑Ç∑ÇÈ
                result[row.idx] = System.Linq.Enumerable.ToArray(row.val);
            }
            return result;
        }

        /// <summary>
        /// Convert two-dimensional array to one-dimensional array
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="src">Enumeration enumeration</param>
        public static T[] ToSingleArray<T>(this IEnumerable<IEnumerable<T>> src)
        {
            var result = new List<T>();
            foreach (var row in src)
            {
                result.AddRange(row);
            }
            return result.ToArray();
        }
        #endregion

        #region Get
        /// <summary>
        /// Extract values vertically from the table
        /// </summary>
        public static IEnumerable<T> TakeVertical<T>(this T[][] jagArray, int column)
        {
            foreach (T[] line in jagArray)
            {   // If the length is insufficient return the default value
                yield return line.Length > column ? line[column] : default(T);
            }
        }

        /// <summary>
        /// Acquire the specified range value from the table
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="src">Table</param>
        /// <param name="top">Start row</param>
        /// <param name="bottom">End row</param>
        /// <param name="left">Start column</param>
        /// <param name="right">End column</param>
        public static IEnumerable<IEnumerable<T>> RangeTake<T>
                                    (this IEnumerable<IEnumerable<T>> src,
                                     Int32 top, Int32 bottom, Int32 left, Int32 right)
        {
            foreach (var item in System.Linq.Enumerable.Select(src, (val, idx) => new { val, idx }))
            {
                if (item.idx >= top && item.idx < bottom)
                {
                    yield return item.val.SkipTake(left, right);
                }
            }
        }

        /// <summary>Get maximum length in Table</summary>
        /// <param name="src">Table</param>
        public static Int32 ColumnsMax<T>(this T[][] src)
            => System.Linq.Enumerable.Max(src, row => row.Length);

        /// <summary>
        /// Extract specified range element from specified position in array
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="src">Array</param>
        /// <param name="skip">Length to skip</param>
        /// <param name="length">Length to take</param>
        public static T[] SkipTake<T>(this IEnumerable<T> src, Int32 skip, Int32 length)
        {
            if (skip < 0) { return new T[0]; }
            var result = new T[length];
            Array.Copy(src is Array ? src as Array : System.Linq.Enumerable.ToArray(src),
                        skip, result, 0, length);
            return result;
        }
        #endregion
    }
}