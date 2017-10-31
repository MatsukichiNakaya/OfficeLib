using System;
using System.Collections.Generic;

namespace OfficeLib
{
    internal static class JagExtention
    {
        #region 結合
        /// <summary> 配列の横結合(全てを含む) </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="srcTable">右側のテーブル</param>
        /// <param name="joinTable">左側のテーブル</param>
        /// <remarks>
        /// SQLのOuterJoinとは少し仕様が異なるが
        /// すべての内容を含むように、配列の要素数が
        /// 多いほうを基準にして横方向で結合する。
        /// 短い配列の存在しない要素はデフォルト値で作成される。
        /// </remarks>
        public static T[][] OuterJoin<T>(this T[][] srcTable, T[][] joinTable)
        {
            T[][] result = null;
            // 初期チェック
            if (srcTable == null) { throw new NullReferenceException(); }
            if (joinTable == null) { return srcTable; }
            try
            {   // 配列の縦の長さを比べて長いほうを取得
                Int32 rows = (srcTable?.Length ?? 0) > (joinTable?.Length ?? 0) ?
                                srcTable?.Length ?? 0 : joinTable?.Length ?? 0;
                result = Join(srcTable, joinTable, rows);
            }
            catch (Exception) { throw; }
            return result;
        }

        /// <summary> 配列の横結合(共通行のみ) </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="srcTable">右側のテーブル</param>
        /// <param name="joinTable">左側のテーブル</param>
        /// <remarks>
        /// SQLのInnerJoinとはかなり仕様が異なるが
        /// 配列の要素数が少ないほうを基準として
        /// 横方向で結合する。
        /// 長いほうの余行は切り捨てられる。
        /// </remarks>
        public static T[][] InnerJoin<T>(this T[][] srcTable, T[][] joinTable)
        {
            T[][] result = null;
            // 初期チェック
            if (srcTable == null) { throw new NullReferenceException(); }
            if (joinTable == null) { return srcTable; }
            try
            {   // 配列の縦の長さを比べて短いほうを取得
                Int32 rows = (srcTable?.Length ?? 0) < (joinTable?.Length ?? 0) ?
                                srcTable?.Length ?? 0 : joinTable?.Length ?? 0;
                result = Join(srcTable, joinTable, rows);
            }
            catch (Exception) { throw; }
            return result;
        }

        /// <summary> 指定された行数で横結合 </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="srcTable">右側のテーブル</param>
        /// <param name="joinTable">左側のテーブル</param>
        /// <param name="rows">指定行</param>
        /// <remarks>
        /// 二つのテーブルに対して指定された行数を持って結合する。
        /// 指定行を超えた場合は、デフォルト値で作成し
        /// あまった行は切り捨てられる。
        /// </remarks>
        private static T[][] Join<T>(T[][] srcTable, T[][] joinTable, Int32 rows)
        {
            T[][] result = null;
            try
            {
                Int32 colSrc = srcTable.ColumnsMax();
                Int32 colJoin = joinTable.ColumnsMax();
                result = new T[rows][];
                for (var row = 0; row < result.Length; row++)
                {
                    result[row] = new T[colSrc + colJoin];
                    if (row < srcTable.Length)
                    {
                        srcTable[row]?.CopyTo(result[row], 0);
                    }
                    if (row < joinTable.Length) { joinTable[row]?.CopyTo(result[row], colSrc); }
                }
            }
            catch (Exception) { throw; }
            return result;
        }

        /// <summary> 配列の縦結合 </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="srcTable">上側のテーブル</param>
        /// <param name="unionTable">下側のテーブル</param>
        /// <remarks>
        /// SQLのUnionとは仕様が異なる。
        /// ベースのテーブルに対して下側にテーブルを結合する。
        /// 各列はそのまま結合される。
        /// </remarks>
        public static T[][] Union<T>(this T[][] srcTable, params T[][][] unionTable)
        {
            var result = new List<T[]>();
            try
            {   // ベースと追加要素を
                var list = new List<T[][]> { srcTable };
                // ひとつのコレクションに集約
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

        #region 分割
        /// <summary> 配列を縦に分割 </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="src">配列の配列</param>
        /// <param name="dividCount">分割数</param>
        public static T[][][] VDividingArray<T>(this T[][] src, Int32 dividCount)
        {
            var result = new List<T[][]>();
            try
            {
                Int32 colLength = src.ColumnsMax();
                Int32 dividCol = colLength / dividCount;
                Int32 pos = 0;
                // 割り切れない場合は終了
                if (colLength % dividCount > 0) { return new T[][][] { src }; }
                // 無限ループになるため終了
                if (dividCol <= 0) { return new T[][][] { src }; }

                while (pos < colLength)
                {
                    var tables = new List<T[]>();
                    for (var row = 0; row < src.Length; row++)
                    {
                        try     // 指定位置から指定数の要素を取得
                        { tables.Add(src[row].SkipTake(pos, dividCol)); }
                        catch (Exception) { break; }
                    }
                    result.Add(tables.ToArray());
                    pos += dividCol;
                }
            }
            catch (Exception) { throw; }
            return result.ToArray();
        }

        /// <summary> 横方向の分割 </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="src">配列の配列</param>
        /// <param name="divid">分割数</param>
        public static T[][][] HDividingArray<T>(this T[][] src, Int32 divid)
        {
            var result = new List<T[][]>();

            Int32 dividRow = src.Length / divid;
            // 割り切れない場合は終了
            if (src.Length % divid > 0) { return new T[][][] { src }; }
            // 無限ループになるため終了
            if (dividRow <= 0) { return new T[][][] { src }; }

            var rowQueue = new Queue<T[]>(src);
            while (rowQueue.Count > 0)
            {
                var row = new List<T[]>();
                for (var i = 0; i < dividRow; i++) // 要素数ごとの塊を作成
                {   // 一行ずつ配列を格納
                    try { row.Add(rowQueue.Dequeue()); }
                    catch (Exception) { break; }
                }
                result.Add(row.ToArray());
            }
            return result.ToArray();
        }
        #endregion

        #region 削除
        /// <summary> 最初の列から指定列数削除 </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="src">配列の配列</param>
        /// <param name="trimCount">削除する列数</param>
        public static T[][] ColumnRemoveStart<T>(this T[][] src, Int32 trimCount)
        {
            T[][] result = new T[src.Length][];
            try
            {
                Int32 colLength = src.ColumnsMax();
                // 全部Trimされる場合は要素がなくなるためnullを返します。
                if (colLength - trimCount <= 0) { return null; }

                for (var row = 0; row < result.Length; row++)
                {
                    result[row] = new T[colLength - trimCount];
                    for (var col = 0; col < result[row].Length; col++)
                    {
                        try { result[row][col] = src[row][col + trimCount]; }
                        // 要素がなくなった場合はリサイズして終了
                        catch (Exception) { Array.Resize(ref result[row], col); break; }
                    }
                }
            }
            catch (Exception) { throw; }
            return result;
        }

        /// <summary> 最後の列から指定列数削除 </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="src">配列の配列</param>
        /// <param name="trimCount">削除する列数</param>
        public static T[][] ColumnRemoveEnd<T>(this T[][] src, Int32 trimCount)
        {
            T[][] result = new T[src.Length][];
            try
            {
                Int32 colLength = src.ColumnsMax();
                // 全部Trimされる場合は要素がなくなるためnullを返します。
                if (colLength - trimCount <= 0) { return null; }

                for (var row = 0; row < result.Length; row++)
                {
                    result[row] = new T[colLength - trimCount];
                    for (var col = 0; col < result[row].Length; col++)
                    {
                        try { result[row][col] = src[row][col]; }
                        // 要素がなくなった場合はリサイズして終了
                        catch (Exception) { Array.Resize(ref result[row], col); break; }
                    }
                }
            }
            catch (Exception) { throw; }
            return result;
        }
        #endregion

        #region 変換
        /// <summary> 矩形配列への変換 </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="src">配列の配列</param>
        public static T[,] ToRectArray<T>(this T[][] src)
        {
            T[,] result = null;
            result = new T[src.Length, src.ColumnsMax()];
            for (var row = 0; row < result.GetLength(0); row++)
            {
                for (var col = 0; col < result.GetLength(1); col++)
                {
                    try               // 要素がある場合はそのまま値を代入
                    { result[row, col] = src[row][col]; }
                    catch (Exception) // 無い場合はデフォルト値
                    { result[row, col] = default(T); }
                }
            }
            return result;
        }

        /// <summary>矩形配列のようにジャグ配列の横長さをそろえる</summary>
        public static T[][] ToRectLikeJagArray<T>(this T[][] src)
        {
            T[][] result = new T[src.Length][];
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

        /// <summary>一次元の配列を二次元の配列に変換する</summary>
        /// <param name="src">配列</param>
        /// <param name="divCount">各次元の要素数</param>
        /// <remarks>
        /// 一次元の要素数を割り切れる数を指定すること
        /// 割ったとき小数点以下は切捨
        /// 最低限divCountで指定した配列数は返ってくる。
        /// </remarks>
        public static T[][] ToJagArray<T>(this IEnumerable<T> src, Int32 divCount)
        {
            T[][] result = new T[divCount][];
            // 二次元目の要素数
            Int32 columnLength = System.Linq.Enumerable.Count(src) / divCount;

            for (var count = 0; count < result.Length; count++)
            {
                result[count] = new T[columnLength];
                // 一定間隔で値を取得して
                src.SkipTake(count * columnLength, columnLength).CopyTo(result[count], 0);
            }
            return result;
        }

        /// <summary>
        /// 列挙型の列挙を二次元配列に変換する
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="src"></param>
        /// <returns></returns>
        public static T[][] ToJagArray<T>(this IEnumerable<IEnumerable<T>> src)
        {
            T[][] result = new T[System.Linq.Enumerable.Count(src)][];
            foreach(var row in System.Linq.Enumerable.Select(src, (val, idx) => new { val, idx }))
            {   // 各行を配列に変換する
                result[row.idx] = System.Linq.Enumerable.ToArray(row.val);
            }
            return result;
        }

        /// <summary>
        /// 二次元配列を一次元配列に変換
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="src">配列の配列</param>
        public static T[] ToSingleArray<T>(this IEnumerable<IEnumerable<T>> src)
        {
            var result = new List<T>();
            foreach (IEnumerable<T> row in src)
            {   // 順繰りに配列へ格納していく
                result.AddRange(row);
            }
            return result.ToArray();
        }
        #endregion

        #region 取得
        /// <summary>
        /// 配列内から縦に値を抜き出す
        /// </summary>
        public static IEnumerable<T> TakeVertical<T>(this T[][] jagArray, int column)
        {
            foreach (T[] line in jagArray)
            {   // 長さが足りない場合はデフォルト値を返す
                yield return line.Length > column ? line[column] : default(T);
            }
        }

        /// <summary>
        /// 配列の配列から指定範囲の値を取得する
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="src">配列</param>
        /// <param name="top">開始行</param>
        /// <param name="bottom">終了行</param>
        /// <param name="left">開始列</param>
        /// <param name="right">終了列</param>
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

        /// <summary>配列中の最大長を取得</summary>
        /// <param name="src">配列の配列</param>
        public static Int32 ColumnsMax<T>(this T[][] src)
            // 配列の長さを取得して最大値を返す
            =>  System.Linq.Enumerable.Max(src, row => row.Length);

        /// <summary>
        /// 配列内の指定位置から指定範囲の要素を取り出す
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="src">配列</param>
        /// <param name="skip">開始位置</param>
        /// <param name="length">取得長さ</param>
        public static T[] SkipTake<T>(this IEnumerable<T> src, Int32 skip, Int32 length)
        {
            if (skip < 0) { return new T[0]; }
            T[] result = new T[length];
            // LinqのSkip().Take()が遅い為、ArrayCopyを使用
            Array.Copy(src is Array ? src as Array : System.Linq.Enumerable.ToArray(src),
                        skip, result, 0, length);
            return result;
        }
        #endregion
    }
}