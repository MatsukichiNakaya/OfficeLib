using System;
using System.Collections.Generic;

namespace OfficeLib
{
    internal static class JagExtention
    {
        #region ����
        /// <summary> �z��̉�����(�S�Ă��܂�) </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="srcTable">�E���̃e�[�u��</param>
        /// <param name="joinTable">�����̃e�[�u��</param>
        /// <remarks>
        /// SQL��OuterJoin�Ƃ͏����d�l���قȂ邪
        /// ���ׂĂ̓��e���܂ނ悤�ɁA�z��̗v�f����
        /// �����ق�����ɂ��ĉ������Ō�������B
        /// �Z���z��̑��݂��Ȃ��v�f�̓f�t�H���g�l�ō쐬�����B
        /// </remarks>
        public static T[][] OuterJoin<T>(this T[][] srcTable, T[][] joinTable)
        {
            T[][] result = null;
            // �����`�F�b�N
            if (srcTable == null) { throw new NullReferenceException(); }
            if (joinTable == null) { return srcTable; }
            try
            {   // �z��̏c�̒������ׂĒ����ق����擾
                Int32 rows = (srcTable?.Length ?? 0) > (joinTable?.Length ?? 0) ?
                                srcTable?.Length ?? 0 : joinTable?.Length ?? 0;
                result = Join(srcTable, joinTable, rows);
            }
            catch (Exception ex) { throw ex; }
            return result;
        }

        /// <summary> �z��̉�����(���ʍs�̂�) </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="srcTable">�E���̃e�[�u��</param>
        /// <param name="joinTable">�����̃e�[�u��</param>
        /// <remarks>
        /// SQL��InnerJoin�Ƃ͂��Ȃ�d�l���قȂ邪
        /// �z��̗v�f�������Ȃ��ق�����Ƃ���
        /// �������Ō�������B
        /// �����ق��̗]�s�͐؂�̂Ă���B
        /// </remarks>
        public static T[][] InnerJoin<T>(this T[][] srcTable, T[][] joinTable)
        {
            T[][] result = null;
            // �����`�F�b�N
            if (srcTable == null) { throw new NullReferenceException(); }
            if (joinTable == null) { return srcTable; }
            try
            {   // �z��̏c�̒������ׂĒZ���ق����擾
                Int32 rows = (srcTable?.Length ?? 0) < (joinTable?.Length ?? 0) ?
                                srcTable?.Length ?? 0 : joinTable?.Length ?? 0;
                result = Join(srcTable, joinTable, rows);
            }
            catch (Exception ex) { throw ex; }
            return result;
        }

        /// <summary> �w�肳�ꂽ�s���ŉ����� </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="srcTable">�E���̃e�[�u��</param>
        /// <param name="joinTable">�����̃e�[�u��</param>
        /// <param name="rows">�w��s</param>
        /// <remarks>
        /// ��̃e�[�u���ɑ΂��Ďw�肳�ꂽ�s���������Č�������B
        /// �w��s�𒴂����ꍇ�́A�f�t�H���g�l�ō쐬��
        /// ���܂����s�͐؂�̂Ă���B
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
            catch (Exception ex) { throw ex; }
            return result;
        }

        /// <summary> �z��̏c���� </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="srcTable">�㑤�̃e�[�u��</param>
        /// <param name="unionTable">�����̃e�[�u��</param>
        /// <remarks>
        /// SQL��Union�Ƃ͎d�l���قȂ�B
        /// �x�[�X�̃e�[�u���ɑ΂��ĉ����Ƀe�[�u������������B
        /// �e��͂��̂܂܌��������B
        /// </remarks>
        public static T[][] Union<T>(this T[][] srcTable, params T[][][] unionTable)
        {
            List<T[]> result = new List<T[]>();
            try
            {
                List<T[][]> list = new List<T[][]>();
                list.Add(srcTable);        // �x�[�X�ƒǉ��v�f��
                list.AddRange(unionTable); // �ЂƂ̃R���N�V�����ɏW��
                foreach (T[][] rows in list)
                {
                    foreach (T[] row in rows) { result.Add(row); }
                }
            }
            catch (Exception ex) { throw ex; }
            return result.ToArray();
        }
        #endregion

        #region ����
        /// <summary> �z����c�ɕ��� </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="src">�z��̔z��</param>
        /// <param name="dividCount">������</param>
        public static T[][][] VDividingArray<T>(this T[][] src, Int32 dividCount)
        {
            List<T[][]> result = new List<T[][]>();
            try
            {
                Int32 colLength = src.ColumnsMax();
                Int32 dividCol = colLength / dividCount;
                Int32 pos = 0;
                // ����؂�Ȃ��ꍇ�͏I��
                if (colLength % dividCount > 0) { return new T[][][] { src }; }
                // �������[�v�ɂȂ邽�ߏI��
                if (dividCol <= 0) { return new T[][][] { src }; }

                while (pos < colLength)
                {
                    List<T[]> tables = new List<T[]>();
                    for (var row = 0; row < src.Length; row++)
                    {
                        try     // �w��ʒu����w�萔�̗v�f���擾
                        { tables.Add(src[row].SkipTake(pos, dividCol)); }
                        catch (Exception) { break; }
                    }
                    result.Add(tables.ToArray());
                    pos += dividCol;
                }
            }
            catch (Exception ex) { throw ex; }
            return result.ToArray();
        }

        /// <summary> �������̕��� </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="src">�z��̔z��</param>
        /// <param name="divid">������</param>
        public static T[][][] HDividingArray<T>(this T[][] src, Int32 divid)
        {
            List<T[][]> result = new List<T[][]>();

            Int32 dividRow = src.Length / divid;
            // ����؂�Ȃ��ꍇ�͏I��
            if (src.Length % divid > 0) { return new T[][][] { src }; }
            // �������[�v�ɂȂ邽�ߏI��
            if (dividRow <= 0) { return new T[][][] { src }; }

            Queue<T[]> rowQueue = new Queue<T[]>(src);
            while (rowQueue.Count > 0)
            {
                List<T[]> row = new List<T[]>();
                for (var i = 0; i < dividRow; i++) // �v�f�����Ƃ̉���쐬
                {   // ��s���z����i�[
                    try { row.Add(rowQueue.Dequeue()); }
                    catch (Exception) { break; }
                }
                result.Add(row.ToArray());
            }
            return result.ToArray();
        }
        #endregion

        #region �폜
        /// <summary> �ŏ��̗񂩂�w��񐔍폜 </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="src">�z��̔z��</param>
        /// <param name="trimCount">�폜�����</param>
        public static T[][] ColumnRemoveStart<T>(this T[][] src, Int32 trimCount)
        {
            T[][] result = new T[src.Length][];
            try
            {
                Int32 colLength = src.ColumnsMax();
                // �S��Trim�����ꍇ�͗v�f���Ȃ��Ȃ邽��null��Ԃ��܂��B
                if (colLength - trimCount <= 0) { return null; }

                for (var row = 0; row < result.Length; row++)
                {
                    result[row] = new T[colLength - trimCount];
                    for (var col = 0; col < result[row].Length; col++)
                    {
                        try { result[row][col] = src[row][col + trimCount]; }
                        // �v�f���Ȃ��Ȃ����ꍇ�̓��T�C�Y���ďI��
                        catch (Exception) { Array.Resize(ref result[row], col); break; }
                    }
                }
            }
            catch (Exception ex) { throw ex; }
            return result;
        }

        /// <summary> �Ō�̗񂩂�w��񐔍폜 </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="src">�z��̔z��</param>
        /// <param name="trimCount">�폜�����</param>
        public static T[][] ColumnRemoveEnd<T>(this T[][] src, Int32 trimCount)
        {
            T[][] result = new T[src.Length][];
            try
            {
                Int32 colLength = src.ColumnsMax();
                // �S��Trim�����ꍇ�͗v�f���Ȃ��Ȃ邽��null��Ԃ��܂��B
                if (colLength - trimCount <= 0) { return null; }

                for (var row = 0; row < result.Length; row++)
                {
                    result[row] = new T[colLength - trimCount];
                    for (var col = 0; col < result[row].Length; col++)
                    {
                        try { result[row][col] = src[row][col]; }
                        // �v�f���Ȃ��Ȃ����ꍇ�̓��T�C�Y���ďI��
                        catch (Exception) { Array.Resize(ref result[row], col); break; }
                    }
                }
            }
            catch (Exception ex) { throw ex; }
            return result;
        }
        #endregion

        #region �ϊ�
        /// <summary> ��`�z��ւ̕ϊ� </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="src">�z��̔z��</param>
        public static T[,] ToRectArray<T>(this T[][] src)
        {
            T[,] result = null;
            result = new T[src.Length, src.ColumnsMax()];
            for (var row = 0; row < result.GetLength(0); row++)
            {
                for (var col = 0; col < result.GetLength(1); col++)
                {
                    try               // �v�f������ꍇ�͂��̂܂ܒl����
                    { result[row, col] = src[row][col]; }
                    catch (Exception) // �����ꍇ�̓f�t�H���g�l
                    { result[row, col] = default(T); }
                }
            }
            return result;
        }

        /// <summary>��`�z��̂悤�ɃW���O�z��̉����������낦��</summary>
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

        /// <summary>�ꎟ���̔z���񎟌��̔z��ɕϊ�����</summary>
        /// <param name="src">�z��</param>
        /// <param name="divCount">�e�����̗v�f��</param>
        /// <remarks>
        /// �ꎟ���̗v�f��������؂�鐔���w�肷�邱��
        /// �������Ƃ������_�ȉ��͐؎�
        /// �Œ��divCount�Ŏw�肵���z�񐔂͕Ԃ��Ă���B
        /// </remarks>
        public static T[][] ToJagArray<T>(this IEnumerable<T> src, Int32 divCount)
        {
            T[][] result = new T[divCount][];
            // �񎟌��ڂ̗v�f��
            Int32 columnLength = System.Linq.Enumerable.Count(src) / divCount;

            for (var count = 0; count < result.Length; count++)
            {
                result[count] = new T[columnLength];
                // ���Ԋu�Œl���擾����
                src.SkipTake(count * columnLength, columnLength).CopyTo(result[count], 0);
            }
            return result;
        }

        /// <summary>
        /// �񋓌^�̗񋓂�񎟌��z��ɕϊ�����
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="src"></param>
        /// <returns></returns>
        public static T[][] ToJagArray<T>(this IEnumerable<IEnumerable<T>> src)
        {
            T[][] result = new T[System.Linq.Enumerable.Count(src)][];
            foreach(var row in System.Linq.Enumerable.Select(src, (val, idx) => new { val, idx }))
            {   // �e�s��z��ɕϊ�����
                result[row.idx] = System.Linq.Enumerable.ToArray(row.val);
            }
            return result;
        }

        /// <summary>
        /// �񎟌��z����ꎟ���z��ɕϊ�
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="src">�z��̔z��</param>
        public static T[] ToSingleArray<T>(this IEnumerable<IEnumerable<T>> src)
        {
            List<T> result = new List<T>();
            foreach (IEnumerable<T> row in src)
            {   // ���J��ɔz��֊i�[���Ă���
                result.AddRange(row);
            }
            return result.ToArray();
        }
        #endregion

        #region �擾
        /// <summary>
        /// �z�������c�ɒl�𔲂��o��
        /// </summary>
        public static IEnumerable<T> TakeVertical<T>(this T[][] jagArray, int column)
        {
            foreach (T[] line in jagArray)
            {   // ����������Ȃ��ꍇ�̓f�t�H���g�l��Ԃ�
                yield return line.Length > column ? line[column] : default(T);
            }
        }

        /// <summary>
        /// �z��̔z�񂩂�w��͈͂̒l���擾����
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="src">�z��</param>
        /// <param name="top">�J�n�s</param>
        /// <param name="bottom">�I���s</param>
        /// <param name="left">�J�n��</param>
        /// <param name="right">�I����</param>
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

        /// <summary>�z�񒆂̍ő咷���擾</summary>
        /// <param name="src">�z��̔z��</param>
        public static Int32 ColumnsMax<T>(this T[][] src)
            // �z��̒������擾���čő�l��Ԃ�
            =>  System.Linq.Enumerable.Max(src, row => row.Length);

        /// <summary>
        /// �z����̎w��ʒu����w��͈̗͂v�f�����o��
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="src">�z��</param>
        /// <param name="skip">�J�n�ʒu</param>
        /// <param name="length">�擾����</param>
        public static T[] SkipTake<T>(this IEnumerable<T> src, Int32 skip, Int32 length)
        {
            if (skip < 0) { return new T[0]; }
            T[] result = new T[length];
            // Linq��Skip().Take()���x���ׁAArrayCopy���g�p
            Array.Copy(src is Array ? src as Array : System.Linq.Enumerable.ToArray(src),
                        skip, result, 0, length);
            return result;
        }
        #endregion
    }
}