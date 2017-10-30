using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace OfficeLib
{
    /// <summary>値の型変換に関する拡張関数クラス</summary>
    internal static class ConvertExtention
    {
        /// <summary>
        /// 値の変換
        /// </summary>
        /// <typeparam name="TOutput">変換後の型</typeparam>
        /// <param name="value">変換もとの値</param>
        public static TOutput To<TOutput>(this Object value)
        {   
            if(value == null) { return default(TOutput); }

            // コンバータを作成して型変換を行う
            var converter = TypeDescriptor.GetConverter(typeof(TOutput));
            return converter != null
                    ? (TOutput)converter.ConvertTo(value, typeof(TOutput))
                    : default(TOutput);
        }

        /// <summary>
        /// 型を指定しての変換
        /// </summary>
        /// <param name="value">変換もとの値</param>
        /// <param name="type">変換後の型</param>
        /// <remarks>戻り値のキャストが必要</remarks>
        public static Object To(this Object value, Type type)
        {
            if (value == null)
            {   // Nullの場合、値型では0, 参照型ではnullを返すようにする
                return type.IsValueType 
                        ? Activator.CreateInstance(type) : null;
            }

            // コンバータを作成して型変換を行う
            var converter = TypeDescriptor.GetConverter(type);
            return converter != null 
                    ? converter.ConvertTo(value, type) 
                    : (type.IsValueType ? Activator.CreateInstance(type) : null);
        }

        /// <summary>
        /// 値の変換 内部で例外処理を行う
        /// </summary>
        /// <typeparam name="TOutput">変換後の型</typeparam>
        /// <param name="value">変換もとの値</param>
        public static TOutput TryTo<TOutput>(this Object value)
        {
            if (value == null) { return default(TOutput); }

            // コンバータを作成して型変換を行う
            var converter = TypeDescriptor.GetConverter(typeof(TOutput));
            try
            {
                return converter != null
                        ? (TOutput)converter.ConvertTo(value, typeof(TOutput))
                        : default(TOutput);
            }
            // 例外時はデフォルト値を返す
            catch (Exception) { return default(TOutput); }
        }

        /// <summary>
        /// 型を指定しての変換 内部で例外処理を行う
        /// </summary>
        /// <param name="value">変換もとの値</param>
        /// <param name="type">変換後の型</param>
        /// <remarks>戻り値のキャストが必要</remarks>
        public static Object TryTo(this Object value, Type type)
        {
            // Nullの場合、値型では0, 参照型ではnullを返すようにする
            Object result = type.IsValueType
                            ? Activator.CreateInstance(type) : null;
            if (value == null) { return result; }
            try
            {
                // コンバータを作成して型変換を行う
                var converter = TypeDescriptor.GetConverter(type);
                return converter != null
                        ? converter.ConvertTo(value, type) : result;
            }
            // 例外時はデフォルト値を返す
            catch (Exception) { return result; }
        }

        /// <summary> 値の変換(一次配列) </summary>
        /// <typeparam name="TInput">変換もとの型</typeparam>
        /// <typeparam name="TOutput">変換後の型</typeparam>
        /// <param name="values">変換もとの値</param>
        public static TOutput[] ConvertAll<TInput, TOutput>(this IEnumerable<TInput> values)
        {
            if (values == null) { return null; }
            return Array.ConvertAll(values.ToArray(), val => val.To<TOutput>());
        }

        /// <summary> 値の変換(二次配列) </summary>
        /// <typeparam name="TInput">変換もとの型</typeparam>
        /// <typeparam name="TOutput">変換後の型</typeparam>
        /// <param name="values">変換もとの値</param>
        public static TOutput[][] ConvertAll<TInput, TOutput>(this IEnumerable<IEnumerable<TInput>> values)
        {
            if (values == null) { return null; }

            var result = new TOutput[values.Count()][];
            Int32 count = 0;
            foreach (var val in values)
            {
                result[count] = val?.ConvertAll<TInput, TOutput>();
                count++;
            }
            return result;
        }

        /// <summary>
        /// 配列をObject型へ変換する
        /// </summary>
        /// <typeparam name="T">変換もとの型</typeparam>
        /// <param name="src">変換を行う配列</param>
        /// <returns>Object型に変換された配列</returns>
        public static Object ToObject<T>(this T[] src)
        {
            return src;
        }

        /// <summary>
        /// Enumerable to array
        /// </summary>
        public static T[] ToArray<T>(this IEnumerable<T> src)
        {
            return Enumerable.ToArray(src);
        }
    }
}


