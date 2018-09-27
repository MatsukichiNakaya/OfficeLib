using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace OfficeLib
{
    /// <summary>Extended function class for type conversion</summary>
    internal static class ConvertExtention
    {
        /// <summary>
        /// Type Conversion
        /// </summary>
        /// <typeparam name="TOutput">Type after conversion</typeparam>
        /// <param name="value">Value of conversion source</param>
        public static TOutput To<TOutput>(this Object value)
        {   
            if(value == null) { return default; }
            try
            {
                TypeConverter converter = TypeDescriptor.GetConverter(typeof(TOutput));

                return converter != null
                            ? (TOutput)converter.ConvertTo(value, typeof(TOutput))
                            : default;
            }
            catch (Exception) { return default; }
        }

        /// <summary>
        /// Type Conversion
        /// </summary>
        /// <param name="value">Value of conversion source</param>
        /// <param name="type">Type after conversion</param>
        /// <remarks>Cast of return value required</remarks>
        public static Object To(this Object value, Type type)
        {
            Object result = type.IsValueType ? Activator.CreateInstance(type) : null;
            if (value == null) { return result; }
            try
            {
                var converter = TypeDescriptor.GetConverter(type);
                return converter != null
                        ? converter.ConvertTo(value, type) : result;
            }
            catch (Exception) { return result; }
        }

#if false
        // 

        /// <summary>
        /// Type Conversion (Perform exception handling)
        /// </summary>
        /// <typeparam name="TOutput">Type after conversion</typeparam>
        /// <param name="value">Value of conversion source</param>
        public static TOutput TryTo<TOutput>(this Object value)
        {
            if (value == null) { return default(TOutput); }
            var converter = TypeDescriptor.GetConverter(typeof(TOutput));
            try
            {
                return converter != null
                        ? (TOutput)converter.ConvertTo(value, typeof(TOutput))
                        : default(TOutput);
            }
            catch (Exception) { return default(TOutput); }
        }

        /// <summary>
        /// Type Conversion (Perform exception handling)
        /// </summary>
        /// <param name="value">Value of conversion source</param>
        /// <param name="type">Type after conversion</param>
        /// <remarks>Cast of return value required</remarks>
        public static Object TryTo(this Object value, Type type)
        {
            Object result = type.IsValueType
                            ? Activator.CreateInstance(type) : null;
            if (value == null) { return result; }
            try
            {
                var converter = TypeDescriptor.GetConverter(type);
                return converter != null
                        ? converter.ConvertTo(value, type) : result;
            }
            catch (Exception) { return result; }
        }
#endif

        /// <summary>Type Conversion of Arrays</summary>
        /// <typeparam name="TInput">Type of conversion source</typeparam>
        /// <typeparam name="TOutput">Type after conversion</typeparam>
        /// <param name="values">Value of conversion source</param>
        public static TOutput[] ConvertAll<TInput, TOutput>(this IEnumerable<TInput> values)
        {
            if (values == null) { return null; }
            return Array.ConvertAll(values.ToArray(), val => val.To<TOutput>());
        }

        /// <summary>Type Conversion of Jag Arrays</summary>
        /// <typeparam name="TInput">Type of conversion source</typeparam>
        /// <typeparam name="TOutput">Type after conversion</typeparam>
        /// <param name="values">Value of conversion source</param>
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
        /// Convert an array to an object type
        /// </summary>
        /// <typeparam name="T">Type of conversion source</typeparam>
        /// <param name="src">Value of conversion source</param>
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


