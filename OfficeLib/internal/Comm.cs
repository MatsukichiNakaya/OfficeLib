using System;
using System.Collections.Generic;

namespace OfficeLib
{
    internal static class Comm
    {
        /// <summary>
        /// Is it an value containing bit
        /// </summary>
        /// <param name="value">value</param>
        /// <param name="bit">bit</param>
        public static Boolean ContainsBitFlag(this UInt32 value, Int32 bit)
            // Confirm whether the specified bit is set in the number
            => (value & (1 << bit - 1)) != 0;

        /// <summary>
        /// Is it an array containing keys
        /// </summary>
        /// <param name="values">value</param>
        /// <param name="key">key</param>
        public static Boolean Contains<T>(this T[] values, T key)
        {
            var checker = new HashSet<T>(values);
            if (checker.Contains(key)) { return true; }
            return false;
        }

        /// <summary>
        /// String is null or empty
        /// </summary>
        /// <param name="value">value</param>
        public static Boolean IsNullOrEmpty(this String value)
        {
            return value == null || value == String.Empty;
        }
    }
}
