using System;

namespace OfficeLib.XLS
{
    /// <summary>Excel extention method class</summary>
    public static class ExcelExtention
    {
        /// <summary>
        /// Convert character string to cell address
        /// </summary>
        /// <param name="value">character string</param>
        /// <returns>Range class object</returns>
        public static Address ToAddress(this String value)
        {
            return new Address(value);
        }

        /// <summary>
        /// Convert numbers to strings in A1 format
        /// </summary>
        /// <param name="value">Value to convert</param>
        public static String ToExcelColumnString(this Int32 value)
        {
            return Address.ToExcelColumnString((UInt32)value);
        }

        /// <summary>
        /// Convert column side strings to numbers
        /// </summary>
        /// <param name="value">column side strings</param>
        /// <returns></returns>
        public static UInt32 ToColumnNumber(this String value)
        {
            return Address.ToColumnNumber(value);
        }
    }
}
