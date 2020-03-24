using System;
using System.Collections.Generic;

namespace OfficeLib.XLS
{
    /// <summary>Cell address class</summary>
    public class Address
    {
        #region --- Constant ---
        /// <summary>Number of characters from alphabet A to Z</summary>
        private const UInt32 ALPHABET_CNT = 26;
        /// <summary>Offset at alphabet conversion</summary>
        private const UInt32 A_OFFSET_NUM = 'A' - 1;     // = 64

#if true
        // Todo : Difference in range by version
        /// <summary>Max of row number(2007-)</summary>
        public static readonly UInt32 MAX_ROW = 1048576;
        /// <summary>Max of column number(2007-)</summary>
        public static readonly UInt32 MAX_COLUMN = 16384;
#else
        // Pre-2003 settings
        /// <summary>Max of row number(-2003)</summary>
        public static readonly UInt32 MAX_ROW_OLD = 65536;
        /// <summary>Max of column number(-2003)</summary>
        public static readonly UInt32 MAX_COLUMN_OLD = 256;
#endif
        #endregion

        #region --- Property ---
        /// <summary>a1 format string</summary>
        public String ReferenceString { get; protected set; }
        /// <summary>Column number</summary>
        public UInt32 Column { get; protected set; }
        /// <summary>Row number</summary>
        public UInt32 Row { get; protected set; }
        #endregion

        #region --- Constructor ---
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="address">cell address string</param>
        public Address(String address)
        {
            var items = ToAddressItems(address);
            this.Column = items.Item1;
            this.Row = items.Item2;
            this.ReferenceString = address;
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="column"></param>
        /// <param name="row"></param>
        public Address(UInt32 column, UInt32 row)
        {
            this.Column = LimitColumnAdjustment(column);
            this.Row = LimitRowAdjustment(row);

            var a1Col = ToExcelColumnString(this.Column);
            this.ReferenceString = String.Format("{0}{1}", a1Col, this.Row);
        }
        #endregion

        #region --- Instance method ---
        /// <summary>
        /// Shift cell address
        /// </summary>
        /// <param name="col">Amount of movement in the column direction</param>
        /// <param name="row">Amount of movement in the row direction</param>
        public void Shift(Int32 col, Int32 row)
        {
            var temp = ((Int32)this.Column + col);
            this.Column = LimitColumnAdjustment(temp < 0 ? 0 : (UInt32)temp);

            temp = ((Int32)this.Row + row);
            this.Row = LimitRowAdjustment(temp < 0 ? 0 : (UInt32)temp);

            var a1Col = ToExcelColumnString(this.Column);
            this.ReferenceString = String.Format("{0}{1}", a1Col, this.Row);
        }

        /// <summary>
        /// Returns a character string of this range
        /// </summary>
        public override String ToString()
        {
            return this.ReferenceString;
        }
        #endregion

        #region --- Static method ---
        /// <summary>
        /// Shifts from the original address to get a new address.
        /// </summary>
        /// <param name="address">Range object</param>
        /// <param name="col">Amount of movement in the column direction</param>
        /// <param name="row">Amount of movement in the row direction</param>
        public static Address Shift(Address address, Int32 col, Int32 row)
        {
            var temp = ((Int32)address.Column + col);
            var c = LimitColumnAdjustment(temp < 0 ? 0 : (UInt32)temp);
            var a1Col = ToExcelColumnString(c);

            temp = ((Int32)address.Row + row);
            var r = LimitColumnAdjustment(temp < 0 ? 0 : (UInt32)temp);

            return new Address(String.Format("{0}{1}", a1Col, r));
        }

        /// <summary>
        /// Convert numbers to strings in A1 format
        /// </summary>
        /// <param name="value">Value to convert</param>
        public static String ToExcelColumnString(UInt32 value)
        {
            var result = String.Empty;
            try {   
                // Continue division until it becomes 0
                var tmpIndex = value;
                for (var i = value; i > 0; i--) {   
                    // Ask for surplus
                    var modIndex = tmpIndex % ALPHABET_CNT;
                    if (modIndex == 0) {   
                        // Set to Z if 0. Subtract 1 so that digits do not increase
                        modIndex = ALPHABET_CNT;
                        tmpIndex--;
                    }
                    tmpIndex = tmpIndex / ALPHABET_CNT;

                    // Convert remainder to alphabet and make it reference character
                    result = Convert.ToChar(A_OFFSET_NUM + modIndex).ToString() + result;
                    if (tmpIndex == 0) { break; }
                }
            }
            catch (Exception) { throw new Exception("ToExcelColumnString: Conversion failed"); }
            return result;
        }

        /// <summary>
        /// Convert column side strings to numbers
        /// </summary>
        /// <param name="columnString">column side strings</param>
        /// <returns></returns>
        public static UInt32 ToColumnNumber(String columnString)
        {
            return ConvertString(0, new Queue<Char>(columnString.ToCharArray()));
        }

        /// <summary>
        /// Convert character string to cell Items
        /// </summary>
        /// <param name="value">cell address string</param>
        /// <returns></returns>
        protected static Tuple<UInt32, UInt32> ToAddressItems(String value)
        {
            UInt32 col = 1;
            UInt32 row = 1;
            try {   // Split into rows and columns
                String[] addr = SplitRC(value);

                // Convert column
                col = ToColumnNumber(addr[Excel.COL]);

                // Convert Row
                if (!UInt32.TryParse(addr[Excel.ROW], out row)) {
                    throw new Exception("ToRange: Row conversion failed");
                }
            }
            catch (Exception) { throw new Exception("ToRange: Address conversion failed"); }

            return new Tuple<UInt32, UInt32>(col, row);
        }

        /// <summary>
        /// Split into rows and columns(A1 format)
        /// </summary>
        /// <param name="cellAddress"></param>
        /// <returns></returns>
        protected static String[] SplitRC(String cellAddress)
        {
            String[] result = null;
            try {
                var pattern = new System.Text.RegularExpressions.Regex(@"\d+");
                String[] temp = pattern.Split(cellAddress);
                result = new String[2];
                result[Excel.ROW] = cellAddress.Substring(
                                        temp[0].Length,
                                        cellAddress.Length - temp[0].Length);
                result[Excel.COL] = temp[0];
            }
            catch (Exception) {
                throw new Exception(
                    "SplitRC: Split failed. Please check the format of the string");
            }
            return result;
        }

        /// <summary>
        /// Convert strings
        /// </summary>
        /// <param name="returnValue">recursive call value(default value 0)</param>
        /// <param name="charQueue">char array</param>
        /// <returns></returns>
        private static UInt32 ConvertString(UInt32 returnValue, Queue<Char> charQueue)
        {
            if (charQueue.Count == 0) {
                return returnValue;
            }
            else {
                var charVal = charQueue.Dequeue(); //Take out one from the queue
                return ConvertString(CalcDesimal(charVal, charQueue.Count) + returnValue, charQueue);
            }
        }

        /// <summary>
        /// Value calculation for character code conversion
        /// </summary>
        /// <param name="charVal">charactor</param>
        /// <param name="count">process count</param>
        /// <returns></returns>
        private static UInt32 CalcDesimal(Char charVal, Int32 count)
        {
            return (UInt32)((Int32)Math.Pow(ALPHABET_CNT, count)
                                            * (charVal - A_OFFSET_NUM));
        }

        /// <summary>
        /// Limit check of row
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        private static UInt32 LimitRowAdjustment(UInt32 row)
        {
            if (MAX_ROW < row) { return MAX_ROW; }
            if (row < 1) { return 1; }

            return row;
        }

        /// <summary>
        /// Limit check of row
        /// </summary>
        /// <param name="column"></param>
        /// <returns></returns>
        private static UInt32 LimitColumnAdjustment(UInt32 column)
        {
            if (MAX_COLUMN < column) { return MAX_COLUMN; }
            if (column < 1) { return 1; }
            return column;
        }
        #endregion
    }
}
