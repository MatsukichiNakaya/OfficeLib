using System;

namespace OfficeLib.XLS
{
    /// <summary></summary>
    public class Color
    {
        private Int32 MAX_VALUE = 255;
        private Int32 _r;
        private Int32 _g;
        private Int32 _b;

        /// <summary>Red</summary>
        public Int32 R {
            get { return this._r; }
            set { this._r = AdjustLimit(value); }
        }
        /// <summary>Green</summary>
        public Int32 G {
            get { return this._g; }
            set { this._g = AdjustLimit(value); }
        }
        /// <summary>Blue</summary>
        public Int32 B {
            get { return this._b; }
            set { this._b = AdjustLimit(value); }
        }

        /// <summary></summary>
        public Color()
        {   // Default Color is White.
            this._r = this.MAX_VALUE;
            this._g = this.MAX_VALUE;
            this._b = this.MAX_VALUE;
        }

        /// <summary></summary>
        public Color(Int32 r, Int32 g, Int32 b)
        {
            this.R = r;
            this.G = g;
            this.B = b;
        }

        /// <summary>
        /// Adjust to 0 - 255
        /// </summary>
        /// <param name="value"></param>
        private Int32 AdjustLimit(Int32 value)
        {
            if(value < 0) { return 0; }
            if(MAX_VALUE < value) { return MAX_VALUE; }
            return value;
        }

        /// <summary>
        /// Get color index
        /// </summary>
        /// <param name="r">Red</param>
        /// <param name="g">Green</param>
        /// <param name="b">Blue</param>
        public static Int32 RGB(Int32 r, Int32 g, Int32 b)
        {
            return r + (b << 8) + (b << 16);
        }

        /// <summary>
        /// Get color index
        /// </summary>
        public Int32 RGB()
        {
            return RGB(this._r, this._g, this._b);
        }
    }
}
