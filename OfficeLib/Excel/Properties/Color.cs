using System;

namespace OfficeLib.Excel
{
    /// <summary></summary>
    public class Color
    {
        private UInt32 MAX_VALUE = 255;
        private UInt32 _r;
        private UInt32 _g;
        private UInt32 _b;

        /// <summary></summary>
        public UInt32 R {
            get { return this._r; }
            set { this._r = this.MAX_VALUE < value ? this.MAX_VALUE : value; }
        }
        /// <summary></summary>
        public UInt32 G {
            get { return this._g; }
            set { this._g = this.MAX_VALUE < value ? this.MAX_VALUE : value; }
        }
        /// <summary></summary>
        public UInt32 B {
            get { return this._b; }
            set { this._b = this.MAX_VALUE < value ? this.MAX_VALUE : value; }
        }

        /// <summary></summary>
        public Color()
        {   // Default Color is White.
            this._r = this.MAX_VALUE;
            this._g = this.MAX_VALUE;
            this._b = this.MAX_VALUE;
        }

        /// <summary></summary>
        public Color(UInt32 r, UInt32 g, UInt32 b)
        {
            this.R = r;
            this.G = g;
            this.B = b;
        }
    }
}
