using System;

namespace OfficeLib.Excel
{
    /// <summary></summary>
    public class Thickness
    {
        /// <summary></summary>
        public Double Left { get; set; }
        /// <summary></summary>
        public Double Top { get; set; }
        /// <summary></summary>
        public Double Right { get; set; }
        /// <summary></summary>
        public Double Bottom { get; set; }

        /// <summary></summary>
        public Double[] Values {
            get {
                return new Double[4] {
                            this.Left,
                            this.Top,
                            this.Right,
                            this.Bottom
                       };
            }
        }

        /// <summary></summary>
        public Thickness()
        {
            this.Left = 0;
            this.Top = 0;
            this.Right = 0;
            this.Bottom = 0;
        }

        /// <summary></summary>
        /// <param name="left"></param>
        /// <param name="top"></param>
        /// <param name="right"></param>
        /// <param name="bottom"></param>
        public Thickness(Double left, Double top, Double right, Double bottom)
        {
            this.Left = left;
            this.Top = top;
            this.Right = right;
            this.Bottom = bottom;
        }
    }
}
