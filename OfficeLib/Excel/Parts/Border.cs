using System;

namespace OfficeLib.XLS
{
    /// <summary>Border class</summary>
    public class Border
    {
        /// <summary>Border index</summary>
        public XlBordersIndex Index { get; private set; }
        /// <summary>Border weight</summary>
        public XlBorderWeight Weight { get; set; }
        /// <summary>Border line style</summary>
        public XlLineStyle LineStyle { get; set; }
        /// <summary>Border color</summary>
        public Int32 Color { get; set; }

        /// <summary></summary>
        public Border()
        {
            this.Index = XlBordersIndex.xlEdgeLeft;
            this.Weight = XlBorderWeight.xlMedium;
            this.LineStyle = XlLineStyle.xlContinuous;
            this.Color = 0;
        }

        /// <summary></summary>
        public Border(XlBordersIndex index)
        {
            this.Index = index;
            this.Weight = XlBorderWeight.xlMedium;
            this.LineStyle = XlLineStyle.xlContinuous;
            this.Color = 0;
        }

        /// <summary></summary>
        /// <param name="index"></param>
        /// <param name="border"></param>
        public static Border Clone(XlBordersIndex index, Border border)
        {
            return new Border(index)
            {
                Weight = border.Weight,
                LineStyle = border.LineStyle,
                Color = border.Color,
            };
        }
    }
}
