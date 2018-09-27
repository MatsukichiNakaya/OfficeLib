using System;

namespace OfficeLib.XLS
{
    /// <summary>
    /// Border Thickness 
    /// </summary>
    public class Thickness
    {
        /// <summary>Left</summary>
        public Border Left { get; set; }
        /// <summary>Top</summary>
        public Border Top { get; set; }
        /// <summary>Right</summary>
        public Border Right { get; set; }
        /// <summary>Bottom</summary>
        public Border Bottom { get; set; }

        /// <summary>4 points param array</summary>
        public Border[] Values {
            get {
                return new Border[4] {
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
            this.Left = new Border(XlBordersIndex.xlEdgeLeft);
            this.Top = new Border(XlBordersIndex.xlEdgeTop);
            this.Right = new Border(XlBordersIndex.xlEdgeRight);
            this.Bottom = new Border(XlBordersIndex.xlEdgeBottom);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="left"></param>
        /// <param name="top"></param>
        /// <param name="right"></param>
        /// <param name="bottom"></param>
        public Thickness(Border left, Border top,
                         Border right, Border bottom)
        {
            // Arrange the index
            this.Left = Border.Clone(XlBordersIndex.xlEdgeLeft, left);
            this.Top = Border.Clone(XlBordersIndex.xlEdgeTop, top);
            this.Right = Border.Clone(XlBordersIndex.xlEdgeRight,right);
            this.Bottom = Border.Clone(XlBordersIndex.xlEdgeBottom, bottom);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="all"></param>
        public Thickness(Border all)
        {
            this.Left = Border.Clone(XlBordersIndex.xlEdgeLeft, all);
            this.Top = Border.Clone(XlBordersIndex.xlEdgeTop, all);
            this.Right = Border.Clone(XlBordersIndex.xlEdgeRight, all);
            this.Bottom = Border.Clone(XlBordersIndex.xlEdgeBottom, all);
        }
    }
}
