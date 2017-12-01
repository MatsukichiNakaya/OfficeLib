using System;
using static OfficeLib.Commands;
using static OfficeLib.XLS.ExcelCommands;

namespace OfficeLib.XLS
{
    /// <summary>
    /// 
    /// </summary>
    public static class ExcelEditExtention
    {
        /// <summary>
        /// Set colors for cell
        /// </summary>
        /// <param name="excel"></param>
        /// <param name="target"></param>
        /// <param name="color"></param>
        /// <remarks>
        /// [2003] and [2007 ～] Check operation
        /// </remarks>
        public static void SetBackgroundColor(this Excel excel, Address target, Color color)
        {
            Object cells = null;
            try
            {
                cells = excel.Sheet.GetProperty(OBJECT_CELL,
                                                    new Object[] {
                                                        target.Row, target.Column
                                                    });
                cells.GetProperty(PROP_INTERIOR)
                     .SetProperty(PROP_COLOR, new Object[] { color.RGB() });
            }
            finally { OfficeCore.ReleaseObject(cells); }
        }

        /// <summary>
        /// Set colors for cell
        /// </summary>
        /// <param name="excel"></param>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <param name="color"></param>
        public static void SetBackgroundColor(this Excel excel, 
                                              Address start, Address end, Color color)
        {
            Object range = null;
            try
            {
                range = excel.GetRange(start.Row, start.Column, end.Row, end.Column);

                range.GetProperty(PROP_INTERIOR)
                     .SetProperty(PROP_COLOR, new Object[] { color.RGB() });
            }
            finally { OfficeCore.ReleaseObject(range); }
        }

        /// <summary>
        /// Set Border
        /// </summary>
        /// <param name="excel"></param>
        /// <param name="target"></param>
        /// <param name="thickness"></param>
        public static void SetBorder(this Excel excel, Range target, Thickness thickness)
        {   // Todo: Experimental content
            Object range = null;
            try
            {
                range = excel.GetRange(target.Start.Row, target.Start.Column,
                                       target.End.Row, target.End.Column);

                foreach (var thk in thickness.Values)
                {
                    range.GetProperty(PROP_BORDERS, new Object[] { thk.Index })
                         .SetProperty(PROP_LINE_STYLE, new Object[] { thk.LineStyle });
                    range.GetProperty(PROP_BORDERS, new Object[] { thk.Index })
                         .SetProperty(PROP_WEIGHT, new Object[] { thk.Weight });
                }
            }
            finally { OfficeCore.ReleaseObject(range); }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="excel"></param>
        /// <param name="target"></param>
        /// <param name="xlWeight"></param>
        /// <param name="xlStyle"></param>
        public static void SetBorder(this Excel excel, Address target,
                                     XlBorderWeight xlWeight, XlLineStyle xlStyle)
        {
            // Todo: Experimental content



        }
    }
}
