using System;

namespace OfficeLib.XLS
{
    /// <summary>
    /// 
    /// </summary>
    public class SheetObject : IDisposable
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheet"></param>
        public SheetObject(Object sheet)
        {
            this.ComObject = sheet;
        }

        /// <summary>
        /// 
        /// </summary>
        public Object ComObject { get; private set; }

        /// <summary>
        /// 
        /// </summary>
        public void Dispose()
        {
            OfficeCore.ReleaseObject(this.ComObject);
        }
    }
}
