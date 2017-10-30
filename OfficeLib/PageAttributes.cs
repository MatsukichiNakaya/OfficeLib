using System;

namespace OfficeLib
{
    /// <summary>Sheet access permission definition</summary>
    public enum EnumSheetPermission : Int32
    {
        /// <summary>Read only</summary>
        Read = 0x0001,              // (0001)
        /// <summary>Write only</summary>
        Write = Read << 1,          // (0010)
        /// <summary>Readable and writable</summary>
        ReadWrite = Read | Write,   // (0011)
    }

    /// <summary>
    /// Attribute settings for reading and writing
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false)]
    public class PageAttribute : Attribute
    {
        /// <summary>Access Permission</summary>
        public EnumSheetPermission Permission { get; private set; }

        /// <summary>Constructor</summary>
        public PageAttribute(EnumSheetPermission SheetParm)
        {
            this.Permission = SheetParm;
        }
    }
}
