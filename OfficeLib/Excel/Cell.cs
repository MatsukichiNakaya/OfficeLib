using System;

namespace OfficeLib.Excel
{
    /// <summary></summary>
    public class Cell<T>
    {
        /// <summary></summary>
        public T Value { get; set; }
        /// <summary></summary>
        public Thickness RuledLine { get; set; }
        /// <summary></summary>
        public Color BackgroundColor { get; set; }
        /// <summary></summary>
        public Color ForegroundColor { get; set; }
        /// <summary></summary>
        public String Fomula { get; set; }
        /// <summary></summary>
        public UInt32 FontSize { get; set; }

#if false
        // Todo :
        /// <summary></summary>
        public Boolean IsBold { get; set; }
        /// <summary></summary>
        public Boolean IsItalic { get; set; }
#endif
        /// <summary></summary>
        public Cell()
        {
            this.RuledLine = new Thickness();
        }
    }
}
