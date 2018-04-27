using System;

namespace OfficeLib.XLS
{
    /// <summary></summary>
    internal class Cell<T>
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
        /// <summary></summary>
        public XlGetValueFormat FormatType { get; set; }

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
            this.Value = default(T);
            this.RuledLine = new Thickness();
            this.Fomula = String.Empty;
            this.FormatType = XlGetValueFormat.xlValue2; 
        }
    }
}
