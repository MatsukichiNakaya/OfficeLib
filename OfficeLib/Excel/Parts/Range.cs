using System;

namespace OfficeLib.XLS
{
    /// <summary></summary>
    public class Range
    {
        /// <summary></summary>
        public Address Start { get; set; }
        /// <summary></summary>
        public Address End { get; set; }

        /// <summary>Constructor</summary>
        /// <param name="rangeString"></param>
        public Range(String rangeString)
        {
            String[] addr = rangeString.Split(':');

            if (1 == addr.Length)
            {
                this.Start = addr[0].ToAddress();
                this.End = addr[0].ToAddress();
            }
            else if (2 == addr.Length)
            {
                this.Start = addr[0].ToAddress();
                this.End = addr[1].ToAddress();
            }
        }

        /// <summary>Constructor</summary>
        /// <param name="target"></param>
        public Range(Address target)
        {
            this.Start = target;
            this.End = target;
        }

        /// <summary>Constructor</summary>
        /// <param name="start"></param>
        /// <param name="end"></param>
        public Range(Address start, Address end)
        {
            this.Start = start;
            this.End = end;
        }

        /// <summary>
        /// Range to String
        /// </summary>
        public override String ToString()
        {
            if (this.Start.ToString().Equals(this.End.ToString()))
            {
                return this.Start.ToString();
            }
            return $"{this.Start.ToString()}:{this.End.ToString()}";
        }
    }
}
