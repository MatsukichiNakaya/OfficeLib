using System;

namespace OfficeLib.EML
{
    /// <summary>
    /// 
    /// </summary>
    public class MailAddress
    {
        /// <summary>
        /// 
        /// </summary>
        public String DisplayName { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public String Address { get; set; }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        public MailAddress(String address)
        {
            this.Address = address;
            this.DisplayName = String.Empty;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="name"></param>
        public MailAddress(String address, String name)
        {
            this.Address = address;
            this.DisplayName = name;
        }
    }
}
