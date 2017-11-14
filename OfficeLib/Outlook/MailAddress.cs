using System;

namespace OfficeLib.EML
{
    /// <summary>
    /// E-Mail Address Class
    /// </summary>
    public class MailAddress
    {
        /// <summary>
        /// Display Name
        /// </summary>
        public String DisplayName { get; set; }

        /// <summary>
        /// Email Address
        /// </summary>
        public String Address { get; set; }

        /// <summary>
        /// Email Recipient Type
        /// </summary>
        public Outlook.OlMailRecipientType Type { get; set; }


        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="address">EMail Address</param>
        public MailAddress(String address)
        {
            this.Address = address;
            this.DisplayName = String.Empty;
            this.Type = Outlook.OlMailRecipientType.olOriginator;
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="address">EMail Address</param>
        /// <param name="name">Display Name</param>
        public MailAddress(String address, String name, Outlook.OlMailRecipientType olType)
        {
            this.Address = address;
            this.DisplayName = name;
            this.Type = olType;
        }

        /// <summary>
        /// To String
        /// </summary>
        public override String ToString()
        {
            return String.Format($"{this.DisplayName} <{this.Address}>");
        }
    }
}
