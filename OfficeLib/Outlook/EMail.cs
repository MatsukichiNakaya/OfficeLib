using System;

namespace OfficeLib.EML
{
    /// <summary>
    /// E-Mail Class
    /// </summary>
    public class EMail
    {
        /// <summary>
        /// 
        /// </summary>
        public String Subject { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public MailAddress From { get; set; }
        
        /// <summary>
        /// 
        /// </summary>
        public MailAddress[] To { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public MailAddress[] CC { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public MailAddress[] BCC { get; set; }
        
        /// <summary>
        /// 
        /// </summary>
        public DateTime SendDate { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public String Body { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public Boolean HasAttachements { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public EMail()
        {
            this.Subject = String.Empty;
            this.From = null;
            this.To = null;
            this.CC = null;
            this.BCC = null;
            this.Body = String.Empty;
            this.HasAttachements = false;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="mailItem"></param>
        public EMail(Object mailItem)
        {
            this.Subject = mailItem.GetProperty("Subject", null).ToString();
            // Todo : MailAddress ->
            this.From = new MailAddress(mailItem.GetProperty("", null).ToString());
            this.To = null;
            this.CC = null;
            this.BCC = null;
            // Todo : MailAddress <-

            this.Body = String.Empty;
            this.HasAttachements = false;
        }
    }
}
