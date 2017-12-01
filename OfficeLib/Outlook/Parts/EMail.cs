using System;
using System.Linq;

namespace OfficeLib.EML
{
    /// <summary>
    /// E-Mail Class
    /// </summary>
    public class EMail
    {
        /// <summary>Subject</summary>
        private const String PROP_SUBJECT = "Subject";

        /// <summary>EMail Address</summary>
        private const String PROP_ADDRESS = "Address";
        /// <summary>EMail Name</summary>
        private const String PROP_NAME = "Name";
        /// <summary>Recipient Type</summary>
        private const String PROP_TYPE = "Type";

        /// <summary>Count</summary>
        private const String PROP_COUNT = "Count";
        /// <summary>Items</summary>
        private const String PROP_ITEMS = "Items";
        /// <summary>Item</summary>
        private const String PROP_ITEM = "Item";

        /// <summary>From</summary>
        private const String PROP_SENDER = "Sender";
        /// <summary>Recipients</summary>
        private const String PROP_RECIPIENTS = "Recipients";

        /// <summary>Body</summary>
        private const String PROP_BODY = "Body";
        /// <summary>Attachement</summary>
        private const String PROP_ATTACHEMENTS = "Attachments";



        /// <summary>
        /// Subject
        /// </summary>
        public String Subject { get; set; }

        /// <summary>
        /// From
        /// </summary>
        public MailAddress From { get; set; }

        /// <summary>
        /// To
        /// </summary>
        public MailAddress[] To { get; set; }

        /// <summary>
        /// CC
        /// </summary>
        public MailAddress[] CC { get; set; }

        /// <summary>
        /// BCC
        /// </summary>
        public MailAddress[] BCC { get; set; }

        /// <summary>
        /// Send Date
        /// </summary>
        public DateTime SendDate { get; set; }

        /// <summary>
        /// Body
        /// </summary>
        public String Body { get; set; }

        /// <summary>
        /// Has Attachements
        /// </summary>
        public Boolean HasAttachements { get; set; }

        /// <summary>
        /// Constructor
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
        /// Constructor
        /// </summary>
        /// <param name="mailItem">Mail object</param>
        public EMail(Object mailItem)
        {
            this.Subject = mailItem.GetProperty(PROP_SUBJECT).ToString();

            this.From = new MailAddress(mailItem.GetProperty(PROP_SENDER).GetProperty(PROP_ADDRESS).ToString(),
                                        mailItem.GetProperty(PROP_SENDER).GetProperty(PROP_NAME).ToString(),
                                        Outlook.OlMailRecipientType.olOriginator);

            MailAddress[] recipients = GetMailAdds(mailItem.GetProperty(PROP_RECIPIENTS));
            this.To = recipients.Where(r => r.Type == Outlook.OlMailRecipientType.olTo).ToArray();
            this.CC = recipients.Where(r => r.Type == Outlook.OlMailRecipientType.olCC).ToArray();
            this.BCC = recipients.Where(r => r.Type == Outlook.OlMailRecipientType.olBCC).ToArray();

            this.Body = mailItem.GetProperty(PROP_BODY).ToString();
            this.HasAttachements = 0 < mailItem.GetProperty(PROP_ATTACHEMENTS)
                                               .GetProperty(PROP_COUNT).To<Int32>();
        }

        /// <summary>
        /// Get Mail Address from mail object
        /// </summary>
        /// <param name="mailAdds">Mail object</param>
        private MailAddress[] GetMailAdds(Object mailAdds)
        {
            var result = new MailAddress[mailAdds.GetProperty(PROP_COUNT).To<Int32>()];

            for (var i = 0; i < result.Length; i++)
            {
                result[i] = new MailAddress(mailAdds.GetProperty(PROP_ITEM, new Object[] { i + 1 })
                                                    .GetProperty(PROP_ADDRESS).ToString(),
                                            mailAdds.GetProperty(PROP_ITEM, new Object[] { i + 1 })
                                                    .GetProperty(PROP_NAME).ToString(),
                                            (Outlook.OlMailRecipientType)mailAdds.GetProperty(PROP_ITEM, new Object[] { i + 1 })
                                                                                 .GetProperty(PROP_TYPE).To<Int32>());
            }
            return result;
        }

        /// <summary>
        /// To String
        /// </summary>
        public override String ToString()
        {
            return new { Sub = this.Subject,
                         Attache = this.HasAttachements }.ToString();
        }
    }
}
