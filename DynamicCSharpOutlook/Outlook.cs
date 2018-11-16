using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace DynamicCSharpOutlook
{
    public class Outlook
    {
        private readonly OutlookOptions _options;
        private dynamic _outlookObj;

        /// <summary>
        /// Version information, use for logging purposes in case of errors
        /// </summary>
        public string Version => $"{_outlookObj.Name} Version {_outlookObj.Version}";


        public Outlook(OutlookOptions options)
        {
            _options = options;

            Type outlookType = Type.GetTypeFromProgID("Outlook.Application", true);
            _outlookObj = Activator.CreateInstance(outlookType);

        }

        /// <summary>
        /// https://www.codeproject.com/Tips/165548/Csharp-Code-Snippet-to-Send-an-Email-with-Attachme
        /// </summary>
        public void SendMail()
        {
            // Create Mail item
            // https://docs.microsoft.com/en-us/office/vba/api/outlook.mailitem
            dynamic mailItemObj = _outlookObj.CreateItem(0);

            // Add recipients
            dynamic recepientsObj = mailItemObj.Recipients;

            foreach (var recipient in _options.Recipients)
            {
                dynamic receipientObj = recepientsObj.Add(recipient);
                receipientObj.Resolve(); // this one resolves the name like clicking on the Check Names
            }

            mailItemObj.Display(); // displaying it first lets Outlook get the signature
            dynamic signatureObj = mailItemObj.HTMLBody; // get the signature from HTML body to be appended later

            if (!string.IsNullOrWhiteSpace(_options.Cc))
                mailItemObj.CC = _options.Cc;

            if (!string.IsNullOrWhiteSpace(_options.Bcc))
                mailItemObj.BCC = _options.Bcc;

            mailItemObj.Subject = _options.Subject;
            mailItemObj.HTMLBody = _options.Body + signatureObj;
            mailItemObj.Importance = _options.Importance;

            //Add attachments
            dynamic attachmentsObj = mailItemObj.Attachments;

            foreach (var attachment in _options.Attachments)
            {
                attachmentsObj.Add(attachment, Type.Missing, mailItemObj.Body.Length + 1, Path.GetFileName(attachment));
            }

            //This is to send
            mailItemObj.Send();

        }
    }

    public class OutlookOptions
    {
        public List<string> Recipients { get; set; }
        public string Subject { get; set; }
        public List<string> Attachments { get; set; }
        public string Body { get; set; }
        public string Cc { get; set; }
        public string Bcc { get; set; }
        /// <summary>
        /// 0 - Low; 1 - Normal (Default); 2 - High
        /// </summary>
        public int Importance { get; set; } = 1;

        public OutlookOptions()
        {
            Recipients = new List<string>();
            Attachments = new List<string>();
        }
    }
}
