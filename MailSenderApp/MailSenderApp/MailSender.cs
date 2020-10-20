namespace MailSenderApp
{
    #region Namespaces.
    using MailSenderApp.Utils;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Net;
    using System.Net.Mail;
    using System.Net.Mime;
    using System.Net.Security;
    using System.Security.Cryptography.X509Certificates;
    using System.Text;
    using System.Text.RegularExpressions;
    #endregion

    public class MailSender
    {
        public MailSender()
        {
        }

        public bool Send(string host, int port)
        {
            MailMessage msg = this.BuildMailMessage();
            if (null == msg)
            {
                return false;
            }

            try
            {
                SmtpClient client = new SmtpClient();
                client.Timeout = 150000;
                client.Host = host;
                client.Port = port;
                client.EnableSsl = true;
                client.UseDefaultCredentials = false;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.Credentials = new NetworkCredential(this.MailFromAddr, this.MailFromPwd);
                ServicePointManager.ServerCertificateValidationCallback = delegate (Object obj, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors) { return true; };
                client.Send(msg);

                return true;
            }
            catch(Exception ex)
            {
                return false;
            }
        }

        public void SendAsync(string host, int port, SendCompletedEventHandler completedEvent)
        {
            MailMessage msg = this.BuildMailMessage();
            if (null == msg)
            {
                return;
            }

            SmtpClient client = new SmtpClient();
            //client.Timeout = 100000;
            client.Host = host;
            client.Port = port;
            client.EnableSsl = true;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.UseDefaultCredentials = false;
            client.Credentials = new NetworkCredential(this.MailFromAddr, this.MailFromPwd);
            client.SendCompleted += new SendCompletedEventHandler(completedEvent);

            client.SendAsync(msg, msg.Body);
        }

        public MailMessage BuildMailMessage()
        {
            if (!this.Validate())
            {
                return null;
            }

            MailMessage msg = new MailMessage();
            msg.Priority = this.MailPriority;
            msg.From = new MailAddress(this.MailFromAddr, this.MailFromDisplayName);
            foreach (var mailTo in this.MailTo)
            {
                msg.To.Add(mailTo);
            }

            if (null != this.MailCc)
            {
                foreach (var mailCc in this.MailCc)
                {
                    msg.CC.Add(mailCc);
                }
            }

            if (null != this.Attachments)
            {
                foreach (var attachment in this.Attachments)
                {
                    msg.Attachments.Add(attachment);
                }
            }

            msg.SubjectEncoding = Encoding.UTF8;
            msg.Subject = this.MailSubject;

            msg.IsBodyHtml = true;
            msg.BodyEncoding = Encoding.UTF8;
            msg.Body = this.MailHtmlBody;

            return msg;
        }

        public string MailFromAddr { get; set; }
        public string MailFromPwd { get; set; }
        public string MailFromDisplayName { get; set; }
        public IEnumerable<string> MailTo { get; set; }
        public IEnumerable<string> MailCc { get; set; }
        public IEnumerable<Attachment> Attachments { get; private set; }
        public string MailSubject { get; set; }
        public string MailHtmlBody { get; set; }
        public MailPriority MailPriority { get; set; }

        public void SetLocalAttachments(IEnumerable<string> attachmentFiles)
        {
            if (null == attachmentFiles)
            {
                return;
            }

            var attachments = new List<Attachment>();
            foreach (var file in attachmentFiles)
            {
                if (File.Exists(file))
                {
                    string exName = file.GetLastSegment('.').ToLower();
                    attachments.Add(new Attachment(file, this.GenMediaType(exName)));
                }
            }

            this.Attachments = attachments;
        }

        private bool Validate()
        {
            if (string.IsNullOrEmpty(this.MailFromAddr) || !this.ValidateMailAddress(this.MailFromAddr))
            {
                return false;
            }

            if (null == this.MailFromPwd)
            {
                return false;
            }

            if (null == this.MailTo || this.MailTo.Count() == 0 || !this.ValidateMailAddresses(this.MailTo))
            {
                return false;
            }

            if (null != this.MailCc && this.MailCc.Count() != 0 && !this.ValidateMailAddresses(this.MailCc))
            {
                return false;
            }

            if (string.IsNullOrEmpty(this.MailFromDisplayName))
            {
                this.MailFromDisplayName = this.MailFromAddr;
            }

            return true;
        }

        private bool ValidateMailAddresses(IEnumerable<string> mailAddresses)
        {
            if (null == mailAddresses)
            {
                return false;
            }

            foreach (var address in mailAddresses)
            {
                if (!this.ValidateMailAddress(address))
                {
                    return false;
                }
            }

            return true;
        }

        private bool ValidateMailAddress(string MailAddress)
        {
            Regex r = new Regex("^\\s*([A-Za-z0-9_-]+(\\.\\w+)*@(\\w+\\.)+\\w{2,5})\\s*$");

            if (r.IsMatch(MailAddress))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private string GenMediaType(string fileExtensionName)
        {
            string mediaType = MediaTypeNames.Application.Octet;
            if (string.IsNullOrEmpty(fileExtensionName))
            {
                return mediaType;
            }

            if ("zip" == fileExtensionName || "rar" == fileExtensionName)
            {
                mediaType = MediaTypeNames.Application.Zip;
            }
            else if ("pdf" == fileExtensionName)
            {
                mediaType = MediaTypeNames.Application.Pdf;
            }
            else if ("xml" == fileExtensionName)
            {
                mediaType = MediaTypeNames.Text.Xml;
            }
            else if ("gif" == fileExtensionName)
            {
                mediaType = MediaTypeNames.Image.Gif;
            }
            else if ("jpg" == fileExtensionName || "jpeg" == fileExtensionName)
            {
                mediaType = MediaTypeNames.Image.Jpeg;
            }
            else if ("png" == fileExtensionName)
            {
                mediaType = "image/png";
            }
            else if ("html" == fileExtensionName)
            {
                mediaType = MediaTypeNames.Text.Html;
            }
            else if ("txt" == fileExtensionName)
            {
                mediaType = MediaTypeNames.Text.Plain;
            }

            return mediaType;
        }

        private void GetFolderAndFileName(string address, out string folderName, out string fileName)
        {
            folderName = string.Empty;
            fileName = string.Empty;
            int lastIndex = address.LastIndexOf('/');
            if (lastIndex >= 0)
            {
                fileName = address.Substring(lastIndex + 1);
                address = address.Remove(lastIndex);
            }

            lastIndex = address.LastIndexOf('/');
            if (lastIndex >= 0)
            {
                folderName = address.Substring(lastIndex + 1);
            }
        }
    }
}
