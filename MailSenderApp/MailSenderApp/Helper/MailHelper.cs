namespace MailSenderApp.Helper
{
    #region Namespaces
    using MailSenderApp.Utils;
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.IO;
    using System.Linq;
    using System.Threading;
    #endregion

    public static class MailHelper
    {
        public static void SendMail(string mailTo, string mailCc, string attachments, string subject, string mbody, ref List<string[]> reportData, bool spec = false)
        {
            if (!spec)
            {
                MailHelper.SendMail(
                    ConfigurationManager.AppSettings["SMTPServer"],
                    Convert.ToInt32(ConfigurationManager.AppSettings["SMTPServerPort"]),
                    ConfigurationManager.AppSettings["MailFromAddr"],
                    ConfigurationManager.AppSettings["MailPassword"],
                    ConfigurationManager.AppSettings["FromDisplayName"],
                    mailTo,
                    mailCc,
                    attachments,
                    subject,
                    mbody,
                    ref reportData);
            }
            else
            {
                MailHelper.SendMail(
                    ConfigurationManager.AppSettings["SMTPServer"],
                    Convert.ToInt32(ConfigurationManager.AppSettings["SMTPServerPort"]),
                    ConfigurationManager.AppSettings["SpecFrom"],
                    ConfigurationManager.AppSettings["SpecPwd"],
                    ConfigurationManager.AppSettings["SpecDisplay"],
                    mailTo,
                    mailCc,
                    attachments,
                    subject,
                    mbody,
                    ref reportData);
            }
        }

        public static bool ValidateAttachments(IEnumerable<string> attachments)
        {
            if (null == attachments)
            {
                return false;
            }

            int attCount = 0;
            foreach (var attachment in attachments)
            {
                if (File.Exists(attachment))
                {
                    attCount++;
                }
            }

            return attCount == attachments.Count();
        }
        
        private const int RETRY_COUNT = 3;
        private const int SLEEPING = 5000;
        private static void SendMail(
            string smtpServer,
            int smtpPort,
            string mailFrom,
            string mailPassword,
            string fromDisplayName,
            string mailTo,
            string mailCc,
            string attachments,
            string subject,
            string mbody,
            ref List<string[]> reportData)
        {
            int loop = MailHelper.RETRY_COUNT;
            var mailSender = new MailSender();
            mailSender.MailFromAddr = mailFrom;
            mailSender.MailFromPwd = mailPassword;
            mailSender.MailFromDisplayName = fromDisplayName;
            mailSender.MailTo = mailTo.ParseTextWithSemicolon();
            mailSender.MailCc = mailCc.ParseTextWithSemicolon();
            mailSender.MailSubject = subject;
            mailSender.MailHtmlBody = mbody;
            var attArr = attachments.ParseTextWithSemicolon();
            mailSender.SetLocalAttachments(attArr);

            bool succ = false;
            while (loop > 0)
            {
                succ = mailSender.Send(smtpServer, smtpPort);
                if (succ)
                {
                    break;
                }

                Thread.Sleep(MailHelper.SLEEPING);
                loop--;
            }

            if (succ)
            {
                var msg = String.Format("{0}, 已发送！", mailTo);
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine(msg);
                Console.ResetColor();
                reportData.Add(new string[] { mailTo, subject, ValidateAttachments(attArr) ? "有" : "无", "已发送" });
            }
            else
            {
                var msg = String.Format("{0}, 发送失败！", mailTo);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(msg);
                Console.ResetColor();
                reportData.Add(new string[] { mailTo, subject, ValidateAttachments(attArr) ? "有" : "无", "发送失败" });
            }
        }
    }
}
