namespace MailSenderApp.Helper
{
    #region Namespaces
    using MailSenderApp.Utils;
    using System;
    using System.Configuration;
    using System.IO;
    #endregion

    public static class FileHelper
    {
        public static bool CheckFileVersion(string fpath, out string message)
        {
            int threshold = Convert.ToInt32(ConfigurationManager.AppSettings["DataFileModifiedWithin"]);
            var minutes = FileHelper.GetModifiedDuration(fpath);
            if (minutes > threshold)
            {
                var curr = DateTime.Now.ToString(Constants.CurrDateTimePattern);
                message = String.Format("[{0}] 文件版本异常：该文件创建时间已逾{1}分钟，请重新下载！" + Constants.CRLF, curr, minutes);

                return false;
            }
            else
            {
                message = String.Format("文件是最新的，该文件的创建（更新）时间在{0}分之内。" + Constants.CRLF, minutes);

                return true;
            }
        }

        private static double GetModifiedDuration(string fpath)
        {
            if (!File.Exists(fpath))
            {
                throw new FileNotFoundException(String.Format("The file '{0}' is not found.", fpath));
            }

            FileInfo fi = new FileInfo(fpath);
            var lastUpdated = fi.LastWriteTime;
            var curr = DateTime.Now;
            var deltaMinutes = (curr - lastUpdated).TotalMinutes;

            return deltaMinutes;
        }
    }
}
