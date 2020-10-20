namespace MailSenderApp.Helper
{
    #region Namespaces
    using System;
    using System.IO;
    #endregion

    public static class FileHelper
    {
        public static bool CheckFileVersion(string fpath, out string message)
        {
            var minutes = FileHelper.GetModifiedDuration(fpath);
            if (minutes > FileHelper.THRESHOLD)
            {
                var curr = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                message = String.Format("[{0}] 文件版本异常：该文件创建时间已逾{1}分钟，请重新下载！\r\n", curr, minutes);

                return false;
            }
            else
            {
                message = String.Format("文件是最新的，该文件的创建（更新）时间在{0}分之内。", minutes);

                return true;
            }
        }

        private const int THRESHOLD = 30;

        private static double GetModifiedDuration(string fpath)
        {
            if (!File.Exists(fpath))
            {
                throw new FileNotFoundException(String.Format("The file '' is not found.", fpath));
            }

            FileInfo fi = new FileInfo(fpath);
            var lastUpdated = fi.LastWriteTime;
            var curr = DateTime.Now;
            var deltaMinutes = (curr - lastUpdated).TotalMinutes;

            return deltaMinutes;
        }
    }
}
