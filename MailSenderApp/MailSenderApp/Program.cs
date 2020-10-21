namespace MailSenderApp
{
    #region Namespaces
    using MailSenderApp.Helper;
    using MailSenderApp.RuleEngine;
    using MailSenderApp.Tools;
    using MailSenderApp.Utils;
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.IO;
    using System.Linq;
    #endregion

    class Program
    {
        static readonly string[] MAIL_REPORT_FIELDS = new string[] { "Mail To", "Subject", "Has Attachments", "Status" };
        static readonly string[] RULE_REPORT_FIELDS = new string[] { "Rule ID", "Rule Description", "Passed", "Failed Message", "Failed Count", "Indexes/MsgTxts" };
        static List<string[]> mailReportData = new List<string[]>();
        static Dictionary<RuleBase, ValidationResult> ruleReportData;

        static void Main(string[] args)
        {
            if (args == null || args.Length == 0)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("缺少必要参数：至少需要一个参数，参数值为\"all\"或\"spec\"。"); // required parameter: "all" or "spec"
                Console.ResetColor();
                return;
            }

            if (args[0].ToLower() == "all")
            {
                var timeTracker = new TimeTracker();
                timeTracker.AddCheckpoint("checkpoint_01");
                JobForAll();
                timeTracker.AddCheckpoint("checkpoint_02");
                Console.WriteLine("执行所用时间：{0}秒\r\n", timeTracker.GetInterval("checkpoint_01", "checkpoint_02"));
            }
            else if (args[0].ToLower() == "spec")
            {
                var timeTracker = new TimeTracker();
                timeTracker.AddCheckpoint("checkpoint_01");
                JobForSpec();
                timeTracker.AddCheckpoint("checkpoint_02");
                Console.WriteLine("执行所用时间：{0}秒\r\n", timeTracker.GetInterval("checkpoint_01", "checkpoint_02"));
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("参数错误：参数值必须为\"all\"或\"spec\"。");
                Console.ResetColor();
            }
        }

        /// <summary>
        /// The job is for all the departments in the company.
        /// Process the personal information data and send mails.
        /// </summary>
        private static void JobForAll()
        {
            // The message will be replaced when new message is generated.
            string logMsg = "Processing, please wait..." + Constants.CRLF;
            logMsg += "1. Read data file;" + Constants.CRLF;
            logMsg += "2. Execute data cleaning;" + Constants.CRLF;
            logMsg += "3. Split data records by departments." + Constants.CRLF;
            Console.WriteLine(logMsg);

            string folderName = String.Empty;
            string errMsg = String.Empty;
            string mReportFile = String.Empty;
            string rReportFile = String.Empty;

            #region Read data file and execute the data cleaning
            var dataFile = Path.GetFullPath(ConfigurationManager.AppSettings["DataFile"]);
            IExcelManager excelMgr = new OXExcelManager();

            try
            {
                bool flag = FileHelper.CheckFileVersion(dataFile, out logMsg);
                Console.WriteLine(logMsg);
                errMsg += logMsg;
                if (!flag)
                {
                    goto End;
                }

                string sheetName = ConfigurationManager.AppSettings["DataFileValidSheetName"];

                #region Fixed WPS Official Issue By 'Resave' operation using Office COM component (Version: 15.0.0.0)
                // wps文件没有完全按照相关协议保存，导致OpenXML读取失败，故在此处进行预处理
                // Error Message: "数据包中不存在指定部件。"
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("正在对xlsx文件进行预处理...");
                Console.ResetColor();
                ExcelManager.ReSave(dataFile, sheetName);
                #endregion

                string[] dFields = excelMgr.ReadFields(dataFile, sheetName);
                string[][] dRows = excelMgr.Read(dataFile, sheetName);
                RecordHelper.DataCleaning(dFields, ref dRows, out ruleReportData);
                var ret = RecordHelper.SplitRecords(dFields, dRows);
                var outRoot = ConfigurationManager.AppSettings["Outputs"];
                folderName = DateTime.Now.ToString("yyyyMMddHHmm");
                var outputs = Path.Combine(outRoot, folderName);
                if (!Directory.Exists(outputs))
                {
                    Directory.CreateDirectory(outputs);
                }

                foreach (var key in ret.Keys)
                {
                    if (ret[key].Count > 0)
                    {
                        ret[key].Insert(0, dFields);
                        var splitedRecs = ret[key].ToArray();
                        excelMgr.Write(splitedRecs, Path.Combine(Path.GetFullPath(outputs), key + Constants.DataFileExtensionName));
                    }
                }
            }
            catch (Exception ex)
            {
                var curr = DateTime.Now.ToString(Constants.CurrDateTimePattern);
                logMsg = String.Format("[{0}]数据清洗或拆分异常，详细信息如下：\r\n{1}", curr, ex.Message);
                Console.WriteLine(errMsg);
                errMsg += logMsg;
                
                goto End;
            }
            finally
            {
                excelMgr.Close();
            }

            #region Generate data cleaning report
            logMsg = "Generate data cleanning report...";
            Console.WriteLine(logMsg);
            rReportFile = ReportHelper.GenerateRReport("数据清洗结果报告", RULE_REPORT_FIELDS, ruleReportData);
            logMsg = String.Format("The data cleaning report has been saved at '{0}'", rReportFile);
            Console.WriteLine(logMsg);
            #endregion

            #endregion

            #region Sending mails according to the mail queque config file
            string[][] rows = null;
            excelMgr = new OXExcelManager();
            try
            {
                var mailQueueFile = Path.Combine(Path.GetFullPath(ConfigurationManager.AppSettings["DataDir"]), ConfigurationManager.AppSettings["MailConfigTemplate"]);
                string[] fields = excelMgr.ReadFields(mailQueueFile, ConfigurationManager.AppSettings["MailQueueSheetName"]);
                if (fields == null)
                {
                    throw new Exception("邮件队列配置文件路径或名称不正确。");
                }
                var mappings = new Dictionary<string, int>();
                for (int i = 0; i < fields.Length; i++)
                {
                    mappings.Add(fields[i], i);
                }

                rows = excelMgr.Read(mailQueueFile, ConfigurationManager.AppSettings["MailQueueSheetName"]);
                var currDate = String.Format("{0}月{1}日", DateTime.Now.Month, DateTime.Now.Day);
                var currTime = String.Format("{0}点", DateTime.Now.Hour);
                for (int i = 0; i < rows.Length; i++)
                {
                    rows[i][mappings["Subject"]] = rows[i][mappings["Subject"]].Replace("{$DATE$}", currDate);
                    rows[i][mappings["Body"]] = rows[i][mappings["Body"]].Replace("{$DATE$}", currDate).Replace("{$TIME$}", currTime);
                    rows[i][mappings["AttachmentPath"]] = rows[i][mappings["AttachmentPath"]].Replace("{$FOLDER$}", folderName);
                }
            }
            catch (Exception ex)
            {
                var curr = DateTime.Now.ToString(Constants.CurrDateTimePattern);
                logMsg = String.Format("[{0}]邮件队列配置异常，详细信息如下：\r\n{1}", curr, ex.Message);
                Console.WriteLine(logMsg);
                errMsg += logMsg;
                goto End;
            }
            finally
            {
                excelMgr.Close();
            }

            logMsg = Constants.SendingMailMessageHeader;
            Console.WriteLine(logMsg);
            foreach (var row in rows)
            {
                var mTo = row[0];
                var mCC = string.Empty;
                var mSubject = row[1];
                var mBody = row[2];
                var mAttach = row[3];

                #region #TODO: Try to move this part to the file MailHelper.cs
                if (MailHelper.ValidateAttachments(mAttach.ParseTextWithSemicolon()))
                {
                    MailHelper.SendMail(mTo, mCC, mAttach, mSubject, mBody, out mailReportData);
                }
                else
                {
                    // Cancel to send mail, if there is no attachment.
                    logMsg = String.Format("{0}, 取消发送！", mTo);
                    Console.WriteLine(logMsg);
                    mailReportData.Add(new string[] { mTo, mSubject, "无", "取消发送" });
                }
                #endregion
            }

            // Logs the successful information
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(Constants.CRLF + "Done!");
            Console.ResetColor();

            #region Generate the mail sending report
            logMsg = "Generating the mail sending report...";
            Console.WriteLine(logMsg);
            mReportFile = ReportHelper.GenerateMReport("邮件发送报告", MAIL_REPORT_FIELDS, mailReportData.ToArray());
            logMsg = String.Format("The sending mail report has been saved at '{0}'", mReportFile);
            Console.WriteLine(logMsg);
            #endregion

            #endregion

            #region Sends confirm mail
            End:
            // Set contents for reporting mail.
            string resSubject = ConfigurationManager.AppSettings["ConfirmMailSubjectForAll"];
            string resAttachments = String.IsNullOrEmpty(rReportFile) ? mReportFile : String.Format("{0};{1}", rReportFile, mReportFile);
            string resBody = string.IsNullOrEmpty(resAttachments) ? string.Empty : "请查看附件以了解更多信息。" + Constants.CRLF;
            resBody += string.IsNullOrEmpty(errMsg) ? string.Empty : "<br>程序执行出现异常，邮件取消发送，详细信息如下：<br>" + errMsg; // use the parameter "errMsg" as the main part of mail body

            logMsg = "确认邮件发送结果：" + Constants.CRLF + Constants.SendingMailMessageHeader;
            Console.WriteLine(logMsg);
            MailHelper.SendMail(ConfigurationManager.AppSettings["ResponseTo"], String.Empty, resAttachments, resSubject, resBody, out mailReportData, spec: true);
            #endregion
        }

        /// <summary>
        /// The job is for office in the company.
        /// Process the personal information data and send mail to the office.
        /// </summary>
        private static void JobForSpec()
        {
            string logMsg = "Processing, please wait...";
            logMsg += "1. Read data file;" + Constants.CRLF;
            logMsg += "2. Execute data cleaning;" + Constants.CRLF;
            logMsg += "3. Execute macro and save as a new data file base on the template." + Constants.CRLF;
            Console.WriteLine(logMsg);

            string errMsg = String.Empty;
            string rReportFile = String.Empty;
            string fileName = String.Empty;
            string outputFile = String.Empty;

            #region Read data file and execute the data cleaning
            var dataFile = Path.GetFullPath(ConfigurationManager.AppSettings["DataFile"]);
            IExcelManager excelDataMgr = new OXExcelManager();
            try
            {
                bool flag = FileHelper.CheckFileVersion(dataFile, out logMsg);
                Console.WriteLine(logMsg);
                errMsg += logMsg;
                if (!flag)
                {
                    goto End;
                }

                string sheetName = ConfigurationManager.AppSettings["DataFileValidSheetName"];

                #region Fixed WPS Official Issue By 'Resave' operation using Office COM component (Version: 15.0.0.0)
                // wps文件没有完全按照相关协议保存，导致OpenXML读取失败，故在此处进行预处理
                // Error Message: "数据包中不存在指定部件。"
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("正在对xlsx文件进行预处理...");
                Console.ResetColor();
                ExcelManager.ReSave(dataFile, sheetName);
                #endregion

                string[] dFields = excelDataMgr.ReadFields(dataFile, sheetName);
                string[][] dRows = excelDataMgr.Read(dataFile, sheetName);
                RecordHelper.DataCleaning(dFields, ref dRows, out ruleReportData); // Execute data cleaning

                var outRoot = ConfigurationManager.AppSettings["Outputs"];
                fileName = String.Format(ConfigurationManager.AppSettings["DataStatisticsFilePattern"], DateTime.Now.ToString("yyyyMMdd"));
                outputFile = Path.Combine(outRoot, fileName + Constants.DataFileExtensionName);
                var xlsxTempl = Path.Combine(ConfigurationManager.AppSettings["DataDir"], ConfigurationManager.AppSettings["XlSXTemplate"]);
                File.Copy(xlsxTempl, outputFile, overwrite: true);

                var records = dRows.ToList();
                records.Insert(0, dFields);
                excelDataMgr.Insert(records.ToArray(), sheetId: 1, srcFile: outputFile);
                //excelDataMgr.Insert(records.ToArray(), sheetName: "明细数据", srcFile: outputFile);
            }
            catch (Exception ex)
            {
                var curr = DateTime.Now.ToString(Constants.CurrDateTimePattern);
                logMsg = String.Format("[{0}] 数据清洗异常，详细信息如下：\r\n{1}", curr, ex.Message);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(logMsg);
                Console.ResetColor();
                errMsg += logMsg;

                goto End;
            }
            finally
            {
                excelDataMgr.Close();
            }

            #region Generate data cleaning report
            logMsg = "Generate data cleanning report...";
            Console.WriteLine(logMsg);
            rReportFile = ReportHelper.GenerateRReport("数据清洗结果报告", RULE_REPORT_FIELDS, ruleReportData);
            logMsg = String.Format("The data cleaning report has been saved at '{0}'", rReportFile);
            Console.WriteLine(logMsg);
            #endregion

            #endregion

            #region Sending mail to office
            try
            {
                string subject = fileName;
                string body = String.Format("{0}：<br>&nbsp;&nbsp;&nbsp;&nbsp;{1}<br>", ConfigurationManager.AppSettings["SpecBody_RecieverNames"], ConfigurationManager.AppSettings["SpecBody_Content"]);
                MailHelper.SendMail(ConfigurationManager.AppSettings["SpecTo"], String.Empty, outputFile, subject, body, out mailReportData, spec: true);
            }
            catch (Exception ex)
            {
                var curr = DateTime.Now.ToString(Constants.CurrDateTimePattern);
                errMsg = String.Format("[{0}] 邮件发送失败，详细信息如下：\r\n{1}", curr, ex.Message);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(errMsg);
                Console.ResetColor();

                goto End;
            }

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(Constants.CRLF + "Done!");
            Console.ResetColor();
            #endregion

            #region Sends confirm mail
            End:
            var resSubject = ConfigurationManager.AppSettings["ConfirmMailSubjectForSpec"];
            var resBody = String.IsNullOrEmpty(errMsg) ? "发送成功！" : errMsg;

            logMsg = "确认邮件发送结果：" + Constants.CRLF + Constants.SendingMailMessageHeader;
            Console.WriteLine(logMsg);
            MailHelper.SendMail(ConfigurationManager.AppSettings["ResponseTo"], String.Empty, rReportFile, resSubject, resBody, out mailReportData, spec: true);
            #endregion
        }
    }
}
