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
    using System.Threading;
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
                Console.WriteLine("执行所用时间：{0}\r\n", timeTracker.GetInterval("checkpoint_01", "checkpoint_02"));
            }
            else if (args[0].ToLower() == "spec")
            {
                var timeTracker = new TimeTracker();
                timeTracker.AddCheckpoint("checkpoint_01");
                JobForSpec();
                timeTracker.AddCheckpoint("checkpoint_02");
                Console.WriteLine("执行所用时间：{0}\r\n", timeTracker.GetInterval("checkpoint_01", "checkpoint_02"));
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("参数错误：参数值必须为\"all\"或\"spec\"。");
                Console.ResetColor();
            }
        }

        private static double CheckDataFile(string fpath)
        {
            FileInfo fi = new FileInfo(fpath);
            var lastUpdated = fi.LastWriteTime;
            var curr = DateTime.Now;
            var deltaMinutes = (curr - lastUpdated).TotalMinutes;

            return deltaMinutes;
        }

        private static void JobForAll()
        {
            string msg = "Processing, please wait...";
            Console.WriteLine(msg);
            string folderName = String.Empty;
            string failureInfo = String.Empty;
            string rReportFile = String.Empty;
            string mReportFile = String.Empty;

            var dataFile = Path.GetFullPath(ConfigurationManager.AppSettings["DataFile"]);
            IExcelManager excelDataMgr = new OXExcelManager();
            //IExcelManager excelDataMgr = new ExcelManager();
            try
            {
                bool flag = FileHelper.CheckFileVersion(dataFile, out failureInfo);
                Console.WriteLine(failureInfo);
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
                        excelDataMgr.Write(splitedRecs, Path.Combine(Path.GetFullPath(outputs), key + ".xlsx"));
                    }
                }

                excelDataMgr.Close();
            }
            catch (Exception ex)
            {
                excelDataMgr.Close();
                var curr = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                failureInfo = String.Format("[{0}]数据清洗或拆分异常，详细信息如下：\r\n{1}", curr, ex.Message);
                Console.WriteLine(failureInfo);
                goto End;
            }

            msg = "Generate data cleanning report...";
            Console.WriteLine(msg);
            rReportFile = ReportHelper.GenerateRReport("数据清洗结果报告", RULE_REPORT_FIELDS, ruleReportData);

            string[][] rows = null;
            IExcelManager excelMgr = new OXExcelManager();
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
                rows = excelMgr.Read(mailQueueFile, "mail_queue");
                excelMgr.Close();
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
                excelMgr.Close();
                var curr = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                failureInfo = String.Format("[{0}]邮件队列配置异常，详细信息如下：\r\n{1}", curr, ex.Message);
                Console.WriteLine(failureInfo);
                goto End;
            }

            msg = "\r\nMail To Address, Status";
            Console.WriteLine(msg);
            foreach (var row in rows)
            {
                var mTo = row[0];
                var mCC = string.Empty;
                var mSubject = row[1];
                var mBody = row[2];
                var mAttach = row[3];
                if (MailHelper.ValidateAttachments(mAttach.ParseTextWithSemicolon()))
                {
                    MailHelper.SendMail(mTo, mCC, mAttach, mSubject, mBody, out mailReportData);
                }
                else
                {
                    // Cancel to send mail, if there is no attachment.
                    msg = String.Format("{0}, 取消发送！", mTo);
                    Console.WriteLine(msg);
                    mailReportData.Add(new string[] { mTo, mSubject, "无", "取消发送" });
                }
            }

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("\nDone!");
            Console.ResetColor();
            Console.WriteLine("Generating the report...");
            Thread.Sleep(1000);
            mReportFile = ReportHelper.GenerateMReport("邮件发送报告", MAIL_REPORT_FIELDS, mailReportData.ToArray());


            End:
            // Set contents for reporting mail.
            string resSubject = "疫情人员信息收集邮件发送报告 - To:所有处室和分支机构";
            string resAttachments = String.IsNullOrEmpty(rReportFile) ? mReportFile : String.Format("{0};{1}", rReportFile, mReportFile);
            string resBody = string.IsNullOrEmpty(resAttachments) ? string.Empty : "请查看附件以了解更多信息。";
            resBody += string.IsNullOrEmpty(failureInfo) ? string.Empty : "<br>取消发送，详细信息如下：<br>" + failureInfo;

            msg = "\r\nMail To Address, Status";
            Console.WriteLine(msg);
            MailHelper.SendMail(ConfigurationManager.AppSettings["ResponseTo"], "", resAttachments, resSubject, resBody, out mailReportData, spec: true);
        }

        private static void JobForSpec()
        {
            string msg = "Processing, please wait...";
            Console.WriteLine(msg);

            string failureInfo = String.Empty;
            string rReportFile = String.Empty;
            string fileName = String.Empty;
            string outputFile = String.Empty;

            var dataFile = Path.GetFullPath(ConfigurationManager.AppSettings["DataFile"]);
            IExcelManager excelDataMgr = new OXExcelManager();
            try
            {
                bool flag = FileHelper.CheckFileVersion(dataFile, out failureInfo);
                Console.WriteLine(failureInfo);
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
                fileName = String.Format("中国信保疫情期间人员信息统计表({0})", DateTime.Now.ToString("yyyyMMdd"));
                outputFile = Path.Combine(outRoot, fileName + ".xlsx");
                var xlsxTempl = Path.Combine(ConfigurationManager.AppSettings["DataDir"], ConfigurationManager.AppSettings["XlSXTemplate"]);
                File.Copy(xlsxTempl, outputFile, overwrite: true);

                var records = dRows.ToList();
                records.Insert(0, dFields);
                excelDataMgr.Insert(records.ToArray(), sheetId: 1, srcFile: outputFile);
                //excelDataMgr.Insert(records.ToArray(), sheetName: "明细数据", srcFile: outputFile);

                excelDataMgr.Close();
            }
            catch (Exception ex)
            {
                excelDataMgr.Close();
                var curr = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                failureInfo = String.Format("[{0}] 数据清洗异常，详细信息如下：\r\n{1}", curr, ex.Message);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(failureInfo);
                Console.ResetColor();
                goto End;
            }

            msg = "Generate data cleanning report...";
            Console.WriteLine(msg);
            rReportFile = ReportHelper.GenerateRReport("数据清洗结果报告", RULE_REPORT_FIELDS, ruleReportData);

            try
            {
                string subject = fileName;
                string body = "梅处、袁帅：<br>&nbsp;&nbsp;&nbsp;&nbsp;今天的数据，请查收。<br>";
                MailHelper.SendMail(ConfigurationManager.AppSettings["SpecTo"], String.Empty, outputFile, subject, body, out mailReportData, spec: true);
            }
            catch (Exception ex)
            {
                var curr = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                failureInfo = String.Format("[{0}] 邮件发送失败，详细信息如下：\r\n{1}", curr, ex.Message);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(failureInfo);
                Console.ResetColor();
                goto End;
            }

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("\nDone!");
            Console.ResetColor();

            End:
            var resSubject = "疫情人员信息收集邮件发送报告 - To:办公室";
            var resBody = String.IsNullOrEmpty(failureInfo) ? "发送成功！" : failureInfo;

            msg = "\r\nMail To Address, Status";
            Console.WriteLine(msg);
            MailHelper.SendMail(ConfigurationManager.AppSettings["ResponseTo"], String.Empty, rReportFile, resSubject, resBody, out mailReportData, spec: true);
        }
    }
}
