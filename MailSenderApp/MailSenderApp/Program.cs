namespace MailSenderApp
{
    #region Namespaces
    using MailSenderApp.RuleEngine;
    using MailSenderApp.Tools;
    using MailSenderApp.Utils;
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Threading;
    #endregion

    class Program
    {
        static readonly string[] reportFields = new string[] { "Mail To", "Subject", "Has Attachments", "Status" };
        static readonly string[] ruleReportFields = new string[] { "Rule ID", "Rule Description", "Passed", "Failed Message", "Failed Count", "Indexes/MsgTxts" };
        static List<string[]> reportData = new List<string[]>();
        static Dictionary<RuleBase, ValidationResult> ruleExecResults;

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
                var minutes = CheckDataFile(dataFile);
                if (minutes > 30)
                {
                    var curr = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                    failureInfo = String.Format("[{0}]文件版本异常：该文件创建时间已逾{1}分钟，请重新下载！\r\n", curr, minutes);
                    Console.WriteLine(failureInfo);
                    goto End;
                }

                Console.WriteLine("文件是最新的，创建/更新时间在{0}分之内", minutes);

                //#region Fixed WPS Official Issue
                //// wps文件没有完全按照相关协议保存，导致OpenXML读取失败，故在此处进行预处理
                //Console.ForegroundColor = ConsoleColor.Yellow;
                //Console.WriteLine("正在对xlsx文件进行预处理...");
                //Console.ResetColor();
                //string sheetName = ConfigurationManager.AppSettings["DataFileValidSheetName"];
                //ExcelManager.ReSave(dataFile, sheetName);
                //#endregion


                string[] dFields = excelDataMgr.ReadFields(dataFile, sheetName);
                string[][] dRows = excelDataMgr.Read(dataFile, sheetName);
                ProcessDataCleaning(dFields, ref dRows, out ruleExecResults);
                var ret = SplitRecords(dFields, dRows);
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
            rReportFile = GenerateRReport("数据清洗结果报告", ruleReportFields, ruleExecResults);

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
                if (ValidateAttachments(mAttach.ParseTextWithSemicolon()))
                {
                    SendMail(mTo, mCC, mAttach, mSubject, mBody);
                }
                else
                {
                    input:
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    var prompt = "没有找到相应的附件，是否继续发送？(Y/N):";
                    var lenOfInput = 10;
                    var lenOfPrompt = Encoding.Default.GetByteCount(prompt) + lenOfInput;
                    Console.Write(prompt);
                    Console.ResetColor();
                    //var input = Console.ReadLine().ToLower();
                    var input = "n";
                    var whitespaces = "";
                    while (lenOfPrompt-- > 0)
                    {
                        whitespaces += " ";
                    }

                    if (input == "y")
                    {
                        SendMail(mTo, mCC, mAttach, mSubject, mBody);
                    }
                    else if (input == "n")
                    {
                        msg = String.Format("{0}, 取消发送！", mTo);
                        Console.WriteLine(msg);
                        reportData.Add(new string[] { mTo, mSubject, "无", "取消发送" });
                    }
                    else
                    {
                        goto input;
                    }
                }
            }

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("\nDone!");
            Console.ResetColor();
            Console.WriteLine("Generating the report...");
            Thread.Sleep(1000);
            mReportFile = GenerateMReport("邮件发送报告", reportFields, reportData.ToArray());

            End:
            string resSubject = "疫情人员信息收集邮件发送报告 - To:所有处室和分支机构";
            string resAttachments = String.IsNullOrEmpty(rReportFile) ? mReportFile : String.Format("{0};{1}", rReportFile, mReportFile);
            string resBody = string.IsNullOrEmpty(resAttachments) ? string.Empty : "请查看附件以了解更多信息。";
            resBody += string.IsNullOrEmpty(failureInfo) ? string.Empty : "<br>取消发送，详细信息如下：<br>" + failureInfo;

            msg = "\r\nMail To Address, Status";
            Console.WriteLine(msg);
            SendMailSpec(ConfigurationManager.AppSettings["ResponseTo"], "", resAttachments, resSubject, resBody);
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
                var minutes = CheckDataFile(dataFile);
                if (minutes > 30)
                {
                    var curr = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                    failureInfo = String.Format("[{0}]文件版本异常：该文件创建时间已逾{1}分钟，请重新下载！\r\n", curr, minutes);
                    Console.WriteLine(failureInfo);
                    goto End;
                }

                Console.WriteLine("文件是最新的，创建/更新时间在{0}分之内", minutes);

                //#region Fixed WPS Official Issue
                //// wps文件没有完全按照相关协议保存，导致OpenXML读取失败，故在此处进行预处理
                //Console.ForegroundColor = ConsoleColor.Yellow;
                //Console.WriteLine("正在对xlsx文件进行预处理...");
                //Console.ResetColor();
                //string sheetName = ConfigurationManager.AppSettings["DataFileValidSheetName"];
                //ExcelManager.ReSave(dataFile, sheetName);
                //#endregion

                string[] dFields = excelDataMgr.ReadFields(dataFile, sheetName);
                string[][] dRows = excelDataMgr.Read(dataFile, sheetName);
                ProcessDataCleaning(dFields, ref dRows, out ruleExecResults); // Execute data cleaning

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
            rReportFile = GenerateRReport("数据清洗结果报告", ruleReportFields, ruleExecResults);
            //Console.WriteLine("Press any key to continue...");
            //Console.ReadKey();

            try
            {
                string subject = fileName;
                string body = "梅处、袁帅：<br>&nbsp;&nbsp;&nbsp;&nbsp;今天的数据，请查收。<br>";
                SendMailSpec(ConfigurationManager.AppSettings["SpecTo"], String.Empty, outputFile, subject, body);
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
            //Console.Write("Press any key to exit...");
            //Console.ReadKey();

            End:
            var resSubject = "疫情人员信息收集邮件发送报告 - To:办公室";
            var resBody = String.IsNullOrEmpty(failureInfo) ? "发送成功！" : failureInfo;

            msg = "\r\nMail To Address, Status";
            Console.WriteLine(msg);
            SendMailSpec(ConfigurationManager.AppSettings["ResponseTo"], String.Empty, rReportFile, resSubject, resBody);
        }

        private static void SendMail(string mailTo, string mailCc, string attachments, string subject, string body)
        {
            SendMailBase(
                ConfigurationManager.AppSettings["MailFromAddr"],
                ConfigurationManager.AppSettings["MailPassword"],
                ConfigurationManager.AppSettings["FromDisplayName"],
                mailTo,
                mailCc,
                attachments,
                subject,
                body);
        }

        private static void SendMailSpec(string mailTo, string mailCc, string attachments, string subject, string body)
        {
            SendMailBase(
                ConfigurationManager.AppSettings["SpecFrom"],
                ConfigurationManager.AppSettings["SpecPwd"],
                ConfigurationManager.AppSettings["SpecDisplay"],
                mailTo,
                mailCc,
                attachments,
                subject,
                body);
        }

        private static void SendMailBase(string mailFrom, string mailPwd, string displayName, string mailTo, string mailCc, string attachments, string subject, string body)
        {
            int loop = 3;
            var mailSender = new MailSender();
            mailSender.MailFromAddr = mailFrom;
            mailSender.MailFromPwd = mailPwd;
            mailSender.MailFromDisplayName = displayName;
            mailSender.MailTo = mailTo.ParseTextWithSemicolon();
            mailSender.MailCc = mailCc.ParseTextWithSemicolon();
            mailSender.MailSubject = subject;
            mailSender.MailHtmlBody = body;
            var attArr = attachments.ParseTextWithSemicolon();
            mailSender.SetLocalAttachments(attArr);

            bool succ = false;
            while (loop > 0)
            {
                succ = mailSender.Send(ConfigurationManager.AppSettings["SMTPServer"], Convert.ToInt32(ConfigurationManager.AppSettings["SMTPServerPort"]));
                if (succ)
                {
                    break;
                }

                Thread.Sleep(5000);
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

        private static string GenerateRReport(string reportTitle, string[] fields, Dictionary<RuleBase, ValidationResult> execResults)
        {
            string reportFileName = "RReport_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".html";
            var reportFile = Path.Combine(ConfigurationManager.AppSettings["ReportDir"], reportFileName);
            string html = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=utf-8\" /></head><body><h1>{0}</h1><table border=\"1\" cellspacing=\"0\" cellpadding=\"5\">{1}</table></body></html>";
            string rdata = "<tr>";
            foreach (var field in fields)
            {
                rdata += String.Format("<td>{0}</td>", field);
            }
            rdata += "</tr>";
            foreach (var key in execResults.Keys)
            {
                rdata += "<tr>";
                rdata += String.Format("<td>{0}</td>", String.Format("{0}", key.ID));
                rdata += String.Format("<td>{0}</td>", key.Description);
                if (execResults[key] != null && execResults[key].FailedRowIndexes != null)
                {
                    rdata += String.Format("<td>{0}</td>", execResults[key].FailedRowIndexes.Length > 0 ? "<font color=\"red\">False</font>" : "<font color=\"green\">True</font>");
                }
                else if (execResults[key] != null && execResults[key].FailedMsgTexts != null)
                {
                    rdata += String.Format("<td>{0}</td>", execResults[key].FailedMsgTexts.Length > 0 ? "<font color=\"red\">False</font>" : "<font color=\"green\">True</font>");
                }
                else
                {
                    rdata += String.Format("<td>{0}</td>", "N/A");
                }

                rdata += String.Format("<td>{0}</td>", key.FailedMessage);

                if (execResults[key] != null && execResults[key].FailedRowIndexes != null)
                {
                    rdata += String.Format("<td>{0}</td>", execResults[key].FailedRowIndexes.Length.ToString());
                }
                else if (execResults[key] != null && execResults[key].FailedMsgTexts != null)
                {
                    rdata += String.Format("<td>{0}</td>", execResults[key].FailedMsgTexts.Length.ToString());
                }
                else
                {
                    rdata += String.Format("<td>{0}</td>", "N/A");
                }

                if (execResults[key] != null && execResults[key].FailedRowIndexes != null && execResults[key].FailedRowIndexes.Length > 0)
                {
                    rdata += String.Format("<td>{0}</td>", String.Join(",", execResults[key].FailedRowIndexes));
                }
                else if (execResults[key] != null && execResults[key].FailedMsgTexts != null && execResults[key].FailedMsgTexts.Length > 0)
                {
                    rdata += String.Format("<td>{0}</td>", String.Join(",", execResults[key].FailedMsgTexts));
                }
                else
                {
                    rdata += String.Format("<td>{0}</td>", "None");
                }
                rdata += "</tr>";
            }

            html = String.Format(html, reportTitle, rdata);
            File.WriteAllText(reportFile, html);
            Thread.Sleep(500);

            Process.Start("iexplore.exe", reportFile);

            return reportFile;
        }

        private static string GenerateMReport(string reportTitle, string[] fields, string[][] data)
        {
            string reportFileName = "MReport_" + DateTime.Now.ToString("yyyyMMdd_hhmmss") + ".html";
            var reportFile = Path.Combine(ConfigurationManager.AppSettings["ReportDir"], reportFileName);
            string html = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=utf-8\" /></head><body><h1>{0}</h1><table border=\"1\" cellspacing=\"0\" cellpadding=\"5\">{1}</table></body></html>";
            string mdata = "<tr>";
            foreach (var field in fields)
            {
                mdata += String.Format("<td>{0}</td>", field);
            }
            mdata += "</tr>";
            foreach (var row in data)
            {
                mdata += "<tr>";
                int colIdx = 0;
                foreach (var col in row)
                {
                    string colPattern = "<td>{0}</td>";
                    if (colIdx == row.Length - 1)
                    {
                        switch (col)
                        {
                            case "已发送":
                                colPattern = "<td><font color=\"green\">{0}</font></td>";
                                break;
                            case "取消发送":
                                colPattern = "<td><font color=\"red\">{0}</font></td>";
                                break;
                            case "发送失败":
                                colPattern = "<td><font color=\"darkorange\">{0}</font></td>";
                                break;
                            default:
                                break;
                        }
                    }
                    mdata += String.Format(colPattern, col);
                    colIdx++;
                }
                mdata += "</tr>";
            }
            html = String.Format(html, reportTitle, mdata);
            File.WriteAllText(reportFile, html);
            Thread.Sleep(500);

            Process.Start("iexplore.exe", reportFile);

            return reportFile;
        }

        private static bool ValidateAttachments(IEnumerable<string> attachments)
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

        private static void ProcessDataCleaning(string[] fields, ref string[][] records, out Dictionary<RuleBase, ValidationResult> results)
        {
            var mappings = new Dictionary<string, int>();
            for (int i = 0; i < fields.Length; i++)
            {
                mappings.Add(fields[i], i);
            }

            new RuleExecutor().Execute(new DataContext() { Mappings = mappings, OriginRecords = records }, ref records, out results);
        }

        private static Dictionary<string, List<string[]>> SplitRecords(string[] fields, string[][] records)
        {
            var ret = new Dictionary<string, List<string[]>>();
            var mappings = new Dictionary<string, int>();
            for (int i = 0; i < fields.Length; i++)
            {
                mappings.Add(fields[i], i);
            }

            string key = String.Empty;
            foreach (var row in records)
            {
                if (row[mappings["单位分类"]] == "总部")
                {
                    key = row[mappings["所属部门"]];
                    if (String.IsNullOrEmpty(key))
                    {
                        throw new Exception("单位分类为总部时，所属部门不能为空");
                    }


                }
                else if (row[mappings["单位分类"]] == "分支")
                {
                    key = row[mappings["所属分支"]];
                    if (String.IsNullOrEmpty(key))
                    {
                        throw new Exception("单位分类为分支时，所属分支不能为空");
                    }
                }

                if (ret.ContainsKey(key))
                {
                    ret[key].Add(row);
                }
                else
                {
                    ret[key] = new List<string[]>();
                    ret[key].Add(row);
                }
            }

            return ret;
        }
    }
}
