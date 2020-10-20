namespace MailSenderApp.Helper
{
    #region Namespaces
    using MailSenderApp.RuleEngine;
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Diagnostics;
    using System.IO;
    using System.Threading;
    #endregion

    public static class ReportHelper
    {
        public static string GenerateRReport(string reportTitle, string[] fields, Dictionary<RuleBase, ValidationResult> execResults)
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

            Process.Start(ConfigurationManager.AppSettings["Browser"], reportFile);

            return reportFile;
        }

        public static string GenerateMReport(string reportTitle, string[] fields, string[][] data)
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

            Process.Start(ConfigurationManager.AppSettings["Browser"], reportFile);

            return reportFile;
        }
    }
}
