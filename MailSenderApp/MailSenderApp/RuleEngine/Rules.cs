namespace MailSenderApp.RuleEngine
{
    #region Namespaces.
    using System;
    using System.Collections.Generic;
    #endregion

    public class Rules
    {
        public List<RuleBase> AllRules = new List<RuleBase>
        {
            //new RuleBase()
            //{
            //    ID = "R1",
            //    Name = "R_CheckDate",
            //    Description = "去掉所有提交时间非当前日期的记录，并将“填报日期”以“提交时间”填充",
            //    FailedMessage = "“提交时间”与当前时间不一致",
            //    TakeActionMethod = (DataContext context, ref string[][] records, out ValidationResult result) =>
            //    {
            //        result = new ValidationResult();
            //        var res = new List<string[]>();
            //        var indexes = new List<int>();
            //        foreach(var row in records)
            //        {
            //            var currDate = DateTime.Now;
            //            var submitDate = row[context.Mappings["提交时间"]];
            //            double oaData;
            //            bool flag = Double.TryParse(submitDate, out oaData);
            //            DateTime actSubmitDate;
            //            if(flag == false)
            //            {
            //                actSubmitDate = DateTime.Parse(submitDate);
            //            }
            //            else
            //            {
            //                actSubmitDate = DateTime.FromOADate(oaData);
            //            }

            //            if(currDate.Date == actSubmitDate.Date)
            //            {
            //                row[context.Mappings["填报日期（已删除）"]] = actSubmitDate.ToString("yyyy年M月d日");
            //                res.Add(row);
            //            }
            //            else
            //            {
            //                var orgIndex = context.GetRecordIndex(row);
            //                indexes.Add(orgIndex);
            //            }
            //        }

            //        records = res.ToArray();
            //        result.FailedRowIndexes = indexes.ToArray();
            //    }
            //},
            new RuleBase()
            {
                ID = "R2",
                Name = "R_CheckNameAndPhoneNumber",
                Description = "通过“姓名”和“手机号”进行查重",
                FailedMessage = "发现重复填报数据",
                TakeActionMethod = (DataContext context, ref string[][] records, out ValidationResult result) =>
                {
                    result = new ValidationResult();
                    var res = new List<string[]>();
                    var dict = new Dictionary<string, Tuple<string, DateTime>>();
                    var indexes = new List<int>();
                    foreach(var row in records)
                    {
                        var submitDate = row[context.Mappings["提交时间"]];
                        double oaData;
                        bool flag = Double.TryParse(submitDate, out oaData);
                        DateTime actSubmitDate;
                        if(flag == false)
                        {
                            actSubmitDate = DateTime.Parse(submitDate);
                        }
                        else
                        {
                            actSubmitDate = DateTime.FromOADate(oaData);
                        }
                        var name = row[context.Mappings["姓名"]];
                        var phone = row[context.Mappings["手机号"]];
                        var single = String.Format("{0}_{1}", name, phone);
                        if(!dict.ContainsKey(single))
                        {
                            res.Add(row);
                            dict.Add(single, Tuple.Create(row[context.Mappings["填写ID"]], actSubmitDate));
                        }
                        else
                        {
                            if(dict[single].Item2 < actSubmitDate)
                            {
                                // Replace the record
                                res.RemoveAt(context.GetRecordIndex(dict[single].Item1, res));
                                var orgIndex = context.GetRecordIndex(dict[single].Item1, context.OriginRecords);
                                indexes.Add(orgIndex);
                                dict[single] = Tuple.Create(row[context.Mappings["填写ID"]], actSubmitDate);
                                res.Add(row);
                            }
                            else
                            {
                                var orgIndex = context.GetRecordIndex(row);
                                indexes.Add(orgIndex);
                            }
                        }
                    }

                    records = res.ToArray();
                    result.FailedRowIndexes = indexes.ToArray();
                }
            },
            new RuleBase()
            {
                ID = "R3",
                Name = "R_HandleTimestamp",
                Description = "转换时间戳时间（提交时间）",
                FailedMessage = "该操作影响所有记录",
                TakeActionMethod = (DataContext context, ref string[][] records, out ValidationResult result) =>
                {
                    result = null;
                    for(int i = 0; i < records.Length; i++)
                    {
                        var submitDate = records[i][context.Mappings["提交时间"]];
                        double oaData;
                        bool flag = Double.TryParse(submitDate, out oaData);
                        DateTime actSubmitDate;
                        if(flag == false)
                        {
                            actSubmitDate = DateTime.Parse(submitDate);
                        }
                        else
                        {
                            actSubmitDate = DateTime.FromOADate(oaData);
                        }
                        records[i][context.Mappings["提交时间"]] = actSubmitDate.ToString("yyyy/MM/dd HH:mm:ss");
                    }
                }
            },
            new RuleBase()
            {
                ID = "R4",
                Name = "R_CheckTemperature",
                Description = "将所有体温读数保留一位小数",
                FailedMessage = "该操作影响所有记录",
                TakeActionMethod = (DataContext context, ref string[][] records, out ValidationResult result) =>
                {
                    result = null;
                    int index = 0;
                    foreach(var row in records)
                    {
                        var temp1 = row[context.Mappings["第一次体温（填写时间：9:30-10:00）"]];
                        var temp2 = row[context.Mappings["第二次体温（填写时间：14:00-14:30）"]];
                        decimal res;
                        bool flag = Decimal.TryParse(temp1, out res);
                        if(flag)
                        {
                            temp1 = Decimal.Round(res, 1).ToString();
                        }
                        flag = Decimal.TryParse(temp2, out res);
                        if(flag)
                        {
                            temp2 = Decimal.Round(res, 1).ToString();
                        }

                        records[index][context.Mappings["第一次体温（填写时间：9:30-10:00）"]] = temp1;
                        records[index++][context.Mappings["第二次体温（填写时间：14:00-14:30）"]] = temp2;
                    }
                }
            },
            new RuleBase()
            {
                ID = "R5",
                Name = "R_CheckReportDate",
                Description = "检查当日是否提交",
                FailedMessage = "“当日未提交",
                TakeActionMethod = (DataContext context, ref string[][] records, out ValidationResult result) =>
                {
                    var fMsgTxts = new List<string>();
                    foreach(var row in context.OriginRecords)
                    {
                        var currDate = DateTime.Now;
                        var submitDate = row[context.Mappings["提交时间"]];
                        
                        double val;
                        var f = Double.TryParse(submitDate, out val);
                        DateTime actSubmitDate = f ? actSubmitDate = DateTime.FromOADate(Convert.ToDouble(submitDate)) : actSubmitDate = Convert.ToDateTime(submitDate);
                        if(currDate.Day != actSubmitDate.Day)
                        {
                            fMsgTxts.Add(row[context.Mappings["姓名"]]);
                        }
                    }

                    result = new ValidationResult(){ /*FailedRowIndexes = fIndexes.ToArray(),*/ FailedMsgTexts = fMsgTxts.ToArray() };
                }
            }
        };
    }
}
