namespace MailSenderApp.Helper
{
    #region Namespaces
    using MailSenderApp.RuleEngine;
    using System;
    using System.Collections.Generic;
    #endregion

    public static class RecordHelper
    {
        public static void DataCleaning(string[] fields, ref string[][] records, out Dictionary<RuleBase, ValidationResult> results)
        {
            var mappings = new Dictionary<string, int>();
            for (int i = 0; i < fields.Length; i++)
            {
                mappings.Add(fields[i], i);
            }

            new RuleExecutor().Execute(new DataContext() { Mappings = mappings, OriginRecords = records }, ref records, out results);
        }

        public static Dictionary<string, List<string[]>> SplitRecords(string[] fields, string[][] records)
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
