namespace MailSenderApp.RuleEngine
{
    #region Namespaces.
    using System.Collections.Generic;
    using System.Linq;
    #endregion

    public class DataContext
    {
        public Dictionary<string, int> Mappings { get; set; }
        public string[][] OriginRecords { get; set; }
        public int GetRecordIndex(string[] row)
        {
            int rIndex = -1;
            for(int i = 0; i < this.OriginRecords.Length; i++)
            {
                if (row[this.Mappings["填写ID"]] == this.OriginRecords[i][this.Mappings["填写ID"]])
                {
                    rIndex = i;
                    break;
                }
            }

            return rIndex;
        }
        public int GetRecordIndex(string id, IEnumerable<string[]> records)
        {
            var rows = records.ToArray();
            int rIndex = -1;
            for (int i = 0; i < rows.Length; i++)
            {
                if (id == rows[i][this.Mappings["填写ID"]])
                {
                    rIndex = i;
                    break;
                }
            }

            return rIndex;
        }
    }
}
