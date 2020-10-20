namespace MailSenderApp.RuleEngine
{
    #region Namespaces.
    using System.Collections.Generic;
    #endregion

    public class RuleBase
    {
        public delegate void TakeAction(DataContext context, ref string[][] records, out ValidationResult result);
        public string ID { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public string FailedMessage { get; set; }
        public TakeAction TakeActionMethod { get; set; }

        public override bool Equals(object obj)
        {
            var param = (RuleBase)obj;

            return this.ID == param.ID && this.Name == param.Name;
        }

        public override int GetHashCode()
        {
            return this.ID.GetHashCode() + this.Name.GetHashCode();
        }
    }
}
