namespace MailSenderApp.RuleEngine
{
    public class ValidationResult
    {
        public int[] FailedRowIndexes { get; set; }
        public string[] FailedMsgTexts { get; set; }
    }
}
