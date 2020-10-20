namespace MailSenderApp.RuleEngine
{
    #region Namespaces.
    using System;
    using System.Collections.Generic;
    #endregion

    public class RuleExecutor
    {
        public RuleExecutor()
        {
        }

        public void Execute(DataContext context, ref string[][] records, out Dictionary<RuleBase, ValidationResult> results)
        {
            results = new Dictionary<RuleBase, ValidationResult>();
            foreach (var rule in new Rules().AllRules)
            {
                ValidationResult result;
                rule.TakeActionMethod(context, ref records, out result);
                results.Add(rule, result);
            }
        }
    }
}
