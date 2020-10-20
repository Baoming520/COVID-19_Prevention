namespace MailSenderApp.Utils
{
    #region Namespace.
    using System;
    #endregion

    [Serializable]
    public class OpenXMLCellReferenceNullException : ApplicationException
    {
        public OpenXMLCellReferenceNullException()
        {
        }

        public OpenXMLCellReferenceNullException(string message) 
            : base(message)
        {
        }

        public OpenXMLCellReferenceNullException(string message, Exception inner) 
            : base(message, inner)
        {
        }
    }
}
