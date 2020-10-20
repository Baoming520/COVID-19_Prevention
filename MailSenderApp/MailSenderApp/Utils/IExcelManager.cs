
namespace MailSenderApp.Utils
{
    #region Namespace.
    #endregion

    public interface IExcelManager
    {
        string[] ReadFields(string fileName, string sheetName);

        string[][] Read(string fileName, string sheetName);

        void Insert(string[][] records, int sheetId, string srcFile);

        void Insert(string[][] records, string sheetName, string srcFile);

        void Write(string[][] records, string filePath);

        void Close();
    }
}
