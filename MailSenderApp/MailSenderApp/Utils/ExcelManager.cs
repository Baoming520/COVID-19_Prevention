namespace MailSenderApp.Utils
{
    #region Namespace.
    using Microsoft.Office.Interop.Excel;
    using System;
    using System.IO;
    using System.Runtime.InteropServices;
    #endregion

    public class ExcelManager : IExcelManager
    {
        public static void ReSave(string fileName, string sheetName)
        {
            var excelApp = new Application();
            excelApp.DisplayAlerts = false;
            try
            {
                if (excelApp == null)
                {
                    return;
                }

                var workbook = excelApp.Workbooks.Open(fileName);
                Sheets sheets = workbook.Worksheets;
                // Use the following method to get worksheet by sheet id:
                // --> Worksheet worksheet = (Worksheet)sheets.get_Item(sheetIndex);
                Worksheet worksheet = (Worksheet)sheets[sheetName];
                if (worksheet == null)
                {
                    return;
                }

                workbook.SaveAs(fileName);
                workbook.Close();
                excelApp.Quit();

                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("\r\n" + ex.Message + "\r\n");
                Console.ResetColor();

                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        public ExcelManager()
        {
            this.app = new Application();
            this.workbook = null;
        }

        public string[] ReadFields(string fileName, string sheetName)
        {
            var worksheet = this.Import(fileName, sheetName);

            return this.ReadFields(fileName, worksheet);
        }

        public string[] ReadFields(string fileName, int sheetId = 1)
        {
            var worksheet = this.Import(fileName, sheetId);

            return this.ReadFields(fileName, worksheet);
        }

        public string[][] Read(string fileName, string sheetName)
        {
            var worksheet = this.Import(fileName, sheetName);

            return this.Read(fileName, worksheet, startIndex: 2);
        }

        public string[][] Read(string fileName, int sheetId = 1)
        {
            var worksheet = this.Import(fileName, sheetId);

            return this.Read(fileName, worksheet, startIndex: 2);
        }

        public void Insert(string[][] records, int sheetId, string srcFile)
        {
            throw new NotImplementedException();
        }

        public void Insert(string[][] records, string sheetName, string srcFile)
        {
            throw new NotImplementedException();
        }

        public void Write(string[][] records, string excelFilePath)
        {
            if (null == records || records.Length == 0)
            {
                return;
            }

            this.app = new Application();
            this.workbook = app.Workbooks.Add(1);
            if (null != this.workbook)
            {
                Worksheet worksheet = workbook.Sheets[1];
                if (null != worksheet)
                {
                    #region Set the format of all the cells.
                    // Set the format of all the cells.
                    int rowCount = records.Length;
                    int colCount = records[0].Length;
                    Range range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[rowCount, colCount]];
                    range.NumberFormatLocal = "@";
                    #endregion

                    for (int i = 0; i < records.Length; i++)
                    {
                        for (int j = 0; j < records[i].Length; j++)
                        {
                            worksheet.Cells[i + 1, j + 1] = records[i][j];
                        }
                    }
                }

                workbook.SaveCopyAs(excelFilePath);
            }

            this.Close();
        }

        public string[] ReadFields(string fileName, Worksheet worksheet)
        {
            string[] fields = null;
            if (null == worksheet)
            {
                return fields;
            }

            fields = new string[this.colCount];
            Range row = worksheet.Rows[1];
            for (int j = 0; j < this.colCount; j++)
            {
                fields[j] = row.Cells[j + 1].Text.ToString();
            }

            return fields;
        }

        public string[][] Read(string fileName, Worksheet worksheet, int startIndex = 2)
        {
            string[][] result = null;
            if (null == worksheet)
            {
                return result;
            }

            int rowCount = this.rowCount - startIndex + 1;
            result = new string[rowCount][];
            for (int i = 0; i < rowCount; i++)
            {
                Range row = worksheet.Rows[i + startIndex];
                if (null == row)
                {
                    return result;
                }

                result[i] = new string[this.colCount];
                for (int j = 0; j < this.colCount; j++)
                {
                    if (null != row.Cells[j + 1].Value)
                    {
                        result[i][j] = row.Cells[j + 1].Value.ToString();
                    }
                    else
                    {
                        result[i][j] = row.Cells[j + 1].Text.ToString();
                    }
                }
            }

            return result;
        }

        public void Close()
        {
            if (null != this.workbook)
            {
                this.workbook.Close(false);
                Marshal.ReleaseComObject(this.workbook);
                this.workbook = null;
            }

            if (null != this.app)
            {
                this.app.Workbooks.Close();
                this.app.Quit();
                Marshal.ReleaseComObject(this.app);
                this.app = null;
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        #region Private members
        private int rowCount;
        private int colCount;
        private Application app;
        private Workbook workbook;

        private Worksheet Import(string fileName, string sheetName)
        {
            try
            {
                if (this.app == null)
                {
                    return null;
                }

                this.workbook = this.app.Workbooks.Open(fileName);
                Sheets sheets = this.workbook.Worksheets;
                Worksheet worksheet = (Worksheet)sheets[sheetName];
                if (worksheet == null)
                {
                    return null;
                }

                this.rowCount = worksheet.UsedRange.Rows.Count;
                this.colCount = worksheet.UsedRange.Columns.Count;

                return worksheet;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }
        }

        private Worksheet Import(string fileName, int sheetIndex = 1)
        {
            try
            {
                if (this.app == null)
                {
                    return null;
                }

                this.workbook = this.app.Workbooks.Open(fileName);
                Sheets sheets = this.workbook.Worksheets;
                Worksheet worksheet = (Worksheet)sheets.get_Item(sheetIndex);
                if (worksheet == null)
                {
                    return null;
                }

                this.rowCount = worksheet.UsedRange.Rows.Count;
                this.colCount = worksheet.UsedRange.Columns.Count;

                return worksheet;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }
        }
        
        #endregion

    }
}
