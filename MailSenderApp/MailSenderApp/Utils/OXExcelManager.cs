
namespace MailSenderApp.Utils
{
    #region Namespace.
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using MailSenderApp.Models;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text.RegularExpressions;
    #endregion

    public class OXExcelManager : IExcelManager
    {
        public OXExcelManager()
        {
            this.fieldCount = 0;
        }

        public string[] ReadFields(string filePath, string sheetName)
        {
            string[] fields = null;
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                return fields;
            }

            this.GenSheetInfo(filePath);

            try
            {
                using (SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.Open(filePath, false))
                {
                    WorkbookPart workbookPart = spreadsheetDoc.WorkbookPart;
                    WorksheetPart worksheetPart = null;
                    string rid = this.GetRid(sheetName);
                    foreach (var sheetPart in workbookPart.WorksheetParts)
                    {
                        if (spreadsheetDoc.WorkbookPart.GetIdOfPart(sheetPart) == rid)
                        {
                            worksheetPart = sheetPart;
                        }
                    }

                    if (null == worksheetPart)
                    {
                        return fields;
                    }

                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                    if (null == sheetData.Elements<Row>() || sheetData.Elements<Row>().Count() == 0)
                    {
                        return fields;
                    }

                    var rows = sheetData.Elements<Row>().ToList();
                    if (null == rows || rows.Count == 0 ||
                        null == rows[0].Elements<Cell>() || rows[0].Elements<Cell>().Count() == 0)
                    {
                        return fields;
                    }

                    var cells = rows[0].Elements<Cell>().ToList();
                    this.fieldCount = cells.Count;
                    this.GenColumnMarks();
                    fields = new string[this.fieldCount];
                    for (int j = 0; j < this.fieldCount; j++)
                    {
                        fields[j] = GetCellValue(spreadsheetDoc, cells[j]);
                    }
                }
            }
            catch (IOException io_ex)
            {
                throw io_ex;
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return fields;
        }

        public string[][] Read(string filePath, string sheetName)
        {
            string[][] data = null;
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                return data;
            }

            try
            {
                using (SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.Open(filePath, false))
                {
                    WorkbookPart workbookPart = spreadsheetDoc.WorkbookPart;
                    WorksheetPart worksheetPart = null;
                    string rid = this.GetRid(sheetName);
                    foreach (var sheetPart in workbookPart.WorksheetParts)
                    {
                        if (spreadsheetDoc.WorkbookPart.GetIdOfPart(sheetPart) == rid)
                        {
                            worksheetPart = sheetPart;
                        }
                    }

                    if (null == worksheetPart)
                    {
                        return data;
                    }

                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    if (null == sheetData.Elements<Row>() || sheetData.Elements<Row>().Count() == 0)
                    {
                        return data;
                    }

                    var rows = sheetData.Elements<Row>().ToList();
                    if (null == rows || rows.Count == 0)
                    {
                        return data;
                    }

                    data = new string[rows.Count - 1][];
                    for (int i = 1; i < rows.Count(); i++)
                    {
                        var cells = rows[i].Elements<Cell>().ToList();
                        if (null == cells || cells.Count == 0)
                        {
                            return data;
                        }

                        int cIndex = 0;
                        data[i - 1] = new string[this.fieldCount];
                        for (int j = 0; j < this.fieldCount; j++)
                        {
                            if (cIndex < cells.Count)
                            {
                                if (cells[cIndex].CellReference == null)
                                {
                                    string errMsg = "Bad format!";
                                    throw new OpenXMLCellReferenceNullException(errMsg);
                                }

                                if (this.SeparateColumnName(cells[cIndex].CellReference) == this.columnMarks[j])
                                {
                                    var scNotationPattern = "^((-?\\d+.?\\d*)[Ee]{1}(-?\\d+))$";
                                    var cellVal = GetCellValue(spreadsheetDoc, cells[cIndex++]);
                                    if (Regex.IsMatch(cellVal, scNotationPattern))
                                    {
                                        cellVal = Convert.ToDecimal(Decimal.Parse(cellVal, System.Globalization.NumberStyles.Float)).ToString();
                                    }
                                    data[i - 1][j] = cellVal;
                                }
                                else
                                {
                                    data[i - 1][j] = string.Empty;
                                }
                            }
                            else
                            {
                                data[i - 1][j] = string.Empty;
                            }
                        }
                    }
                }
            }
            catch (IOException io_ex)
            {
                throw io_ex;
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return data;
        }

        public void Insert(string[][] records, string sheetName, string srcFile)
        {
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(srcFile, true))
            {
                IEnumerable<Sheet> sheets = doc.WorkbookPart.Workbook.Descendants<Sheet>().Where(sh => sh.Name == sheetName);
                if (sheets.Count() == 0)
                {
                    throw new KeyNotFoundException("Sheet name is not found.");
                }
                WorksheetPart worksheetPart = (WorksheetPart)doc.WorkbookPart.GetPartById(sheets.First().Id);
                Worksheet worksheet = worksheetPart.Worksheet;
                SheetData sheetData = worksheet.GetFirstChild<SheetData>();

                if (null == records || records.Length == 0)
                {
                    return;
                }

                for (int i = 0; i < records.Length; i++)
                {
                    if (null == records[i] || records[i].Length == 0)
                    {
                        return;
                    }

                    Row row = new Row();
                    for (int j = 0; j < records[i].Length; j++)
                    {
                        Cell cell = new Cell();
                        cell.DataType = CellValues.String;
                        cell.CellValue = !string.IsNullOrEmpty(records[i][j]) ? new CellValue(records[i][j]) : new CellValue(string.Empty);
                        row.AppendChild(cell);
                    }

                    sheetData.Append(row);
                }

                worksheet.Save();
                doc.Save();
            }
        }

        public void Insert(string[][] records, int sheetId, string srcFile)
        {
            this.GenSheetInfo(srcFile);

            // Read xlsx original file. 
            using (var doc = SpreadsheetDocument.Open(srcFile, true))
            {
                WorkbookPart workbookPart = doc.WorkbookPart;
                Worksheet worksheet = null;
                string rid = this.GetRid(sheetId);
                foreach (var worksheetPart in workbookPart.WorksheetParts)
                {
                    if (doc.WorkbookPart.GetIdOfPart(worksheetPart) == rid)
                    {
                        worksheet = worksheetPart.Worksheet;
                    }
                    else
                    {
                        var ws = worksheetPart.Worksheet;
                        var rows = ws.Descendants<Row>();
                        if (rows != null)
                        {
                            rows.SelectMany(s => s.Elements<Cell>())
                                .Where(s => s.CellFormula != null)
                                .ToList()
                                .ForEach(s => s.CellFormula.CalculateCell = true);
                        }
                    }
                }

                // Inserted data
                SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                //sheetData.RemoveAllChildren(); // Remove all the child nodes from the sheetData node.
                if (null == records || records.Length == 0)
                {
                    return;
                }

                for (int i = 0; i < records.Length; i++)
                {
                    if (null == records[i] || records[i].Length == 0)
                    {
                        return;
                    }

                    Row row = new Row();
                    for (int j = 0; j < records[i].Length; j++)
                    {
                        Cell cell = new Cell();
                        cell.DataType = CellValues.String;
                        cell.CellValue = !string.IsNullOrEmpty(records[i][j]) ? new CellValue(records[i][j]) : new CellValue(string.Empty);
                        row.AppendChild(cell);
                    }

                    sheetData.Append(row);
                }

                worksheet.Save();
                doc.Save();
            }
        }

        public void Write(string[][] records, string filePath)
        {
            using (SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookpart = spreadsheetDoc.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);
                Sheets sheets = spreadsheetDoc.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
                Sheet sheet = new Sheet()
                {
                    Id = spreadsheetDoc.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Sheet1"
                };

                sheets.Append(sheet);

                if (null == records || records.Length == 0)
                {
                    return;
                }

                for (int i = 0; i < records.Length; i++)
                {
                    if (null == records[i] || records[i].Length == 0)
                    {
                        return;
                    }

                    Row row = new Row();
                    for (int j = 0; j < records[i].Length; j++)
                    {
                        Cell cell = new Cell();
                        cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                        cell.CellValue = !string.IsNullOrEmpty(records[i][j]) ? new CellValue(records[i][j]) : new CellValue(string.Empty);
                        row.AppendChild(cell);
                    }

                    sheetData.Append(row);
                }

                try
                {
                    workbookpart.Workbook.Save();
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        public void Close()
        {
            GC.Collect();
        }

        #region Others useful methods.
        public string[][] ReadBase(string filePath, int colLength = -1, int startRowIndex = 0, int rowLength = -1, int sheetIndex = 1)
        {
            string[][] data = null;
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                return data;
            }

            this.GenSheetInfo(filePath);

            if (null == this.sheetInfos || this.sheetInfos.Count == 0)
            {
                return data;
            }

            try
            {
                using (SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.Open(filePath, false))
                {
                    WorkbookPart workbookPart = spreadsheetDoc.WorkbookPart;
                    WorksheetPart worksheetPart = null;
                    string rid = this.GetRid(sheetIndex);
                    foreach (var sheetPart in workbookPart.WorksheetParts)
                    {
                        if (spreadsheetDoc.WorkbookPart.GetIdOfPart(sheetPart) == rid)
                        {
                            worksheetPart = sheetPart;
                        }
                    }

                    if (null == worksheetPart)
                    {
                        return data;
                    }

                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                    if (null == sheetData.Elements<Row>() || sheetData.Elements<Row>().Count() == 0)
                    {
                        return data;
                    }

                    var rows = sheetData.Elements<Row>().ToList();
                    if (null == rows || rows.Count == 0)
                    {
                        return data;
                    }

                    if (rowLength == -1)
                    {
                        rowLength = rows.Count - startRowIndex;
                    }

                    if (rows.Count - startRowIndex < rowLength)
                    {
                        return data;
                    }

                    data = new string[rowLength][];
                    for (int i = startRowIndex; i < rowLength + startRowIndex; i++)
                    {
                        var cells = rows[i].Elements<Cell>().ToList();
                        if (null == cells || cells.Count == 0)
                        {
                            return data;
                        }

                        int cIndex = 0;
                        if (colLength == -1 && i == 0)
                        {
                            colLength = cells.Count;
                        }

                        if (colLength == -1)
                        {
                            throw new Exception("");
                        }

                        this.fieldCount = colLength;
                        this.GenColumnMarks();
                        data[i - startRowIndex] = new string[colLength];
                        for (int j = 0; j < colLength; j++)
                        {
                            if (cIndex < cells.Count)
                            {
                                if (cells[cIndex].CellReference == null)
                                {
                                    string errMsg = "Bad format!";
                                    throw new OpenXMLCellReferenceNullException(errMsg);
                                }

                                if (this.SeparateColumnName(cells[cIndex].CellReference) == this.columnMarks[j])
                                {
                                    data[i - startRowIndex][j] = GetCellValue(spreadsheetDoc, cells[cIndex++]);
                                }
                                else
                                {
                                    data[i - startRowIndex][j] = string.Empty;
                                }
                            }
                            else
                            {
                                data[i - startRowIndex][j] = string.Empty;
                            }
                        }
                    }
                }
            }
            catch (IOException io_ex)
            {
                throw io_ex;
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return data;
        }

        #endregion

        #region Private members.
        //private const int SheetIndex = 1;
        //private const string SheetName = "收集表"; 
        private int fieldCount;
        private string[] columnMarks;
        private List<OXSheetInfo> sheetInfos;
        #endregion

        #region Private methods.
        private string GetCellValue(SpreadsheetDocument spreadSheetDoc, Cell cell)
        {
            if (null != cell.DataType && CellValues.InlineString == cell.DataType.Value)
            {
                return cell.InlineString.InnerText;
            }

            if (null == cell.CellValue)
            {
                return string.Empty;
            }

            string value = cell.CellValue.InnerText;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return spreadSheetDoc.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements.GetItem(int.Parse(value)).InnerText;
            }

            return value;
        }

        private void GenColumnMarks()
        {
            if (this.fieldCount == 0)
            {
                return;
            }

            this.columnMarks = new string[this.fieldCount];
            for (int i = 0; i < fieldCount; i++)
            {
                var hdd = new Hexadoubledecimal(i + 1);
                if (null == hdd)
                {
                    this.columnMarks = null;

                    break;
                }

                this.columnMarks[i] = hdd.HDDValue;
            }
        }

        private string SeparateColumnName(StringValue cellReference)
        {
            if (null == cellReference)
            {
                return null;
            }

            string columnName = string.Empty;
            for (int i = 0; i < cellReference.Value.Length; i++)
            {
                if (cellReference.Value[i] >= 65 && cellReference.Value[i] <= 90)
                {
                    columnName += cellReference.Value[i];
                }
            }

            return columnName;
        }

        private int GetColumnIndex(string columnName)
        {
            int index = -1;
            if (string.IsNullOrEmpty(columnName))
            {
                return index;
            }

            var hdd = new Hexadoubledecimal(columnName);
            if (null != hdd)
            {
                index = hdd.DecValue - 1;
            }

            return index;
        }

        private void GenSheetInfo(string filePath)
        {
            this.sheetInfos = new List<OXSheetInfo>();
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                return;
            }

            try
            {
                using (SpreadsheetDocument mySpreadsheet = SpreadsheetDocument.Open(filePath, false))
                {
                    OpenXmlElement sheets = mySpreadsheet.WorkbookPart.Workbook.Sheets;

                    // For each sheet, display the sheet information.
                    foreach (OpenXmlElement sheet in sheets)
                    {
                        var sheetInfo = new OXSheetInfo();
                        foreach (OpenXmlAttribute attr in sheet.GetAttributes())
                        {
                            switch (attr.LocalName)
                            {
                                case "name":
                                    sheetInfo.Name = attr.Value;
                                    break;
                                case "sheetId":
                                    int val;
                                    bool flag = Int32.TryParse(attr.Value, out val);
                                    sheetInfo.SheetId = flag ? val : 0;
                                    break;
                                case "id":
                                    if (attr.Prefix == "r")
                                    {
                                        sheetInfo.Rid = attr.Value;
                                    }

                                    break;
                                default:
                                    break;
                            }
                        }

                        this.sheetInfos.Add(sheetInfo);
                    }
                }
            }
            catch (IOException ioEx)
            {
                throw ioEx;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private string GetRid(int sheetId)
        {
            if (null == this.sheetInfos || this.sheetInfos.Count == 0)
            {
                return string.Empty;
            }

            foreach (var sheetInfo in this.sheetInfos)
            {
                if (sheetId == sheetInfo.SheetId)
                {
                    return sheetInfo.Rid;
                }
            }

            return string.Empty;
        }

        private string GetRid(string sheetName)
        {
            if (null == this.sheetInfos || this.sheetInfos.Count == 0)
            {
                return string.Empty;
            }

            foreach (var sheetInfo in this.sheetInfos)
            {
                if (sheetName == sheetInfo.Name)
                {
                    return sheetInfo.Rid;
                }
            }

            return string.Empty;
        }
        #endregion
    }

    /// <summary>
    /// OpenXML sheet information data.
    /// </summary>
    public struct OXSheetInfo
    {
        public string Name { get; set; }
        public int SheetId { get; set; }
        public string Rid { get; set; }
    }
}
