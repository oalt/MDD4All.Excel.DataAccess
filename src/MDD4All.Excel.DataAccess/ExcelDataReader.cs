using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using MDD4All.Excel.DataAccess.Contracts;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace MDD4All.Excel.DataAccess
{
    public class ExcelDataReader : ISpreadsheetDataReader
    {

        private SpreadsheetDocument _spreadsheetDocument = null;

        private WorkbookPart _workbookPart;
        private Sheets _sheetCollection;

        private Dictionary<string, SheetData> _sheetDictionary = new Dictionary<string, SheetData>();

        public string GetCellContent(string spreadsheetName, int rowIndex, string columnName)
        {
            string result = "";

            try
            {
                SheetData sheet = _sheetDictionary[spreadsheetName];

                if (sheet != null)
                {
                    IEnumerable<Row> rows = sheet.Elements<Row>().Where(r => r.RowIndex == rowIndex);

                    if (rows.Count() == 0)
                    {
                        // A cell does not exist at the specified row.
                        result = "";
                    }

                    IEnumerable<Cell> cells = rows.First().Elements<Cell>().Where(c => string.Compare(c.CellReference.Value, columnName + rowIndex, true) == 0);
                    if (cells.Count() == 0)
                    {
                        // A cell does not exist at the specified column, in the specified row.
                        result = "";
                    }

                    Cell cell = cells.First();

                    if (cell.DataType != null)
                    {
                        if (cell.DataType == CellValues.SharedString)
                        {
                            int id;
                            if (Int32.TryParse(cell.InnerText, out id))
                            {
                                SharedStringItem item = _workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
                                if (item.Text != null)
                                {
                                    //code to take the string value  
                                    result = item.Text.Text;
                                }
                                else if (item.InnerText != null)
                                {
                                    result = item.InnerText;
                                }
                                else if (item.InnerXml != null)
                                {
                                    result = item.InnerXml;
                                }
                            }
                        }
                        else
                        {
                            result = cell.CellValue.InnerText;
                        }
                    }
                    else
                    {
                        result = cell.InnerText;
                    }
                }
            }
            catch(Exception exception)
            {
                Debug.WriteLine(exception);
            }
            return result;
        }

        public List<string> GetSpreadsheetNames()
        {
            List<string> result = new List<string>();

            foreach(Sheet sheet in _sheetCollection)
            {
                result.Add(sheet.Name);
            }

            return result;
        }

        

        public void OpenFile(string filename)
        {
            _spreadsheetDocument = SpreadsheetDocument.Open(filename, false);

            _workbookPart = _spreadsheetDocument.WorkbookPart;

            _sheetCollection = _workbookPart.Workbook.GetFirstChild<Sheets>();

            _sheetDictionary = new Dictionary<string, SheetData>();

            foreach(Sheet sheet in _sheetCollection)
            {
                string name = sheet.Name;

                Debug.WriteLine(name);

                if(!_sheetDictionary.ContainsKey(name))
                {
                    Worksheet theWorksheet = ((WorksheetPart)_workbookPart.GetPartById(sheet.Id)).Worksheet;

                    _sheetDictionary.Add(name, theWorksheet.GetFirstChild<SheetData>());
                }
            }
        }
    }
}
