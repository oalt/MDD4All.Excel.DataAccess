using System.Collections.Generic;

namespace MDD4All.Excel.DataAccess.Contracts
{
    public interface ISpreadsheetDataReader
    {
        void OpenFile(string filename);

        List<string> GetSpreadsheetNames();

        /// <summary>
        /// Get a cell content as a string.
        /// </summary>
        /// <param name="spreadsheetName">The spradsheet name</param>
        /// <param name="rowIndex">The 1based row index</param>
        /// <param name="columnName">The column name (e.g. "A", "B", "AA" etc.)</param>
        /// <returns>The value or an empty string if the cell is empty or can not be read.</returns>
        string GetCellContent(string spreadsheetName, int rowIndex, string columnName);
    }
}
