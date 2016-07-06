using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace Import
{
  public class ExcelImporter
  {
    private readonly Workbook _workbook = new Workbook();

    public void Import(DataTable dataTable)
    {
      var sheet = _workbook.Worksheets[0];
      sheet.Cells.ImportDataTable(dataTable, true, 0, 0, dataTable.Rows.Count, dataTable.Columns.Count, true, "dd.mm.yyyy");
      sheet.AutoFitColumns();
    }

    public void Save(string fileName)
    {
      _workbook.Save(fileName, SaveFormat.Excel97To2003);
    }

    #region private methods
    /// <summary>
    /// Static constructor
    /// </summary>
    static ExcelImporter()
    {
    }

    #endregion
  }
}
