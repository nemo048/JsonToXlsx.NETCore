using System.Collections.Generic;
using System.Data;
using System.IO;
using OfficeOpenXml;

namespace JsonToXlsx.NETCore.Extensions
{
    public static class DataTableExtensions
    {
	    private static readonly object _writeToFileLock = new object();
	    
        public static string WriteToXlsx(this DataTable table, string filePath, bool printHeaders = true)
		{
			FileInfo fileInfo = new FileInfo(filePath);
			
			if (table.Rows.Count == 0)
			{
				return filePath;
			}

			lock (_writeToFileLock)
            {
                ExcelPackage excelPackage;
                ExcelWorksheet excelWorksheet;

                if (fileInfo.Exists)
                {
	                excelPackage = new ExcelPackage(fileInfo);
                    excelWorksheet = excelPackage.Workbook.Worksheets["Лист 1"];
                    excelWorksheet.Cells[excelWorksheet.Dimension.End.Row + 1, excelWorksheet.Dimension.Start.Column]
	                    .LoadFromDataTable(table, false);

                    excelPackage.Save();
                }
                else
                {
	                excelPackage = new ExcelPackage();
                    excelWorksheet = excelPackage.Workbook.Worksheets.Add("Лист 1");
                    excelWorksheet.Cells.LoadFromDataTable(table, printHeaders);

                    excelPackage.SaveAs(fileInfo);
                }

                excelWorksheet.Dispose();
                excelPackage.Dispose();
            }

            return filePath;
        }
    }
}