using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ClosedXML.Excel;

namespace ClosedXMLLibrary
{
    public class Convert
    {
        public static void XlsxToCsv(string xlsx, string csv)
        {
            StringBuilder csvContents = new StringBuilder();
            string cellValue;
            using (var workbook = new XLWorkbook(xlsx))
            {
                var worksheet = workbook.Worksheet(1);
                var lastCol = worksheet.LastColumnUsed();

                foreach (var row in worksheet.RowsUsed())
                {
                    foreach (var col in worksheet.ColumnsUsed())
                    {
                        cellValue = row.Cell(col.ColumnNumber()).GetString();
                        if (cellValue.Contains("\""))
                            cellValue = cellValue.Replace("\"", "\"\"");
                        if (cellValue.Contains(','))
                            cellValue = "\"" + cellValue + "\"";
                        csvContents.Append(cellValue);
                        if (col != lastCol)
                            csvContents.Append(",");
                        else
                            csvContents.AppendLine();                        
                    }
                    Console.Write(".");
                }
            }
            File.WriteAllText(xlsx.Replace(".xlsx", ".csv"), csvContents.ToString());
        }

        public static void XlsxToCsv(string xlsx)
        {
            XlsxToCsv(xlsx, xlsx.Replace("xlsx", "csv"));
        }
    }
}
