using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace ClosedXMLLibrary
{
    public class Hello
    {
        public Hello()
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Hello Sheet");
                worksheet.Cell("A1").Value = "Hello wordl";
                workbook.SaveAs("Hello.xlsx");
            }
        }
    }
}
