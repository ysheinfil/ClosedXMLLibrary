using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            var xlsx = @"C:\inputs\CFC CENSUS 3-18-19.xlsx";
            var csv = @"C:\inputs\CFC CENSUS 3-18-19.csv";
            ClosedXMLLibrary.Convert.XlsxToCsv(xlsx);            
            Console.ReadKey();

        }
    }
}
