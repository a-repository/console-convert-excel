using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelConversion
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileOutPath = @"D:\test_csv.csv";
            string fileInPath = @"D:\test_excel.xlsm";
            //instaniate class
            ConvertExcel convert = new ConvertExcel();
            //read from excel
            LocationInfo locInfo = convert.ReadInfoFromExcel(fileInPath);
            //write out to CSV
            convert.WriteToCSV(locInfo, fileOutPath);
        }
    }
}
