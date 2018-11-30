using ExcelDataReader;
using Nortal.Utilities.Csv;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ExcelConversion
{
    public class ConvertExcel
    {
        public LocationInfo ReadInfoFromExcel(string fileInPath)
        {
            LocationInfo locInfo = new LocationInfo();

            using (var stream = File.Open(fileInPath, FileMode.Open, FileAccess.Read))
            {
                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx, *.xlsm)
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var dataSet = reader.AsDataSet();
                    
                    DataTable workSheetCoverLocInfo = dataSet.Tables[1];
                    
                    RowColIndexes propertyIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "Property Description:");
                    RowColIndexes bedRoomIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "Bedrooms:");
                    RowColIndexes bathRoomIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "Bathrooms:");
                    RowColIndexes dateOnMktIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "Days on the Market:");
                    RowColIndexes addressIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "Address:");
                    RowColIndexes cityStateIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "City, State:");
                    

                    locInfo = workSheetCoverLocInfo.AsEnumerable().Select(x => new LocationInfo
                    {
                        property = x.Table.Rows[propertyIndexes.RowIndex + 1][propertyIndexes.ColIndex + 2].ToString().Trim()/* + x.Table.Rows[propertyIndexes.rowIndex + 1][propertyIndexes.colIndex + 1].ToString().Trim()*/,
                        bedrooms = x.Table.Rows[bedRoomIndexes.RowIndex + 1][bedRoomIndexes.ColIndex].ToString().Trim(),
                        bathrooms = x.Table.Rows[bathRoomIndexes.RowIndex + 1][bathRoomIndexes.ColIndex].ToString().Trim(),
                        dateOnMkt = x.Table.Rows[dateOnMktIndexes.RowIndex + 1][dateOnMktIndexes.ColIndex + 3].ToString().Trim() + x.Table.Rows[dateOnMktIndexes.RowIndex][dateOnMktIndexes.ColIndex + 2].ToString().Trim(),
                        address = x.Table.Rows[addressIndexes.RowIndex][addressIndexes.ColIndex].ToString().Trim(),
                        cityState = x.Table.Rows[cityStateIndexes.RowIndex + 1][cityStateIndexes.ColIndex].ToString().Trim(),

                    }).First();
                }
            }

            return locInfo;
        }

        private RowColIndexes GetTableRowColIndexesForExactMatch(DataTable workSheetCoverLocInfo, string searchText)
        {
            RowColIndexes returnRCIndexes = new RowColIndexes();
            int rowIndex = -1;

            //LINQ answer on Stack Overflow

            var rowIndexArray = workSheetCoverLocInfo
             .Rows
             .Cast<DataRow>()
             .Where(r => r.ItemArray.Any(c => Regex.IsMatch(c.ToString().Trim(), searchText, RegexOptions.IgnoreCase)))
             .Select(r => r.Table.Rows.IndexOf(r)).ToArray();

            if (rowIndexArray.Length > 0)
            {
                var rowCol = rowIndexArray[0];
                rowIndex = rowCol;
            }

            int colIndex = 0;
            if (rowIndex >= 0)
            {
                foreach (var dc in workSheetCoverLocInfo.Rows[rowIndex].ItemArray)
                {
                    if (dc != DBNull.Value)
                    {
                        if (Regex.IsMatch(dc.ToString().Trim(), searchText, RegexOptions.IgnoreCase))
                        {
                            break;
                        }
                    }
                    colIndex++;
                }
            }
            else
            {
                colIndex = -1;
            }

            return new RowColIndexes { RowIndex = rowIndex, ColIndex = colIndex };
        }


        private int GetTableRowIndexForContainsText(DataTable workSheetCoverLocInfo, string searchText)
        {
            int returnRowIndex = -1;
            var rowIndex = workSheetCoverLocInfo
             .Rows
             .Cast<DataRow>()
             .Where(r => r.ItemArray.Any(c => Regex.IsMatch(c.ToString().Trim(), searchText, RegexOptions.IgnoreCase)))
             .Select(r => r.Table.Rows.IndexOf(r)).ToArray();

            if (rowIndex.Length > 0)
            {
                returnRowIndex = rowIndex[0];
            }
            return returnRowIndex;
        }

        public void WriteToCSV(LocationInfo locInfo, string fileOutPath)
        {
            using (var writer = new StringWriter())
            {
                var csv = new CsvWriter(writer, new CsvSettings());
                csv.WriteLine(locInfo.property, locInfo.bedrooms, locInfo.bathrooms, locInfo.dateOnMkt, locInfo.address, locInfo.cityState);
                File.WriteAllText(fileOutPath, writer.ToString());
            }
        }
    }

}

