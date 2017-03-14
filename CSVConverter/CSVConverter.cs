using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace CSVConverter
{
    public static class CSVConverter
    {
        public static void Convert(params FileInfo[] files)
        {
            foreach (ExcelPackage package in GetPackages(files))
            {
                ExcelWorksheet[] worksheets = package.Workbook.Worksheets.Cast<ExcelWorksheet>().ToArray();

                foreach (ExcelWorksheet w in worksheets)
                {
                    Page Page = new Page(w);
                }
            }
        }

        private static IEnumerable<ExcelPackage> GetPackages(FileInfo[] files)
        {
            IList<ExcelPackage> result = new List<ExcelPackage>();

            foreach (FileInfo file in files)
                try
                {
                    result.Add(new ExcelPackage(file));
                }
                catch (Exception exc)
                {
                    Console.Error.WriteLine(string.Concat(file.FullName, " - ", exc.Message));
                }

            return result;
        }
    }
}
