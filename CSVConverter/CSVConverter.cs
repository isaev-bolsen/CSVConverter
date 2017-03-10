using OfficeOpenXml.Core.ExcelPackage;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace CSVConverter
{
    public static  class CSVConverter
    {
        public static void Convert(params FileInfo[] files)
        {
            Page[] pages = Convert(files.Select(f => new ExcelPackage(f))).ToArray();
        }

        private static IEnumerable<Page> Convert(IEnumerable<ExcelPackage> files)
        {
            return files.SelectMany(f => f.Workbook.Worksheets.Cast<ExcelWorksheet>()).Select(w => new Page(w));
        }
    }
}
