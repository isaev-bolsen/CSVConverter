using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;

namespace CSVConverter
{
    public static class CSVConverter
    {
        public static void Convert(params FileInfo[] files)
        {
            foreach (FileInfo file in files)
                try
                {
                    string Filename = Path.GetFileNameWithoutExtension(file.Name);
                    ExcelWorksheet[] worksheets = new ExcelPackage(file).Workbook.Worksheets.Cast<ExcelWorksheet>().ToArray();
                    if (worksheets.Count() == 1) new Page(worksheets.Single()).Save(GetResultPath(file.Directory, Filename));
                    else foreach (ExcelWorksheet ws in worksheets)
                            new Page(ws).Save(GetResultPath(file.Directory, string.Join("_", Filename, ws.Name)));
                }
                catch (Exception exc)
                {
                    Console.Error.WriteLine(string.Concat(file.FullName, " - ", exc.Message));
                }
        }

        private static FileInfo GetResultPath(DirectoryInfo directory, string fileName)
        {
            return new FileInfo(Path.Combine(directory.FullName, "CSV", fileName + ".csv"));
        }
    }
}
