using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace CSVConverter
{
    public static class CSVConverter
    {
        public static IEnumerable<FileInfo> Convert(params FileInfo[] files)
        {
            return files.SelectMany(ConvertFile).ToArray();
        }

        private static IEnumerable<FileInfo> ConvertFile(FileInfo file)
        {
            string Filename = Path.GetFileNameWithoutExtension(file.Name);
            ExcelWorksheet[] worksheets = new ExcelPackage(file).Workbook.Worksheets.Cast<ExcelWorksheet>().ToArray();
            if (worksheets.Count() == 1) yield return new Page(worksheets.Single()).Save(GetResultPath(file.Directory, Filename));
            else foreach (ExcelWorksheet ws in worksheets)
                    yield return new Page(ws).Save(GetResultPath(file.Directory, string.Join("_", Filename, ws.Name)));
        }

        private static FileInfo GetResultPath(DirectoryInfo directory, string fileName)
        {
            return new FileInfo(Path.Combine(directory.FullName, "CSV", fileName + ".csv"));
        }
    }
}
