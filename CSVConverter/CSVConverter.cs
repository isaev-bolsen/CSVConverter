using OfficeOpenXml.Core.ExcelPackage;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace CSVConverter
{
    public static class CSVConverter
    {
        public static void Convert(params FileInfo[] files)
        {
            foreach (ExcelPackage package in GetPackages(files))
            {
                foreach (XElement shitElement in
                    package.Workbook.WorkbookXml.Descendants(package.Workbook.WorkbookXml.Document.Elements().First().Name.Namespace + "sheet"))
                {
                    IEnumerable<XAttribute> attributes = shitElement.Attributes();
                    foreach (XAttribute attr in attributes) attr.Remove();
                    shitElement.Add(attributes.Take(2));
                }

                ExcelWorksheet[] worksheets = package.Workbook.Worksheets.Cast<ExcelWorksheet>().ToArray();
                var props = package.Package.PackageProperties;

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
