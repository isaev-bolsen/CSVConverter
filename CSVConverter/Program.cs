using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml.Core.ExcelPackage;
using System.IO;

namespace CSVConverter
{
    class Program
    {
        private static List<FileInfo> _filesToProcess = new List<FileInfo>();

        static void Main(string[] args)
        {
            foreach (string path in args) ParsePath(path);
        }

        private static void ParsePath(string path)
        {
            try
            {
                FileInfo FileInfo = new FileInfo(path);
                if (FileInfo.Exists)
                {
                    _filesToProcess.Add(FileInfo);
                    Console.WriteLine("File found: " + FileInfo.FullName);
                    return;
                }
            }
            catch { }
            Console.WriteLine("File NOT found: " + path);
        }
    }
}
