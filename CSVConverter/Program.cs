using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace CSVConverter
{
    class Program
    {
        private static IEnumerable<string> paths;

        [STAThread]
        static void Main(string[] args)
        {
            if (args.Any()) paths = args;
            else
            {
                OpenFileDialog OpenFileDialog = new OpenFileDialog
                {
                    CheckFileExists = true,
                    Multiselect = true
                };
                switch (OpenFileDialog.ShowDialog())
                {
                    case DialogResult.OK:
                    case DialogResult.Yes:
                        paths = OpenFileDialog.FileNames;
                        break;
                    default: return;
                }
            }

            CSVConverter.Convert(paths.Select(p => new FileInfo(p)).Where(f => f.Exists).ToArray());
        }
    }
}
