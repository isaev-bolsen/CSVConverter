using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace CSVConverter
{
    class Page
    {
        private const string separator = ";";

        private readonly int _rowsCount;
        private readonly int _columnsCount;
        private readonly ExcelWorksheet _sheet;

        public Page(ExcelWorksheet sheet)
        {
            _rowsCount = sheet.Dimension.End.Row;
            _columnsCount = sheet.Dimension.End.Column;
            _sheet = sheet;
        }

        public FileInfo Save(FileInfo path)
        {
            if (!path.Directory.Exists) path.Directory.Create();

            using (StreamWriter stream = path.CreateText())
            {
                for (int i = 1; i <= _rowsCount; ++i) stream.WriteLine(GetLine(i));
                stream.Close();
            }

            return path;
        }

        private string GetLine(int i)
        {
            return string.Join(separator, GetRow(i));
        }

        private IEnumerable<string> GetRow(int i)
        {
            for (int j = 1; j <= _columnsCount; ++j) yield return GetValueForOutput(i, j);
        }

        private string GetValueForOutput(int i, int j)
        {
            string value = GetValue(i, j);
            if (value == null) return string.Empty;
            else if (value.Contains(separator)) return string.Concat("\"", value, "\"");
            else return value;
        }

        private string GetValue(int i, int j)
        {
            var cells = _sheet.Cells[i, j];
            if (cells.Value != null && cells.Style.Numberformat.Format == "m\\/d\\/yyyy h:mm AM/PM")
                return DateTime.FromOADate((double)cells.Value).ToString();
            else return cells.Text;
        }
    }
}
