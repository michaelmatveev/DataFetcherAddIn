using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace DataFetcherAddIn.DfFields
{
    public class DfTextFromExcel : DfExcel
    {
        private static Regex processTextFinder = new Regex("ParseText=\"(.*?)\"");
        protected readonly bool _processText;

        public DfTextFromExcel(string switcher, AliasParser parser) : base(switcher, parser)
        {
            _processText = bool.Parse(GetProperty(_switcher, processTextFinder));
        }

        public override void Update(Microsoft.Office.Interop.Word.Range range)
        {
            Microsoft.Office.Interop.Excel.Application app;
            app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = false;
            var workbook = app.Workbooks.Open(_pathToFile);
            var worksheet = (Worksheet)workbook.Worksheets[_worksheet];
            Microsoft.Office.Interop.Excel.Range excelRange = worksheet.Cells[_row, _column];
            string value = excelRange.Text;
            range.Text = _parser.Parse(value ?? string.Empty);
            workbook.Close(false);
         //   app.Workbooks.Close();
            app.Quit();
        }
    }
}
