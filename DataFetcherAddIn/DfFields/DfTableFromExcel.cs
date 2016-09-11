using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace DataFetcherAddIn.DfFields
{
    public class DfTableFromExcel : DfExcel
    {
        private static Regex tableWidthFinder = new Regex("TableWidth=\"(.*?)\"");
        private static Regex tableHeightFinder = new Regex("TableHeight=\"(.*?)\"");

        protected readonly int _tableWidth;
        protected readonly int _tableHeight;

        public DfTableFromExcel(string switcher, AliasParser parser) : base(switcher, parser)
        {
            _tableWidth =  ExpressionEvaluator.Eval(GetProperty(_switcher, tableWidthFinder));
            _tableHeight = ExpressionEvaluator.Eval(GetProperty(_switcher, tableHeightFinder));
        }

        public override void Update(Microsoft.Office.Interop.Word.Range range)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = false;
            var workbook = app.Workbooks.Open(_pathToFile);
            var worksheet = (Worksheet)workbook.Worksheets[_worksheet];
            Microsoft.Office.Interop.Excel.Range tl = worksheet.Cells[_row, _column];
            var tableRange = worksheet.Range[tl, tl.Offset[_tableHeight - 1, _tableWidth - 1]];
            tableRange.Copy();
            if (range.Tables.Count == 1)
            {
                range.Tables[1].Delete();
            }
            range.PasteExcelTable(false, false, false);
            //app.Workbooks.Close();
            workbook.Close(false);
            app.Quit();
        }

    }
}
