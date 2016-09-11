using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;

namespace DataFetcherAddIn.DfFields
{
    public class DfDiagrammFromExcel : DfExcel
    {
        private static Regex nameFinder = new Regex("Name=\"(.*?)\"");
        protected readonly string _name;

        public DfDiagrammFromExcel(string switcher, AliasParser parser) : base(switcher, parser, bypassRowAndCol:true)
        {
            _name = GetProperty(_switcher, nameFinder);
        }

        public override void Update(Microsoft.Office.Interop.Word.Range range)
        {
            Microsoft.Office.Interop.Excel.Application app;
            app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = false;
            var workbook = app.Workbooks.Open(_pathToFile);
            var worksheet = (Worksheet)workbook.Worksheets[_worksheet];
            var chart = (ChartObject)worksheet.ChartObjects(_name);
            //chart.Chart.ChartArea.Copy();
            //range.PasteSpecial();
            chart.CopyPicture();
            range.Paste();
            workbook.Close(false);
            app.Quit();
        }
    }
}
