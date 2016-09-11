using System.Text.RegularExpressions;

namespace DataFetcherAddIn.DfFields
{
    public abstract class DfExcel : DfBase
    {
        private static Regex pathFileFinder = new Regex("File=\"(.*?)\"");
        private static Regex workSheetFinder = new Regex("Sheet=\"(.*?)\"");
        private static Regex rowFinder = new Regex("Row=\"(.*?)\"");
        private static Regex columnFinder = new Regex("Column=\"(.*?)\"");

        protected readonly AliasParser _parser;
        protected readonly string _pathToFile;
        protected readonly string _worksheet;
        protected readonly int _row;
        protected readonly int _column;

        public DfExcel(string switcher, AliasParser parser, bool bypassRowAndCol = false) : base(parser.Parse(switcher))
        {
            _parser = parser;

            _pathToFile = GetProperty(_switcher, pathFileFinder);
            _worksheet = GetProperty(_switcher, workSheetFinder);
            if (bypassRowAndCol)
            {
                return;
            }
            _row = ExpressionEvaluator.Eval(GetProperty(_switcher, rowFinder));
            _column = ExpressionEvaluator.Eval(GetProperty(_switcher, columnFinder));
        }

        protected static string GetProperty(string switcher, Regex propertyFinder)
        {
            var match = propertyFinder.Match(switcher);
            if (match.Success && match.Groups.Count > 1)
            {
                return match.Groups[1].Value;
            }
            return string.Empty;
        }
    }
}
