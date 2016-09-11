using Microsoft.Office.Interop.Word;

namespace DataFetcherAddIn.DfFields
{
    public class DfAlias : DfBase
    {
        private readonly AliasParser _parser;
        public DfAlias(string switcher, AliasParser parser) : base(switcher)
        {
            _parser = parser;
        }

        public override void Update(Range range)
        {
            range.Text = _parser.Parse(_switcher);
        }
    }
}
