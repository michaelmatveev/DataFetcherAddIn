using Microsoft.Office.Interop.Word;

namespace DataFetcherAddIn.DfFields
{
    public class DfTest : DfBase
    {
        public DfTest(string switcher) : base(switcher)
        {
        }

        public override void Update(Range range)
        {
            range.Text = "Updated from Addin";
        }
    }
}
