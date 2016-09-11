using Microsoft.Office.Interop.Word;

namespace DataFetcherAddIn.DfFields
{
    public abstract class DfBase
    {
        protected readonly string _switcher;
        public DfBase(string switcher)
        {
            _switcher = switcher;
        }

        public abstract void Update(Range range);
    }
}
