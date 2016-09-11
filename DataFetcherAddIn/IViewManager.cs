namespace DataFetcherAddIn
{
    public interface IViewManager
    {
        void AddView(IView view, string caption);
        void RemoveView(IView view);
    }
}
