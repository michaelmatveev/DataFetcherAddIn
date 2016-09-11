using StructureMap;

namespace DataFetcherAddIn
{
    public abstract class DocumentController : IController
    {
        protected IContainer Container { get; set; }

        protected DocumentController()
        {
            Container = new Container();
        }

        public T GetInstance<T>()
        {
            return Container.GetInstance<T>();
        }

    }
}
