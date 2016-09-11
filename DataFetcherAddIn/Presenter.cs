namespace DataFetcherAddIn
{
    public abstract class Presenter<V>
        where V : IView
    {
        protected readonly V View;
        protected readonly IController Controller;

        public Presenter(V view, IController controller)
        {
            View = view;
            Controller = controller;
        }

        //public virtual void Run()
        //{
        //}

        //public virtual void Stop()
        //{
        //}

    }
}
