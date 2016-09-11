namespace DataFetcherAddIn
{
    public class SubstitutePresenter : Presenter<ISubstituteView>
    {
        public SubstitutePresenter(ISubstituteView view, IController controller) 
            : base(view, controller)
        {
            view.AliasUpdated += (s, e) => Controller.GetInstance<IDataHolderStorage>().SetHolder((DataHolder)s);

            view.SubstitutionSelected += (s, e) => 
            {
                if(e.InsertAsField)
                {
                    Controller.GetInstance<IDocumentService>().InsertPropertyField(e.SelectedAlias);
                }
                else
                {
                    Controller.GetInstance<IDocumentService>().InsertAlias(e.SelectedAlias);
                }
            };
        }

        public void Run()
        {
            View.ShowIt(Controller.GetInstance<IDataHolderStorage>().GetHolder());
        }

        public void Stop()
        {
            View.HideIt();
        }

    }
}
