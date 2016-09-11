using System;

namespace DataFetcherAddIn
{
    public class SubstitutionSelectedEventArg : EventArgs
    {
        public Alias SelectedAlias { get; set; }
        public bool InsertAsField { get; set; }
    }

    public interface ISubstituteView : IView
    {
        void ShowIt(DataHolder dataHolder);
        void HideIt();

        event EventHandler AliasUpdated;
        event EventHandler<SubstitutionSelectedEventArg> SubstitutionSelected;        
    }
}
