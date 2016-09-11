using System;

namespace DataFetcherAddIn
{
    public class ShowSubstitutionArg : EventArgs
    {
        public int MainWindowHandle { get; set; }
    } 

    interface IMainView : IView
    {
        event EventHandler RefreshAllDataRequested;
        bool ShowCodes { get; }
        event EventHandler ToggleShowMode;

        event EventHandler<ShowSubstitutionArg> ShowSubstitutions;
        //event EventHandler HideSubstitutions;
        event EventHandler InsertText;
        event EventHandler InsertTable;
        event EventHandler InsertDiagramm;

        void SetControlState(bool enabled);
    }
}
