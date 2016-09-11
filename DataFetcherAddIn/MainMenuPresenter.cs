namespace DataFetcherAddIn
{
    class MainMenuPresenter : Presenter<IMainView>
    {
        private const string FileNamePlaceHolder = "Вставьте путь к файлу";
        private const string WorksheetPlaceHolder = "Вставьте название листа";
        private const string RowPlaceHolder = "Номер строки";
        private const string ColumnPlaceHolder = "Номер столбца";
        private const string ParseHolder = "True - заменить подстановки в полученном текст";
        private const string TableWidthPlaceHolder = "Ширина таблицы в столбцах";
        private const string TableHeightPlaceHolder = "Высота таблицы в строках";
        private const string DiagrammNamePlaceHolder = "Имя диаграммы";

        public static int AttachedDocuments;

        public MainMenuPresenter(IMainView view, IController controller)
            : base(view, controller)
        {
            var currentHandle = Globals.ThisAddIn.Application.ActiveWindow.Hwnd;

            View.RefreshAllDataRequested +=
                (s, e) =>
                {
                    var holder = Controller.GetInstance<IDataHolderStorage>().GetHolder();
                    Controller.GetInstance<DataFetcher>().RefreshAllData(holder);
                };
            View.ToggleShowMode += (s, e) => Controller.GetInstance<IDocumentService>().ShowFieldCodes(View.ShowCodes);

            View.ShowSubstitutions += (s, e) =>
            {
                if (e.MainWindowHandle == currentHandle)
                {
                    Controller.GetInstance<SubstitutePresenter>().Run();
                }
            };

            View.InsertText += (s, e) => 
                Controller
                .GetInstance<IDocumentService>()
                .InsertTextFromExcelField(FileNamePlaceHolder, WorksheetPlaceHolder, 1, 1, true);
            View.InsertTable += (s, e) =>
                Controller
                .GetInstance<IDocumentService>()
                .InsertTableFromExcelField(FileNamePlaceHolder, WorksheetPlaceHolder, 1, 1, 1, 1);
            View.InsertDiagramm += (s, e) =>
                Controller
                .GetInstance<IDocumentService>()
                .InsertDiagrammFromExcelField(FileNamePlaceHolder, WorksheetPlaceHolder, DiagrammNamePlaceHolder);
        }

        public void Run()
        {
            AttachedDocuments++;
            View.SetControlState(enabled: true);
        }

        public void Stop()
        {
            Controller.GetInstance<SubstitutePresenter>().Stop();
            if (--AttachedDocuments == 0)
            {
                View.SetControlState(enabled: false);
            }
        }
    }
}
