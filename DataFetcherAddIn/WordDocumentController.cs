using DataFetcherAddIn.DfFields;
using DataFetcherAddIn.Vidgets;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;

namespace DataFetcherAddIn
{
    public class WordDocumentController : DocumentController
    {
        public MyRibbon MainRibbon { get; set; }
        public CustomTaskPaneCollection PaneCollection { get; set; }
        public Document CurrentDocument { get; set; }
        public IDocumentService DocumentInsertService { get; set; }

        public void Configure()
        {
            Container.Configure(_ =>
            {
                _.For<IController>()
                    .Use(this);

                _.For<IMainView>()
                    .Use(MainRibbon);

                _.ForConcreteType<MainMenuPresenter>()
                    .Configure
                    .Singleton();

                _.For<IViewManager>()
                    .Use<PanesManager>()
                    .Ctor<CustomTaskPaneCollection>()
                    .Is(PaneCollection)
                    .Singleton();

                _.For<ISubstituteView>()
                    .Use<SubstituteControl>();

                _.ForConcreteType<SubstitutePresenter>()
                    .Configure
                    .Singleton();

                _.For<DataFetcher>()
                    .Use<DataFetcher>()
                    .Ctor<Document>()
                    .Is(CurrentDocument);

                _.For<IDataHolderStorage>()
                    .Use<DataHolderStorage>()
                    .Ctor<Document>()
                    .Is(CurrentDocument);
                    //.Singleton();

                _.For<IDocumentService>()
                    .Use(DocumentInsertService);
            });
        }

        private const string FiedDefaultValue = "Чтобы обновить это поле откройте вкладку \"Сбор данных\" и нажмите кнопку обновить";
        public void PrepareCustomProperties()
        {
            var hs = GetInstance<IDataHolderStorage>();
            hs.CreateOrUpdateCustomProperty(FieldsCodes.DfAlias, FiedDefaultValue);
            hs.CreateOrUpdateCustomProperty(FieldsCodes.DfTextFromExcel, FiedDefaultValue);
            hs.CreateOrUpdateCustomProperty(FieldsCodes.DfTableFromExcel, FiedDefaultValue);
            hs.CreateOrUpdateCustomProperty(FieldsCodes.DfDiagrammFromExcel, FiedDefaultValue);
        }
    }
}
