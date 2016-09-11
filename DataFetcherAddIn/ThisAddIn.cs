using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using DataFetcherAddIn.Vidgets;

namespace DataFetcherAddIn
{
    public partial class ThisAddIn
    {
        private static Dictionary<Word.Document, DocumentController> DocumentControllers = new Dictionary<Word.Document, DocumentController>();

        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.DocumentChange += Application_DocumentChange;
            //Application.DocumentOpen += Application_DocumentOpen;
            Application.DocumentBeforeClose += Application_DocumentBeforeClose;
 //           Application.DocumentBeforeSave += Application_DocumentBeforeSave;
        }

        private void Application_DocumentChange()
        {
            foreach (Word.Document document in Application.Documents)
            {
                if (!DocumentControllers.ContainsKey(document))
                {
                    CreateDocumentController(document);
                }
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void Application_DocumentOpen(Word.Document document)
        {
            CreateDocumentController(document);
        }

        private void Application_DocumentBeforeClose(Word.Document document, ref bool cancel)
        {
            var controller = DocumentControllers[document];
            controller.GetInstance<MainMenuPresenter>().Stop();
            //CreateOrUpdateDefaultCustomPart(document, controller.SelectedDate, controller.ObjectName);
            DocumentControllers.Remove(document);
        }

        //private void Application_DocumentBeforeSave(Word.Document document, ref bool SaveAsUI, ref bool Cancel)
        //{
        //    //var controller = DocumentControllers[document];
        //    //CreateOrUpdateDefaultCustomPart(document, controller.SelectedDate, controller.ObjectName);
        //}

        private void CreateDocumentController(Word.Document document)
        {
            var controller = new WordDocumentController()
            {
                MainRibbon = Globals.Ribbons.MyRibbon,
                CurrentDocument = document,
                DocumentInsertService = new DocumentInsertService(document, Application.Selection),
                PaneCollection = CustomTaskPanes
            };
            controller.Configure();
            controller.PrepareCustomProperties();            
            DocumentControllers.Add(document, controller);
            controller.GetInstance<MainMenuPresenter>().Run();
        }

    }
}
