using DataFetcherAddIn.DfFields;
using Microsoft.Office.Interop.Word;

namespace DataFetcherAddIn
{
    public class DocumentInsertService : IDocumentService
    {
        private readonly Document _document;
        private readonly Selection _selection;
        public DocumentInsertService(Document document, Selection selection)
        {
            _document = document;
            _selection = selection; 
        }

        public void ShowFieldCodes(bool showCodes)
        {
            foreach(Field field in _document.Fields)
            {
                field.ShowCodes = showCodes;
            }
        }
          
        public void InsertAlias(Alias alias)
        {
            _selection.Range.Text = alias.Name;
        }

        public void InsertPropertyField(Alias alias)
        {
            SetField(AliasParser.GetFieldText(alias));
        }

        public void InsertTextFromExcelField(string fileName, string sheet, int row, int column, bool parseText)
        {
            var fieldText = $"{FieldsCodes.DfTextFromExcel} File=\"{fileName}\" Sheet=\"{sheet}\" Row=\"{row}\" Column=\"{column}\" ParseText=\"{parseText}\" ";
            SetField(fieldText);
        }

        public void InsertTableFromExcelField(string fileName, string sheet, int row, int column, int width, int height)
        {
            var fieldText = $"{FieldsCodes.DfTableFromExcel} File=\"{fileName}\" Sheet=\"{sheet}\" Row=\"{row}\" Column=\"{column}\" TableWidth=\"{width}\" TableHeight=\"{height}\" ";
            SetField(fieldText);
        }

        public void InsertDiagrammFromExcelField(string fileName, string sheet, string name)
        {
            var fieldText = $"{FieldsCodes.DfDiagrammFromExcel} File=\"{fileName}\" Sheet=\"{sheet}\" Name=\"{name}\" ";
            SetField(fieldText);
        }


        private void SetField(string fieldText)
        {
            _document
                .Fields
                .Add(_selection.Range, WdFieldType.wdFieldDocProperty, fieldText)
                .ShowCodes = false;
        }
    }
}
