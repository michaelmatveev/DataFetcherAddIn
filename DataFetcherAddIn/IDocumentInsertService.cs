namespace DataFetcherAddIn
{
    public interface IDocumentService
    {
        void ShowFieldCodes(bool showCodes);

        void InsertAlias(Alias alias);
        void InsertPropertyField(Alias alias);

        void InsertTextFromExcelField(string fileName, string sheet, int row, int column, bool parseText);
        void InsertTableFromExcelField(string fileName, string sheet, int row, int column, int width, int height);
        void InsertDiagrammFromExcelField(string fileName, string sheet, string name);
    }
}
