namespace DataFetcherAddIn
{
    interface IDataHolderStorage
    {
        void CreateOrUpdateCustomProperty(string propertyName, string value = "");
        DataHolder GetHolder();
        void SetHolder(DataHolder holder);
    }
}
