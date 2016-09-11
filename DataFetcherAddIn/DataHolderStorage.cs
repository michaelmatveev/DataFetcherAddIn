using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.IO;
using System.Xml.Serialization;

namespace DataFetcherAddIn
{
    public class DataHolderStorage : IDataHolderStorage
    {
        private readonly Document _document;
        public DataHolderStorage(Document document)
        {
            _document = document;
        }

        private const string CustomPropertyPartId = "dfStoragePartId";

        public DataHolder GetHolder()
        {
            try
            {
                var customPartId = _document.CustomDocumentProperties[CustomPropertyPartId];
                var part = _document.CustomXMLParts.SelectByID(customPartId);
                if (part == null)
                {
                    SetHolder(DataHolder.Default);
                    return GetHolder();
                }
                var serializer = new XmlSerializer(typeof(DataHolder));
                using (var reader = new StringReader(part.DocumentElement.XML))
                {
                    return (DataHolder)serializer.Deserialize(reader);
                }
            }
            catch (Exception)
            {
                SetHolder(DataHolder.Default);
                return GetHolder();
            }
        }

        public void SetHolder(DataHolder holder)
        {
            var serializer = new XmlSerializer(typeof(DataHolder));
            using (var writer = new StringWriter())
            {
                serializer.Serialize(writer, holder);
                var part = _document.CustomXMLParts.Add(writer.ToString());
                CreateOrUpdateCustomProperty(CustomPropertyPartId, part.Id);
            }
        }

        public void CreateOrUpdateCustomProperty(string propertyName, string value = "")
        {
            try
            {
                _document.CustomDocumentProperties.Add(propertyName, false,
                    MsoDocProperties.msoPropertyTypeString, value);
            }
            catch (Exception)
            {
                _document.CustomDocumentProperties[propertyName] = value;
            }

        }

    }
}
