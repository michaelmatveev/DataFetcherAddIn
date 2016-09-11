using DataFetcherAddIn.DfFields;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace DataFetcherAddIn
{
    public class DataFetcher
    {
        private readonly Document _document;
        private static Dictionary<string, Func<string, DataHolder, DfBase>> _mapper 
            = new Dictionary<string, Func<string, DataHolder, DfBase>>()
        {
            { FieldsCodes.DfTest,               (s, h) => new DfTest(s) },
            { FieldsCodes.DfAlias,              (s, h) => new DfAlias(s, new AliasParser(h)) },
            { FieldsCodes.DfTextFromExcel,      (s, h) => new DfTextFromExcel(s, new AliasParser(h)) },
            { FieldsCodes.DfTableFromExcel,     (s, h) => new DfTableFromExcel(s, new AliasParser(h)) },
            { FieldsCodes.DfDiagrammFromExcel,  (s, h) => new DfDiagrammFromExcel(s, new AliasParser(h)) }
        };

        public DataFetcher(Document document)
        {
            _document = document;
        }

        private static Regex fieldFinder = new Regex(@"\s*DOCPROPERTY\s*df\w*");
        private static Regex propertyFinder = new Regex(@"df\w*");
        private static Regex switcherFinder = new Regex(@"df\w*(.*)\\\*");

        private static string GetCustomTag(string tagText)
        {
            var match = fieldFinder.Match(tagText);
            if(match.Success)
            {
                return propertyFinder.Match(match.Value).Value;
            }
            return string.Empty;
        }

        private static string GetSwitcher(string tagText)
        {
            var match = switcherFinder.Match(tagText);
            if (match.Success && match.Groups.Count > 1)
            {
                return match.Groups[1].Value;
            }
            return string.Empty;
        }

        public void RefreshAllData(DataHolder holder)
        {
            foreach(Field field in _document.Fields)
            {
                RefreshField(field, holder);
            }
        }

        public void RefreshField(Field field, DataHolder holder)
        {
            var customTag = GetCustomTag(field.Code.Text);
            Func<string, DataHolder, DfBase> dfUpdaterFactory;
            if (_mapper.TryGetValue(customTag, out dfUpdaterFactory))
            {
                var switcher = GetSwitcher(field.Code.Text).Trim();
                var dfUpdater = dfUpdaterFactory.Invoke(switcher, holder);
                dfUpdater.Update(field.Result);
            }
            else
            {
                field.Update();
            }
        }

    }
}
