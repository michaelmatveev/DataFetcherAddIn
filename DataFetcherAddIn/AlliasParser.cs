using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace DataFetcherAddIn
{
    public class AliasParser
    {
        public static List<Alias> SupportedMappings =
            new List<Alias>
            {
                new Alias("[дата]")
                {
                    Sample = "[дата]",
                    Description = "Дата в длинном формате",
                    MatchExpression = @"\[дата\]",
                    ValueGetter = (d, s) => d.Date.ToString("dd MMMM yyyy")
                },
                new Alias("[дт]")
                {
                    Sample = "[дт]",
                    Description = "Дата",
                    MatchExpression = @"\[дт\]",
                    ValueGetter = (d, s) => d.Date.ToString("dd.MM.yyyy")
                },
                new Alias("[гг]")
                {
                    Sample = "[гг]",
                    Description = "Год (2 символа)",
                    MatchExpression = @"\[гг\]",
                    ValueGetter = (d, s) => d.Date.ToString("yy")
                },
                new Alias("[гггг]")
                {
                    Sample = "[гггг]",
                    Description = "Год",
                    MatchExpression = @"\[гггг\]",
                    ValueGetter = (d, s) => d.Date.ToString("yyyy")
                },
                new Alias("[мм]")
                {
                    Sample = "[мм]",
                    Description = "Номер месяца (с ведущим нулем)",
                    MatchExpression = @"\[мм\]",
                    ValueGetter = (d, s) => d.Date.ToString("MM")
                },
                new Alias("[нм]")
                {
                    Sample = "[нм]",
                    Description = "Номер месяца",
                    MatchExpression = @"\[нм\]",
                    ValueGetter = (d, s) => d.Date.Month.ToString()
                },
                new Alias("[мм+1]")
                {
                    Sample = "[мм+1]",
                    Description = "Номер следующего месяца (с ведущим нулем)",
                    MatchExpression = @"\[мм\+1\]",
                    ValueGetter = (d, s) => d.Date.AddMonths(1).ToString("MM")
                },
                new Alias("[мм-1]")
                {
                    Sample = "[мм-1]",
                    Description = "Номер предыдущего месяца (с ведущим нулем)",
                    MatchExpression = @"\[мм-1\]",
                    ValueGetter = (d, s) => d.Date.AddMonths(-1).ToString("MM")
                },
                new Alias("[нм+1]")
                {
                    Sample = "[нм+1]",
                    Description = "Номер следующего месяца",
                    MatchExpression = @"\[нм\+1\]",
                    ValueGetter = (d, s) => d.Date.AddMonths(1).Month.ToString()
                },
                new Alias("[нм-1]")
                {
                    Sample = "[нм-1]",
                    Description = "Номер предыдущего месяца",
                    MatchExpression = @"\[нм-1\]",
                    ValueGetter = (d, s) => d.Date.AddMonths(-1).Month.ToString()
                },
                new Alias("[мммм]")
                {
                    Sample = "[мммм]",
                    Description = "Название месяца по русски",
                    MatchExpression = @"\[мммм\]",
                    ValueGetter = (d, s) => d.Date.ToString("MMMM")
                },
                new Alias("[дд]")
                {
                    Sample = "[дд]",
                    Description = "Номер дня в месяце (с ведущим нулем)",
                    MatchExpression = @"\[дд\]",
                    ValueGetter = (d, s) => d.Date.AddMonths(-1).ToString("dd")
                },
                new Alias("[нд]")
                {
                    Sample = "[нд]",
                    Description = "Номер дня в месяце",
                    MatchExpression = @"\[нд\]",
                    ValueGetter = (d, s) => d.Date.Day.ToString()
                },
                new Alias("[двм]")
                {
                    Sample = "[двм]",
                    Description = "Дней в месяце",
                    MatchExpression = @"\[двм\]",
                    ValueGetter = (d, s) => DateTime.DaysInMonth(d.Date.Year, d.Date.Month).ToString()
                },
                new Alias("[псевдоним1]")
                {
                    Sample = "[псевдоним1]",
                    Description = "Псевдоним 1",
                    MatchExpression = @"\[псевдоним1\]",
                    ValueGetter = (d, s) => d.Alias1 ?? string.Empty
                },
                new Alias("[псевдоним2]")
                {
                    Sample = "[псевдоним2]",
                    Description = "Псевдоним 2",
                    MatchExpression = @"\[псевдоним2\]",
                    ValueGetter = (d, s) => d.Alias2 ?? string.Empty
                },
                new Alias("[псевдоним3]")
                {
                    Sample = "[псевдоним3]",
                    Description = "Псевдоним 3",
                    MatchExpression = @"\[псевдоним3\]",
                    ValueGetter = (d, s) => d.Alias3 ?? string.Empty
                },
                new Alias("[псевдоним4]")
                {
                    Sample = "[псевдоним4]",
                    Description = "псевдоним 4",
                    MatchExpression = @"\[псевдоним4\]",
                    ValueGetter = (d, s) => d.Alias4 ?? string.Empty
                },
            };

        public static string GetFieldText(Alias alias)
        {
            return $"dfAlias {alias.Name}";
        }

        private readonly DataHolder _dataHolder;
        public AliasParser(DataHolder dataHolder)
        {
            _dataHolder = dataHolder;
        }

        public string Parse(string aliasText)
        {
            foreach(var m in SupportedMappings)
            {
                aliasText = Regex.Replace(aliasText, m.MatchExpression, m.ValueGetter(_dataHolder, string.Empty));
            }
            return aliasText;
        }
    }
}
