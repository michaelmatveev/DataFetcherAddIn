using System;
using System.Data;

namespace DataFetcherAddIn
{
    public static class ExpressionEvaluator
    {
        public static int Eval(string expression)
        {
            var loDataTable = new DataTable();
            var loDataColumn = new DataColumn("Eval", typeof(double), expression);
            loDataTable.Columns.Add(loDataColumn);
            loDataTable.Rows.Add(0);
            return Convert.ToInt32(Math.Round((double)loDataTable.Rows[0]["Eval"]));
        }
    }
}
