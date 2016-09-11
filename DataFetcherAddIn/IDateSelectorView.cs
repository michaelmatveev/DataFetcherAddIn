using System;

namespace DataFetcherAddIn
{
    public interface IDateSelectorView : IView
    {
        DateTime GetSelectedDate(DateTime currentDate);
    }
}
