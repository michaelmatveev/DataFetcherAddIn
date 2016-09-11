using System;
using System.ComponentModel;

namespace DataFetcherAddIn
{
    public class Alias : INotifyPropertyChanged
    {
        public Alias(string name, string value = null)
        {
            _name = name;
            _value = value;
        }

        private string _name;     
        public string Name
        {
            get
            {
                return _name;
            }
            set
            {
                if (_name != value)
                {
                    _name = value;
                    OnNotifyPropertyChanged(nameof(Name));
                }
            }
        }

        private string _value;
        public string Value
        {
            get
            {
                return _value;
            }
            set
            {
                if(_value != value)
                {
                    _value = value;
                    OnNotifyPropertyChanged(nameof(Value));
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnNotifyPropertyChanged(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        public string Sample { get; set; }
        public string MatchExpression { get; set; }
        public string Description { get; set; }
        public Func<DataHolder, string, string> ValueGetter { get; set; }
    }
}
