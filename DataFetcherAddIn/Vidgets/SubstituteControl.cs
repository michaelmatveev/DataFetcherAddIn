using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DataFetcherAddIn.Vidgets
{
    public partial class SubstituteControl : UserControl, ISubstituteView
    {
        private readonly IViewManager _manager;
        private BindingList<Alias> _aliases;

        public SubstituteControl(IViewManager manager)
        {
            _manager = manager;
            _aliases = new BindingList<Alias>(AliasParser.SupportedMappings);
            
            InitializeComponent();
            dataGridView.AutoGenerateColumns = false;
            dataGridView.DataSource = _aliases;
        }

        public void ShowIt(DataHolder dataHolder)
        {
            dateTimePicker.Value = dataHolder.Date;
            comboBoxMode.SelectedIndex = dataHolder.InsertMode;
            _manager.AddView(this, "Подстановки");
            FillAliases(dataHolder);
        }

        public void HideIt()
        {
            _manager.RemoveView(this);
        }

        public event EventHandler<SubstitutionSelectedEventArg> SubstitutionSelected;
        protected void OnSubstitutionSelected()
        {
            var arg = new SubstitutionSelectedEventArg
            {
                SelectedAlias = (Alias)dataGridView.SelectedRows[0].DataBoundItem,
                InsertAsField = comboBoxMode.SelectedIndex == 0
            };
            SubstitutionSelected?.Invoke(this, arg);
        }

        private void FillAliases(DataHolder holder)
        {
            var parser = new AliasParser(holder);
            foreach(var a in _aliases)
            {
                a.Value = parser.Parse(a.Sample);
            }
        }

        public event EventHandler AliasUpdated;
        protected void OnAliasUpdated(DataHolder holder)
        {            
            AliasUpdated?.Invoke(holder, EventArgs.Empty);
        }

        private DataHolder GetDataHolder()
        {
            return new DataHolder
            {
                Date = dateTimePicker.Value,
                Alias1 = textBoxAlias1.Text,
                Alias2 = textBoxAlias2.Text,
                Alias3 = textBoxAlias3.Text,
                Alias4 = textBoxAlias4.Text,
                InsertMode = comboBoxMode.SelectedIndex
            };
        }

        private void dateTimePicker_ValueChanged(object sender, EventArgs e)
        {
            var holder = GetDataHolder();            
            FillAliases(holder);
            OnAliasUpdated(holder);
        }

        private void dataGridView_DoubleClick(object sender, EventArgs e)
        {
            OnSubstitutionSelected();
        }

        private void comboBoxMode_SelectedIndexChanged(object sender, EventArgs e)
        {
            var holder = GetDataHolder();
            OnAliasUpdated(holder);
        }

        private void dataGridView_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            var cell = dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex];
            cell.ToolTipText = _aliases[e.RowIndex].Description;
        }

    }
}
