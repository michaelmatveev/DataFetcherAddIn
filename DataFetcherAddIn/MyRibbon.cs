using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace DataFetcherAddIn
{
    public partial class MyRibbon : IMainView
    {
        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        public event EventHandler RefreshAllDataRequested;
        protected void OnRefreshAllDataRequested()
        {
            RefreshAllDataRequested?.Invoke(this, EventArgs.Empty);
        }

        public event EventHandler ToggleShowMode;
        protected void OnToggleShowMode()
        {
            ToggleShowMode?.Invoke(this, EventArgs.Empty);
        }

        public bool ShowCodes
        {
            get
            {
                return toggleButtonShowFields.Checked;
            }
        }

        public event EventHandler<ShowSubstitutionArg> ShowSubstitutions;
        protected void OnShowSubstitution()
        {            
            ShowSubstitutions?.Invoke(this, new ShowSubstitutionArg
            {
                MainWindowHandle = Globals.ThisAddIn.Application.ActiveWindow.Hwnd
            });
        }

        public event EventHandler InsertText;
        protected void OnInsertText()
        {
            InsertText?.Invoke(this, EventArgs.Empty);
        }

        public event EventHandler InsertTable;
        protected void OnInsertTable()
        {
            InsertTable?.Invoke(this, EventArgs.Empty);
        }

        public event EventHandler InsertDiagramm;
        protected void OnInsertDiagramm()
        {
            InsertDiagramm?.Invoke(this, EventArgs.Empty);
        }

        private void buttonRefreshData_Click(object sender, RibbonControlEventArgs e)
        {
            OnRefreshAllDataRequested();
        }

        public void SetControlState(bool enabled)
        {
            foreach (var item in this.group1.Items.Union(group2.Items))
            {
                item.Enabled = enabled;
            }
        }

        private void toggleButtonShowPanel_Click(object sender, RibbonControlEventArgs e)
        {
            OnShowSubstitution();
         }

        private void buttonTextFromExcel_Click(object sender, RibbonControlEventArgs e)
        {
            OnInsertText();
        }

        private void buttonTable_Click(object sender, RibbonControlEventArgs e)
        {
            OnInsertTable();
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            OnInsertDiagramm();
        }

        private void toggleButtonShowFields_Click(object sender, RibbonControlEventArgs e)
        {
            OnToggleShowMode();
        }
    }

}
