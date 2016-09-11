using Microsoft.Office.Tools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DataFetcherAddIn.Vidgets
{
    public class PanesManager : IViewManager
    {
        private readonly CustomTaskPaneCollection _panesCollection;

        public PanesManager(CustomTaskPaneCollection collection)
        {
            _panesCollection = collection;
        }

        public void AddView(IView view, string caption)
        {
            var panel = _panesCollection.SingleOrDefault(p => p.Control == view as Control);
            if (panel == null)
            {
                panel = _panesCollection.Add(view as UserControl, caption);
                panel.Width = 320;
            }
            panel.Visible = true;
        }

        public void RemoveView(IView view)
        {
            var panel = _panesCollection.SingleOrDefault(p => p.Control == view as Control);
            if (panel != null)
            {
                panel.Visible = false;
            }
        }
    }
}
