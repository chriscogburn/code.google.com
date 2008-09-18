using System;
using System.Windows.Forms;

namespace WordAddinExample
{
    public partial class AdminTaskPane : UserControl
    {
        public AdminTaskPane()
        {
            InitializeComponent();
        }

        private void InsertDateButton_Click(object sender, EventArgs e)
        {
            InsertDate();
        }

        public void InsertDate()
        {
            var dt = calendar.SelectionRange.Start.ToShortDateString();
            Globals.ThisAddIn.Application.Selection.Range.Text = dt + Environment.NewLine;
        }
    }
}