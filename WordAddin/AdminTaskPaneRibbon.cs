using System;
using Microsoft.Office.Tools.Ribbon;
using WordAddinExample.Properties;
using WordAddinExample.Properties;

namespace WordAddinExample
{
    public partial class AdminTaskPaneRibbon : OfficeRibbon
    {
        public AdminTaskPaneRibbon()
        {
            InitializeComponent();
        }

        private void AdminTaskPaneRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            // we can check our last saved status here to show the taskpane 
            SetTaskPaneDisplay(Settings.Default.StartupDisplay);
            Close += AdminTaskPaneRibbon_Close;
        }

        private void AdminTaskPaneRibbon_Close(object sender, EventArgs e)
        {
            // save any setting changes
            Settings.Default.StartupDisplay = ShowTaskPane.Checked;
            Settings.Default.Save();
        }

        private void ShowTaskPane_Click(object sender, RibbonControlEventArgs e)
        {
            // this is just normal "figure out what we set" code
            var show = ((RibbonToggleButton) sender).Checked;
            SetTaskPaneDisplay(show);
        }

        private void SetTaskPaneDisplay(bool show)
        {
            Globals.ThisAddIn.AdminTaskPaneVisible = show;
            UpdateToggleButton(show);

            // now try and sync all open task panes
            if (!Globals.ThisAddIn.AdminTaskPaneVisible)
                Globals.ThisAddIn.RemoveAllTaskPanes();
            else
                Globals.ThisAddIn.AddAllTaskPanes();
        }

        public void UpdateToggleButton(bool visible)
        {
            ShowTaskPane.Image = (visible ? Resources.television_delete : Resources.television_add);
            ShowTaskPane.Label = (visible ? "Hide" : "Show");
            ShowTaskPane.Checked = visible;

            InsertDate.Enabled = visible;
        }

        private void InsertDate_Click(object sender, RibbonControlEventArgs e)
        {
            var ctl = Globals.ThisAddIn.AdminTaskPaneControl;
            // just to be safe
            if (ctl != null)
                ctl.InsertDate();
        }
    }
}