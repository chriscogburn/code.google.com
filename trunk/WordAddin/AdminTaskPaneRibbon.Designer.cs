using WordAddinExample;
using WordAddinExample.Properties;

namespace WordAddinExample
{
    partial class AdminTaskPaneRibbon
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = new Microsoft.Office.Tools.Ribbon.RibbonTab();
            this.group1 = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            this.ShowTaskPane = new Microsoft.Office.Tools.Ribbon.RibbonToggleButton();
            this.InsertDate = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "MySyncTab";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.ShowTaskPane);
            this.group1.Items.Add(this.InsertDate);
            this.group1.Label = "Sync Group";
            this.group1.Name = "group1";
            // 
            // ShowTaskPane
            // 
            this.ShowTaskPane.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ShowTaskPane.Image = Resources.television_add;
            this.ShowTaskPane.Label = "Show";
            this.ShowTaskPane.Name = "ShowTaskPane";
            this.ShowTaskPane.ShowImage = true;
            this.ShowTaskPane.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.ShowTaskPane_Click);
            // 
            // InsertDate
            // 
            this.InsertDate.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.InsertDate.Image = Resources.calendar_add;
            this.InsertDate.Label = "Insert Date";
            this.InsertDate.Name = "InsertDate";
            this.InsertDate.ShowImage = true;
            this.InsertDate.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.InsertDate_Click);
            // 
            // AdminTaskPaneRibbon
            // 
            this.Name = "AdminTaskPaneRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonUIEventArgs>(this.AdminTaskPaneRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton ShowTaskPane;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton InsertDate;
    }
}

namespace WordAddinExample
{
    partial class ThisRibbonCollection : Microsoft.Office.Tools.Ribbon.RibbonReadOnlyCollection
    {
        internal AdminTaskPaneRibbon AdminTaskPaneRibbon
        {
            get { return this.GetRibbon<AdminTaskPaneRibbon>(); }
        }
    }
}
