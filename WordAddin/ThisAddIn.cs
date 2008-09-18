using System;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using Office = Microsoft.Office.Core;

namespace WordAddinExample
{
    public partial class ThisAddIn
    {
        private const string TaskPaneTitle = "Select a date";

        public bool AdminTaskPaneVisible { get; set; }

        /// <summary>
        /// AdminTaskPane from the active document window.
        /// </summary>
        public AdminTaskPane AdminTaskPaneControl
        {
            get
            {
                if (Globals.ThisAddIn.Application.Documents.Count <= 0) return null;

                // find the one that is currently active; return it
                foreach (var taskPane in CustomTaskPanes)
                    if (taskPane.Window == Globals.ThisAddIn.Application.ActiveDocument.ActiveWindow)
                        return (AdminTaskPane) taskPane.Control;

                return null;
            }
        }

        /// <summary>
        /// Add a task pane to all open document windows
        /// </summary>
        public void AddAllTaskPanes()
        {
            // no open documents? return
            if (Globals.ThisAddIn.Application.Documents.Count <= 0) return;

            // If ShowAllWindowsInTaskbar is enabled then 
            // each open document has its own window.  
            if (Application.ShowWindowsInTaskbar)
                foreach (Document document in Application.Documents)
                    AddTaskPane(document);

            else // If ShowAllWindowsInTaskbar is not selected  
                // Word displays each document in the same window.  
                if (!AdminTaskPaneVisible)
                    AddTaskPane(Application.ActiveDocument);

            AdminTaskPaneVisible = true;
        }

        /// <summary>
        /// Add a task pane to the Document
        /// </summary>
        /// <param name="document"></param>
        public void AddTaskPane(Document document)
        {
            // Create a new custom task pane and add it to the 
            // collection of custom task panes belonging to this add-in
            // The first two arguments of the Add method specify a control to add
            // to the custom task pane and the title to display on the task pane. 
            // The third argument, which is optional, specifies the 
            // parent window for the custom task pane. 
            var taskPaneContainer = CustomTaskPanes.Add(new AdminTaskPane(),
                                                        TaskPaneTitle, document.ActiveWindow);

            // Display the custom task pane.
            taskPaneContainer.Visible = true;

            // sync event
            taskPaneContainer.VisibleChanged += taskPaneContainer_VisibleChanged;
        }

        /// <summary>
        /// Need to handle when the pane is closed via the [X]
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void taskPaneContainer_VisibleChanged(object sender, EventArgs e)
        {
            var visible = ((CustomTaskPane) sender).Visible;
            if (visible)
                AddAllTaskPanes();
            else
                RemoveAllTaskPanes();
            // update the ribbon
            Globals.Ribbons.AdminTaskPaneRibbon.UpdateToggleButton(visible);
        }

        /// <summary>
        /// Remove only our custom Task Pane from all open windows (documents)
        /// </summary>
        public void RemoveAllTaskPanes()
        {
            // return if no docs open
            if (Globals.ThisAddIn.Application.Documents.Count <= 0) return;

            // loop through ALL custom task panes
            for (int i = CustomTaskPanes.Count; i > 0; i--)
            {
                var taskPaneControl = CustomTaskPanes[i - 1];
                // only remove ours 
                if (taskPaneControl.Title == TaskPaneTitle)
                    CustomTaskPanes.RemoveAt(i - 1);
            }

            AdminTaskPaneVisible = false;
        }

        /// <summary>
        /// Show all of our admin Task Panes on all open windows (documents)
        /// </summary>
        public void ToggleAdminTaskPanes(bool show)
        {
            // return if no docs open
            if (Globals.ThisAddIn.Application.Documents.Count <= 0) return;
            // loop through ALL custom task panes
            foreach (var taskPane in CustomTaskPanes)
                taskPane.Visible = show;

            AdminTaskPaneVisible = show;
        }

        /// <summary>
        /// Remove all of TaskPanes that lost their owner
        /// </summary>
        private void RemoveOrphanedTaskPanes()
        {
            // loop through ALL custom task panes
            for (int i = CustomTaskPanes.Count; i > 0; i--)
            {
                var taskPaneControl = CustomTaskPanes[i - 1];
                // no associated window? remove it
                if (taskPaneControl.Window == null)
                    CustomTaskPanes.Remove(taskPaneControl);
            }
        }

        /// <summary>
        /// Attach the TaskPane each time a document is opened.
        /// </summary>
        /// <param name="document"></param>
        private void Application_DocumentOpen(Document document)
        {
            RemoveOrphanedTaskPanes();
            if (AdminTaskPaneVisible && Application.ShowWindowsInTaskbar)
                AddTaskPane(document);
        }

        /// <summary>
        /// Attach the TaskPane each time a new document is created.
        /// </summary>
        /// <param name="document"></param>
        private void Application_NewDocument(Document document)
        {
            if (AdminTaskPaneVisible && Application.ShowWindowsInTaskbar)
                AddTaskPane(document);
        }

        /// <summary>
        /// Remove any lost TaskPanes on the document change event.
        /// </summary>
        private void Application_DocumentChange()
        {
            RemoveOrphanedTaskPanes();
        }

        /// <summary>
        /// Sync our events for the lifecycle of the Addin.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Application.DocumentOpen += Application_DocumentOpen;
            Application.DocumentChange += Application_DocumentChange;
            ((ApplicationEvents4_Event) Application).NewDocument += Application_NewDocument;
        }

        /// <summary>
        /// Remove any remaining TaskPanes for our app when the addin is unloaded.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            RemoveAllTaskPanes();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion
    }
}