namespace WordAddinExample
{
    partial class AdminTaskPane
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
            this.label1 = new System.Windows.Forms.Label();
            this.InsertDateButton = new System.Windows.Forms.Button();
            this.calendar = new System.Windows.Forms.MonthCalendar();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(3, 6);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(156, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Whiz-Bang Date Picker Thingie";
            this.label1.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // InsertDateButton
            // 
            this.InsertDateButton.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.InsertDateButton.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.InsertDateButton.Location = new System.Drawing.Point(39, 196);
            this.InsertDateButton.Name = "InsertDateButton";
            this.InsertDateButton.Size = new System.Drawing.Size(141, 23);
            this.InsertDateButton.TabIndex = 2;
            this.InsertDateButton.Text = "Insert Date";
            this.InsertDateButton.UseVisualStyleBackColor = true;
            this.InsertDateButton.Click += new System.EventHandler(this.InsertDateButton_Click);
            // 
            // calendar
            // 
            this.calendar.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.calendar.Location = new System.Drawing.Point(2, 22);
            this.calendar.Name = "calendar";
            this.calendar.TabIndex = 3;
            // 
            // AdminTaskPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.calendar);
            this.Controls.Add(this.InsertDateButton);
            this.Controls.Add(this.label1);
            this.Name = "AdminTaskPane";
            this.Size = new System.Drawing.Size(231, 318);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button InsertDateButton;
        private System.Windows.Forms.MonthCalendar calendar;
    }
}