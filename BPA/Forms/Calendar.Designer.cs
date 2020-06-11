namespace BPA.Forms
{
    partial class Calendar
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.Calendar_Control = new System.Windows.Forms.MonthCalendar();
            this.Date_TextBox = new System.Windows.Forms.TextBox();
            this.Ok_Button = new System.Windows.Forms.Button();
            this.Cancel_Button = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // Calendar_Control
            // 
            this.Calendar_Control.FirstDayOfWeek = System.Windows.Forms.Day.Monday;
            this.Calendar_Control.Location = new System.Drawing.Point(12, 44);
            this.Calendar_Control.MaxSelectionCount = 1;
            this.Calendar_Control.MinDate = new System.DateTime(1900, 1, 1, 0, 0, 0, 0);
            this.Calendar_Control.Name = "Calendar_Control";
            this.Calendar_Control.ShowWeekNumbers = true;
            this.Calendar_Control.TabIndex = 0;
            this.Calendar_Control.TabStop = false;
            // 
            // Date_TextBox
            // 
            this.Date_TextBox.Location = new System.Drawing.Point(12, 12);
            this.Date_TextBox.Name = "Date_TextBox";
            this.Date_TextBox.Size = new System.Drawing.Size(186, 20);
            this.Date_TextBox.TabIndex = 1;
            this.Date_TextBox.TextChanged += new System.EventHandler(this.Date_TextBox_TextChanged);
            this.Date_TextBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Date_TextBox_KeyDown);
            this.Date_TextBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.Date_TextBox_KeyPress);
            // 
            // Ok_Button
            // 
            this.Ok_Button.Location = new System.Drawing.Point(12, 209);
            this.Ok_Button.Name = "Ok_Button";
            this.Ok_Button.Size = new System.Drawing.Size(75, 23);
            this.Ok_Button.TabIndex = 2;
            this.Ok_Button.Text = "Ok";
            this.Ok_Button.UseVisualStyleBackColor = true;
            // 
            // Cancel_Button
            // 
            this.Cancel_Button.Location = new System.Drawing.Point(121, 209);
            this.Cancel_Button.Name = "Cancel_Button";
            this.Cancel_Button.Size = new System.Drawing.Size(75, 23);
            this.Cancel_Button.TabIndex = 3;
            this.Cancel_Button.Text = "Отмена";
            this.Cancel_Button.UseVisualStyleBackColor = true;
            // 
            // Calendar
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(208, 236);
            this.Controls.Add(this.Cancel_Button);
            this.Controls.Add(this.Ok_Button);
            this.Controls.Add(this.Date_TextBox);
            this.Controls.Add(this.Calendar_Control);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "Calendar";
            this.Text = "Календарь";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MonthCalendar Calendar_Control;
        private System.Windows.Forms.TextBox Date_TextBox;
        private System.Windows.Forms.Button Ok_Button;
        private System.Windows.Forms.Button Cancel_Button;
    }
}