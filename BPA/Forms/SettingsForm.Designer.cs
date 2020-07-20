namespace BPA.Forms
{
    partial class SettingsForm
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
            this.components = new System.ComponentModel.Container();
            this.ProductCalendarLabel = new System.Windows.Forms.Label();
            this.ProductCalendar_Path_TextBox = new System.Windows.Forms.TextBox();
            this.Budget_Path_TextBox = new System.Windows.Forms.TextBox();
            this.BudgetLabel = new System.Windows.Forms.Label();
            this.Decision_Path_TextBox = new System.Windows.Forms.TextBox();
            this.DecisionLabel = new System.Windows.Forms.Label();
            this.PriceListMT_Path_TextBox = new System.Windows.Forms.TextBox();
            this.PriceListMTlabel = new System.Windows.Forms.Label();
            this.ProductCalendar_SetPath_Button = new System.Windows.Forms.Button();
            this.Budget_SetPath_Button = new System.Windows.Forms.Button();
            this.Decision_SetPath_Button = new System.Windows.Forms.Button();
            this.PriceListMT_SetPath_Button = new System.Windows.Forms.Button();
            this.Cancel_Button = new System.Windows.Forms.Button();
            this.Ok_Button = new System.Windows.Forms.Button();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.errorProvider = new System.Windows.Forms.ErrorProvider(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider)).BeginInit();
            this.SuspendLayout();
            // 
            // ProductCalendarLabel
            // 
            this.ProductCalendarLabel.AutoSize = true;
            this.ProductCalendarLabel.Location = new System.Drawing.Point(6, 15);
            this.ProductCalendarLabel.Name = "ProductCalendarLabel";
            this.ProductCalendarLabel.Size = new System.Drawing.Size(89, 13);
            this.ProductCalendarLabel.TabIndex = 0;
            this.ProductCalendarLabel.Text = "Product Calendar";
            // 
            // ProductCalendar_Path_TextBox
            // 
            this.ProductCalendar_Path_TextBox.Location = new System.Drawing.Point(101, 12);
            this.ProductCalendar_Path_TextBox.Name = "ProductCalendar_Path_TextBox";
            this.ProductCalendar_Path_TextBox.ReadOnly = true;
            this.ProductCalendar_Path_TextBox.Size = new System.Drawing.Size(486, 20);
            this.ProductCalendar_Path_TextBox.TabIndex = 1;
            // 
            // Budget_Path_TextBox
            // 
            this.Budget_Path_TextBox.Location = new System.Drawing.Point(101, 38);
            this.Budget_Path_TextBox.Name = "Budget_Path_TextBox";
            this.Budget_Path_TextBox.ReadOnly = true;
            this.Budget_Path_TextBox.Size = new System.Drawing.Size(486, 20);
            this.Budget_Path_TextBox.TabIndex = 3;
            // 
            // BudgetLabel
            // 
            this.BudgetLabel.AutoSize = true;
            this.BudgetLabel.Location = new System.Drawing.Point(6, 41);
            this.BudgetLabel.Name = "BudgetLabel";
            this.BudgetLabel.Size = new System.Drawing.Size(41, 13);
            this.BudgetLabel.TabIndex = 2;
            this.BudgetLabel.Text = "Budget";
            // 
            // Decision_Path_TextBox
            // 
            this.Decision_Path_TextBox.Location = new System.Drawing.Point(101, 64);
            this.Decision_Path_TextBox.Name = "Decision_Path_TextBox";
            this.Decision_Path_TextBox.ReadOnly = true;
            this.Decision_Path_TextBox.Size = new System.Drawing.Size(486, 20);
            this.Decision_Path_TextBox.TabIndex = 5;
            // 
            // DecisionLabel
            // 
            this.DecisionLabel.AutoSize = true;
            this.DecisionLabel.Location = new System.Drawing.Point(6, 67);
            this.DecisionLabel.Name = "DecisionLabel";
            this.DecisionLabel.Size = new System.Drawing.Size(48, 13);
            this.DecisionLabel.TabIndex = 4;
            this.DecisionLabel.Text = "Decision";
            // 
            // PriceListMT_Path_TextBox
            // 
            this.PriceListMT_Path_TextBox.Location = new System.Drawing.Point(101, 90);
            this.PriceListMT_Path_TextBox.Name = "PriceListMT_Path_TextBox";
            this.PriceListMT_Path_TextBox.ReadOnly = true;
            this.PriceListMT_Path_TextBox.Size = new System.Drawing.Size(486, 20);
            this.PriceListMT_Path_TextBox.TabIndex = 7;
            // 
            // PriceListMTlabel
            // 
            this.PriceListMTlabel.AutoSize = true;
            this.PriceListMTlabel.Location = new System.Drawing.Point(6, 93);
            this.PriceListMTlabel.Name = "PriceListMTlabel";
            this.PriceListMTlabel.Size = new System.Drawing.Size(63, 13);
            this.PriceListMTlabel.TabIndex = 6;
            this.PriceListMTlabel.Text = "PriceListMT";
            // 
            // ProductCalendar_SetPath_Button
            // 
            this.ProductCalendar_SetPath_Button.Location = new System.Drawing.Point(614, 12);
            this.ProductCalendar_SetPath_Button.Name = "ProductCalendar_SetPath_Button";
            this.ProductCalendar_SetPath_Button.Size = new System.Drawing.Size(34, 20);
            this.ProductCalendar_SetPath_Button.TabIndex = 8;
            this.ProductCalendar_SetPath_Button.Text = "...";
            this.ProductCalendar_SetPath_Button.UseVisualStyleBackColor = true;
            this.ProductCalendar_SetPath_Button.Click += new System.EventHandler(this.ProductCalendar_SetPath_Button_Click);
            // 
            // Budget_SetPath_Button
            // 
            this.Budget_SetPath_Button.Location = new System.Drawing.Point(614, 38);
            this.Budget_SetPath_Button.Name = "Budget_SetPath_Button";
            this.Budget_SetPath_Button.Size = new System.Drawing.Size(34, 20);
            this.Budget_SetPath_Button.TabIndex = 9;
            this.Budget_SetPath_Button.Text = "...";
            this.Budget_SetPath_Button.UseVisualStyleBackColor = true;
            this.Budget_SetPath_Button.Click += new System.EventHandler(this.Budget_SetPath_Button_Click);
            // 
            // Decision_SetPath_Button
            // 
            this.Decision_SetPath_Button.Location = new System.Drawing.Point(614, 64);
            this.Decision_SetPath_Button.Name = "Decision_SetPath_Button";
            this.Decision_SetPath_Button.Size = new System.Drawing.Size(34, 20);
            this.Decision_SetPath_Button.TabIndex = 10;
            this.Decision_SetPath_Button.Text = "...";
            this.Decision_SetPath_Button.UseVisualStyleBackColor = true;
            this.Decision_SetPath_Button.Click += new System.EventHandler(this.Decision_SetPath_Button_Click);
            // 
            // PriceListMT_SetPath_Button
            // 
            this.PriceListMT_SetPath_Button.Location = new System.Drawing.Point(614, 90);
            this.PriceListMT_SetPath_Button.Name = "PriceListMT_SetPath_Button";
            this.PriceListMT_SetPath_Button.Size = new System.Drawing.Size(34, 20);
            this.PriceListMT_SetPath_Button.TabIndex = 11;
            this.PriceListMT_SetPath_Button.Text = "...";
            this.PriceListMT_SetPath_Button.UseVisualStyleBackColor = true;
            this.PriceListMT_SetPath_Button.Click += new System.EventHandler(this.PriceListMT_SetPath_Button_Click);
            // 
            // Cancel_Button
            // 
            this.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Cancel_Button.Location = new System.Drawing.Point(431, 116);
            this.Cancel_Button.Name = "Cancel_Button";
            this.Cancel_Button.Size = new System.Drawing.Size(75, 23);
            this.Cancel_Button.TabIndex = 12;
            this.Cancel_Button.Text = "Cancel";
            this.Cancel_Button.UseVisualStyleBackColor = true;
            // 
            // Ok_Button
            // 
            this.Ok_Button.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Ok_Button.Location = new System.Drawing.Point(512, 116);
            this.Ok_Button.Name = "Ok_Button";
            this.Ok_Button.Size = new System.Drawing.Size(75, 23);
            this.Ok_Button.TabIndex = 13;
            this.Ok_Button.Text = "Ok";
            this.Ok_Button.UseVisualStyleBackColor = true;
            this.Ok_Button.Click += new System.EventHandler(this.Ok_Button_Click);
            // 
            // openFileDialog
            // 
            this.openFileDialog.Filter = "Excel files|*.xls?|All files|*.*";
            // 
            // errorProvider
            // 
            this.errorProvider.BlinkStyle = System.Windows.Forms.ErrorBlinkStyle.NeverBlink;
            this.errorProvider.ContainerControl = this;
            // 
            // SettingsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(660, 146);
            this.Controls.Add(this.Ok_Button);
            this.Controls.Add(this.Cancel_Button);
            this.Controls.Add(this.PriceListMT_SetPath_Button);
            this.Controls.Add(this.Decision_SetPath_Button);
            this.Controls.Add(this.Budget_SetPath_Button);
            this.Controls.Add(this.ProductCalendar_SetPath_Button);
            this.Controls.Add(this.PriceListMT_Path_TextBox);
            this.Controls.Add(this.PriceListMTlabel);
            this.Controls.Add(this.Decision_Path_TextBox);
            this.Controls.Add(this.DecisionLabel);
            this.Controls.Add(this.Budget_Path_TextBox);
            this.Controls.Add(this.BudgetLabel);
            this.Controls.Add(this.ProductCalendar_Path_TextBox);
            this.Controls.Add(this.ProductCalendarLabel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SettingsForm";
            this.Text = "Настройки";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.SettingsForm_FormClosing);
            this.Load += new System.EventHandler(this.SettingsForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label ProductCalendarLabel;
        private System.Windows.Forms.TextBox ProductCalendar_Path_TextBox;
        private System.Windows.Forms.TextBox Budget_Path_TextBox;
        private System.Windows.Forms.Label BudgetLabel;
        private System.Windows.Forms.TextBox Decision_Path_TextBox;
        private System.Windows.Forms.Label DecisionLabel;
        private System.Windows.Forms.TextBox PriceListMT_Path_TextBox;
        private System.Windows.Forms.Label PriceListMTlabel;
        private System.Windows.Forms.Button ProductCalendar_SetPath_Button;
        private System.Windows.Forms.Button Budget_SetPath_Button;
        private System.Windows.Forms.Button Decision_SetPath_Button;
        private System.Windows.Forms.Button PriceListMT_SetPath_Button;
        private System.Windows.Forms.Button Cancel_Button;
        private System.Windows.Forms.Button Ok_Button;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.ErrorProvider errorProvider;
    }
}