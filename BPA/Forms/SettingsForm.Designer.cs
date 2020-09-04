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
            this.Budget_Path_TextBox = new System.Windows.Forms.TextBox();
            this.BudgetLabel = new System.Windows.Forms.Label();
            this.Decision_Path_TextBox = new System.Windows.Forms.TextBox();
            this.DecisionLabel = new System.Windows.Forms.Label();
            this.PriceListMT_Path_TextBox = new System.Windows.Forms.TextBox();
            this.PriceListMTlabel = new System.Windows.Forms.Label();
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
            // Budget_Path_TextBox
            // 
            this.Budget_Path_TextBox.Location = new System.Drawing.Point(98, 6);
            this.Budget_Path_TextBox.Name = "Budget_Path_TextBox";
            this.Budget_Path_TextBox.ReadOnly = true;
            this.Budget_Path_TextBox.Size = new System.Drawing.Size(486, 20);
            this.Budget_Path_TextBox.TabIndex = 3;
            this.Budget_Path_TextBox.Visible = false;
            // 
            // BudgetLabel
            // 
            this.BudgetLabel.AutoSize = true;
            this.BudgetLabel.Location = new System.Drawing.Point(3, 9);
            this.BudgetLabel.Name = "BudgetLabel";
            this.BudgetLabel.Size = new System.Drawing.Size(41, 13);
            this.BudgetLabel.TabIndex = 2;
            this.BudgetLabel.Text = "Budget";
            this.BudgetLabel.Visible = false;
            // 
            // Decision_Path_TextBox
            // 
            this.Decision_Path_TextBox.Location = new System.Drawing.Point(98, 32);
            this.Decision_Path_TextBox.Name = "Decision_Path_TextBox";
            this.Decision_Path_TextBox.ReadOnly = true;
            this.Decision_Path_TextBox.Size = new System.Drawing.Size(486, 20);
            this.Decision_Path_TextBox.TabIndex = 5;
            this.Decision_Path_TextBox.Visible = false;
            // 
            // DecisionLabel
            // 
            this.DecisionLabel.AutoSize = true;
            this.DecisionLabel.Location = new System.Drawing.Point(3, 35);
            this.DecisionLabel.Name = "DecisionLabel";
            this.DecisionLabel.Size = new System.Drawing.Size(48, 13);
            this.DecisionLabel.TabIndex = 4;
            this.DecisionLabel.Text = "Decision";
            this.DecisionLabel.Visible = false;
            // 
            // PriceListMT_Path_TextBox
            // 
            this.PriceListMT_Path_TextBox.Location = new System.Drawing.Point(98, 58);
            this.PriceListMT_Path_TextBox.Name = "PriceListMT_Path_TextBox";
            this.PriceListMT_Path_TextBox.ReadOnly = true;
            this.PriceListMT_Path_TextBox.Size = new System.Drawing.Size(486, 20);
            this.PriceListMT_Path_TextBox.TabIndex = 7;
            this.PriceListMT_Path_TextBox.Visible = false;
            // 
            // PriceListMTlabel
            // 
            this.PriceListMTlabel.AutoSize = true;
            this.PriceListMTlabel.Location = new System.Drawing.Point(3, 61);
            this.PriceListMTlabel.Name = "PriceListMTlabel";
            this.PriceListMTlabel.Size = new System.Drawing.Size(63, 13);
            this.PriceListMTlabel.TabIndex = 6;
            this.PriceListMTlabel.Text = "PriceListMT";
            this.PriceListMTlabel.Visible = false;
            // 
            // Budget_SetPath_Button
            // 
            this.Budget_SetPath_Button.Location = new System.Drawing.Point(611, 6);
            this.Budget_SetPath_Button.Name = "Budget_SetPath_Button";
            this.Budget_SetPath_Button.Size = new System.Drawing.Size(34, 20);
            this.Budget_SetPath_Button.TabIndex = 9;
            this.Budget_SetPath_Button.Text = "...";
            this.Budget_SetPath_Button.UseVisualStyleBackColor = true;
            this.Budget_SetPath_Button.Visible = false;
            this.Budget_SetPath_Button.Click += new System.EventHandler(this.Budget_SetPath_Button_Click);
            // 
            // Decision_SetPath_Button
            // 
            this.Decision_SetPath_Button.Location = new System.Drawing.Point(611, 32);
            this.Decision_SetPath_Button.Name = "Decision_SetPath_Button";
            this.Decision_SetPath_Button.Size = new System.Drawing.Size(34, 20);
            this.Decision_SetPath_Button.TabIndex = 10;
            this.Decision_SetPath_Button.Text = "...";
            this.Decision_SetPath_Button.UseVisualStyleBackColor = true;
            this.Decision_SetPath_Button.Visible = false;
            this.Decision_SetPath_Button.Click += new System.EventHandler(this.Decision_SetPath_Button_Click);
            // 
            // PriceListMT_SetPath_Button
            // 
            this.PriceListMT_SetPath_Button.Location = new System.Drawing.Point(611, 58);
            this.PriceListMT_SetPath_Button.Name = "PriceListMT_SetPath_Button";
            this.PriceListMT_SetPath_Button.Size = new System.Drawing.Size(34, 20);
            this.PriceListMT_SetPath_Button.TabIndex = 11;
            this.PriceListMT_SetPath_Button.Text = "...";
            this.PriceListMT_SetPath_Button.UseVisualStyleBackColor = true;
            this.PriceListMT_SetPath_Button.Visible = false;
            this.PriceListMT_SetPath_Button.Click += new System.EventHandler(this.PriceListMT_SetPath_Button_Click);
            // 
            // Cancel_Button
            // 
            this.Cancel_Button.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Cancel_Button.Location = new System.Drawing.Point(428, 80);
            this.Cancel_Button.Name = "Cancel_Button";
            this.Cancel_Button.Size = new System.Drawing.Size(75, 23);
            this.Cancel_Button.TabIndex = 12;
            this.Cancel_Button.Text = "Cancel";
            this.Cancel_Button.UseVisualStyleBackColor = true;
            // 
            // Ok_Button
            // 
            this.Ok_Button.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.Ok_Button.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Ok_Button.Location = new System.Drawing.Point(509, 80);
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
            this.ClientSize = new System.Drawing.Size(659, 106);
            this.Controls.Add(this.Ok_Button);
            this.Controls.Add(this.Cancel_Button);
            this.Controls.Add(this.PriceListMT_SetPath_Button);
            this.Controls.Add(this.Decision_SetPath_Button);
            this.Controls.Add(this.Budget_SetPath_Button);
            this.Controls.Add(this.PriceListMT_Path_TextBox);
            this.Controls.Add(this.PriceListMTlabel);
            this.Controls.Add(this.Decision_Path_TextBox);
            this.Controls.Add(this.DecisionLabel);
            this.Controls.Add(this.Budget_Path_TextBox);
            this.Controls.Add(this.BudgetLabel);
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
        private System.Windows.Forms.TextBox Budget_Path_TextBox;
        private System.Windows.Forms.Label BudgetLabel;
        private System.Windows.Forms.TextBox Decision_Path_TextBox;
        private System.Windows.Forms.Label DecisionLabel;
        private System.Windows.Forms.TextBox PriceListMT_Path_TextBox;
        private System.Windows.Forms.Label PriceListMTlabel;
        private System.Windows.Forms.Button Budget_SetPath_Button;
        private System.Windows.Forms.Button Decision_SetPath_Button;
        private System.Windows.Forms.Button PriceListMT_SetPath_Button;
        private System.Windows.Forms.Button Cancel_Button;
        private System.Windows.Forms.Button Ok_Button;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.ErrorProvider errorProvider;
    }
}