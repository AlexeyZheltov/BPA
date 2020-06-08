namespace BPA.Forms
{
    partial class ProcessBar
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ProcessBar));
            this.LabelNameProcess = new System.Windows.Forms.Label();
            this.LabelNameTask = new System.Windows.Forms.Label();
            this.LabelComplete = new System.Windows.Forms.Label();
            this.btnCancelBox = new System.Windows.Forms.PictureBox();
            this.label2 = new System.Windows.Forms.Label();
            this.LabelTimeLost = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.progressBar = new BPA.Controls.ProgressBarControl();
            ((System.ComponentModel.ISupportInitialize)(this.btnCancelBox)).BeginInit();
            this.SuspendLayout();
            // 
            // LabelNameProcess
            // 
            this.LabelNameProcess.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.LabelNameProcess.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.LabelNameProcess.Location = new System.Drawing.Point(4, 9);
            this.LabelNameProcess.Margin = new System.Windows.Forms.Padding(3, 0, 0, 0);
            this.LabelNameProcess.Name = "LabelNameProcess";
            this.LabelNameProcess.Size = new System.Drawing.Size(455, 14);
            this.LabelNameProcess.TabIndex = 1;
            this.LabelNameProcess.Text = "Описание процесса";
            // 
            // LabelNameTask
            // 
            this.LabelNameTask.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.LabelNameTask.Location = new System.Drawing.Point(4, 32);
            this.LabelNameTask.Margin = new System.Windows.Forms.Padding(0);
            this.LabelNameTask.Name = "LabelNameTask";
            this.LabelNameTask.Size = new System.Drawing.Size(344, 18);
            this.LabelNameTask.TabIndex = 3;
            this.LabelNameTask.Text = "Выполняемая задача";
            // 
            // LabelComplete
            // 
            this.LabelComplete.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.LabelComplete.AutoSize = true;
            this.LabelComplete.BackColor = System.Drawing.Color.Transparent;
            this.LabelComplete.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.LabelComplete.Location = new System.Drawing.Point(431, 54);
            this.LabelComplete.Margin = new System.Windows.Forms.Padding(0);
            this.LabelComplete.Name = "LabelComplete";
            this.LabelComplete.Size = new System.Drawing.Size(29, 16);
            this.LabelComplete.TabIndex = 4;
            this.LabelComplete.Text = "0%";
            // 
            // btnCancelBox
            // 
            this.btnCancelBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancelBox.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnCancelBox.ErrorImage = null;
            this.btnCancelBox.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelBox.Image")));
            this.btnCancelBox.Location = new System.Drawing.Point(461, 1);
            this.btnCancelBox.Name = "btnCancelBox";
            this.btnCancelBox.Size = new System.Drawing.Size(26, 25);
            this.btnCancelBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.btnCancelBox.TabIndex = 5;
            this.btnCancelBox.TabStop = false;
            this.btnCancelBox.Click += new System.EventHandler(this.CancelBox_Click);
            // 
            // label2
            // 
            this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label2.Location = new System.Drawing.Point(363, 32);
            this.label2.Margin = new System.Windows.Forms.Padding(0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(62, 18);
            this.label2.TabIndex = 6;
            this.label2.Text = "Осталось:";
            this.label2.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // LabelTimeLost
            // 
            this.LabelTimeLost.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.LabelTimeLost.Location = new System.Drawing.Point(424, 32);
            this.LabelTimeLost.Margin = new System.Windows.Forms.Padding(0);
            this.LabelTimeLost.Name = "LabelTimeLost";
            this.LabelTimeLost.Size = new System.Drawing.Size(60, 18);
            this.LabelTimeLost.TabIndex = 7;
            this.LabelTimeLost.Text = "-";
            this.LabelTimeLost.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(488, 78);
            this.panel1.TabIndex = 9;
            // 
            // progressBar
            // 
            this.progressBar.BackColor = System.Drawing.Color.LightGray;
            this.progressBar.BackColorProgressLeft = System.Drawing.Color.ForestGreen;
            this.progressBar.BackColorProgressRight = System.Drawing.Color.DarkSeaGreen;
            this.progressBar.BorderColor = System.Drawing.Color.Gray;
            this.progressBar.IsFreeze = false;
            this.progressBar.Location = new System.Drawing.Point(7, 52);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(417, 20);
            this.progressBar.Step = 10;
            this.progressBar.TabIndex = 8;
            this.progressBar.Text = "progressBar";
            this.progressBar.Value = 0;
            this.progressBar.ValueMaximum = 100;
            this.progressBar.ValueMinimum = 0;
            // 
            // ProcessBar
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(488, 78);
            this.ControlBox = false;
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.LabelTimeLost);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnCancelBox);
            this.Controls.Add(this.LabelComplete);
            this.Controls.Add(this.LabelNameTask);
            this.Controls.Add(this.LabelNameProcess);
            this.Controls.Add(this.panel1);
            this.Cursor = System.Windows.Forms.Cursors.Default;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ProcessBar";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Процесс";
            this.TopMost = true;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.ProcessBar_FormClosing);
            ((System.ComponentModel.ISupportInitialize)(this.btnCancelBox)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label LabelNameProcess;
        private System.Windows.Forms.Label LabelNameTask;
        private System.Windows.Forms.Label LabelComplete;
        private System.Windows.Forms.PictureBox btnCancelBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label LabelTimeLost;
        private Controls.ProgressBarControl progressBar;
        private System.Windows.Forms.Panel panel1;
    }
}