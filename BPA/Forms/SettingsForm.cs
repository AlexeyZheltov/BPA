using BPA.Properties;
using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using BPA.Modules;

namespace BPA.Forms
{
    public partial class SettingsForm : Form
    {
        readonly Settings settings = Properties.Settings.Default;
        bool _decision_flag, _budget_flag, _price_flag;
        public SettingsForm(BPASettingEnum typeSetting)
        {
            InitializeComponent();

            _budget_flag = ((int)typeSetting & 0b0001) > 0;
            _decision_flag = ((int)typeSetting & 0b0010) > 0;
            _price_flag = ((int)typeSetting & 0b0100) > 0;
        }

        private void SettingsForm_Load(object sender, EventArgs e)
        {
            int h_index = 0;
            Budget_Path_TextBox.Text = settings.BudgetPath;
            if (_budget_flag)
            {
                BudgetLabel.Visible = true;
                Budget_Path_TextBox.Visible = true;
                Budget_SetPath_Button.Visible = true;
                
                h_index++;
                ValidatePath(Budget_Path_TextBox);
            }

            Decision_Path_TextBox.Text = settings.DecisionPath;
            if (_decision_flag)
            {
                DecisionLabel.Visible = true;
                Decision_Path_TextBox.Visible = true;
                Decision_SetPath_Button.Visible = true;
                
                DecisionLabel.Top = 9 + 26 * h_index;
                Decision_Path_TextBox.Top = 6 + 26 * h_index;
                Decision_SetPath_Button.Top = 6 + 26 * h_index;
                h_index++;
                ValidatePath(Decision_Path_TextBox);
            }

            PriceListMT_Path_TextBox.Text = settings.PriceListMTPath;
            if (_price_flag)
            {
                PriceListMTlabel.Visible = true;
                PriceListMT_Path_TextBox.Visible = true;
                PriceListMT_SetPath_Button.Visible = true;
                
                PriceListMTlabel.Top = 9 + 26 * h_index;
                PriceListMT_Path_TextBox.Top = 6 + 26 * h_index;
                PriceListMT_SetPath_Button.Top = 6 + 26 * h_index;
                h_index++;
                ValidatePath(PriceListMT_Path_TextBox);
            }

            this.Height = 93 + 26 * (h_index - 1);
        }

        private void ValidatePath(TextBox textBox)
        {
            if (textBox.Text == "") errorProvider.SetError(textBox, "Файл не выбран");
            else if (!File.Exists(textBox.Text)) errorProvider.SetError(textBox, "Указанный файл не существует");
            else errorProvider.SetError(textBox, "");
        }

        private void SettingsForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing) DialogResult = DialogResult.Cancel;
        }

        private void Budget_SetPath_Button_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK) Budget_Path_TextBox.Text = openFileDialog.FileName;
            ValidatePath(Budget_Path_TextBox);
        }

        private void Decision_SetPath_Button_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK) Decision_Path_TextBox.Text = openFileDialog.FileName;
            ValidatePath(Decision_Path_TextBox);
        }

        private void PriceListMT_SetPath_Button_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK) PriceListMT_Path_TextBox.Text = openFileDialog.FileName;
            ValidatePath(PriceListMT_Path_TextBox);
        }

        private void Ok_Button_Click(object sender, EventArgs e)
        {
            settings.BudgetPath = Budget_Path_TextBox.Text;
            settings.DecisionPath = Decision_Path_TextBox.Text;
            settings.PriceListMTPath = PriceListMT_Path_TextBox.Text;
            settings.Save();
        }
    }
}
