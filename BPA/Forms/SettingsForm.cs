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

namespace BPA.Forms
{
    public partial class SettingsForm : Form
    {
        readonly Settings settings = Properties.Settings.Default;
        public SettingsForm()
        {
            InitializeComponent();
        }

        private void SettingsForm_Load(object sender, EventArgs e)
        {
            ProductCalendar_Path_TextBox.Text = settings.ProductCalendarPath;
            Decision_Path_TextBox.Text = settings.DecisionPath;
            Budget_Path_TextBox.Text = settings.BudgetPath;
            PriceListMT_Path_TextBox.Text = settings.PriceListMTPath;

            //Проверка на валидность путей
            ValidatePath(ProductCalendar_Path_TextBox);
            ValidatePath(Decision_Path_TextBox);
            ValidatePath(Budget_Path_TextBox);
            ValidatePath(PriceListMT_Path_TextBox);
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

        private void ProductCalendar_SetPath_Button_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK) ProductCalendar_Path_TextBox.Text = openFileDialog.FileName;
            ValidatePath(ProductCalendar_Path_TextBox);
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
            settings.ProductCalendarPath = ProductCalendar_Path_TextBox.Text;
            settings.BudgetPath = Budget_Path_TextBox.Text;
            settings.DecisionPath = Decision_Path_TextBox.Text;
            settings.PriceListMTPath = PriceListMT_Path_TextBox.Text;
        }
    }
}
