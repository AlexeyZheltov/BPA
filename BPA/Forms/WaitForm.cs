using System;
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
    public partial class WaitForm : Form
    {
        string text;
        public WaitForm()
        {
            InitializeComponent();
        }

        private void textTimer_Tick(object sender, EventArgs e)
        {
            textLabel.Text = textLabel.Text.Length < text.Length + 4 ? $"{textLabel.Text}." : text;
        }

        private void WaitForm_Load(object sender, EventArgs e)
        {
            text = textLabel.Text;
        }
    }
}
