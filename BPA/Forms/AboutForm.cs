using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Deployment.Application;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BPA.Forms
{
    public partial class AboutForm : Form
    {
        public AboutForm()
        {
            InitializeComponent();
            linkLabel1.LinkClicked += linkLabel1_LinkClicked;
            try
            {
                Version ver;
                try
                {
                    ver = ApplicationDeployment.CurrentDeployment.CurrentVersion;
                } catch
                {
                    ver = Assembly.GetExecutingAssembly().GetName().Version;
                }
                string[] verArr = ver.ToString().Split(new char[] { '.' });
                string version = String.Join(".", verArr);
                if (verArr != null)
                    if(verArr.Length >= 2)
                        label2.Text = $"v.{ verArr[0] }.{ verArr[1] }";
                    else if (verArr.Length == 1)
                        label2.Text = $"v.{ verArr[0] }";
            } catch
            {
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://micro-solution.ru/");

        }
    }
}
