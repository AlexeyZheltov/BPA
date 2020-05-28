using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BPA.Model;
using Microsoft.Office.Tools.Ribbon;

namespace BPA
{
    public partial class RibbonBPA
    {
        private void RibbonBPA_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Supercategory supercategory = new Supercategory();
            supercategory.SetProperty(1);

        }
    }
}
