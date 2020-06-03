using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BPA.Model;
using BPA.Modules;

using Microsoft.Office.Tools.Ribbon;

namespace BPA
{
    public partial class RibbonBPA
    {
        private void RibbonBPA_Load(object sender, RibbonUIEventArgs e)
        {

        }

        /// <summary>
        /// кнопка загрузки
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            FileCalendar fileCalendar = new FileCalendar();

            for (int i = 0; i < fileCalendar.LastRow; i++)
            {

            }



            Modules.ProductCalendar calendar = new Modules.ProductCalendar();
            calendar.LoadCalendar();
        }

        /// <summary>
        /// кнопка обновления
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Modules.ProductCalendar calendar = new Modules.ProductCalendar();
            calendar.UpdateCalendar();
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {


        }
    }
}
