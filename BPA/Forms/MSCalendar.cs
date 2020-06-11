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
    public partial class MSCalendar : Form
    {
        public MSCalendar()
        {
            InitializeComponent();
            SelectedDate = Calendar_Control.SelectionStart.Date;
        }

        public DateTime SelectedDate { get; private set; }

        private void Calendar_Control_DateSelected(object sender, DateRangeEventArgs e)
        {
            SelectedDate = e.Start.Date;
        }
    }
}
