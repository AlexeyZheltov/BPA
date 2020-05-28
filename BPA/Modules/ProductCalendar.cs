using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BPA.Modules
{
    class ProductCalendar
    {
        readonly Microsoft.Office.Interop.Excel.Application ex = new Microsoft.Office.Interop.Excel.Application();
        Workbook WB;

        public Workbook Open
        {
            get
            {
                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    InitialDirectory = Globals.ThisWorkbook.Application.ActiveWorkbook.Path,
                    Filter = "Excel files (*.xls*)|*.xls*",
                    Title = "Выберите файл календаря"
                };

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;
                    return WB = ex.Workbooks.Open(filePath);
                }
                else
                {
                    return null;
                }
            }
        }
    }
    
}
