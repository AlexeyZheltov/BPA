using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace BPA.Modules
{
    class ExcelWindows : IWin32Window
    {
        Excel.Window window;
        public ExcelWindows(ThisWorkbook wb) => window = wb.Windows[1];

        public IntPtr Handle => (IntPtr)window.Hwnd;
    }
}
