using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.Modules
{
    class ProductCalendar
    {

        public void Method()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                InitialDirectory = Globals.ThisWorkbook.Application.ActiveWorkbook.Path,
                Filter = "Word files (*.doc*)|*.doc*",
                Title = "Выберите файл, который необходимо заполнить"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                WordFiller wordFiller = new WordFiller(openFileDialog.FileName, new Field().GetFields());
                ProcessBar processBar = new ProcessBar("Заполнение документа", wordFiller.CountActions);
                processBar.Show();
                wordFiller.ActionStart += processBar.TaskStart;
                wordFiller.ActionDone += processBar.TaskDone;
                processBar.CancelClick += wordFiller.Cancel;
                wordFiller.FillTemplate();
            }
        }
    }
    }
}
