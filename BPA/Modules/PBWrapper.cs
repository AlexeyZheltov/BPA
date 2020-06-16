using BPA.Forms;
using BPA.Modules;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.Modules
{
    class PBWrapper : IDisposable
    {
        public bool IsCancel { get; private set; } = false;

        ProcessBar bar;
        string taskText, caption;

        public PBWrapper(string caption, string taskText)
        {
            this.taskText = taskText;
            this.caption = caption;
        }

        public void Start(int amount)
        {
            bar = new ProcessBar(caption, amount);
            bar.CancelClick += Cancel;
            bar.Show(new ExcelWindows(Globals.ThisWorkbook));
        }

        public void Action(string Index) => bar.TaskStart(taskText.Replace("[Index]", Index));

        public void Done(int value) => bar.TaskDone(value);

        void Cancel() => IsCancel = true;

        public void Dispose()
        {
            bar.Close();
        }
    }
}
