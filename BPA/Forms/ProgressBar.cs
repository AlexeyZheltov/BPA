using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

namespace BPA.Forms
{
    /// <summary>
    /// Форма прогрессбара
    /// </summary>
    public partial class ProcessBar : Form
    {
        #region Переменные

        private readonly Stopwatch Timer;
        private readonly Stopwatch TimerTask;

        #endregion

        #region Свойства

        /// <summary>
        /// Дочерний прогрессбар
        /// </summary>
        public ProcessBar SubBar;

        /// <summary>
        /// Заголовок прогрессбара, то что будет написано в названии формы
        /// </summary>
        public string Title
        {
            get => title;
            private set
            {
                title = value;
                LabelNameProcess.Text = title;
            }
        }
        private string title;

        /// <summary>
        /// Уровень прогрессбара
        /// </summary>
        public int Level { get; set; } = 1;

        /// <summary>
        /// Название задачи
        /// </summary>
        public string TaskName
        {
            get => taskName;
            private set
            {
                taskName = value;
                LabelNameTask.Text = taskName;
            }
        }
        private string taskName;

        /// <summary>
        /// Начальная позиция прогресбара
        /// </summary>
        public int Start
        {
            get
            {
                return start;
            }
            set
            {
                start = value;
                progressBar.ValueMaximum = start;
            }
        }
        private int start;

        /// <summary>
        /// Шаг прогрессбара
        /// </summary>
        public int Step
        {
            get => step;
            set
            {
                step = value;
                progressBar.Step = step;
                progressBar.Step = step;
            }
        }
        private int step;

        /// <summary>
        /// Общее количество итераций
        /// </summary>
        public int Count
        {
            get => count;
            set
            {
                count = value;
                progressBar.ValueMaximum = count;
            }
        }
        private int count;

        /// <summary>
        /// Текущее значение прогрессбара
        /// </summary>
        public int Value
        {
            get => _value;
            set
            {
                _value = value;
                progressBar.Value = _value;

                if (_value == Count) Close();
                if (_value == 0) return;
                LabelComplete.Text = Count == 0 ? "0%" : ((double)_value / (double)Count).ToString("##0%");
                TimeSpan ts = TimeSpan.FromMilliseconds(Timer.Elapsed.TotalMilliseconds / _value * (Count - _value));
                if (ts.Days > 0)
                {
                    LabelTimeLost.Text = String.Format("{0:#0}дн. {1:#0}ч", ts.Days, ts.Hours);
                }
                else if (ts.Hours > 0)
                {
                    LabelTimeLost.Text = String.Format("{0:#0}ч {1:#0}мин", ts.Hours, ts.Minutes);
                }
                else if (ts.Minutes > 0)
                {
                    LabelTimeLost.Text = String.Format("{0:#0}:{1:00}", ts.Minutes, ts.Seconds);
                }
                else
                {
                    LabelTimeLost.Text = String.Format("{0:#0} сек", ts.Seconds);
                }

                if (index == 0 || TimerTask.ElapsedMilliseconds > 80)
                {
                    Application.DoEvents();
                    TimerTask.Restart();
                }
                index++;
            }
        }
        private int _value;
        private int index = 0;

        /// <summary>
        /// Нажата кнопка отмены
        /// </summary>
        public bool IsCancel { get; set; } = false;

        #endregion

        #region События
      
        /// <summary>
        /// Событие клика кнопки отмены
        /// </summary>
        public event CancelProccess CancelClick;
        public delegate void CancelProccess();

        /// <summary>
        /// Событие выполнения задачи
        /// </summary>
        public event MoveProgress TaskDoned;
        public delegate void MoveProgress(int count);

        /// <summary>
        /// Событие увеличения размеров прогрессбара
        /// </summary>
        public event ResizeProgressPlus ResisingPlus;
        public delegate void ResizeProgressPlus(int count);

        /// <summary>
        /// Событие уменьшения размеров прогрессбара
        /// </summary>
        public event ResizeProgressMinus ResisingMinus;
        public delegate void ResizeProgressMinus(int count);

        #endregion

        #region События формы
  
        /// <summary>
        /// Закрытие формы прогрессбара
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ProcessBar_FormClosing(object sender, FormClosingEventArgs e)
        {
            Timer.Stop();
            IsCancel = true;
            
        }

        /// <summary>
        /// Кнопка закрытия формы
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CancelBox_Click(object sender, EventArgs e)
        {
            Cancel();
            CancelClick?.Invoke();
        }

        #endregion

        /// <summary>
        /// Инициализация прогрессбара
        /// </summary>
        /// <param name="title">Описание глобальной задачи</param>
        /// <param name="count">Количество итераций</param>
        /// <param name="start">Старт</param>
        /// <param name="step">Шаг</param>
        public ProcessBar(string title, int count, int start = 0, int step = 1)
        {
            InitializeComponent();
            Title = title;
            Count = count;
            Start = start;
            Step = step;

            TaskName = "Запуск прогрессбара";
            Timer = Stopwatch.StartNew();
            TimerTask = Stopwatch.StartNew();
        }

        #region Методы

        /// <summary>
        /// Старт задачи
        /// </summary>
        /// <param name="taskName">Наименование задачи</param>
        public void TaskStart(string taskName)
        {
            TaskName = taskName;
        }

        /// <summary>
        /// Завершение задачи
        /// </summary>
        /// <param name="count">вес задачи</param>
        public void TaskDone(int count = 1)
        {
            Value += count;
            TaskDoned?.Invoke(count);
        }

        /// <summary>
        /// Отмена операции
        /// </summary>
        public void Cancel()
        {
            IsCancel = true;
            SubBar?.Cancel();
            Close();
        }

        /// <summary>
        /// Добавление дочернего прогрессбара
        /// </summary>
        /// <param name="title"></param>
        /// <param name="count"></param>
        /// <param name="start"></param>
        /// <param name="step"></param>
        public void AddSubBar(string title, int count, int start = 0, int step = 1)
        {
            ResizePlus(count);

            SubBar = new ProcessBar(title, count, start, step);
            SubBar.Show();
            SubBar.btnCancelBox.Visible = false;
            SubBar.Level = Level++;
            SubBar.Top = Top + Height;
            SubBar.Left = Left;
            Application.DoEvents();
            SubBar.TaskDoned += TaskDone;
            SubBar.CancelClick += Cancel;
            CancelClick += SubBar.Cancel;
            SubBar.FormClosing += SubBar_FormClosing;

            SubBar.ResisingPlus += ResizePlus;
            SubBar.ResisingMinus += ResizeMinus;
        }

        private void SubBar_FormClosing(object sender, FormClosingEventArgs e)
        {
            ResizeMinus(SubBar.Count);
        }

        /// <summary>
        /// Увеличение размеров прогрессбара
        /// </summary>
        /// <param name="count">количество в дочернем прогрессбаре</param>
        public void ResizePlus(int count)
        {
            progressBar.IsFreeze = true;
            Count *= count;
            Value *= count;
            progressBar.IsFreeze = false;
            ResisingPlus?.Invoke(count);
        }

        /// <summary>
        /// Уменьшение размеров прогрессбара
        /// </summary>
        /// <param name="count">количество в дочернем прогрессбаре</param>
        public void ResizeMinus(int count)
        {
            progressBar.IsFreeze = true;
            Value /= count;
            Count /= count;
            progressBar.IsFreeze = false;
            ResisingMinus?.Invoke(count);
        }

        #endregion

        Point oldPos;
        bool isDragging = false;
        Point oldMouse;
        private void MyForm_MouseDown(object sender, MouseEventArgs e)
        {
            this.isDragging = true;
            this.oldPos = this.Location;
            this.oldMouse = e.Location;
        }

        private void MyForm_MouseMove(object sender, MouseEventArgs e)
        {
            if (this.isDragging)
            {
                this.Location = new Point(oldPos.X + (e.X - oldMouse.X), oldPos.Y + (e.Y - oldMouse.Y));
            }
        }

        private void MyForm_MouseUp(object sender, MouseEventArgs e)
        {
            this.isDragging = false;
        }
    }
}