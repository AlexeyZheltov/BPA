using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BPA.Controls
{
    public class ProgressBarControl : Control
    {
        #region -- Переменные --

        Animation ProgressAnim = new Animation();

        #endregion

        #region -- Свойства --

        public bool IsFreeze
        {
            get
            {
                return _IsFreese;
            }
            set
            {
                _IsFreese = value;
                if (!IsFreeze) Invalidate();
            }
        }
        private bool _IsFreese = false;

        public Color BorderColor { get; set; } = Color.DarkGray;
        public Color BackColorProgressLeft { get; set; } = Color.Green;
        public Color BackColorProgressRight { get; set; } = Color.GreenYellow;

        private int _value = 0;
        public int Value
        {
            get => _value;
            set
            {
                if (value >= ValueMinimum && value <= ValueMaximum)
                {
                    _value = value;

                    if (Animator.IsWork == false)
                    {
                        ProgressAnim.Value = _value;
                        if (!IsFreeze) Invalidate();
                    }
                    else
                    {
                        if (!IsFreeze) ProgressAction(_value);
                    }
                }
                else
                {
                    value = _value;
                }
            }
        }

        private int _valueMinimum = 0;
        public int ValueMinimum
        {
            get => _valueMinimum;
            set
            {
                if (value < ValueMaximum)
                {
                    _valueMinimum = value;

                    if (_valueMinimum > Value)
                    {
                        Value = _valueMinimum;
                        if (!IsFreeze) Invalidate();
                    }
                }
                else
                {
                    value = _valueMinimum;
                }
            }
        }

        private int _valueMaximum = 100;
        public int ValueMaximum
        {
            get => _valueMaximum;
            set
            {
                if (value > ValueMinimum)
                {
                    _valueMaximum = value;
                    if (!IsFreeze) Invalidate();
                }
                else
                {
                    value = _valueMaximum;
                }
            }
        }

        public int Step { get; set; } = 10;

        #endregion

        public ProgressBarControl()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw | ControlStyles.SupportsTransparentBackColor | ControlStyles.UserPaint, true);
            DoubleBuffered = true;

            Size = new Size(200, 20);

            BackColor = Color.LightGray;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            if (IsFreeze) return;
            Graphics graph = e.Graphics;
            graph.SmoothingMode = SmoothingMode.HighQuality;
            graph.InterpolationMode = InterpolationMode.HighQualityBicubic;
            graph.Clear(BackColor);

            Rectangle rectBase = new Rectangle(0, 0, Width - 1, Height - 1);
            Rectangle rectProgress = new Rectangle(
                rectBase.X,
                rectBase.Y,
                CalculateProgressRectSize(rectBase),
                rectBase.Height);

            

            // Рисуем основу
            DrawBase(graph, rectBase);

            // Рисуем прогресс
            DrawProgress(graph, rectProgress);

            // Рисуем обводку
            DrawBorder(graph, rectBase);
        }

        private int CalculateProgressRectSize(Rectangle rect)
        {
            int margin = ValueMaximum - ValueMinimum;
            return rect.Width * (int)ProgressAnim.Value / margin;
        }

        #region -- Рисование объектов --

        private void DrawBase(Graphics graph, Rectangle rect)
        {
            graph.FillRectangle(new SolidBrush(BackColor), rect);
        }

        private void DrawBorder(Graphics graph, Rectangle rect)
        {
            graph.DrawRectangle(new Pen(BorderColor), rect);
        }

        private void DrawProgress(Graphics graph, Rectangle rect)
        {
            if (rect.Width > 0)
            {
                LinearGradientBrush LGB = new LinearGradientBrush(rect, BackColorProgressLeft, BackColorProgressRight, 0, true);

                graph.DrawRectangle(new Pen(LGB), rect);
                graph.FillRectangle(LGB, rect);
            }
        }

        #endregion

        #region -- Запуск анимаций --

        private void ProgressAction(int PIXELS)
        {
            ProgressAnim = new Animation("ProgressBar_" + Handle, Invalidate, ProgressAnim.Value, PIXELS);

            ProgressAnim.StepDivider = 8;
            Animator.Request(ProgressAnim, true);
        }

        #endregion

        #region -- Public методы --

        public bool PerformStep()
        {
            if (Value < ValueMaximum)
            {
                if (Value + Step >= ValueMaximum)
                {
                    Value = ValueMaximum;
                    return false;
                }
                else
                {
                    Value += Step;
                    return true;
                }
            }
            else
            {
                return false;
            }
        }

        public void ResetProgress()
        {
            Value = ValueMinimum;
        }

        #endregion
    }
    public class Animation
    {
        public string ID { get; set; }

        public float Value;

        public float StartValue;

        private float targetValue;
        public float TargetValue
        {
            get => targetValue;
            set
            {
                targetValue = value;
                Reverse = value < Value ? true : false;
            }
        }

        public float Volume;

        public bool Reverse = false;

        public AnimationStatus Status { get; set; }
        public enum AnimationStatus
        {
            Requested,
            Active,
            Completed
        }

        private float p15, p30, p70, p85;

        public int StepDivider = 11;

        private float Step()
        {
            float basicStep = Math.Abs(Volume) / StepDivider; // Math.Abs - превращает числа 0< в >0
            float resultStep = 0;

            if (Reverse == false)
            {
                if (Value <= p15 || Value >= p85)
                {
                    resultStep = basicStep / 3.5f;
                }
                else if (Value <= p30 || Value >= p70)
                {
                    resultStep = basicStep / 2f;
                }
                else if (Value > p30 && Value < p70)
                {
                    resultStep = basicStep;
                }
            }
            else
            {
                if (Value >= p15 || Value <= p85)
                {
                    resultStep = basicStep / 3.5f;
                }
                else if (Value >= p30 || Value <= p70)
                {
                    resultStep = basicStep / 2f;
                }
                else if (Value < p30 && Value > p70)
                {
                    resultStep = basicStep;
                }
            }

            return Math.Abs(resultStep);
        }

        private float ValueByPercent(float Percent)
        {
            float COEFF = Percent / 100;
            float VolumeInPercent = Volume * COEFF;
            float ValueInPercent = StartValue + VolumeInPercent;

            return ValueInPercent;
        }

        public delegate void ControlMethod();
        private ControlMethod InvalidateControl;

        public void UpdateFrame()
        {
            Status = AnimationStatus.Active;

            if (Reverse == false)
            {
                if (Value <= targetValue)
                {
                    Value += Step();

                    if (Value >= targetValue)
                    {
                        Value = targetValue;
                        Status = AnimationStatus.Completed;
                    }
                }
            }
            else
            {
                if (Value >= targetValue)
                {
                    Value -= Step();

                    if (Value <= targetValue)
                    {
                        Value = targetValue;
                        Status = AnimationStatus.Completed;
                    }
                }
            }

            InvalidateControl.Invoke();
        }

        public Animation() { }

        public Animation(string ID, ControlMethod InvalidateControl, float Value, float TargetValue)
        {
            this.ID = ID;

            this.InvalidateControl = InvalidateControl;

            this.Value = Value;
            this.TargetValue = TargetValue;

            StartValue = Value;
            Volume = TargetValue - Value;

            p15 = ValueByPercent(15);
            p30 = ValueByPercent(30);
            p70 = ValueByPercent(70);
            p85 = ValueByPercent(85);
        }
    }

    public static class Animator
    {
        public static List<Animation> AnimationList = new List<Animation>();

        public static int Count()
        {
            return AnimationList.Count;
        }

        private static Thread AnimatorThread;

        private static double Interval;

        public static bool IsWork = false;

        public static void Start()
        {
            IsWork = true;
            Interval = 14; // FPS ~66

            AnimatorThread = new Thread(AnimationInvoker)
            {
                IsBackground = true,
                Name = "UI Animation"
            };

            AnimatorThread.Start();
        }

        private static void AnimationInvoker()
        {
            while (IsWork)
            {
                AnimationList.RemoveAll(a => a == null || a.Status == Animation.AnimationStatus.Completed);

                Parallel.For(0, Count(), index =>
                {
                    AnimationList[index].UpdateFrame();
                });

                Thread.Sleep((int)Interval);
            }
        }

        public static void Request(Animation Anim, bool ReplaceIfExists = true)
        {
            Debug.WriteLine("Запуск анимации: " + Anim.ID + "| TargetValue: " + Anim.TargetValue);
            Anim.Status = Animation.AnimationStatus.Requested;

            Animation dupAnim = GetDuplicate(Anim);

            if (dupAnim != null)
            {
                if (ReplaceIfExists == true)
                {
                    dupAnim.Status = Animation.AnimationStatus.Completed;
                }
                else
                {
                    return;
                }
            }

            AnimationList.Add(Anim);
        }

        private static Animation GetDuplicate(Animation Anim)
        {
            return AnimationList.Find(a => a.ID == Anim.ID);
        }
    }
}