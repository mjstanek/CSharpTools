using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataverseMappingRemover.UI.Components
{
    public class RoundedButton : Button
    {
        public Color BorderColor { get; set; } = Color.Gray;
        public int BorderRadius { get; set; } = 8;
        public int BorderThickness { get; set; } = 2;

        public RoundedButton()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint |
                ControlStyles.OptimizedDoubleBuffer |
                ControlStyles.ResizeRedraw |
                ControlStyles.UserPaint |
                ControlStyles.Selectable, true);

            TabStop = true;
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
            BackColor = Color.White;
            ForeColor = Color.Black;
        }

        protected override bool IsInputKey(Keys keyData)
        {
            return keyData is Keys.Up or Keys.Down or Keys.Left or Keys.Right
                ? true
                : base.IsInputKey(keyData);
        }

        protected override bool ProcessDialogKey(Keys keyData)
        {
            System.Diagnostics.Debug.WriteLine($"ProcessDialogKey: {keyData}");
            if (keyData == Keys.Right || keyData == Keys.Up)
            {
                SelectNextControl(this, true, true, true, true);
                return true;
            }
            else if (keyData == Keys.Left || keyData == Keys.Down)
            {
                SelectNextControl(this, false, true, true, true);
                return true;
            }
            return base.ProcessDialogKey(keyData);
        }

        protected override void OnKeyDown(KeyEventArgs kevent)
        {
            System.Diagnostics.Debug.WriteLine($"Key Down: {kevent.KeyCode}");
            base.OnKeyDown(kevent);
            if (kevent.KeyCode == Keys.Right || kevent.KeyCode == Keys.Down)
            {
                SelectNextControl(this, true, true, true, true);
                kevent.Handled = true;
                kevent.SuppressKeyPress = true;
            }
            else if (kevent.KeyCode == Keys.Left || kevent.KeyCode == Keys.Up)
            {
                SelectNextControl(this, false, true, true, true);
                kevent.Handled = true;
                kevent.SuppressKeyPress = true;
            }
        }

        protected override void OnResize(System.EventArgs e)
        {
            base.OnResize(e);
            using var path = GetRoundRectangle(ClientRectangle, BorderRadius);
            Region = new Region(path);
            Invalidate();
        }
        protected override void OnPaint(System.Windows.Forms.PaintEventArgs pevent)
        {
            pevent.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

            using var bgBrush = new SolidBrush(Parent?.BackColor ?? SystemColors.Control);
            pevent.Graphics.FillRectangle(bgBrush, ClientRectangle);

            using var path = GetRoundRectangle(ClientRectangle, BorderRadius);
            using var fillBrush = new SolidBrush(BackColor);
            pevent.Graphics.FillPath(fillBrush, path);

            var rectangle = ClientRectangle;
            rectangle.Width -= 1;
            rectangle.Height -= 1;
            using var borderPath = GetRoundRectangle(rectangle, BorderRadius);
            using var pen = new Pen(BorderColor, BorderThickness);
            pevent.Graphics.DrawPath(pen, borderPath);

            TextRenderer.DrawText(pevent.Graphics, Text, Font, ClientRectangle, ForeColor,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);

            if (Focused)
            {
                var focusRect = Rectangle.Inflate(ClientRectangle, -3, -3);
                using var focusPen = new Pen(ControlPaint.Dark(Color.White))
                {
                    DashStyle = System.Drawing.Drawing2D.DashStyle.Dot,
                    Width = 2
                };
                using var focusPath = GetRoundRectangle(focusRect, Math.Max(BorderRadius - 2, 2));
                pevent.Graphics.DrawPath(focusPen, focusPath);
            }
        }

        private System.Drawing.Drawing2D.GraphicsPath GetRoundRectangle(Rectangle rect, int radius)
        {
            var path = new System.Drawing.Drawing2D.GraphicsPath();
            int diameter = radius * 2;
            if (diameter > rect.Width) diameter = rect.Width;
            if (diameter > rect.Height) diameter = rect.Height;
            path.AddArc(rect.X, rect.Y, diameter, diameter, 180, 90);
            path.AddArc(rect.Right - diameter, rect.Y, diameter, diameter, 270, 90);
            path.AddArc(rect.Right - diameter, rect.Bottom - diameter, diameter, diameter, 0, 90);
            path.AddArc(rect.X, rect.Bottom - diameter, diameter, diameter, 90, 90);
            path.CloseFigure();
            return path;
        }
    }
}
