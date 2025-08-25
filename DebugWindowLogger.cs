// DebugWindowLogger.cs (完整取代)
using System;
using System.Drawing;
using System.Windows.Forms;

namespace TouchKeyBoard
{
    public class DebugWindowLogger : IDisposable
    {
        private Form _debugWindow;
        private TextBox _debugTextBox;
        private PictureBox _pictureBox; // 新增 PictureBox
        private bool _isDisposed = false;

        public DebugWindowLogger(Point location)
        {
            _debugWindow = new Form
            {
                Text = "TouchKeyBoard Debug Window",
                Size = new Size(500, 600), // 增加高度以容納圖片
                Location = location,
                TopMost = true,
                FormBorderStyle = FormBorderStyle.SizableToolWindow,
                BackColor = Color.Black
            };


            _debugTextBox = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill, // 將填滿剩餘空間
                BackColor = Color.Black,
                ForeColor = Color.Lime,
                Font = new Font("Consolas", 10, FontStyle.Regular),
                ReadOnly = true,
                WordWrap = true
            };

            _debugWindow.Controls.Add(_debugTextBox);
            _debugWindow.Controls.Add(_pictureBox); // 將 PictureBox 加入
            _debugWindow.Show();
        }

        public void Log(string message)
        {
            if (_isDisposed || _debugTextBox == null) return;

            string timestamp = DateTime.Now.ToString("HH:mm:ss");
            string logMessage = $"[{timestamp}] {message}";

            if (_debugTextBox.InvokeRequired)
            {
                _debugTextBox.Invoke(new Action(() => AppendText(logMessage)));
            }
            else
            {
                AppendText(logMessage);
            }
        }

       

        private void AppendText(string text)
        {
            _debugTextBox.AppendText(text + Environment.NewLine);
            _debugTextBox.SelectionStart = _debugTextBox.Text.Length;
            _debugTextBox.ScrollToCaret();
        }

        public void Dispose()
        {
            if (!_isDisposed)
            {
                _pictureBox?.Image?.Dispose();
                _debugWindow?.Close();
                _debugWindow?.Dispose();
                _isDisposed = true;
            }
        }
    }
}