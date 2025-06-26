using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace SamsungStatements
{
    public class LogManager
    {
        private readonly TextBox logTextBox;
        private readonly Dispatcher dispatcher;

        public LogManager(TextBox textBox)
        {
            logTextBox = textBox;
            dispatcher = textBox.Dispatcher;
        }

        public void LogMessage(string message)
        {
            var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            var logEntry = $"[{timestamp}] {message}";

            // UI 스레드에서 안전하게 로그 추가
            dispatcher.BeginInvoke(new Action(() =>
            {
                logTextBox.AppendText(logEntry + Environment.NewLine);
                logTextBox.ScrollToEnd();
            }));
        }
    }
} 