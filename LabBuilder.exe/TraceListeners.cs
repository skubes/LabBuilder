using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Windows.Controls;

namespace LabBuilder
{
    public class UITraceListener : TraceListener
    {

        private TextBox output;
        private char[] trimChars = { '0', ' ', ':' };

        public UITraceListener(TextBox output)
        {

            base.Name = "Trace";
            this.output = output;
        }
        public override void Write(string message)
        {
            if (message == null) throw new ArgumentNullException("message");
            if (message.StartsWith("EnvironmentGenLoggerPrefix", StringComparison.Ordinal))
            {
                message = message.Replace("EnvironmentGenLoggerPrefix", "");
                message = message.TrimEnd(trimChars) + ": ";
                message = string.Format(CultureInfo.CurrentCulture, "[{0}] <TID:{1:D3}> ", DateTime.Now.ToString(), Thread.CurrentThread.ManagedThreadId) + message;
            }

            Action append = delegate()
            {
                output.AppendText(message);
                output.CaretIndex = output.Text.Length;
                output.ScrollToEnd();
            };
            output.Dispatcher.BeginInvoke(append);

        }
        public override void WriteLine(string message)
        {

            Write(message + Environment.NewLine);
        }

    }

    public class TxtFileTraceListener : TraceListener
    {

        private StreamWriter sw; 
        private char[] trimChars = { '0', ' ', ':' };
        private string _logFilePath;
        public string LogFilePath
        {
            get
            {
                return _logFilePath;
            }
        }

        public TxtFileTraceListener()
        {

            base.Name = "TempFileTrace";
            _logFilePath = Path.Combine(Path.GetTempPath(), "LabBuilder_" + Path.GetRandomFileName() + ".log");
            sw = new StreamWriter(_logFilePath);

        }

        public override void Flush()
        {
            base.Flush();
            if (sw != null) sw.Flush();
        }
        public override void Close()
        {
            if (sw != null) sw.Close();
        
        }

        public override void Write(string message)
        {
            if (message == null) throw new ArgumentNullException("message");
            if (message.StartsWith("EnvironmentGenLoggerPrefix", StringComparison.Ordinal))
            {
                message = message.Replace("EnvironmentGenLoggerPrefix", "");
                message = message.TrimEnd(trimChars) + ": ";
                message = string.Format(CultureInfo.CurrentCulture, "[{0}] <TID:{1:D3}> ", DateTime.Now.ToString(), Thread.CurrentThread.ManagedThreadId) + message;
            }
                sw.Write(message);
        }
        public override void WriteLine(string message)
        {

            Write(message + Environment.NewLine);
        }

    }
}
