using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
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
}
