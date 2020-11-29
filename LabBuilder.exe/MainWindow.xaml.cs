using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Diagnostics;
using System.ComponentModel;
using System.Threading;
using System.Globalization;

namespace LabBuilder
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, IDisposable
    {
        public MainWindow()
        {
            InitializeComponent();

            //setup tracing
            UIListener = new UITraceListener(LogTextBlock);
            ts.Listeners.Remove("Default");
            ts.Listeners.Add(UIListener);
            ts.Switch.Level = SourceLevels.Information;
            ts.TraceInformation("Welcome to Lab Builder v1");
            try
            {
                ts.TraceInformation("Domain Detected: {0}", DirectoryEntryContext.DefaultDomain);

                tabControl.Visibility = Visibility.Collapsed;

                // setup background worker objects
                _bw_create = new BackgroundWorker();
                _bw_create.WorkerReportsProgress = true;
                _bw_create.WorkerSupportsCancellation = true;
                _bw_create.ProgressChanged += bw_create_ProgressChanged;
                _bw_create.RunWorkerCompleted += bw_create_RunWorkerCompleted;
                _bw_create.DoWork += bw_create_DoWork;

                _bw_send = new BackgroundWorker();
                _bw_send.WorkerReportsProgress = true;
                _bw_send.WorkerSupportsCancellation = true;
                _bw_send.ProgressChanged += _bw_send_ProgressChanged;
                _bw_send.RunWorkerCompleted += _bw_send_RunWorkerCompleted;
                _bw_send.DoWork += _bw_send_DoWork;

                settingsGrid.DataContext = gs;

#if DEBUG
                debugOn.IsChecked = true;
#endif

                 topOU = new DirectoryEntryContext();
                OUBrowser.Items.Add(topOU);
                ExchangeDBListBox.ItemsSource = ExchangeDatabases.DBs;

            }
            catch (Exception e)
            {
                Exception inner = e.InnerException;
                string innerstring = "n/a";
                if (inner != null)
                {
                    innerstring = inner.ToString();
                }
                ts.TraceEvent(TraceEventType.Critical, 0, "Error starting up. \n{0}\nInner Exception: {1}", e.ToString(), innerstring);
                ts.TraceEvent(TraceEventType.Critical, 0, "Is this program running under a domain account?\nIf so possible authentication (Kerberos?) or DNS issue");
            }

        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
            
        }

        private void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (topOU != null)
                {
                    topOU.Dispose();
                }

            }
        }

        void _bw_send_DoWork(object sender, DoWorkEventArgs e)
        {
            var args = e.Argument as GenMailArgs;
            Gen.Mail(args, sender, e);
        }

        void _bw_send_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (userClosed) this.Close();
            StartMailGenButton.IsEnabled = true;
            Settings_Button.IsEnabled = true;
            sendMailProgress.Visibility = Visibility.Hidden;
            if (e.Cancelled)
                SendMailStatusTextBlock.Text = LabBuilder.Properties.Resources.OpStopString;

            if (!(e.Error == null))
            {
                ts.TraceEvent(TraceEventType.Error, 0, "Exception during operation.  {0}", e.Error.ToString());
            }

        }

        private void _bw_send_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            var progress = e.UserState as GenMailProgress;


            mailboxTextBlock.Text = string.Format(CultureInfo.CurrentCulture, LabBuilder.Properties.Resources.PopulationStatusFormatText,
          progress.numberOfMailboxes, progress.messagesSent, progress.errorsSending);

            // update progress bar

            sendMailProgress.Value = ((double)progress.messagesSent) / progress.numberOfMailsToSend * 100;

            if (progress.cancelled == 1)
            {
                SendMailStatusTextBlock.Text = LabBuilder.Properties.Resources.OpCancelPending;
            }
            else if (progress.cancelled == 2)
            {
                SendMailStatusTextBlock.Text = LabBuilder.Properties.Resources.OpCancelDueToFailure;
            }



        }

        private void Create_Users_Click(object sender, RoutedEventArgs e)
        {

            var sel = OUBrowser.SelectedItem as DirectoryEntryContext;
            if (sel != null)
            {
                ts.TraceEvent(TraceEventType.Verbose, 0, "OU selected: {0}", sel.DE.Path);

                if (MessageBox.Show(string.Format(CultureInfo.CurrentCulture, LabBuilder.Properties.Resources.ConfirmUserCreationWindowTextFormat, usersToCreate.Text, sel.ObjectName), LabBuilder.Properties.Resources.ConfirmationWindowTitle, MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.No)
                    return;

                var args = new GenUsersArgs();
                args.DefaultPassword = defaultPassword.Text;
                args.DEPath = sel.DE.Path;
                args.UserCount = int.Parse(usersToCreate.Text,CultureInfo.CurrentCulture);

                progressPanel.Visibility = Visibility.Visible;
                UserProgressBar.Visibility = Visibility.Visible;
                _bw_create.RunWorkerAsync(args);
                CreateUsersButton.IsEnabled = false;
            }
        }

        private void bw_create_DoWork(object sender, DoWorkEventArgs e)
        {
            var args = e.Argument as GenUsersArgs;
            Gen.Users(args.UserCount, args.DEPath, args.DefaultPassword, sender, e);

        }

        private void bw_create_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            UserProgressBar.Value = e.ProgressPercentage;
            var userStateText = e.UserState as string;
            UserProgressText.Text = userStateText;
        }

        private void bw_create_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (userClosed) this.Close();
            CreateUsersButton.IsEnabled = true;
            UserProgressBar.Visibility = Visibility.Hidden;


        }

        static UITraceListener UIListener;
        static BackgroundWorker _bw_create;
        static BackgroundWorker _bw_send;
        internal static TraceSource ts = new TraceSource("EnvironmentGenLoggerPrefix");
        static bool ouSelected = false;
        static bool userClosed = false;
        DirectoryEntryContext topOU;

        static GenMailSettings gs = new GenMailSettings();

        private void Verbose_Unchecked(object sender, RoutedEventArgs e)
        {
            ts.Switch.Level = SourceLevels.Information;
        }

        private void Verbose_Checked(object sender, RoutedEventArgs e)
        {
            ts.Switch.Level = SourceLevels.All;
        }

        private void OUContextMenu_Click(object sender, RoutedEventArgs e)
        {

        }

        private void Cancel_Create_Click(object sender, RoutedEventArgs e)
        {
            if (_bw_create.IsBusy)
            {
                _bw_create.CancelAsync();
                UserProgressText.Text = LabBuilder.Properties.Resources.OpCancelRequested;
            }
        }


        private void Create_Mailboxes_Click(object sender, RoutedEventArgs e)
        {
            if (OUBrowser.SelectedItem is DirectoryEntryContext && ExchangeDBListBox.SelectedItem is ExchangeDatabase)
            {
                var sel = OUBrowser.SelectedItem as DirectoryEntryContext;
                var dbsel = ExchangeDBListBox.SelectedItem as ExchangeDatabase;

                //  make powershell command
                string powershellcommand = "";
                string powershell_command_format = "";

                if (sel.DE.Path.Length > 0)
                {
                    string friendlyOUPath = "'" + sel.DE.Path.Substring(7) + "'";
                    powershell_command_format = @"
Get-User -OrganizationalUnit {0} -resultsize {1} -Filter {{-not(alias -like '*' )}} | Select-Object -ExpandProperty samAccountName | Enable-Mailbox -Database '{2}' | ft
";
                    if (checkBoxMailboxLimit.IsChecked == false)
                    {
                        powershellcommand = string.Format(CultureInfo.InvariantCulture, powershell_command_format, friendlyOUPath, "unlimited", dbsel.DatabaseName);
                    }
                    else
                    {
                        powershellcommand = string.Format(CultureInfo.InvariantCulture, powershell_command_format, friendlyOUPath, MailboxLimit.Text, dbsel.DatabaseName);
                    }
                }
                else
                {
                    // no OU selected, it's root
                    powershell_command_format = @"
Get-User -resultsize {0} -Filter {{-not(alias -like '*' )}} | Select-Object -ExpandProperty samAccountName | Enable-Mailbox -Database '{1}' | ft
";
                    if (checkBoxMailboxLimit.IsChecked == false)
                    {
                        powershellcommand = string.Format(CultureInfo.InvariantCulture, powershell_command_format, "unlimited", dbsel.DatabaseName);
                    }
                    else
                    {
                        powershellcommand = string.Format(CultureInfo.InvariantCulture, powershell_command_format, MailboxLimit.Text, dbsel.DatabaseName);
                    }
                }

                ts.TraceEvent(TraceEventType.Verbose, 0, "Gonna create mailboxes now in: {0}", sel.ObjectName);

                string beginexchsession = string.Format(CultureInfo.InvariantCulture, "Import-PSSession (New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://{0}/PowerShell/ -Authentication Kerberos)", dbsel.CurrentServer);

                ts.TraceEvent(TraceEventType.Information, 0, "Powershell command: \n\n{0}", beginexchsession + powershellcommand);
                Process psp = null;
                try
                {
                    psp = new Process();
                    psp.StartInfo.FileName = "powershell.exe";
                    psp.StartInfo.Arguments = @"-ExecutionPolicy Bypass -NoExit -Command """ + beginexchsession + powershellcommand + @"""";
                    psp.StartInfo.WindowStyle = ProcessWindowStyle.Normal;
                    psp.Start();
                }
               finally
                {
                    if (psp != null)
                    psp.Dispose();
                }
               
            }
            else
            {
                if (!(ExchangeDBListBox.SelectedItem is ExchangeDatabase))
                    ts.TraceEvent(TraceEventType.Information, 0, "Select an Exchange Database to start creating mailboxes.");
            }
        }

        private void checkBoxMailboxLimit_Checked(object sender, RoutedEventArgs e)
        {
            MailboxLimit.IsEnabled = true;
            MailboxLimit.Focus();
        }

        private void OUBrowser_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (ouSelected == false)
            {
                tabControl.Visibility = Visibility.Visible;
                ouSelected = true;
            }
            if (OUBrowser.SelectedItem is DirectoryEntryContext)
                CreateMbxInfoTextBlock.Text = string.Format(CultureInfo.CurrentCulture, LabBuilder.Properties.Resources.MailboxCreationInfo, ((DirectoryEntryContext)OUBrowser.SelectedItem).ObjectName);
            else CreateMbxInfoTextBlock.Text = "";
        }

        private void checkBoxMailboxLimit_Unchecked(object sender, RoutedEventArgs e)
        {
            MailboxLimit.IsEnabled = false;
            MailboxLimit.Text = "";
        }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            if (!(_bw_create == null) && _bw_create.IsBusy)
            {
                _bw_create.CancelAsync();
                userClosed = true;
                e.Cancel = true;
            }
            if (!(_bw_send == null) && _bw_send.IsBusy)
            {
                _bw_send.CancelAsync();
                userClosed = true;
                e.Cancel = true;
            }
        }

        private void Start_Mail_Click(object sender, RoutedEventArgs e)
        {

            if (OUBrowser.SelectedItem is DirectoryEntryContext)
            {

                var selectedOUDEContext = OUBrowser.SelectedItem as DirectoryEntryContext;
                if (MessageBox.Show(string.Format(CultureInfo.CurrentCulture, LabBuilder.Properties.Resources.ConfirmMailboxPopulationWindowTextFormat, selectedOUDEContext.ObjectName), LabBuilder.Properties.Resources.ConfirmationWindowTitle, MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.No)
                    return;
                var args = new GenMailArgs();
                args.selectedOUPath = selectedOUDEContext.DE.Path;
                args.DefaultPassword = gs.DefaultPassword;
                args.InitialItems = gs.InitialItems;
                args.Threads = gs.Threads;
                args.maxAdditionalRecips = gs.MaxAdditionalRecips;
                args.maxAttachments = gs.MaxAttachments;
                args.percentChanceOfAttachments = gs.PercentChanceOfAttachments;
                args.percentChanceOfExtraRecips = gs.PercentChanceOfExtraRecips;

                _bw_send.RunWorkerAsync(args);
                StartMailGenButton.IsEnabled = false;
                Settings_Button.IsEnabled = false;
                SendMailStatusTextBlock.Text = "";
                sendMailProgress.Value = 0;
                sendMailProgress.Visibility = Visibility.Visible;


            }
        }

        private void Cancel_Mail_Click(object sender, RoutedEventArgs e)
        {
            if (_bw_send.IsBusy)
            {
                SendMailStatusTextBlock.Text = LabBuilder.Properties.Resources.OpCancelRequested;
                _bw_send.CancelAsync();
            }
        }

        private void Refresh_DBs_Click(object sender, RoutedEventArgs e)
        {
            ExchangeDBListBox.ItemsSource = null;
            ExchangeDBListBox.ItemsSource = ExchangeDatabases.DBs;
        }

        private void Settings_Button_Click(object sender, RoutedEventArgs e)
        {
            var sw = new SettingsWindow();
            sw.Owner = this;
            sw.DataContext = gs;
            sw.ShowDialog();
        }

        private void refreshOUsButton_Click(object sender, RoutedEventArgs e)
        {
            OUBrowser.Items.Clear();
            ouSelected = false;
            tabControl.Visibility = Visibility.Hidden;
            OUBrowser.Items.Add(topOU);

        }

    }


}
