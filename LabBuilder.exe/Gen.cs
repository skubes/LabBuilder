using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Runtime.Serialization.Formatters.Binary;
using System.IO.Compression;
using System.IO;
using System.Text.RegularExpressions;
using System.DirectoryServices;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Threading;
using System.Net.Mail;
using System.Net;
using System.Globalization;

namespace LabBuilder
{

    public static class Gen
    {

        // public methods

        public static void Users(int userCount, string dePath, string defaultPassword, object sender, CancelEventArgs doWorkEventArgs)
        {
            int successCount = 0;
            int iterationCount = 0;
            int consecutiveFailures = 0;
            var bw = sender as BackgroundWorker;
            WorkerWrapper ww = null;
            if (bw == null)
            {
                ww = new WorkerWrapper();
            }
            else
            {
                ww = new WorkerWrapper(bw);
            }
            var de = new DirectoryEntry(dePath);

            ww.ReportProgress(0, "Initializing...");
            var r = new Random();

           

            // load and decompress name list if necessary
            if (_names == null)
            {
                MemoryStream ms = null;
                try
                {
                    ms = new MemoryStream(LabBuilder.Properties.Resources.imdbActresses_txt);

                    using (DeflateStream ds = new DeflateStream(ms, CompressionMode.Decompress))
                    {
                        ms = null;
                        _names = (List<string>)_bf.Deserialize(ds);
                    }
                }

                finally
                {
                    if (ms != null)
                        ms.Dispose();
                }


                _ts.TraceEvent(TraceEventType.Verbose, 0, "Loaded {0:N0} possible names...", _names.Count);
            }
            _ts.TraceEvent(TraceEventType.Information, 0, "Creating {0} users...", userCount);
            while (successCount < userCount && consecutiveFailures < MAX_CONSECUTIVE_FAILURES_CREATING_USERS)
            {
                // check for cancellation
                if (ww.CancellationPending)
                {
                    doWorkEventArgs.Cancel = true;
                    break;
                }
           
              

                ww.ReportProgress((int)((double)successCount / userCount * 100), string.Format(CultureInfo.CurrentCulture, "Created {0:N0} of {1:N0} users", successCount, userCount));


                iterationCount++;

                // setup new user info

                string name = _names[r.Next(_names.Count)];

                string firstname;
                string lastname;
                string SAMAccountName;
                string wholeName;

                // remove parens (with contents) from  name
                name = Regex.Replace(name, @"\([^)]+\)", string.Empty);

                // if name has comma, get firstname and last name
                // TODO: Middle name processing?

                if (name.Contains(','))
                {
                    int commapos = name.IndexOf(',');
                    firstname = name.Substring(commapos + 1).Trim();
                    lastname = name.Substring(0, commapos).Trim();

                }
                else // no comma, no first name
                {
                    lastname = name.Trim();
                    firstname = "";
                }


                if (firstname.Length > 0)
                {
                    SAMAccountName = firstname + "." + lastname;
                    wholeName = firstname + " " + lastname;
                }
                else
                {
                    SAMAccountName = lastname;
                    wholeName = lastname;
                }

                // Name may have accented chars, replace with same non-accented characters for SAMAccountName
                // trick I learned from Stack OVerflow, no idea why Cyrillic
                SAMAccountName = Encoding.ASCII.GetString(Encoding.GetEncoding("Cyrillic").GetBytes(SAMAccountName));

                // replace all sequential non-word characters with a period
                SAMAccountName = Regex.Replace(SAMAccountName, @"[^\w]+", ".");

                // 20 char limit for SAMAccountName
                if (SAMAccountName.Length > 20) SAMAccountName = SAMAccountName.Substring(0, 20);

                // SAMAccountName cannot end (or begin) with a dot

                SAMAccountName = SAMAccountName.Trim('.');

                // make sure SAMAccountName is unique (for this run of the program)
                // not checking with AD for performance reasons, will just get object exists error
                // and move on. (Don't optimize the uncommon cases!!)

                if (_usedNames.Contains(SAMAccountName)) continue;
                _usedNames.Add(SAMAccountName);

                RandomAttributes ra = new RandomAttributes();

                _ts.TraceEvent(TraceEventType.Verbose, 0,
                    "Creating name: {0}, SAMAccountName: {1}, Title: {2}, Office Phone: {3}, Street: {4}",
                   name, SAMAccountName, ra.Title, ra.OfficePhone, ra.StreetAddress);

                try
                {
                    DirectoryEntry newuser = de.Children.Add
                        ("CN=" + wholeName, "user");
                    newuser.Properties["samAccountName"].Value = SAMAccountName;
                    newuser.Properties["userPrincipalName"].Value = SAMAccountName + "@" + DirectoryEntryContext.DefaultDomain;

                    if (firstname.Length > 0)
                        newuser.Properties["givenName"].Value = firstname;

                    newuser.Properties["SN"].Value = lastname;
                    newuser.Properties["displayName"].Value = wholeName;
                    newuser.Properties["l"].Value = ra.City;
                    newuser.Properties["c"].Value = ra.Country;
                    newuser.Properties["department"].Value = ra.Department;
                    newuser.Properties["division"].Value = ra.Division;
                    newuser.Properties["employeeID"].Value = ra.EmployeeID;
                    newuser.Properties["homePhone"].Value = ra.HomePhone;
                    newuser.Properties["physicalDeliveryOfficeName"].Value = ra.Office;
                    newuser.Properties["telephoneNumber"].Value = ra.OfficePhone;
                    newuser.Properties["o"].Value = ra.Organization;
                    newuser.Properties["postalCode"].Value = ra.PostalCode;
                    newuser.Properties["st"].Value = ra.State;
                    newuser.Properties["streetAddress"].Value = ra.StreetAddress;
                    newuser.Properties["title"].Value = ra.Title;
                    newuser.Properties["extensionAttribute13"].Value = "{" + defaultPassword;

                    newuser.CommitChanges();

                    newuser.Invoke("SetPassword", new object[] { defaultPassword });
                    // get existing value
                    int val = (int)newuser.Properties["userAccountControl"].Value;

                    //set flags to enable account and disable password expiry
                    newuser.Properties["userAccountControl"].Value = (val & ~0x2) | 0x10000;

                    newuser.CommitChanges();
                    newuser.Close();

                    successCount++;
                    consecutiveFailures = 0;

                }
                catch (COMException ce)
                {

                    if (!(ce.ErrorCode == unchecked((int)0x80071392) || ce.ErrorCode == unchecked((int)0x8007202F)))
                    {
                        // It's not "The object already exists" or "A constraint violation occurred"
                        // which is expected error if user exists.  If user exists in another OU the
                        // constaint error is given.

                        consecutiveFailures++;
                        _ts.TraceEvent(TraceEventType.Error, 0, "creating user {0}. {1}: {2}", name, ce.ErrorCode, ce.ToString());
                    }
                }
                catch (Exception e)
                {
                    consecutiveFailures++;
                    _ts.TraceEvent(TraceEventType.Error, 0, "creating user {0}: {1}", name, e.ToString());
                }

            }
            if (ww.CancellationPending)
            {
                ww.ReportProgress(0, string.Format(CultureInfo.CurrentCulture, "Cancelled.  Created {0:N0} users.", successCount));
            }
            else if (consecutiveFailures == MAX_CONSECUTIVE_FAILURES_CREATING_USERS)
            {
                ww.ReportProgress(0, string.Format(CultureInfo.CurrentCulture, "Stopped due to too many failures.  Created {0:N0} users.", successCount));
                _ts.TraceEvent(TraceEventType.Critical, 0, "Stopped due to too many failures.  Created {0:N0} users.", successCount);
            }
            else
            {
                ww.ReportProgress(0, string.Format(CultureInfo.CurrentCulture, "Complete.  Created {0:N0} users.", successCount));
                _ts.TraceEvent(TraceEventType.Information, 0, "Complete.  Created {0:N0} users.", successCount);
            }

            // clean up non-memory resources
            if (de != null)
                de.Dispose();
        }
        public static void Users(int usersToCreate, string ouPath, string defaultPassword)
        {
            // called from powershell / api (non-UI)

            Users(usersToCreate, ouPath, defaultPassword, null, null);
            _ts.Flush();
        }

        public static void Mail(GenMailArgs args, object sender, DoWorkEventArgs doWorkEventArgs)
        {
            _ts.TraceEvent(TraceEventType.Information, 0, "Gen.Mail started.");
            if (args == null) throw new ArgumentNullException("args");
            // set connection limit higher
            ServicePointManager.DefaultConnectionLimit = 100;

            _ts.TraceEvent(TraceEventType.Information, 0, "Initializing...");
            _mailUsers = MailUsers.GetAllOUMailUsers(args.SelectedOUPath);
            _ts.TraceEvent(TraceEventType.Verbose, 0, "Got list of {0} mail users.", _mailUsers.Count);

            // lower maxadditionalrecips if there are not enough users
            if (_mailUsers.Count - 1 < args.MaxAdditionalRecipients) args.MaxAdditionalRecipients = _mailUsers.Count - 1;

            // load and decompress word list if necessary
            loadWordListIfNecessary();
            setupObjectsForNewRun(sender, doWorkEventArgs);

            // random instance to generate seeds for other random instances
            var rs = new Random();

            populateMailsRemainingDictionary(args.InitialItems);

            // check for cancelation after potentially long running
            // operations
            if (ShouldCancel) return;


            createSenderAssignments(args, rs);

            if (ShouldCancel) return;

            var tasks = _senderAssignments.Count;
            _mailProgress.numberOfSenders = tasks;
            _mailProgress.numberOfMailboxes = _mailUsers.Count;
            _mailProgress.numberOfMailsToSend = _mailCounter;


            var actions = new Action[tasks];
            int actionIndex = 0;

            foreach (var senderGenMailUser in _senderAssignments.Keys)
            {
                if (ShouldCancel) return;

                var smfua = new SendMailFromUserArgs();

                smfua.User = senderGenMailUser;
                smfua.randomInstance = new Random(rs.Next());
                smfua.defaultPassword = args.DefaultPassword;
                smfua.percentChanceOfAttachments = args.PercentChanceOfAttachments;
                smfua.maxAttachments = args.MaxAttachments;


                actions[actionIndex] = () =>
                {
                    SendMailFromUser(smfua);
                };

                actionIndex++;

            }


            _ts.TraceEvent(TraceEventType.Information, 0, "Starting send mail threads...");


            if (_mailCounter < args.InitialItems * _mailUsers.Count)
                _ts.TraceEvent(TraceEventType.Information, 0, "Since some items will be received by multiple mailboxes,\n  {0:N0} items per mailbox can be achieved by sending {1:N0} item(s) total.", args.InitialItems, _mailCounter);
            var opts = new ParallelOptions();

            //sanity check thread count

            if (args.Threads > 0 && args.Threads <= 100)
            {
                opts.MaxDegreeOfParallelism = args.Threads;
            }
            else
            {
                opts.MaxDegreeOfParallelism = 3;
            }


            // invoke  threads.  This blocks until all complte

            Parallel.Invoke(opts, actions);

            // clear sender assignments

            _senderAssignments.Clear();

            if (_mailProgress.cancelled > 0)
            {
                if (_mailProgress.cancelled == 2)
                    _ts.TraceEvent(TraceEventType.Warning, 0, "Mailbox population cancelled due to too many consecutive errors.");
                if (_mailProgress.cancelled == 1)
                    _ts.TraceEvent(TraceEventType.Information, 0, "Mailbox population cancelled by user.");
            }
            else
            {
                _ts.TraceEvent(TraceEventType.Information, 0, "Done sending mail.");
                if (_mailCounter < args.InitialItems * _mailUsers.Count)
                    _ts.TraceEvent(TraceEventType.Information, 0, "Some messages were sent to multiple mailboxes.\n  Sent {0:N0} item(s) total in order for each mailbox to receive {1:N0} item(s).", _mailCounter, args.InitialItems);

            }

        }
        public static void Mail(GenMailArgs args)
        {
            // called from powershell / api (non-UI)

            Mail(args, null, null);
            // flush log after run
            _ts.Flush();
        }
       

        public static string InitializeFileLogging()
        {
            var tl = new TxtFileTraceListener();
            _ts.Listeners.Add(tl);
            // to do:  set trace level based on parameter
            _ts.Switch.Level = SourceLevels.Verbose;
            return tl.LogFilePath;

        }

        // private methods

        private static void createSenderAssignments(GenMailArgs args, Random rs)
        {
            _ts.TraceEvent(TraceEventType.Information, 0, "Creating sender assignments...");
            foreach (GenMailUser mu in _mailUsers)
            {
                var to = new MailAddress(mu.smtpAddress, mu.displayName);

                while (_mailsRemainingForRecipient[mu.smtpAddress] > 0)
                {
                    if (ShouldCancel) return;

                    // keep track of # of emails to send for accurate progress later...

                    _mailCounter++;

                    // get random user for from address

                    GenMailUser rmu;
                    do
                    {
                        rmu = _mailUsers[rs.Next(_mailUsers.Count)];
                    }
                    while (rmu == mu && _mailUsers.Count > 1);

                    var recipList = new List<MailAddress>() { to };

                    // decrement mail count
                    _mailsRemainingForRecipient[mu.smtpAddress]--;

                    // optionally add some random recipients if there are enough people in the mailusers list

                    if (args.MaxAdditionalRecipients > 0 && rs.Next(1, 101) > (100 - args.PercentChanceOfExtraRecipients))
                    {
                        // add 1 - configured max
                        var extraRecips = rs.Next(1, args.MaxAdditionalRecipients + 1);

                        // keep track of additional recips that we don't allow dupes

                        var genMailRecips = new List<GenMailUser>();

                        for (int i = 0; i < extraRecips; i++)
                        {
                            GenMailUser recip;
                            do
                            {
                                recip = _mailUsers[rs.Next(_mailUsers.Count)];
                            }
                            while (recip == mu || genMailRecips.Contains(recip)); // don't want duplicates 

                            if (_mailsRemainingForRecipient[recip.smtpAddress] > 0)
                            {
                                recipList.Add(new MailAddress(recip.smtpAddress, recip.displayName));
                                _mailsRemainingForRecipient[recip.smtpAddress]--;
                                genMailRecips.Add(recip);
                            }
                        }
                    }

                    if (_senderAssignments.ContainsKey(rmu))
                    {
                        _senderAssignments[rmu].Add(recipList);
                    }
                    else
                    {
                        var assignments = new List<List<MailAddress>>();
                        assignments.Add(recipList);
                        _senderAssignments.Add(rmu, assignments);
                    }

                }
            }
            // free up memory 
            _mailsRemainingForRecipient.Clear();
        }

        private static void populateMailsRemainingDictionary(int initialItems)
        {
            foreach (GenMailUser mu in _mailUsers)
            {
                if (ShouldCancel) return;
                _mailsRemainingForRecipient.Add(mu.smtpAddress, initialItems);
            }
        }

        private static void setupObjectsForNewRun(object sender, DoWorkEventArgs dea)
        {
            var bw = sender as BackgroundWorker;

            if (sender == null)
            {
                _ww = new WorkerWrapper();
            }
            else
            {
                _ww = new WorkerWrapper(bw);
            }
            _doWorkEventArgs = dea;
            _mailCounter = 0;
            // clear dictionaries in case process was interrupted
            // before completion on previous run
            _senderAssignments.Clear();
            _mailsRemainingForRecipient.Clear();
            _mailProgress = new GenMailProgress();
        }

        private static void loadWordListIfNecessary()
        {
            if (_words == null)
            {
                _ts.TraceEvent(TraceEventType.Verbose, 0, "Need to decompress word dictionary...");
                MemoryStream ms = null;
                try
                {
                    ms = new MemoryStream(LabBuilder.Properties.Resources.words_txt);
                    using (DeflateStream ds = new DeflateStream(ms, CompressionMode.Decompress))
                    {
                        ms = null;
                        _words = (List<string>)_bf.Deserialize(ds);
                    }
                }

                finally
                {
                    if (ms != null)
                        ms.Dispose();
                }
                _ts.TraceEvent(TraceEventType.Verbose, 0, "Initialized {0:N0} possible words...", _words.Count);

            }
        }


        // this method is called from Parallel Invoke, so multiple threads will be executing this concurrently
        private static void SendMailFromUser(SendMailFromUserArgs args)
        {
            var fromMu = args.User;

            // make sure creds are set and setup SMTP client instance for sender to use this session..

            if (!(fromMu.networkCred is NetworkCredential))
            {
                fromMu.networkCred = new NetworkCredential(fromMu.sAMAccountName, args.defaultPassword, DirectoryEntryContext.DefaultDomain);

            }

            using (var curSmtpClient = new SmtpClient(MailUsers.dblist[fromMu.exchangeDBIndex].CurrentServer))
            {
                curSmtpClient.Credentials = fromMu.networkCred;

                var from = new MailAddress(fromMu.smtpAddress, fromMu.displayName);
                var r = args.randomInstance;
                var mailBodyStyle = new BodyCompositionRules();

                // only trace here if worker is not being cancelled.  Otherwise if verbose tracing is enabled it can hang UI.
                if (!_ww.CancellationPending) _ts.TraceEvent(TraceEventType.Verbose, 0, "Starting sending mail from {0}", fromMu.displayName);

                foreach (var mal in _senderAssignments[fromMu])
                {
                    // check for user cancellation 

                    if (ShouldCancel)
                    {
                        Interlocked.CompareExchange(ref _mailProgress.cancelled, 1, 0);
                        _ww.ReportProgress(0, _mailProgress);
                        return;
                    }

                    // check for too many errors

                    if (_mailProgress.consecutiveErrorsSending >= MAX_CONSECUTIVE_FAILURES_SENDING_MAIL)
                    {
             

                        if (Interlocked.Exchange(ref _mailProgress.cancelled, 2) == 0)
                        {
                            //only trace this once, if not already pending cancel
                            _ts.TraceEvent(TraceEventType.Critical, 0, "Stopping after {0} consecutive errors", MAX_CONSECUTIVE_FAILURES_SENDING_MAIL);
                        }

                        _ww.ReportProgress(0, _mailProgress);
                        return;

                    }

                    // compose message

                    var message = new MailMessage(from, mal[0]);
                    if (mal.Count > 1)
                    {
                        for (int i = 1; i < mal.Count; i++)
                        {
                            message.To.Add(mal[i]);
                        }
                    }
                    message.Body = getRandomBody(r, mailBodyStyle);
                    message.Subject = getRandomSubject(r);

                    // attachments

                    if (r.Next(1, 101) > (100 - args.percentChanceOfAttachments))
                    {
                        message.Attachments.Add(getRandomAttachment(r, mailBodyStyle));
                        if (r.Next(1, 101) > 50) // 50 percent chance of more than one attachment.
                        {
                            var attachments = r.Next(1, args.maxAttachments + 1);

                            for (int i = 1; i < attachments; i++) // start i=1 because we already have 1 attachment
                            {
                                message.Attachments.Add(getRandomAttachment(r, mailBodyStyle));
                            }
                        }

                    }

                    // send  

                    try
                    {
                        curSmtpClient.Send(message);
                        Interlocked.Increment(ref _mailProgress.messagesSent);
                        Interlocked.Exchange(ref _mailProgress.consecutiveErrorsSending, 0);

                    }
                    catch (Exception e)
                    {
                        Interlocked.Increment(ref _mailProgress.errorsSending);
                        Interlocked.Increment(ref _mailProgress.consecutiveErrorsSending);
                        _ts.TraceEvent(TraceEventType.Error, 0, "sending message from user {0}.  {1}", fromMu.displayName, e.ToString());
                    }
                    finally
                    {
                        if (message != null)
                            message.Dispose();
                    }


                    _ww.ReportProgress(0, _mailProgress);
                }

                Interlocked.Increment(ref _mailProgress.sendersDone);
                _ts.TraceEvent(TraceEventType.Verbose, 0, "Finished sending from {0}", fromMu.displayName);
                _ww.ReportProgress(0, _mailProgress);
            }

        }

        private static Attachment getRandomAttachment(Random r, BodyCompositionRules bcr)
        {
            string attachmentName = getRandomSubject(r).Replace("'", "").TrimEnd();
            Attachment attach;
            if (r.Next(1, 101) > 30)
            {
                string attachmentText = getRandomBody(r, bcr);
                var stream = new MemoryStream(Encoding.UTF8.GetBytes(attachmentText));
                attach = new Attachment(stream, attachmentName + ".txt");
            }
            else
            {
                // 30 percent chance we want binary attach
                // make attach 900 bytes - ~1.5 MB
                byte[] bytes = new byte[r.Next(900, 1500000)];
                r.NextBytes(bytes);
                var stream = new MemoryStream(bytes);

                attach = new Attachment(stream, attachmentName + ".binary");

            }
            return attach;
        }


        private static string getRandomBody(Random r, BodyCompositionRules bcr)
        {
            StringBuilder sb = new StringBuilder();
            int numwords = r.Next(10, MAX_WORDS_IN_BODY + 1);
            int wordsInSentence = 1;
            int sentencesInParagraph = 1;
            bcr.currentSentenceLength = r.Next(3, 18);
            bcr.currentParagraphLength = r.Next(2, 6);

            for (int i = 0; i < numwords; i++)
            {
                string word = _words[r.Next(_words.Count)];
                if (wordsInSentence == 1)
                {
                    string capitalFirstLetter = word.Substring(0, 1).ToUpper(System.Globalization.CultureInfo.CurrentCulture);

                    sb.Append(capitalFirstLetter);
                    sb.Append(word.Substring(1));
                }
                else
                {
                    sb.Append(word);
                }

                if (wordsInSentence % bcr.currentSentenceLength == 0)
                {
                    sb.Append(". ");
                    bcr.currentSentenceLength = r.Next(3, 18);
                    wordsInSentence = 1;
                    sentencesInParagraph++;
                }
                else
                {
                    if (i == numwords - 1) // last word we are going to add put a period.
                    {
                        sb.Append(".");
                    }
                    else
                    {
                        sb.Append(" ");
                    }

                    wordsInSentence++;
                }
                if (sentencesInParagraph % bcr.currentParagraphLength == 0)
                {
                    sb.Append("\r\n\r\n");
                    bcr.currentParagraphLength = r.Next(2, 6);
                    sentencesInParagraph = 1;
                }

            }

            return sb.ToString();
        }

        private static string getRandomSubject(Random r)
        {
            StringBuilder sb = new StringBuilder();
            int numwords = r.Next(2, MAX_WORDS_IN_SUBJECT + 1);
            for (int i = 0; i < numwords; i++)
            {
                sb.Append(_words[r.Next(_words.Count)] + " ");
            }

            return sb.ToString();
        }

        // class members

        const int MAX_CONSECUTIVE_FAILURES_CREATING_USERS = 500;
        const int MAX_CONSECUTIVE_FAILURES_SENDING_MAIL = 1000;
        const int MAX_WORDS_IN_SUBJECT = 7;
        const int MAX_WORDS_IN_BODY = 300;

        static TraceSource _ts = MainWindow.ts;
        static BinaryFormatter _bf = new BinaryFormatter();
        static List<string> _names;
        static List<string> _words;
        static HashSet<string> _usedNames = new HashSet<string>();

        static List<GenMailUser> _mailUsers = new List<GenMailUser>();

        static Dictionary<GenMailUser, List<List<MailAddress>>> _senderAssignments = new Dictionary<GenMailUser, List<List<MailAddress>>>();
        static Dictionary<string, int> _mailsRemainingForRecipient = new Dictionary<string, int>();
        static int _mailCounter;
        static WorkerWrapper _ww = null;
        static DoWorkEventArgs _doWorkEventArgs = null;
        static GenMailProgress _mailProgress;

        /// <summary>
        ///   Gets a value indicating whether the worker thread should be canceled.
        ///   If so,  DoWorkEventArgs.Cancel is also set to true
        /// </summary>
        public static bool ShouldCancel
        {
            get
            {
                // check for cancelation
                if (_ww.CancellationPending == true)
                {
                    _doWorkEventArgs.Cancel = true;
                    return true;
                }
                return false;
            }
        }

     
    }
    class WorkerWrapper
    {
        BackgroundWorker _bw = null;

        public WorkerWrapper(BackgroundWorker b)
        {
            _bw = b;
        }

        public WorkerWrapper() { }

        public void ReportProgress(int percent, object state)
        {
            if (_bw != null)
                _bw.ReportProgress(percent, state);
        }
        public bool CancellationPending
        {
            get
            {
                if (_bw != null)
                {
                    return _bw.CancellationPending;
                }
                else return false;

            }

        }
    }
    class BodyCompositionRules
    {
        public int currentSentenceLength;
        public int currentParagraphLength;
    }
    public class GenUsersArgs
    {
        public string DEPath { get; set; }
        public int UserCount { get; set; }
        public string DefaultPassword { get; set; }

    }

    public class GenMailArgs
    {
        public string SelectedOUPath { get; set; }
        public string DefaultPassword { get; set; }
        public int InitialItems { get; set; }
        public int Threads { get; set; }
        public int PercentChanceOfExtraRecipients { get; set; }
        public int MaxAdditionalRecipients { get; set; }
        public int PercentChanceOfAttachments { get; set; }
        public int MaxAttachments { get; set; }
    }

    public class GenMailSettings
    {
        int initialItems;
        string defaultPassword;
        int threads;
        int percentChanceOfExtraRecipients;
        int maxAdditionalRecipients;
        int percentChanceOfAttachments;
        int maxAttachments;

        public GenMailSettings()
        {
            defaultPassword = "kvs";
            initialItems = 100;
            threads = 4;
            percentChanceOfAttachments = 10;
            percentChanceOfExtraRecipients = 20;
            maxAdditionalRecipients = 5;
            maxAttachments = 5;

        }
        public int InitialItems
        {
            get { return initialItems; }
            set
            {
                initialItems = value;
            }
        }

        public string DefaultPassword
        {
            get { return defaultPassword; }
            set
            {
                defaultPassword = value;
            }
        }

        public int Threads
        {
            get { return threads; }

            set
            {
                threads = value;
            }
        }
        public int PercentChanceOfExtraRecipients
        {
            get { return percentChanceOfExtraRecipients; }
            set
            {
                percentChanceOfExtraRecipients = value;
            }
        }
        public int MaxAdditionalRecipients
        {
            get { return maxAdditionalRecipients; }
            set
            {
                maxAdditionalRecipients = value;
            }
        }
        public int PercentChanceOfAttachments
        {
            get { return percentChanceOfAttachments; }
            set
            {
                percentChanceOfAttachments = value;
            }
        }
        public int MaxAttachments
        {
            get { return maxAttachments; }
            set
            {
                maxAttachments = value;

            }
        }

    }


    class SendMailFromUserArgs
    {

        public GenMailUser User;

        public Random randomInstance;
        public string defaultPassword;
        public int percentChanceOfAttachments;
        public int maxAttachments;
    }


    class GenMailUser
    {
        public string smtpAddress;
        public int exchangeDBIndex;
        public string password;
        public string displayName;
        public string sAMAccountName;
        public NetworkCredential networkCred;
    }

    class GenMailProgress
    {
        public int messagesSent = 0;
        public int sendersDone = 0;
        public int errorsSending = 0;
        public int consecutiveErrorsSending = 0;
        public int cancelled = 0;
        public int numberOfMailboxes = 0;
        public int numberOfSenders = 0;
        public int numberOfMailsToSend = 0;

    }

}

