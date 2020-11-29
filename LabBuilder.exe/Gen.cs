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
    public class Gen
    {
        public static void Users(int userCount, string dePath, string defpwd, object sender, DoWorkEventArgs dwea)
        {
            int successCount = 0;
            int iterationCount = 0;
            int consecutiveFailures = 0;
            var bw = sender as BackgroundWorker;
            var de = new DirectoryEntry(dePath);

            // without this AuthenticationType property, one of my dev boxes would 
            // prompt for smart card during "setpassword" call
            // seemed to be from re-binding to LDAP using SSL
            // see: https://social.technet.microsoft.com/Forums/scriptcenter/en-US/d2c9347c-cfd6-44a8-8baf-7c5202999177/adsi-setpassword-requests-smart-card-pin?forum=ITCG
            // https://gitlab.i.caplette.org/open-source/labbuilder/issues/1
            de.AuthenticationType = AuthenticationTypes.Secure | AuthenticationTypes.Signing | AuthenticationTypes.Sealing;

            bw.ReportProgress(0, "Initializing...");
            var r = new Random();

            ts.TraceEvent(TraceEventType.Verbose, 0, "Creating {0} users...", userCount);

            // load and decompress name list if necessary
            if (names == null)
            {
                MemoryStream ms = null;
                try
                {
                    ms = new MemoryStream(LabBuilder.Properties.Resources.imdbActresses_txt);

                    using (DeflateStream ds = new DeflateStream(ms, CompressionMode.Decompress))
                    {
                        ms = null;
                        names = (List<string>)bf.Deserialize(ds);
                    }
                }

                finally
                {
                    if (ms != null)
                        ms.Dispose();
                }


                ts.TraceEvent(TraceEventType.Verbose, 0, "Initialized {0:N0} possible names...", names.Count);
            }
            while (successCount < userCount && consecutiveFailures < maxConsecutiveFailuresCreatingUsers)
            {
                // check for cancellation
                if (bw.CancellationPending)
                {
                    dwea.Cancel = true;
                    break;
                }


                bw.ReportProgress((int)((double)successCount / userCount * 100), string.Format(CultureInfo.CurrentCulture, "Created {0:N0} of {1:N0} users", successCount, userCount));


                iterationCount++;

                // setup new user info

                string name = names[r.Next(names.Count)];

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

                if (usedNames.Contains(SAMAccountName)) continue;
                usedNames.Add(SAMAccountName);

                RandomAttributes ra = new RandomAttributes();

                ts.TraceEvent(TraceEventType.Verbose, 0,
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
                    newuser.Properties["extensionAttribute13"].Value = "{" + defpwd;

                    newuser.CommitChanges();

                    newuser.Invoke("SetPassword", new object[] { defpwd });
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
                        ts.TraceEvent(TraceEventType.Error, 0, "creating user {0}. {1}: {2}", name, ce.ErrorCode, ce.ToString());
                    }
                }
                catch (Exception e)
                {
                    consecutiveFailures++;
                    ts.TraceEvent(TraceEventType.Error, 0, "creating user {0}: {1}", name, e.ToString());
                }
               
            }
            if (bw.CancellationPending)
            {
                bw.ReportProgress(0, string.Format(CultureInfo.CurrentCulture, "Cancelled.  Created {0:N0} users.", successCount));
            }
            else if (consecutiveFailures == maxConsecutiveFailuresCreatingUsers)
            {
                bw.ReportProgress(0, string.Format(CultureInfo.CurrentCulture, "Stopped due to too many failures.  Created {0:N0} users.", successCount));
            }
            else
            {
                bw.ReportProgress(0, string.Format(CultureInfo.CurrentCulture, "Complete.  Created {0:N0} users.", successCount));
            }

            // clean up non-memory resources
            if (de != null)
                de.Dispose();
        }

        public static void Mail(GenMailArgs args, object sender, DoWorkEventArgs dwea)
        {
            ts.TraceEvent(TraceEventType.Verbose, 0, "Gen.Mail started.");

            // set connection limit higher
            ServicePointManager.DefaultConnectionLimit = 100;

            ts.TraceEvent(TraceEventType.Information, 0, "Initializing...");
            mailUsers = MailUsers.GetAllOUMailUsers(args.selectedOUPath);
            ts.TraceEvent(TraceEventType.Verbose, 0, "Got list of {0} users.", mailUsers.Count);

            // lower maxadditionalrecips if there are not enough users
            if (mailUsers.Count - 1 < args.maxAdditionalRecips) args.maxAdditionalRecips = mailUsers.Count - 1;

            // load and decompress word list if necessary
            if (words == null)
            {
                LoadDictionary();

            }
            ts.TraceEvent(TraceEventType.Information, 0, "Preparing recipient lists...");

            var progress = new GenMailProgress();
            var bw = sender as BackgroundWorker;
            mailCounter = 0;
            // clear dictionaries in case process was interrupted
            senderAssignments.Clear();
            mailsRemainingForRecipient.Clear();

            // random instance to generate seeds for other random instances
            var rs = new Random();

            // populate mails remaining dictionary

            foreach (GenMailUser mu in mailUsers)
            {
                if (bw.CancellationPending == true)
                {

                    dwea.Cancel = true;
                    return;
                }

                mailsRemainingForRecipient.Add(mu.smtpAddress, args.InitialItems);
            }

            foreach (GenMailUser mu in mailUsers)
            {
                var to = new MailAddress(mu.smtpAddress, mu.displayName);

                while (mailsRemainingForRecipient[mu.smtpAddress] > 0)
                {
                    if (bw.CancellationPending == true)
                    {

                        dwea.Cancel = true;
                        return;
                    }

                    // keep track of # of emails to send for accurate progress later...

                    mailCounter++;

                    // get random user for from address

                    GenMailUser rmu;
                    do
                    {
                        rmu = mailUsers[rs.Next(mailUsers.Count)];
                    }
                    while (rmu == mu && mailUsers.Count > 1);

                    var recipList = new List<MailAddress>() { to };

                    // decrement mail count
                    mailsRemainingForRecipient[mu.smtpAddress]--;

                    // optionally add some random recipients if there are enough people in the mailusers list

                    if (args.maxAdditionalRecips > 0 && rs.Next(1, 101) > (100 - args.percentChanceOfExtraRecips))
                    {
                        // add 1 - configured max
                        var extraRecips = rs.Next(1, args.maxAdditionalRecips + 1);

                        // keep track of additional recips that we don't allow dupes

                        var genMailRecips = new List<GenMailUser>();

                        for (int i = 0; i < extraRecips; i++)
                        {
                            GenMailUser recip;
                            do
                            {
                                recip = mailUsers[rs.Next(mailUsers.Count)];
                            }
                            while (recip == mu || genMailRecips.Contains(recip)); // don't want duplicates 

                            if (mailsRemainingForRecipient[recip.smtpAddress] > 0)
                            {
                                recipList.Add(new MailAddress(recip.smtpAddress, recip.displayName));
                                mailsRemainingForRecipient[recip.smtpAddress]--;
                                genMailRecips.Add(recip);
                            }
                        }
                    }

                    if (senderAssignments.ContainsKey(rmu))
                    {
                        senderAssignments[rmu].Add(recipList);
                    }
                    else
                    {
                        var assignments = new List<List<MailAddress>>();
                        assignments.Add(recipList);
                        senderAssignments.Add(rmu, assignments);
                    }

                }
            }
            // free up memory 
            mailsRemainingForRecipient.Clear();

            var tasks = senderAssignments.Count;
            progress.numberOfSenders = tasks;
            progress.numberOfMailboxes = mailUsers.Count;
            progress.numberOfMailsToSend = mailCounter;


            var actions = new Action[tasks];
            int actionIndex = 0;

            foreach (var senderGenMailUser in senderAssignments.Keys)
            {
                if (bw.CancellationPending == true)
                {

                    dwea.Cancel = true;
                    return;
                }

                var smfua = new SendMailFromUserArgs();

                smfua.backgroundWorker = bw;
                smfua.doWorkEventArgs = dwea;

                smfua.progress = progress;
                smfua.User = senderGenMailUser;
                smfua.randomInstance = new Random(rs.Next());
                smfua.defaultPassword = args.DefaultPassword;
                smfua.percentChanceOfAttachments = args.percentChanceOfAttachments;
                smfua.maxAttachments = args.maxAttachments;


                actions[actionIndex] = () =>
                {
                    SendMailFromUser(smfua);
                };

                actionIndex++;

            }


            ts.TraceEvent(TraceEventType.Information, 0, "Starting send mail threads...");


            if (mailCounter < args.InitialItems * mailUsers.Count)
                ts.TraceEvent(TraceEventType.Information, 0, "Since some items will be received by multiple mailboxes,\n  {0:N0} items per mailbox can be achieved by sending {1:N0} item(s) total.", args.InitialItems, mailCounter);
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

            senderAssignments.Clear();

            if (dwea.Cancel)
            {
                if (progress.cancelled == 2)
                    ts.TraceEvent(TraceEventType.Information, 0, "Mailbox population cancelled due to too many consecutive errors.");
                if (progress.cancelled == 1)
                    ts.TraceEvent(TraceEventType.Information, 0, "Mailbox population cancelled by user.");
            }
            else
            {
                ts.TraceEvent(TraceEventType.Information, 0, "Done sending mail.");
                if (mailCounter < args.InitialItems * mailUsers.Count)
                    ts.TraceEvent(TraceEventType.Information, 0, "Some messages were sent to multiple mailboxes.\n  Sent {0:N0} item(s) total in order for each mailbox to receive {1:N0} item(s).", mailCounter, args.InitialItems);

            }

        }

        public static void LoadDictionary()
        {
            ts.TraceEvent(TraceEventType.Verbose, 0, "Decompress and load word dictionary...");
            MemoryStream ms = null;
            try
            {
                ms = new MemoryStream(LabBuilder.Properties.Resources.words_txt);
                using (DeflateStream ds = new DeflateStream(ms, CompressionMode.Decompress))
                {
                    ms = null;
                    words = (List<string>)bf.Deserialize(ds);
                }
            }

            finally
            {
                if (ms != null)
                    ms.Dispose();
            }
            ts.TraceEvent(TraceEventType.Verbose, 0, "Initialized {0:N0} possible words...", words.Count);
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

            using (var curSmtpClient = new SmtpClient(MailUsers.dblist[fromMu.exchangeDBIndex].CurrentServer, 587))
            {
                curSmtpClient.Credentials = fromMu.networkCred;

                var from = new MailAddress(fromMu.smtpAddress, fromMu.displayName);
                var progress = args.progress;

                var bw = args.backgroundWorker;
                var dwea = args.doWorkEventArgs;
                var r = args.randomInstance;
                var mailBodyStyle = new BodyCompositionRules();

                // only trace here if worker is not being cancelled.  Otherwise if verbose tracing is enabled it can hang UI.
                if (!dwea.Cancel) ts.TraceEvent(TraceEventType.Verbose, 0, "Starting sending mail from {0}", fromMu.displayName);

                foreach (var mal in senderAssignments[fromMu])
                {
                    // check for user cancellation 

                    if (bw.CancellationPending)
                    {
                        if (!dwea.Cancel) dwea.Cancel = true;

                        Interlocked.CompareExchange(ref progress.cancelled, 1, 0);

                        bw.ReportProgress(0, progress);
                        return;

                    }

                    // check for too many errors

                    if (progress.consecutiveErrorsSending >= maxConsecutiveFailuresSending)
                    {
                        if (!dwea.Cancel) dwea.Cancel = true;

                        if (Interlocked.Exchange(ref progress.cancelled, 2) == 0)
                        {
                            //only trace this once, if not already pending cancel
                            ts.TraceEvent(TraceEventType.Critical, 0, "Stopping after {0} consecutive errors", maxConsecutiveFailuresSending);
                        }

                        bw.ReportProgress(0, progress);
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
                        Interlocked.Increment(ref progress.messagesSent);
                        Interlocked.Exchange(ref progress.consecutiveErrorsSending, 0);

                    }
                    catch (Exception e)
                    {
                        Interlocked.Increment(ref progress.errorsSending);
                        Interlocked.Increment(ref progress.consecutiveErrorsSending);
                        ts.TraceEvent(TraceEventType.Error, 0, "sending message from user {0}.  {1}", fromMu.displayName, e.ToString());
                    }
                    finally
                    {
                        if (message != null)
                            message.Dispose();
                    }


                    bw.ReportProgress(0, progress);
                }

                Interlocked.Increment(ref progress.sendersDone);
                ts.TraceEvent(TraceEventType.Verbose, 0, "Finished sending from {0}", fromMu.displayName);
                bw.ReportProgress(0, progress);
            }

        }

        public static Attachment getRandomAttachment(Random r, BodyCompositionRules bcr)
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


        public static string getRandomBody(Random r, BodyCompositionRules bcr)
        {
            StringBuilder sb = new StringBuilder();
            int numwords = r.Next(10, maxWordsInBody + 1);
            int wordsInSentence = 1;
            int sentencesInParagraph = 1;
            bcr.currentSentenceLength = r.Next(3, 18);
            bcr.currentParagraphLength = r.Next(2, 6);

            for (int i = 0; i < numwords; i++)
            {
                string word = words[r.Next(words.Count)];
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

        public static string getRandomSubject(Random r)
        {
            StringBuilder sb = new StringBuilder();
            int numwords = r.Next(2, maxWordsInSubject + 1);
            for (int i = 0; i < numwords; i++)
            {
                sb.Append(words[r.Next(words.Count)] + " ");
            }

            return sb.ToString();
        }
        readonly static TraceSource ts = MainWindow.ts;
        static BinaryFormatter bf = new BinaryFormatter();
        static List<string> names;
        static List<string> words;
        static HashSet<string> usedNames = new HashSet<string>();
        const int maxConsecutiveFailuresCreatingUsers = 500;
        const int maxConsecutiveFailuresSending = 1000;
        static List<GenMailUser> mailUsers = new List<GenMailUser>();
        const int maxWordsInSubject = 7;
        const int maxWordsInBody = 300;
        static Dictionary<GenMailUser, List<List<MailAddress>>> senderAssignments = new Dictionary<GenMailUser, List<List<MailAddress>>>();
        static Dictionary<string, int> mailsRemainingForRecipient = new Dictionary<string, int>();
        private static int mailCounter;

    }

    public class BodyCompositionRules
    {
        public int currentSentenceLength;
        public int currentParagraphLength;
    }
    public class GenUsersArgs
    {
        public string DEPath;
        public int UserCount;
        public string DefaultPassword;

    }

    public class GenMailArgs
    {
        public string selectedOUPath;
        public string DefaultPassword;
        public int InitialItems;
        public int Threads;
        public int percentChanceOfExtraRecips;
        public int maxAdditionalRecips;
        public int percentChanceOfAttachments;
        public int maxAttachments;
    }

    public class GenMailSettings
    {
        int initialItems;
        string defaultPassword;
        int threads;
        int percentChanceOfExtraRecips;
        int maxAdditionalRecips;
        int percentChanceOfAttachments;
        int maxAttachments;

        public GenMailSettings()
        {
            defaultPassword = "kvs";
            initialItems = 100;
            threads = 4;
            percentChanceOfAttachments = 10;
            percentChanceOfExtraRecips = 20;
            maxAdditionalRecips = 5;
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
        public int PercentChanceOfExtraRecips
        {
            get { return percentChanceOfExtraRecips; }
            set
            {
                percentChanceOfExtraRecips = value;
            }
        }
        public int MaxAdditionalRecips
        {
            get { return maxAdditionalRecips; }
            set
            {
                maxAdditionalRecips = value;
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
        public GenMailProgress progress;
        public BackgroundWorker backgroundWorker;
        public DoWorkEventArgs doWorkEventArgs;
        public Random randomInstance;
        public string defaultPassword;
        public int percentChanceOfAttachments;
        public int maxAttachments;
    }


    public class GenMailUser
    {
        public string smtpAddress;
        public int exchangeDBIndex;
        public string exchangeDBName;
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
