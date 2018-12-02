

using System;
using System.Collections.Generic;
using System.Diagnostics;
using static LabBuilder.ADHelper;

using System.Windows;

namespace LabBuilder
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {

        static commandLineArgs argsobj;
        private void go(object sender, StartupEventArgs e)
        {

            if (e.Args.GetLength(0) > 0)
            {

                initCommandMode();
                processArguments(e.Args);
                QuitCleanly();

            }

            else
            {
                this.StartupUri = new System.Uri("MainWindow.xaml", System.UriKind.Relative);
            }

        }

         void QuitCleanly()
        {
            ts.Close();
            Shutdown();
        }

        static void initCommandMode()
        {

            argsobj = new commandLineArgs();
            ts = LabBuilder.MainWindow.ts;
            var txtlogger = new TxtFileTraceListener();
            ts.Listeners.Add(txtlogger);
            ts.Switch.Level = SourceLevels.Information;
        }

        static void processArguments(string[] args)
        {

            foreach (var arg in args)
            {
                ts.TraceEvent(TraceEventType.Information, 0, "Attempting to process argument [{0}]", arg);
                foreach (var progArg in _argumentNames)
                {
                    if (arg.StartsWith("/" + progArg + ":", StringComparison.OrdinalIgnoreCase))
                    {
                        // found named argument

                        var argparts = arg.Split(':');
                        var argval = argparts[1];
                        for (var i = 2; i < argparts.Length; ++i)
                        {
                            // starting i = 2 because want to ignore 
                            // first 2 elements.  If more than 2 parts exist
                            // it's because another colon exists but was removed
                            // in the split operation. Add it back.
                            argval = argval + ":" + argparts[i];


                        }
                        _arguments[progArg] = argval;

                        // don't look for this partiular arg anymore:
                        _argumentNames.Remove(progArg);
                        ts.TraceEvent(TraceEventType.Information, 0, "Processed named argument: [{0}], with value: [{1}]", progArg, argval);
                        // found match so move on to next command line arg
                        break;
                    }
                    else if (arg.ToUpperInvariant() == "/" + progArg.ToUpperInvariant() )
                    {
                        // found switch argument
                        _arguments[progArg] = "on";
                        ts.TraceEvent(TraceEventType.Information, 0, "Processed switch argument: [{0}]", progArg);
                        break;
                    }
                }
            }

            // first process arguments that are shared between modes

            if (_arguments.ContainsKey("debug"))
            {
                // switch arg, if it's there it's 'on'
                ts.Switch.Level = SourceLevels.Verbose;
            }
            if (_arguments.ContainsKey("mode"))
            {
                switch (_arguments["mode"])
                {
                    case "m":
                    case "mail":
                        argsobj.mode = GenMode.Mail;
                        break;
                    case "u":
                    case "user":
                    case "users":
                        argsobj.mode = GenMode.Users;
                        break;
                    default:
                        ts.TraceEvent(TraceEventType.Error, 0, "mode argument invalid. Must be /mode:mail or /mode:users, or shortened versions");
                        return;


                }

            }
            else
            {
                ts.TraceEvent(TraceEventType.Error, 0, "No mode argument given");
                return;
            }

            if (_arguments.ContainsKey("oupath"))
            {
                _arguments["oupath"] = fixupOuPath(_arguments["oupath"]);
                if (!isValidOUPath(_arguments["oupath"]))
                {
                    ts.TraceEvent(TraceEventType.Error, 0, "Invalid oupath argument given");
                    return;
                }
            }
            else
            {
                ts.TraceEvent(TraceEventType.Error, 0, "No oupath argument given");
                return;
            }


          
            switch (argsobj.mode)
            {
                case GenMode.Mail:
                    processMailArgs();
                    break;
                case GenMode.Users:
                    processUsersArgs();
                    break;
            }
        }


        static int getIntArgValueorDefault(string keyName, int def)
        {

            if (_arguments.ContainsKey(keyName))
            {
                if (int.TryParse(_arguments[keyName], out int number))
                {
                    return number;
                }
            }
            return def;

        }
        static string getStringArgValueorDefault(string keyName, string def)
        {

            if (_arguments.ContainsKey(keyName))
            {

                if (_arguments[keyName].Length > 0)
                {
                    return _arguments[keyName];
                }
            }
            return def;

        }

        static void processMailArgs()
        {
            argsobj.mailArgs.InitialItems = getIntArgValueorDefault("itemspermailbox", 10);
            argsobj.mailArgs.Threads = getIntArgValueorDefault("threads", 4);
            argsobj.mailArgs.PercentChanceOfAttachments = getIntArgValueorDefault("percentchanceofattachments", 10);
            argsobj.mailArgs.PercentChanceOfExtraRecipients = getIntArgValueorDefault("percentchanceofextrarecips", 20);
            argsobj.mailArgs.MaxAttachments = getIntArgValueorDefault("maxattachments", 5);
            argsobj.mailArgs.MaxAdditionalRecipients = getIntArgValueorDefault("maxadditionalrecips", 10);
            argsobj.mailArgs.DefaultPassword = getStringArgValueorDefault("password", "kvs");
            argsobj.mailArgs.SelectedOUPath = _arguments["oupath"];

            Gen.Mail(argsobj.mailArgs, null, null);


        }
        static void processUsersArgs()
        {
            int number;
            if (_arguments.ContainsKey("userstocreate"))
            {

                if (!int.TryParse(_arguments["userstocreate"], out number))
                {
                    ts.TraceEvent(TraceEventType.Error, 0, "Invalid value for 'userstocreate'");
                    return;
                }
            }
            else
            {
                ts.TraceEvent(TraceEventType.Error, 0, "Argument 'userstocreate' not provided");
                return;
            }

            Gen.Users(number,
                _arguments["oupath"],
                getStringArgValueorDefault("password", "kvs"),
                null, null);
        }

        static HashSet<string> createArgumentNames()
        {

            // add/remove arguments to look for
           
            var arguments = new HashSet<string>();
            arguments.Add("mode");
            arguments.Add("itemspermailbox");
            arguments.Add("userstocreate");
            arguments.Add("oupath");
            arguments.Add("percentchanceofattachments");
            arguments.Add("percentchanceofextrarecips");
            arguments.Add("threads");
            arguments.Add("password");
            arguments.Add("maxattachments");
            arguments.Add("maxadditionalrecips");
            arguments.Add("debug");

            return arguments;


        }

        static HashSet<string> _argumentNames = createArgumentNames();
        static Dictionary<string, string> _arguments = new Dictionary<string, string>();

        static TraceSource ts = null;

        class commandLineArgs
        {
            public GenMailArgs mailArgs = new GenMailArgs();
            public GenMode mode;
        }
    }


    enum GenMode
    {
        Users, Mail
    }
}


