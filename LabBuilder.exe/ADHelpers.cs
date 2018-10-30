using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.DirectoryServices;
using System.Windows.Input;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;

namespace LabBuilder
{
    class DirectoryEntryContext : IDisposable
    {
        static readonly DirectoryEntry _rootde = new DirectoryEntry("LDAP://RootDSE");
        static readonly SortOption so = new SortOption("name", SortDirection.Ascending);
         const string LDAPfilter = "(objectClass=organizationalUnit)";

        readonly DirectoryEntry _currentde;
        readonly DirectorySearcher _ds;
        List<DirectoryEntryContext> _childOUs = new List<DirectoryEntryContext>();

        public DirectoryEntryContext()
        {
            _currentde = new DirectoryEntry();
            _ds = new DirectorySearcher(_currentde);
            _ds.Filter = LDAPfilter;
            _ds.SearchScope = SearchScope.OneLevel;
            _ds.Sort = so;
        }
        public DirectoryEntryContext(DirectoryEntry de)
        {
            _currentde = de;
            _ds = new DirectorySearcher(_currentde);
            _ds.Filter = LDAPfilter;
            _ds.SearchScope = SearchScope.OneLevel;
            _ds.Sort = so;
        }
        public static string DefaultDomain = DefaultNamingContext.ToLower(System.Globalization.CultureInfo.InvariantCulture).Replace("dc=", "").Replace(',', '.');

        public static string DefaultNamingContext
        {
            get
            {
                return _rootde.Properties["defaultNamingContext"].Value.ToString();
            }
        }

        public static string ConfigurationContext
        {
            get
            {
                return _rootde.Properties["configurationNamingContext"].Value.ToString();


            }
        }
        public List<DirectoryEntryContext> ChildOUs
        {
            get
            {
                _childOUs.Clear();
                SearchResultCollection res = _ds.FindAll();
                foreach (SearchResult s in res)
                {
                    _childOUs.Add(new DirectoryEntryContext(s.GetDirectoryEntry()));
                }
                return _childOUs;

            }
        }
        public string ObjectName
        {
            get
            {
                if (_currentde.SchemaClassName == "domainDNS")
                {
                    //it's the root domain so format like a dns domain name.
                    return DefaultDomain;
                }
                else
                {
                    return _currentde.Name.Substring(3);//don't want OU= etc
                }
            }
        }

        public DirectoryEntry DE
        {
            get
            {
                return _currentde;
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);


        }

        protected void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_currentde != null)
                    _currentde.Dispose();
                if (_ds != null)
                    _ds.Dispose();
            }
        }

    }


    class MailUsers
    {
        public static List<GenMailUser> GetAllOUMailUsers(string selOUPath)
        {
            DirectoryEntry de = null;
            DirectorySearcher ds = null;
            var _userList = new List<GenMailUser>(); 
            try
            {
                de = new DirectoryEntry(selOUPath);
                ds = new DirectorySearcher(de);
                ds.Filter = LDAPfilter;
                ds.PageSize = pageSize;
                ds.PropertiesToLoad.Add("mail");
                ds.PropertiesToLoad.Add("homeMDB");
                ds.PropertiesToLoad.Add("extensionAttribute13");
                ds.PropertiesToLoad.Add("displayName");
                ds.PropertiesToLoad.Add("sAMAccountName");

                SearchResultCollection res = ds.FindAll();


                foreach (SearchResult s in res)
                {
                    var tmpuser = new GenMailUser();
                    try
                    {

                        tmpuser.smtpAddress = (string)s.Properties["mail"][0];
                        tmpuser.sAMAccountName = (string)s.Properties["sAMAccountName"][0];
                        if (s.Properties.Contains("extensionAttribute13"))
                        {
                            tmpuser.password = ((string)s.Properties["extensionAttribute13"][0]).TrimStart('{');
                            tmpuser.networkCred = new System.Net.NetworkCredential(tmpuser.sAMAccountName, tmpuser.password, DirectoryEntryContext.DefaultDomain);
                        }
                        tmpuser.displayName = (string)s.Properties["displayName"][0];
                        string exchangeDB = (string)s.Properties["homeMDB"][0];

                        // convert from DN to display name
                        int commapos = exchangeDB.IndexOf(',');
                        exchangeDB = exchangeDB.Substring(3, commapos - 3);

                        // see if DB is in our list
                        tmpuser.exchangeDBIndex = dblist.FindIndex(database =>
                        {
                            return database.DatabaseName == exchangeDB;
                        });

                        // if we don't have DB in the list save name
                        if (tmpuser.exchangeDBIndex == -1)
                        {

                            tmpuser.exchangeDBName = exchangeDB;
                        }

                        _userList.Add(tmpuser);
                    }
                    catch (Exception e)
                    {
                        // swallow but trace
                        if (tmpuser.smtpAddress != null)
                            ts.TraceEvent(TraceEventType.Verbose, 0, "Error adding user {0}", tmpuser.smtpAddress);
                        ts.TraceEvent(TraceEventType.Verbose, 0, e.ToString());
                    }
                }
            }
            finally
            {
                if (de != null)
                    de.Dispose();
                if (ds != null)
                    ds.Dispose();
            }
            return _userList;
        }
        // mail enabled users who are not disabled
        // bitmask via: https://ctogonewild.com/2009/09/03/bitmask-searches-in-ldap/
        //
        const string LDAPfilter = "(&(objectClass=user)(objectCategory=person)(mailNickName=*)(!userAccountControl:1.2.840.113556.1.4.803:=2)) ";
        // pagesize must be 1000 to facilitate returning all records
        const int pageSize = 1000;
        public static List<ExchangeDatabase> dblist = ExchangeDatabases.DBs;
        readonly static TraceSource ts = MainWindow.ts;

    }

    class ExchangeDatabase
    {
        public string DatabaseName { get; set; }
        public string CurrentServer { get; set; }
        public string DatabaseToolTip { get; set; }
        public DirectoryEntry ServerDE { get; set; }
    }
    class ExchangeDatabases
    {
        static List<ExchangeDatabase> _exchangeDBs = new List<ExchangeDatabase>();
        static DirectorySearcher _ds = CreateAndInitializeDS();
       
        static DirectorySearcher CreateAndInitializeDS ()
        {
            var ds = new DirectorySearcher(new DirectoryEntry("LDAP://" + DirectoryEntryContext.ConfigurationContext));
            ds.Filter = "(objectClass=msExchPrivateMDB)";
            ds.PropertiesToLoad.Add("msExchOwningServer");
            ds.PropertiesToLoad.Add("displayName");
            return ds;
        }

        public static List<ExchangeDatabase> DBs
        {
            get
            {
                // refresh list each time called
                _exchangeDBs.Clear();
                SearchResultCollection res = _ds.FindAll();
                foreach (SearchResult s in res)
                {
                    ExchangeDatabase exdb = new ExchangeDatabase();

                    exdb.DatabaseName = (string)s.Properties["displayName"][0];
                    exdb.CurrentServer = (string)s.Properties["msExchOwningServer"][0];
                    exdb.ServerDE = new DirectoryEntry(exdb.CurrentServer);

                    int commapos = exdb.CurrentServer.IndexOf(',');
                    exdb.CurrentServer = exdb.CurrentServer.Substring(3, commapos - 3);
                    exdb.DatabaseToolTip = "Currently on server " + exdb.CurrentServer;

                    _exchangeDBs.Add(exdb);
                }

                return _exchangeDBs;
            }
        }


    }


}
