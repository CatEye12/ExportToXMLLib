using ExportToXMLLib.Properties;
using ExportToXMLLib;
using System.Collections.Generic;
using System.Data.Linq;
 

namespace ExportToXMLLib
{
    class DBConnection : DataContext
    {
        public IEnumerable<View_Part> ViewParts { get { return this.GetTable<View_Part>(); } }

        private static DBConnection Instance = null;


        private DBConnection(): base(ConnStr)
        {

        }

        private static string ConnStr { get { return Settings.Default.DBConnectionString; }  }  

        public static DBConnection DBProp
        {
            get
            {
                if (Instance == null) { return Instance = new DBConnection(); }
                return Instance;
            }
        }  

    }
}
