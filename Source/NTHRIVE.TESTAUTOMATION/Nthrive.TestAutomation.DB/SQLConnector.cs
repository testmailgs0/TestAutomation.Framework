using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Nthrive.TestAutomation.DB
{
    public static class SQLConnector
    {
        private static SqlConnection _conn;
        private static SqlCommand _sqlcmd;
        public static DataTable _reconReports;
        public static DataTable GetDataFromDB(String Query,String DBName,String SourceType)
        {
            string connetionString = null;
            DataTable DBOutput = new DataTable("DB_Data");
            if (SourceType.Equals("Source"))
            {
               connetionString = "Data Source='analyticsproddbs01.database.windows.net';Initial Catalog='" + DBName + "';User ID='aasuser';Password='AuS!9OqRYP8'";            
            }
            else if (SourceType.Equals("Target"))
            {
                connetionString = "Data Source = 'RCS12CPMDXADB07\\RCSApp07'; Initial Catalog = 'SUMMIT_W'; Integrated Security = True;";
            }
            _conn = new SqlConnection(connetionString);
            try
            {
                _conn.Open();
                DBOutput=GetDataFromQuery(Query);
                _conn.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Connection Failed");
                throw ex;
            }
            return DBOutput;
        }
        private static DataTable GetDataFromQuery(String Query)
        {
            try
            {
                _reconReports = new DataTable();
                _sqlcmd = new SqlCommand(Query, _conn);
                _sqlcmd.CommandTimeout = 120;
                SqlDataAdapter mda = new SqlDataAdapter(_sqlcmd);
                mda.Fill(_reconReports);
            }
            catch (Exception)
            {
                throw;
            }
            return _reconReports;
        }

    }
}
