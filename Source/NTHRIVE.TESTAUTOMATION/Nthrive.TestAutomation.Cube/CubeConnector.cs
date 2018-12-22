using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AnalysisServices.AdomdClient;
using Microsoft.AnalysisServices.SPClient;

namespace Nthrive.TestAutomation.Cube
{
    public static  class CubeConnector
    {
        public static DataTable GetDataFromCube(String Querry)
        {
            DataSet LoadedData = new DataSet();
            DataTable LoadedRecords = new DataTable();
            try
            {
                using (AdomdConnection conn = new AdomdConnection("Data Source=asazure://southcentralus.asazure.windows.net/analyticsdevas01;Provider=MSOLAP;Initial Catalog=cb_000002_Claims; MDX Compatibility=1;User ID=karan.prakash@nthrive.com"))
                //String ConnectionString = System.Configuration.ConfigurationSettings.AppSettings["ServerName"] + ";" + System.Configuration.ConfigurationSettings.AppSettings["Provider"]+ ";" + System.Configuration.ConfigurationSettings.AppSettings["Database"] + System.Configuration.ConfigurationSettings.AppSettings["DBName"] + ";" + System.Configuration.ConfigurationSettings.AppSettings["MDX_Compatibility"] + System.Configuration.ConfigurationSettings.AppSettings["CompatibilityType"] + ";" + System.Configuration.ConfigurationSettings.AppSettings["CredentialType"] + System.Configuration.ConfigurationSettings.AppSettings["UserName"];
                //using (AdomdConnection conn = new AdomdConnection(ConnectionString))
                {
                    //Provider=MSOLAP
                    conn.Open();
                    var mdxQuery = new StringBuilder();
                    mdxQuery.Append(Querry);
                    using (AdomdCommand cmd = new AdomdCommand(mdxQuery.ToString(), conn))
                    {
                        LoadedData.EnforceConstraints = false;
                        LoadedData.Tables.Add(LoadedRecords);
                        LoadedRecords.Load(cmd.ExecuteReader());
                        CleanUpMeasures(LoadedRecords);
                        CleanUpDimension(LoadedRecords);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return LoadedRecords;
        }

        private static DataTable CleanUpMeasures(DataTable RawData)
        {
            foreach (DataColumn dc in RawData.Columns)
            {
                dc.ColumnName = dc.ColumnName.Replace("[Measures].", " ");
                dc.ColumnName = dc.ColumnName.Replace("[", " ");
                dc.ColumnName = dc.ColumnName.Replace("]", " ");
                dc.ColumnName.Trim();
            }
            return RawData;
        }

        private static DataTable CleanUpDimension(DataTable RawData)
        {
            foreach (DataColumn dc in RawData.Columns)
            {
                if(dc.ColumnName.Contains("MEMBER_CAPTION"))
                {
                   List<String> AllValues = dc.ColumnName.Split('.').ToList();
                    dc.ColumnName = AllValues[2].ToString();
                    dc.ColumnName.Trim();
                }
            }
            return RawData;
        }
    }
}
