using DataTable = System.Data.DataTable;
using Microsoft.Office.Interop.Excel;
//using TestAutomation.Interfaces;
using System;
using System.Data;
using System.Data.OleDb;
using System.Runtime.InteropServices;

namespace TestAutomation.TestDataProvider
{
    public sealed class ExcelDataProvider
    {
        #region Members

        /// <summary>
        /// The _excel object
        /// </summary>
        private string _excelObject = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";

        /// <summary>
        /// The _filepath
        /// </summary>
        private string _filepath = "File\\Data.xlsx";

        /// <summary>
        /// The _con
        /// </summary>
        private OleDbConnection _con = null;

        #endregion Members

        #region Properties

        /// <summary>
        /// Gets the connection.
        /// </summary>
        /// <value>
        /// The connection.
        /// </value>
        private OleDbConnection Connection
        {
            get
            {
                if (_con == null)
                {
                    OleDbConnection con = new OleDbConnection { ConnectionString = ConnectionString };
                    _con = con;
                }
                return _con;
            }
        }

        /// <summary>
        /// Gets the connection string.
        /// </summary>
        /// <value>
        /// The connection string.
        /// </value>
        private string ConnectionString
        {
            get
            {
                if (_filepath != string.Empty)
                {
                    return string.Format(_excelObject, _filepath);
                }
                else
                {
                    return string.Empty;
                }
            }
        }

        #endregion Properties

        #region Methods

        /// <summary>
        /// Gets the test data.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        /// <param name="rowHeaderIndex">Index of the row header.</param>
        /// <param name="startColumnFrom">The start column from.</param>
        /// <returns></returns>
        public DataSet GetTestData(string filePath, int rowHeaderIndex = 1, int startColumnFrom = 1)
        {
            _filepath = filePath;
            OnConnectionStringChanged();
            return ExcelAsDataSet();
        }

        /// <summary>
        /// Excels as data set.
        /// </summary>
        /// <returns></returns>
        private DataSet ExcelAsDataSet()
        {
            OpenWithInterop();
            DataTable mySheets = GetSchema();
            DataSet ds = new DataSet();

            for (int i = 0; i < mySheets.Rows.Count; i++)
            {
                string sheetName = mySheets.Rows[i]["TABLE_NAME"].ToString();
                var dt = ReadTableFromSheetName(sheetName);
                if (dt != null)
                {
                    dt.TableName = sheetName.Substring(0, sheetName.Length - 1);
                    ds.Tables.Add(dt);
                }
            }
            return ds;
        }

        /// <summary>
        /// Gets the schema.
        /// </summary>
        /// <returns></returns>
        private DataTable GetSchema()
        {
            DataTable dtSchema = null;
            if (Connection.State != ConnectionState.Open) Connection.Open();
            dtSchema = Connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            return dtSchema;
        }

        /// <summary>
        /// Called when [connection string changed].
        /// </summary>
        private void OnConnectionStringChanged()
        {
            if (Connection != null && !Connection.ConnectionString.Equals(ConnectionString))
            {
                if (Connection.State == ConnectionState.Open)
                    Connection.Close();
                Connection.Dispose();
                _con = null;

            }
        }

        /// <summary>
        /// Opens the with interop.
        /// </summary>
        private void OpenWithInterop()
        {
            var xlApp = new Application();
            var xlWorkBook = xlApp.Workbooks.Open(_filepath);
            xlApp.CalculateFull();

            xlWorkBook.Close(true);
            xlApp.Quit();

            ReleaseObject(xlWorkBook);
            ReleaseObject(xlApp);
        }

        /// <summary>
        /// Reads the name of the table from sheet.
        /// </summary>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <returns></returns>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2100:Review SQL queries for security vulnerabilities")]
        private DataTable ReadTableFromSheetName(string sheetName)
        {
            if (Connection.State != ConnectionState.Open)
            {
                Connection.Open();
            }
            string cmdText = "Select * from [{0}]";

            OleDbCommand cmd = new OleDbCommand(string.Format(cmdText, sheetName)) { Connection = Connection };
            OleDbDataAdapter adpt = new OleDbDataAdapter(cmd);

            DataTable returnTable = new DataTable();

            adpt.Fill(returnTable);
            return returnTable;
        }

        /// <summary>
        /// Releases the object.
        /// </summary>
        /// <param name="obj">The object.</param>
        private void ReleaseObject(object obj)
        {
            try
            {
                Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }

        #endregion Methods
    }
}
