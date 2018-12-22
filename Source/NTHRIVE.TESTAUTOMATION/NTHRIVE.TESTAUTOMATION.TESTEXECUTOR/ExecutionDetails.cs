using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.IO;
using System.Data;
using System.Windows.Forms;
using System.Collections;
using System.ComponentModel;
using Nthrive.TestAutomation.TestDataProvider;
using Nthrive.TestAutomation.BussinessObjects;
using Nthrive.TestAutomation.Util;
using Nthrive.TestAutomation.DB;
using Nthrive.TestAutomation.Cube;
using System.Configuration;
using System.Collections.Specialized;
using Nthirve.TestAutomation.ExtendedReports;
using Nthrive.TestAutomation.Logging;


namespace Nthrive.TestAutomation.TestExecutor
{
    public class ExecutionDetails : BindableBase
    {
        #region Commands
        private ICommand _runtest;
        private ICommand _exit;
        private ICommand _BrowseTcase;
        private RelayCommand _loadTCases;
        private DataTable _loadData;
        private DataTable ResultSet = new DataTable();
        #endregion
        #region Properties
        public ICommand Run
        {
            get
            {
                if (_runtest == null)
                {
                    _runtest = new RelayCommand(
                        param => this.RunTest()
                    );
                }
                return _runtest;
            }
        }
        public ICommand ExitApp
        {
            get
            {
                if (_exit == null)
                {
                    _exit = new RelayCommand(
                        param => this.Exit()
                    );
                }
                return _exit;
            }
        }
        public ICommand BrowseTCase
        {
            get
            {
                if (_BrowseTcase == null)
                {
                    _BrowseTcase = new RelayCommand(
                        param => this.BrowseCases()
                    );
                }
                return _BrowseTcase;
            }
        }
        public String TestCasesPath
        {
            get { return ApplicationArguments.TestCasePath; }
            set {
                SetProperty(ref ApplicationArguments.TestCasePath, value); 
                OnPropertyChanged("TestCasesPath"); 
            }
        }
        public DataTable LoadedCases
        {
            get{ return this._loadData;}
            set{
                if (value != this._loadData)
                {
                    this._loadData = value;
                    SetProperty(ref _loadData, value);
                    OnPropertyChanged("LoadedCases"); 
                }
            }
        }
        public RelayCommand LoadedTCases
        {
            get
            {
                if (_loadTCases == null)
                {
                    _loadTCases = new RelayCommand(
                        param => this.LoadCases()
                    );
                }
                return _loadTCases;
            }
        }
        public bool DB_Validation
        {
            get { return ApplicationArguments.Is_DB_Validation; }
            set
            {
                SetProperty(ref ApplicationArguments.Is_DB_Validation, value);
                OnPropertyChanged("DB_Validation");
            }
        }
        public bool DBToDB_Validation
        {
            get { return ApplicationArguments.Is_DB_To_DBValidation; }
            set
            {
                SetProperty(ref ApplicationArguments.Is_DB_To_DBValidation, value);
                OnPropertyChanged("DBToDB_Validation");
            }
        }
        public bool CubeToDB_Validation
        {
            get { return ApplicationArguments.Is_DB_To_CubeValidation; }
            set
            {
                SetProperty(ref ApplicationArguments.Is_DB_To_CubeValidation, value);
                OnPropertyChanged("CubeToDB_Validation");
            }
        }
        public bool UI_Validation
        {
            get { return ApplicationArguments.Is_UIValidation; }
            set
            {
                SetProperty(ref ApplicationArguments.Is_UIValidation, value);
                OnPropertyChanged("UI_Validation");
            }
        }
        #endregion
        #region  Methods
        public void RunTest()
        {
            List<String> Errors = new List<String>();
            try
            {
                if (ApplicationArguments.Is_DB_Validation)
                {
                    ReportModule.HtmlReportPath = "Reports\\TestReport_DB_Validation.html";
                    ReportModule.extent = ReportModule.GenerateReports();
                    List<TestCase> TCases = new CreateCases().TCases(ApplicationArguments.TestCasePath);
                    foreach (TestCase tc in TCases)
                    {
                        if (tc.Category == "Acceptance")
                        {
                            DataTable Test = SQLConnector.GetDataFromDB(tc.SQL_Query, ConfigurationManager.AppSettings["DBName"].ToString(), "Source");
                            if (Test.Rows.Count > 0)
                            {
                                Errors.Add("Test Case" + tc.TestCaseID + "got Failed");
                            }

                        }
                        else if (tc.Category.Equals("Regression"))
                        {
                            Errors = DataComparer.GetMismatchedData(new OpenXmlDataProvider().GetTestData("Test_Cases\\file1.xlsx", 1, 1).Tables[0], new OpenXmlDataProvider().GetTestData("Test_Cases\\file2.xlsx", 1, 1).Tables[0]);
                        }
                        if (Errors.Count > 0)
                        {
                            tc.Result = "Fail";
                            ReportModule.ExtentReportTestFail(ReportModule.extent, ReportModule.HtmlReportPath, tc.TestCaseID, tc.Category);
                        }
                        else if (Errors.Count <= 0)
                        {
                            ReportModule.ExtentReportTestPass(ReportModule.extent, tc.TestCaseID, tc.Category);
                        }
                    }
                    Errors.Clear();
                    ResultSet = ReportModule.CreateReports(TCases);
                    RemoveUnwantedColumns(ResultSet);
                    LoadedCases.Clear();
                    LoadCases();
                    MailSender.SendReports(); 
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("Fatal Error Occured: Check Logs For Details");
                Logger.HandleException(ex);
            }
            finally
            {
              
            }
        }
        private void RemoveUnwantedColumns(DataTable ResultSet)
        {
            try
            {
                if (ResultSet.Columns.Contains("SQL_Query"))
                {
                    ResultSet.Columns.Remove("SQL_Query");
                }
                if (ResultSet.Columns.Contains("MDX_Query"))
                {
                    ResultSet.Columns.Remove("MDX_Query");
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        public void Exit()
        {
            try
            {
                System.Windows.Application.Current.Shutdown();
            }
            catch (Exception ex )
            {
                Logger.HandleException(ex);
            }
        }
        public void BrowseCases()
        {
            try
            {
                using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
                {
                    dialog.ShowDialog();
                    TestCasesPath = dialog.SelectedPath+"\\Test Data.xlsx";
                }
            }
            catch (Exception ex)
            {
                Logger.HandleException(ex);
            }

        }
        public void LoadCases()
        {
            try
            {
                if (ResultSet.Rows.Count>0)
                {
                    LoadedCases = ResultSet;
                }
                else
                {
                    DataSet grdData = new OpenXmlDataProvider().GetTestData(TestCasesPath, 1, 1);
                    LoadedCases = grdData.Tables[0];
                }
                
            }
            catch (Exception ex)
            {
                Logger.HandleException(ex);
            }
        }
        #endregion
    }
}
