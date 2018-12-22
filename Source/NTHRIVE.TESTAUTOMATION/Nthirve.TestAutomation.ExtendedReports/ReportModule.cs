using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using Nthrive.TestAutomation.BussinessObjects;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;


namespace Nthirve.TestAutomation.ExtendedReports
{
    public static class ReportModule
    {
        public static ExtentReports extent;
        public static String HtmlReportPath = String.Empty;
        public static ExtentReports GenerateReports()
        {
            ExtentHtmlReporter htmlReporter = new ExtentHtmlReporter(HtmlReportPath);
            htmlReporter.Configuration().Theme = AventStack.ExtentReports.Reporter.Configuration.Theme.Dark;
            ExtentReports extent = new ExtentReports();
            extent.AddSystemInfo("Invoked by user: ", Environment.UserName);
            extent.AddSystemInfo(" Domain Name: " + Environment.UserDomainName, " Machine Name: " + Environment.MachineName);
            extent.AttachReporter(htmlReporter);
            return extent;
        }
        public static void ExtentReportTestPass(ExtentReports extent, string TestCaseName, string TestMethodName)
        {
            var test = extent.CreateTest(TestCaseName, TestMethodName);
            test.Log(Status.Pass, "Detail is: Test case is Passed");
            extent.Flush();
        }
        public static void ExtentReportTestFail(ExtentReports extent, string capturedimagepath, string TestCaseName, string TestMethodName)
        {
            var test = extent.CreateTest(TestCaseName, TestMethodName);
            test.Log(Status.Fail, "Detail is: Test case is failed");
            test.Fail("details", MediaEntityBuilder.CreateScreenCaptureFromPath(capturedimagepath).Build());
            extent.Flush();
        }
        public static DataTable CreateReports(List<TestCase> Tcases)
        {
            DataTable Result = new DataTable(typeof(TestCase).Name);
            //Get all the properties
            PropertyInfo[] Props = typeof(TestCase).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {
                //Defining type of data column gives proper data table 
                var type = (prop.PropertyType.IsGenericType && prop.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>) ? Nullable.GetUnderlyingType(prop.PropertyType) : prop.PropertyType);
                //Setting column names as Property names
                Result.Columns.Add(prop.Name, type);
            }
            foreach (TestCase tc in Tcases)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                    //inserting property values to datatable rows
                    values[i] = Props[i].GetValue(tc, null);
                }
                Result.Rows.Add(values);
            }
            //put a breakpoint here and check datatable
            return Result;
        }
    }
}

