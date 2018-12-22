using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Nthrive.TestAutomation.TestDataProvider;

namespace Nthrive.TestAutomation.BussinessObjects
{
    public class TestCase
    {
        public string TestCaseID { get; set; }
        public string Description { get; set; }
        public string Category { get; set; }
        public string SQL_Query { get; set; }
        public string MDX_Query { get; set; }
        public string Result { get; set; }
    }

    public class CreateCases
    {
        public List<TestCase> TCases(String FilePath)
        {
            List<TestCase> Cases = new List<TestCase>();
            try
            {
                DataSet _metaData = new OpenXmlDataProvider().GetTestData(FilePath, 1, 1);
                foreach (DataTable dt in _metaData.Tables)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        TestCase testCase = new TestCase();
                        {
                            testCase.TestCaseID = dr["TestCaseID"].ToString();
                            testCase.Description = dr["Description"].ToString();
                            testCase.Category = dr["Category"].ToString();
                            testCase.SQL_Query = dr["SQL_Query"].ToString();
                            testCase.MDX_Query = dr["MDX_Query"].ToString();
                            testCase.Result = "Pass";
                            if (Cases.Count < 1)
                            {
                                Cases.Add(testCase);
                            }
                            else
                            {
                                if (CheckDuplicateEntry(testCase.TestCaseID, Cases) == false)
                                {
                                    Cases.Add(testCase);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                throw ;
            }
            return Cases;
        }


        private bool CheckDuplicateEntry(string p, List<TestCase> Cases)
        {
            foreach (var obj in Cases)
            {
                if (p.Equals(obj.TestCaseID))
                {
                    return true;
                }
            }
            return false;
        }

    }
}
