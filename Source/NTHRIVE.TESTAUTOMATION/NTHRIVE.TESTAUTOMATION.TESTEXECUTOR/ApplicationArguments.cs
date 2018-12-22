using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Nthrive.TestAutomation.BussinessObjects;

namespace Nthrive.TestAutomation.TestExecutor
{
    public static class ApplicationArguments
    {
        public static String TestCasePath = string.Empty;
        public static List<TestCase> TestCasesToBeRun { get; set; } 
        public static String ReleasePath=AppDomain.CurrentDomain.BaseDirectory;
        public static bool Is_DB_Validation = false;
        public static bool Is_DB_To_DBValidation = false;
        public static bool Is_DB_To_CubeValidation = false;
        public static bool Is_UIValidation = false;
        
    }
}
