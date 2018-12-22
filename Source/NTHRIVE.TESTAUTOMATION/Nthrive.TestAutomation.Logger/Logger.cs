using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Nthrive.TestAutomation.Logging
{
    public static class Logger
    {
       private static log4net.ILog log = log4net.LogManager.GetLogger(typeof(Logger));
        public static void Testing()
        {
            
            try
            {
                string str = String.Empty;
                string subStr = str.Substring(0, 4); //this line will create error as the string "str" is empty.  
            }
            catch (Exception ex)
            {
                log.Error("Error Message: " + ex.Message.ToString(), ex);
            }
        }

        public static void HandleException(Exception ex)
        {
           // log4net.Config.BasicConfigurator.Configure();
            log.Error(ex.Message);
        }
    }
}
