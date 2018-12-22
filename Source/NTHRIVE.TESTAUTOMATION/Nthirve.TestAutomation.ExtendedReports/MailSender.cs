using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Nthirve.TestAutomation.ExtendedReports
{
   public static class MailSender
    {
        //private static string Styling = "<style>#customers {font-family: 'Trebuchet MS', Arial, Helvetica, sans-serif; font-size:10px; border-collapse: collapse; width: 100%; border:2px;} #customers table, th, td { border: 1px solid black;} #customers td, #customers th { border: 1px solid #ddd; padding: 8px; }#customers tr:nth-child(even){background-color: #f2f2f2;} #customers tr:hover {background-color: #ddd;}#customers th {    padding-top: 12px; padding-bottom: 12px; text-align: center; background-color: #4CAF50; color: white;}</style>";
        public static void SendReports()
        {

            Microsoft.Office.Interop.Outlook.Application oApp = new Microsoft.Office.Interop.Outlook.Application();
            Microsoft.Office.Interop.Outlook.MailItem oMsg = (Microsoft.Office.Interop.Outlook.MailItem)oApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
            //oMsg.HTMLBody = "Hello Team,<br/>" + System.Environment.NewLine + Styling + "<br/>Please review below data validation reports.<br/><br/>" + System.Environment.NewLine + GetDataTableHTML(ReconData) + System.Environment.NewLine + "<br/>Thanks,<br/>Automation System";
            oMsg.HTMLBody = "Automation Testing Report";
            #region Attachment Code
            String FilePath = @"C:\Users\karan.prakash\source\repos\Nthrive.TestAutomation\NTHRIVE.TESTAUTOMATION\NTHRIVE.TESTAUTOMATION.TESTEXECUTOR\BIN\DEBUG\Reports\TestReport.html";
            oMsg.Attachments.Add(FilePath, Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, 1, "TestReport.html");
            #endregion
            oMsg.Subject = "Automated Testing Reports [ARA-Acceptance]";
            #region Recipient Code
            // Add a recipient.
            Microsoft.Office.Interop.Outlook.Recipients oRecips = (Microsoft.Office.Interop.Outlook.Recipients)oMsg.Recipients;
            // Change the recipient in the next line if necessary.
            Microsoft.Office.Interop.Outlook.Recipient oRecip = (Microsoft.Office.Interop.Outlook.Recipient)oRecips.Add("karan.prakash@nthrive.com");
            oRecips.Add(ConfigurationSettings.AppSettings["CCEmails"].ToString());
            oRecip.Resolve();
            #endregion
            oMsg.Send();
            #region CleanUp
            oRecip = null;
            oRecips = null;
            oMsg = null;
            oApp = null;
            #endregion

        }
    }
}
