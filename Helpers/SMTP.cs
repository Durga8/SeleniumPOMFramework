using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using NUnit.Framework;
using NUnit.Framework.Interfaces;
using SeleniumFrameWork.Helpers;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using SeleniumFrameWork.Base;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SeleniumFrameWork.Helpers
{
    public class SMTP:ReportingHelpers
    {
        
        //SMTP Authentication- Whereby an SMTP client may log in using an authentication mechanism chosen among supported by SMTP servers.
        public static void smtpMailConfiguration(string htmlBody,string displayName, string subject, string recipientGroup)
        {
            try
            {
                // Create the Outlook application.
                Outlook.Application oApp = new Outlook.Application();
                // Create a new mail item.
                Outlook.MailItem oMsg = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
                // Set HTMLBody. 
                //add the body of the email
                oMsg.HTMLBody = htmlBody;
                //Add an attachment.
                String sDisplayName = displayName;
                int iPosition = (int)oMsg.Body.Length + 1;
                int iAttachType = (int)Outlook.OlAttachmentType.olByValue;
                string fileName = reportZipath+".zip";
                if (File.Exists(fileName))
                {
                    
                    //now attached the file
                    Outlook.Attachment oAttach = oMsg.Attachments.Add
                                             (fileName, iAttachType, iPosition, sDisplayName);
                }
                else
                {
                    Logger.log("No  HTML Report File Exists");
                }
                
                
                //Subject line
                oMsg.Subject = subject;
                // Add a recipient.
                Outlook.Recipients oRecips = (Outlook.Recipients)oMsg.Recipients;
                // Change the recipient in the next line if necessary.
                Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(recipientGroup);
                oRecip.Resolve();
                // Send.
                oMsg.Send();
                // Clean up.
                oRecip = null;
                oRecips = null;
                oMsg = null;
                oApp = null;
                Logger.log("Mail Sent sucessfully");

            }//end of try block
            catch (Exception ex)
            {
                Logger.log("Error While Sending Email::" + ex.Message);
            }//end of catch 

        }
       

        
    }

   
    
}
