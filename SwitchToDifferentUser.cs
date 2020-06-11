using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumFrameWork.Base
{
    public class SwitchToDifferentUser
    {
        //Running the incognito Window ...Opening the browser in IE through Batch file and can execute rest of the operations through selenium
        public SecureString MakeSecureString(string text)
        {
            SecureString secure = new SecureString();
            foreach (char c in text)
            {
                secure.AppendChar(c);
            }

            return secure;
        }

        public void RunAs(string path, string username, string password)
        {
            ProcessStartInfo myProcess = new ProcessStartInfo(path);
            myProcess.WorkingDirectory = @"C:\Program Files\internet explorer";
            myProcess.UserName = username;
            myProcess.Password = MakeSecureString(password);

            myProcess.Domain = "MCHP-MAIN";
            myProcess.LoadUserProfile = true;
            myProcess.UseShellExecute = false;
            Process.Start(myProcess);


        }

        public void switchDifferentUser()
        {
            SwitchToDifferentUser run = new SwitchToDifferentUser();

            //run.RunAs(@"C:\Program Files\internet explorer\iexplore.exe", "X17052", "Microchip3#");
            run.RunAs(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe", "X17052", "Microchip3#");
        }



    }
}

