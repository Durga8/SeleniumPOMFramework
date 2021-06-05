using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;

namespace SeleniumFrameWork.Helpers
{
    public class WordHelpers
    {
        public string GetTextFromWord(string filename)
        {
            string userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name.Split('\\').Last();
            Logger.log(userName);
            string downloadpath = $@"C:\Users\{userName}\Downloads\";
            string fullfilepath = downloadpath + filename;
            Logger.log(downloadpath);
            //@"C:\Users\Administrator\Downloads\filename"
            StringBuilder text = new StringBuilder();
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            object miss = System.Reflection.Missing.Value;
            object path = fullfilepath;
            object readOnly = true;
            Microsoft.Office.Interop.Word.Document docs = word.Documents.Open(ref path, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);

            for (int i = 0; i < docs.Paragraphs.Count; i++)
            {
                //text.Append(" \r\n " + docs.Paragraphs[i + 1].Range.Text.ToString());
                text.Append(docs.Paragraphs[i + 1].Range.Text.ToString().Trim());
            }

            Logger.log(text.ToString());

            if (Regex.IsMatch(text.ToString(), "^[a-zA-Z0-9]*$"))
            {
                Logger.log("Alphanumeric String");
            }
            else
            {
                Logger.log("Non-Alphanumeric String");
            }

            return text.ToString();

        }

    }
}
