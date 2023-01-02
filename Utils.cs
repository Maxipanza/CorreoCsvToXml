using Microsoft.VisualBasic.FileIO;
using System.Globalization;
using System.Text.RegularExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mime;
using System.Globalization;
using System.Text.RegularExpressions;
using Microsoft.VisualBasic.FileIO;

namespace CorreoCsvToXml
{
    internal class Utils
    {
        public static string Capitalize(string s)
        {
            Regex regex = new Regex(@"[a-zA-Z]");
            if (!(regex.IsMatch(s)))
            {
                return s;
            }
            s = s.ToLower();
            CultureInfo cultureInfo = Thread.CurrentThread.CurrentCulture;
            TextInfo textInfo = cultureInfo.TextInfo;
            return textInfo.ToTitleCase(s);
        }
        public static TextFieldParser ReadCSVFromDirectory(string filename)
        {
            var csvFile = @$"{GetProjectDirectory()}\{filename}";
            //https://stackoverflow.com/questions/6542996/how-to-split-csv-whose-columns-may-contain-comma
            return new TextFieldParser(csvFile);
        }
        public static string GetProjectDirectory()
        {
            // This will get the current WORKING directory (i.e. \bin\Debug)
            string workingDirectory = Environment.CurrentDirectory;
            // This will get the current PROJECT directory
            return Directory.GetParent(workingDirectory).Parent.Parent.FullName;
        }
    }
}
