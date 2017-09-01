using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Microsys.EAI.Framework.PasswordManager
{
    public static class LogHelper
    {

        public enum EntryType
        {
            Information,
            Error
        }

        public static void Write(EntryType entryType, string message, params object[] paramList)
        {
            message = string.Format(message, paramList);

            Write(entryType, message);
        }

        public static void Write(EntryType entryType, string message)
        {

            string filePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Microsys.EAI.Framework.PasswordManager.log");

            using (StreamWriter streamWriter = File.AppendText(filePath))
            {
                streamWriter.WriteLine("{0} - {1} - {2}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss:fff"), entryType, message);
            }
        }

    }
}
