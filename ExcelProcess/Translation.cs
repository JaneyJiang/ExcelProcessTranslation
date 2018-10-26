using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelProcess
{
    class Translation
    {
        public static string DoTranslation(Dictionary<string, string> dict, string sent)
        {
            foreach (string key in dict.Keys)
            {
                if (sent.Contains(key))
                    sent = sent.Replace(key, dict[key]);
            }
            return sent;
        }

        public static string DoTranslation(string sent, string lang)
        {
            return sent;
        }
        
    }
}
