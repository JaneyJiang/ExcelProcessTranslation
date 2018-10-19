using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace FuncLibrary
{
    public class StringTool
    {
        //去掉字符串中的空白字符：匹配符\s(匹配任何空白字符，包括空格，制表符，换页符等，与[\f\n\t\r\v]等效)
        //如果只是去掉首位空白用input.Trim(),去掉首尾和句中的空格用input.Replace(" ","");
        public static string blankomit(string input)
        {
            string output = Regex.Replace(input, @"\s", "");
            return output;
        }

        //去除字符串里的回车，换行，标签，以及纯标点符号。
        public static string preprocess(string input)
        {
            return puncProcess(enterOmit(tagOmit(input))).Trim();
        }


        //去除回车，换行。
        public static string enterOmit(string input)
        {
            return input.Replace("\n", "").Replace("\r", "").Trim();
        }
        //去除包含在[]内的标签。
        public static string tagOmit(string input)
        {
            return Regex.Replace(input, @"\[.*?\]", "").Trim();
        }

        //去除全是标点的句子。
        public static string puncProcess(string input)
        {
            int puncCount = 0;
            foreach (char c in input)
            {
                if (Char.IsPunctuation(c))
                {
                    puncCount++;
                }
                else
                {
                    return input;
                }

            }

            return "";
        }
        public static void MinFirst(string a, string b, out string min, out string max)
        {
            if (a.Length <= b.Length)
            {
                min = a;
                max = b;
            }
            else
            {
                min = b;
                max = a;
            }
        }
    }
}
