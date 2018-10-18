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
            else {
                min = b;
                max = a;
            }
        }
    }
   
    
    public class Levenshtein
    {
        //算法原理参考http://blog.jobbole.com/75496/
        /*
            步骤
                说明
                1
                设置n为字符串s的长度。(“GUMBO”)
                设置m为字符串t的长度。(“GAMBOL”)
                如果n等于0，返回m并退出。
                如果m等于0，返回n并退出。
                构造两个向量v0[m+1] 和v1[m+1]，串联0..m之间所有的元素。
                2
                初始化 v0 to 0..m。字符从1开始，初始化的值是从空字符串到生成字符需要的编辑距离。同理每次v1[0]初始化的值也是符合这个原理的，只是
                3
                检查 s (i from 1 to n) 中的每个字符。
                4
                检查 t (j from 1 to m) 中的每个字符
                5
                如果 s[i] 等于 t[j]，则编辑代价为 0；
                如果 s[i] 不等于 t[j]，则编辑代价为1。
                6
                设置单元v1[j]为下面的最小值之一：
                a、紧邻该单元上方+1：v1[j-1] + 1
                b、紧邻该单元左侧+1：v0[j] + 1
                c、该单元对角线上方和左侧+cost：v0[j-1] + cost
                7
                在完成迭代 (3, 4, 5, 6) 之后，v1[m]便是编辑距离的值。
         */
        //这是求编辑距离，这里用两个列向量替代了原来的矩阵，并以短的长度作为列向量的大小。
        public static int distance(string src, string tgt)
        {
            src = StringTool.blankomit(src);
            tgt = StringTool.blankomit(tgt);
            int row = 0;
            int col = 0;
            //我们把字符串调整为长字符串在上端比较，短字符放侧边比较。这样可以较小开辟空间的大小。
            if (src.Length < tgt.Length)
            {
                string tmp = src;
                src = tgt;
                tgt = tmp;
            }
            row = tgt.Length;
            col = src.Length;           
            if(col == 0)
                return row;
            if (row == 0)
                return col;
            int[] v0 = new int[row + 1];
            int[] v1 = new int[row + 1]; 
            for (int i = 0; i <= row; i++)
                v0[i]=i;
            int cost = 0;
            for (int i = 1; i <= col; i++)
            { 
                v1[0] = i;
                for (int j = 1; j <= row; j++)
                {
                    if (src[i-1].Equals(tgt[j-1]))
                        cost = 0;
                    else
                        cost = 1;
                    //这里三个分别对应，插入，删除和替换。
                    v1[j] = Math.Min(Math.Min(v1[j - 1] + 1, v0[j] + 1), v0[j - 1] + cost);
                }
                v0 = (int[])v1.Clone();
            }
            return v1[row];
        }

        //这是求最长公共子串的长度。
        public static int LCS(string src, string tgt)
        {
            int row = 0;
            int col = 0;
            //我们把字符串调整为长字符串在上端比较，短字符放侧边比较。这样可以较小开辟空间的大小。
            if (src.Length < tgt.Length)
            {
                string tmp = src;
                src = tgt;
                tgt = tmp;
            }
            row = tgt.Length;
            col = src.Length;

            if (row == 0 || col == 0)
                return 0;
            int[] v0 = new int[row+1];
            int[] v1 = new int[row+1];
            for (int i = 0; i <= row; i++)
                v0[i] = 0;
            for (int i = 1; i <= col; i++)
            {
                v1[0] = 0;
                for (int j = 1; j <= row; j++)
                {
                    if (src[i - 1].Equals(tgt[j - 1]))
                        v1[j] = v0[j - 1] + 1;
                    else
                        v1[j] = Math.Max(Math.Max(v1[j-1],v0[j-1]),v0[j]);
                }
                v0 = (int[])v1.Clone();
            }
            return v1[row];
        }
        public static float Similarity(string src, string tgt)
        {
            int ld = distance(src, tgt);
            int lcs = LCS(src, tgt);
            return ((float)lcs) / (ld + lcs);

        }
        //两个字符串前缀匹配的长度
        public static int PrefixMatch(string src, string tgt)
        {
            if (src.Length == 0 || tgt.Length == 0)
                return 0;
            int n = Math.Min(src.Length, tgt.Length);
            int count = 0;
            while (count < n)
            {
                if (src[count].Equals(tgt[count]))
                    count += 1;
            }
            return count;

        }
        //Jaro-Winkler 计算公式dw=dj+L∗P(1−dJ)dw=dj+L∗P(1−dJ)


      /*  public static float jaro_winkler(string src, string tgt)
        {
            int ld = distance(src, tgt);
            int lcs = LCS(src, tgt);
            float d = (1 / 3) * ((float)lcs / src.Length + (float)lcs / tgt.Length + (float)(lcs - ld) / lcs);
            return d;
        }
        */

   
        private static readonly double mWeightThreshold = 0.7;

        /* Size of the prefix to be concidered by the Winkler modification. 
         * Winkler's paper used a default value of 4
         */
        private static readonly int mNumChars = 4;//我觉得可以根据句子的长度来设置这个值。
        public static double Jaro_Winkler(string aString1, string aString2)
        {

            aString1 = StringTool.blankomit(aString1);
            aString2 = StringTool.blankomit(aString2);

            string minStr=null;
            string maxStr=null;
            StringTool.MinFirst(aString1, aString2, out minStr, out maxStr);

            int lLen1 = minStr.Length;
            int lLen2 = maxStr.Length;
            if (lLen1 == 0)
                return lLen2 == 0 ? 1.0 : 0.0;//两个字符串长度都为0的时候判断为完全相似1，否则只有一个为0，则判断为完全不相似0


            //这一段是求匹配字符个数的，其中这里有个窗口大小来求匹配字符个数，超过窗口大小的，就算存在相同的字符也不计数。
            int lSearchRange = Math.Max(0, Math.Max(lLen1, lLen2) / 2 - 1);

            // default initialized to false
            bool[] lMatched1 = new bool[lLen1];
            bool[] lMatched2 = new bool[lLen2];

            
            int lNumCommon = 0;
            //对第一个字符依次后移，匹配另一个字符的字符段里面的字符，如果之前匹配成功过，则跳过，直到字符段窗口结束或者遇见相同字符。
            for (int i = 0; i < lLen1; ++i)//因为我调整过字符串的长度，所以len1是短，len2是长。
            {
                int lStart = Math.Max(0, i - lSearchRange);
                int lEnd = Math.Min(i + lSearchRange + 1, lLen2);
                for (int j = lStart; j < lEnd; ++j)
                {
                    if (lMatched2[j]) continue;
                    if (minStr[i] != maxStr[j])
                        continue;
                    lMatched1[i] = true;
                    lMatched2[j] = true;
                    ++lNumCommon;
                    break;
                }
            }
            //


            if (lNumCommon == 0)
            {
                return 0.0;//在窗口范围内的匹配字符为零则返回相似度为0；
            }


            //求换位的数目，即匹配字符是否是按相同顺序，如果顺序相同，则换位数目为0，如果顺序不同则需要计算换位数目。
            //先位移到第一个字符串的第一个匹配字符哪里，对第二个字符串进行匹配，不匹配就后移。这时候，两个标签都在匹配字符那里，这时候，我们比较这两个匹配字符是否相同
            //如果不等则增加换位数。然后再对第一个字符串的第二个匹配字符进行匹配。
            //因为换位是字符对，所以要把结果除以2.
            //即若在字符串的第i位出现了a,b，在第j位又出现了b,a，则表示两者出现了换位。
            int lNumHalfTransposed = 0;
            int k = 0;
            for (int i = 0; i < lLen1; ++i)
            {
                if (!lMatched1[i]) continue;
                while (!lMatched2[k]) ++k;
                if (minStr[i] != maxStr[k])
                    ++lNumHalfTransposed;
                ++k;
            }
            // System.Diagnostics.Debug.WriteLine("numHalfTransposed=" + numHalfTransposed);
            int lNumTransposed = lNumHalfTransposed / 2;

            // System.Diagnostics.Debug.WriteLine("numCommon=" + numCommon + " numTransposed=" + numTransposed);
            double lNumCommonD = lNumCommon;
            double lWeight = (lNumCommonD / lLen1
                             + lNumCommonD / lLen2
                             + (lNumCommon - lNumTransposed) / lNumCommonD) / 3.0;

            if (lWeight <= mWeightThreshold)
            {
                return lWeight;
            }

            //这是求最大前缀匹配长度
            int lMax = Math.Min(mNumChars, Math.Min(minStr.Length, maxStr.Length));
            int lPos = 0;
            while (lPos < lMax && minStr[lPos] == maxStr[lPos])
                ++lPos;
            //

            if (lPos == 0) return lWeight;
            return lWeight + 0.1 * lPos * (1.0 - lWeight);

        }



    }
}
