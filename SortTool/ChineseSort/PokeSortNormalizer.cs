using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace ChineseSort
{
    public class PokeSortNormalizer
    {
        private static Dictionary<char, string> normalizeTable = new Dictionary<char, string>()
        {
            // ・複合文字の展開
            // ・記号の削除
            // ・半角用代替文字の全角Unicode変換
            // ・ひらがな→カタカナ
            // の正規化を行う。
        	{'Æ', "AE" },
            {'Œ', "OE" },
            {'⑩', "ER"},
            {'⑪', "RE"},
            {'⑫', "R"},
            {'⒅', "E"},
            {'‘', ""},
            {'\'', ""},
            {'“', ""},
            {'"', ""},
            {'„', ""},
            {'@', ""},
            {'#', ""},
            {'$', ""},
            {'%', ""},
            {'&', ""},
            {'(', ""},
            {')', ""},
            {'*', ""},
            {'+', ""},
            {',', ""},
            {'.', ""},
            {':', ""},
            {';', ""},
            {'=', ""},
            {'¡', ""},
            {'!', ""},
            {'¿', ""},
            {'?', ""},
            {'~', ""},
            {'・', ""},
            //{'♀', ""},
            //{'♂', ""},
            {'°', ""},
            {'⒆', "PK"},
            {'⒇', "MN"},
            {'⑭', "♂"},
            {'⑮', "♀"},
            // スペースとして扱われるもの（スペースを無視せず、単語ごとにソートする）
            {' ', " "},
            {'　', " "}, // 全角スペース
            {'-', " "},
            {'/', " "},
            //日本语的，平假名替换成片假名就没有在这里列出了，如果有需要可以加上
        };
        Regex removeDigitSeparatorRegex;
        Regex numberRegex;

        const int maxDigitsInNumber = 10;

        public PokeSortNormalizer(string language)
        {
            removeDigitSeparatorRegex = GetRemoveDigitSeparatorRegex(language);
            numberRegex = new Regex(@"\d+");
        }
        private Regex GetRemoveDigitSeparatorRegex(string language)
        {
            switch (language)
            {
                case "German":
                    return new Regex(@"(\d)[ \.](\d{3})");
                case "French":
                case "Spanish":
                    return new Regex(@"(\d) (\d{3})");
                case "Italian":
                    return new Regex(@"(\d)\.(\d{3})");

                case "JPN":
                case "Korean":
                case "English":
                case "CHS"://simplifiedChinese,added by iQue
                case "CHT"://traditionalChinese,added by iQue
                default:
                    return new Regex(@"(\d),(\d{3})");
            }
        }

        private string RemoveDigitSeparator(string s)
        {
            return removeDigitSeparatorRegex.Replace(s, (Match m) => m.Groups[1].Value + m.Groups[2].Value);
        }

        // 数値に、maxDigitsInNumber 文字になるまで0を付加する
        private string ZeroPaddedNumber(string s)
        {
            return numberRegex.Replace(s, (Match m) => m.Value.PadLeft(maxDigitsInNumber, '0'));
        }

        public string Normalize(string word)
        {
            // <TODO> タグが入ってたらNG
            StringBuilder normalizedWord = new StringBuilder();
            word = ZeroPaddedNumber(RemoveDigitSeparator(word));
            foreach (char c in word)
            {
                string replacement;
                if (normalizeTable.TryGetValue(c, out replacement))
                {
                    normalizedWord.Append(replacement);
                }
                else
                {
                    normalizedWord.Append(Char.ToUpper(c));
                  }
            }

            return normalizedWord.ToString();
        }

    }
}
