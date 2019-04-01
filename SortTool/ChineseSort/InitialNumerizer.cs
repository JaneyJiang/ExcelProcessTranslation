using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChineseSort
{
    //factory pattern,and use private class to prevent other class 
    //to call the abstract class's implementation
    public abstract class InitialNumerizer
    {
        //word's Numerizer not defined
        public const int Undefined = -1;
        //word's Numerizer is not unique
        public const int UnUnique = -2;
        // initialCount which Numerizer can recoginzed.
        public abstract int InitialCount { get; }

        // 文字cのイニシャル番号を取得。値域は[0, InitialCount)
        public abstract int GetInitialNumber(char c);

        //at first use language,then change to culture_name , 
        //please note the difference between langue and culture_name
        //eg:CHT<->zh-TW, CHS<->zh-CN.Korean<->ko-KR,They have global standards.

        public static InitialNumerizer Create(string culture_name)
        {
            switch (culture_name)
            {
                case "zh-CN"://simplifiedChinese,added by iQue
                    return new InitialNumberizerSimplifiedChinese();
                case "zh-TW"://traditionalChinese,added by iQue
                    return new InitialNumberizerTraditionalChinese();
                default:
                    throw new ArgumentException();
            }
        }

    }

}
