using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChineseSort
{
    public class SorterItem<IdType>
    {
        public IdType id;
        public string word;
        public string normalized_word;
        public int initial;
        public SorterItem(IdType id, string word, string normalized_word)
        {
            this.id = id;
            this.word = word;
            this.normalized_word = normalized_word;
            this.initial = InitialNumerizer.Undefined;
        }

        //变量initial 指定首字母发音，一般在多音字才会出现 inital = InitialAlphabet.upper()-'A'
        public SorterItem(IdType id, string word, string normalized_word, int initial)
        {
            this.id = id;
            this.word = word;
            this.normalized_word = normalized_word;
            this.initial = initial;
        }

    }
    public class PokeMessageSorter<IdType> where IdType:IComparable<IdType>
    {
        private StringComparer comparer;
        private List<SorterItem<IdType>> list;
        private PokeSortNormalizer normalizer;
        private bool sorted = false;
        public PokeMessageSorter(string cultureName)
        {
            CultureInfo culture = new CultureInfo(cultureName);
            comparer = StringComparer.Create(culture, true);
            list = new List<SorterItem<IdType>>();

            normalizer = new PokeSortNormalizer(cultureName);
        }
        public void AddMessage(IdType id, string word, char initial)//指定首字母发音
        {
            string normalizedWord = normalizer.Normalize(word);
            list.Add(new SorterItem<IdType>(id, word, normalizedWord, Char.ToUpper(initial)-'A'));
            sorted = false;
        }
        public void AddMessage(IdType id, string word)
        {
            string normalizedWord = normalizer.Normalize(word);
            list.Add(new SorterItem<IdType>(id, word, normalizedWord));
            sorted = false;
        }
        public IList<SorterItem<IdType>> SortedMessages
        {
            get
            {
                if (!sorted)
                {
                    list.Sort(Compare);
                }
                return new ReadOnlyCollection<SorterItem<IdType>>(list);
            }
        }

        private int Compare(SorterItem<IdType> lhs, SorterItem<IdType> rhs)
        {
            int result;
            result = comparer.Compare(lhs.normalized_word, rhs.normalized_word);
            if (result == 0)
            {
                result = lhs.id.CompareTo(rhs.id);
            }
            return result;
        }
    }
}
