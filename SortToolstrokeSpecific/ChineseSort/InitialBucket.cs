using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChineseSort
{
    //为了保证排序结果按照initial来进行分类和排序。
    public class InitialBuckets<IdType>
    {
        private InitialNumerizer initialNumerizer;
        private List<SorterItem<IdType>>[] initialBucket;

        // 分類の数（イニシャル数 + undefined initial个数+ unUnique initial个数）
        public int BucketCount { get { return initialNumerizer.InitialCount + 1 + 1; } }
        // 未定義イニシャルのインデックス(BucketCount-1)
        public int UnknownInitialIndex { get { return 0; } }

        public int UnUniqueInitialIndex { get { return BucketCount - 1; } }

        public InitialBuckets(string language)
        {
            initialNumerizer = InitialNumerizer.Create(language);
            initialBucket = new List<SorterItem<IdType>>[BucketCount];
            for (int i = 0; i < BucketCount; i++)
            {
                initialBucket[i] = new List<SorterItem<IdType>>();
            }
        }

        // イニシャルで分類されたSorterItemを取得。UnknownInitialIndex = BucketCount-1が未定義イニシャル
        public ReadOnlyCollection<SorterItem<IdType>> GetItemWithInitial(int index)
        {
            return new ReadOnlyCollection<SorterItem<IdType>>(initialBucket[index]);
        }

        private int GetIndex(string word)
        {//规则就是，不认识的标记为0，不唯一的标记到最后一位。,因为initialNumber是从0开始的，所以返回的时候需要+1
            if (word != "")
            {
                int initial = initialNumerizer.GetInitialNumber(word[0]);
                if (initial == InitialNumerizer.Undefined)
                {
                    return UnknownInitialIndex;
                }
                else if (initial == InitialNumerizer.UnUnique)
                {
                    return UnUniqueInitialIndex;
                }
                else {
                    return initial+1;
                }
            }
            else
            {
                return UnknownInitialIndex;
            }
        }

        public void Add(SorterItem<IdType> sorterItem)
        {
            int index = 0;
            if (sorterItem.initial == InitialNumerizer.Undefined)
            {
                index = GetIndex(sorterItem.normalized_word);
            } else if (sorterItem.initial == InitialNumerizer.UnUnique)
            {
                index = UnUniqueInitialIndex;
            }
            else {
                index = sorterItem.initial + 1;
            }
            initialBucket[index].Add(sorterItem);
        }
    }
}
