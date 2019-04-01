using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using ChineseSort;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Collections;

namespace SortTool
{
    /// <summary>
    /// Sort.xaml 的交互逻辑
    /// </summary>
    public partial class Sort : Window
    {
        private string file_suffix;
        private string culture_name;
        private DataTable sort_out;
        private DataTable sort_in;
        private Dictionary<string, string> sortedMap;
        private ArrayList colnames = new ArrayList();//分别对应的名字[Id, words, sortedNo,stroke]
        public Sort()
        {
            InitializeComponent();
            // InitLanguageTag();
            InitData();
        }
        private void InitData()
        {
            file_suffix = "xlsx";
            selected_culture.Content = "zh-CN";
            culture_name = selected_culture.Content.ToString();
            textBox.Text = "selected culture:" + culture_name + "\n";
        }
        private void LoadToSortFile(object sender, RoutedEventArgs e)
        {
            string[] tableTags = { "Id", "toSort" };
            try { 
                LoadFile(tableTags);
                //textBox.Text += "载入字典" + this.dict.Count.ToString() + "条";
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
            }
        }
        private void LoadFile(string[] tableTags)
        {
            var filePath = CommonTools.OpenOneFile(file_suffix);
            if (filePath != null)
            {
                ExcelTagChoose tagChoose = new ExcelTagChoose(tableTags);

                tagChoose.sendTable += new ExcelTagChoose.Tabledelegate(tagChooseText_clicked);
                textBox.Text += "Loading " + filePath + "\n";
                tagChoose.SetData(filePath);
                tagChoose.Show();
            }
        }


        private void tagChooseText_clicked(DataTable tb)
        {
            textBox.Text += "File loaded! ";
            sort_in = tb;
            foreach (DataColumn dc in tb.Columns)
            {
                colnames.Add(dc.ColumnName.ToString());
            }
            textBox.Text += "input row lines:" + tb.Rows.Count.ToString() + "\n";

        }

        private void ChangeCulture(object sender, RoutedEventArgs e)
        {
            if (selected_culture.Content.ToString() == "zh-TW")
            {
                selected_culture.Content = "zh-CN";
            }
            else
            {
                selected_culture.Content = "zh-TW";
            }
            culture_name = selected_culture.Content.ToString();
            textBox.Text += "selected culture:" + culture_name + "\n";
        }

        private void SortAndSave(object sender, RoutedEventArgs e)
        {
            try
            {
                DataTable sorted = DoSort(sort_in);
                sort_out = mergeBack(sort_in, sorted);

                //这张表是一张干净的，不包含空格，只包含笔顺和次序的表
                CommonTools.SaveDataTable2Excel(sorted, "保存笔顺表");
                textBox.Text += "output row lines:" + sorted.Rows.Count.ToString() + "\n";

                DataRow[] selected = sorted.Select("initial = 0");
                if (isChineseWordExist(selected, 1))
                {
                    MessageBox.Show("Check Unidentify Chinese Words First!");
                    return;
                }
                //根据原表的结构填回，只包含次序，不包含笔顺。
                CommonTools.SaveDataTable2Excel(sort_out, "保存分类排序表");
                textBox.Text += "output row lines:" + sort_out.Rows.Count.ToString() + "\n";
                

            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
            }           

        }

        bool isChineseWordExist(DataRow[] rows, int tb_idx)//检查是否指定列有中文出现在首字。
        {
            int chfrom = Convert.ToInt32("4e00", 16);
            int chend = Convert.ToInt32("9fa5", 16);
            int code;
            DataRow dr;
            for(int i=0;i< rows.Length;i++)
            {
                dr = rows[i];
                code = Char.ConvertToUtf32(dr[tb_idx].ToString(), 0);
                if (code >= chfrom && code <= chend)
                {
                    return true;
                }
            }
            return false;
        }
        static string ConvertSubstituteTag(string s)
        {
            return s.Replace("[Character2:PKMN ]", "⒆⒇")
                .Replace("[Character1:male ]", "⑭")
                .Replace("[Character1:female ]", "⑮");
        }

        private DataTable DoSort(DataTable dt)
        {
            if (dt.Columns.Count < 2)
            {
                throw new Exception("input data is less than two columns!");
            }
            var sorter = new PokeMessageSorter<string>(selected_culture.Content.ToString());
            for (int i = 0; i < dt.Rows.Count; i++)
            {              
                   if (dt.Rows[i][0] == DBNull.Value || dt.Rows[i][1] == DBNull.Value) continue;
                   string label = dt.Rows[i][0].ToString();
                   string word = dt.Rows[i][1].ToString().TrimEnd('\r', '\n');
                   if (word == "-") continue;
                   label = ConvertSubstituteTag(label);
                   word = ConvertSubstituteTag(word);
                    //TODO:对多音字的判别还没做int initial = CHSFilter.FindFilter(word);
                   sorter.AddMessage(label, word);
            }
            var initialBuckets = new InitialBuckets<string>(culture_name);
            foreach (var item in sorter.SortedMessages)
            { initialBuckets.Add(item); }
            DataTable sorted = dt.Clone();//Copy拷贝数据和结构，Clone拷贝结构
            DataColumn dc = new DataColumn("sortedNo", Type.GetType("System.Int16"));
            sorted.Columns.Add(dc);
            colnames.Add("sortedNo");
            dc = new DataColumn("initial", Type.GetType("System.Int16"));
            sorted.Columns.Add(dc);
            colnames.Add("initial");

            int orderedNo = 1;
           // for (int initial = 0,orderedNo = 1; initial < initialBuckets.BucketCount; initial++)
            for (int initial = 0; initial < initialBuckets.BucketCount; initial++)
                {
                foreach (var item in initialBuckets.GetItemWithInitial(initial))
                {
                    DataRow dr = sorted.NewRow();
                    dr[0] = item.id;
                    dr[1] = item.word;
                    dr[2] = orderedNo;
                    dr[3] = initial;

                    orderedNo += 1;
                    sorted.Rows.Add(dr);
                }
            }
                       
            return sorted;
        }

        private DataTable mergeBack(DataTable org, DataTable sorted )
        {
            sortedMap = CommonTools.DataTableToDict(sorted,0,2);//0,2对应相应提取的列的信息来制作字典对应。
            string id;
            string sortedNo;
            DataColumn dc = new DataColumn("sortedNo", Type.GetType("System.Int16"));
            org.Columns.Add(dc);
            for (int i = 0; i < org.Rows.Count; i++)
            {
                if (org.Rows[i][0] == DBNull.Value)
                    continue;
                id = org.Rows[i][0].ToString();
                if (sortedMap.ContainsKey(id))
                {
                    sortedNo = sortedMap[id];
                    org.Rows[i][2] = sortedNo;
                }
            }
            return org;

        }
    }
}
