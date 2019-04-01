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
            string[] tableTags = { "Label", "toSort" };
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
                sort_out = DoSort(sort_in);
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
            }

            CommonTools.SaveDataTable2Excel(sort_out, "保存分类排序表");
            textBox.Text += "output row lines:" + sort_out.Rows.Count.ToString() + "\n";

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
            DataColumn dc = new DataColumn("initial", Type.GetType("System.Int16"));
            sorted.Columns.Add(dc);
            //dc = new DataColumn("OrderedNo", Type.GetType("System.Int16"));
            //sorted.Columns.Add(dc);
            
           // for (int initial = 0,orderedNo = 1; initial < initialBuckets.BucketCount; initial++)
            for (int initial = 0; initial < initialBuckets.BucketCount; initial++)
                {
                foreach (var item in initialBuckets.GetItemWithInitial(initial))
                {
                    DataRow dr = sorted.NewRow();
                    dr[0] = item.id;
                    dr[1] = item.word;
                    dr[2] = initial;
                    //dr[3] = orderedNo;
                    //orderedNo += 1;
                    sorted.Rows.Add(dr);
                }
            }

            return sorted;
        }
    }
}
