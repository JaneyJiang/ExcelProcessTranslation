using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using Microsoft.Win32;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Data;
using FuncLibrary;
using System.Collections;
using System.IO;
using System.Runtime.InteropServices;

namespace UI_Design
{
    /// <summary>
    /// Similarity.xaml 的交互逻辑
    /// </summary>
    /// 

    enum InputType
    {
        DictionaryFile,
        IndexFile,
        TranslationFile,
        MappingFile,
    }
    public delegate void SendTable(DataTable tb);


    public partial class Similarity : Window
    {

        //private TagChoose tagChoose;
        private string excelSuffix;
        private Dictionary<string, string> dict;
        private DataTable idxMap;
        private DataTable single;
        private double simRate;
        private double lenRate;
        private string identify;

        public Similarity()
        {
            InitializeComponent();
            InitData();
        }

        private void InitData()
        {
            textBox.Text = "";
            excelSuffix = "xlsx";
            dict= null;
            idxMap = null;
            single = null;
            simRate = 0.85;//相似度大小
            lenRate = 0.65;//句子长度比
            identify = "identify";
        }
     
        private void tagChooseDict_clicked(DataTable dt)
        {
            textBox.Text += "File loaded\n";
            try
            {
                dict = CommonTools.DataTableToDict(dt);
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
            }
        }

        private void LoadFile(string[] tableTags, InputType type)
        {
            var filePath = CommonTools.OpenOneFile(excelSuffix);
            if (filePath != null)
            {
                TagChoose tagChoose = new TagChoose(tableTags);
                switch (type)//不用这个长度，我们在调用的时候指定一个类型，根据类型使用回调函数TODO
                {
                    case InputType.DictionaryFile:
                        tagChoose.sendTable += new TagChoose.Tabledelegate(tagChooseDict_clicked);
                        break;
                    case InputType.TranslationFile:
                        tagChoose.sendTable += new TagChoose.Tabledelegate(tagChooseText_clicked);
                        break;
                    default:
                        tagChoose.sendTable += new TagChoose.Tabledelegate(tagChooseText_clicked);
                        break;
                }
                textBox.Text += "Loading " + filePath + "\n";
                tagChoose.SetData(filePath);
                tagChoose.Show();
            }
        }

        private void LoadDictFile(object sender, RoutedEventArgs e)
        {
            string[] tableTags = { "待翻译文本", "目标文本" };
            LoadFile(tableTags,InputType.DictionaryFile);
        }
        private void LoadTranslationFile(object sender, RoutedEventArgs e)
        {
            string[] tableTags = { "UniqueId", "Source", "Target" };
            LoadFile(tableTags,InputType.TranslationFile);
        }
        

        private void tagChooseText_clicked(DataTable tb)
        {
            textBox.Text += "File loaded\n";
            try
            {
                single = tb.Clone();//复制输入tb的结构
                DataTable sortTable = SortAccordingToLength(tb, 1);//根据指定列的长度进行排序
                Dictionary<string, string> indexMap = SingleExtraction(sortTable, single);
                //CommonTools.SaveDataTable2Excel(single);
                idxMap = CommonTools.DictToDataTable<string>(indexMap);
                //CommonTools.SaveDataTable2Excel(map);
            }
            catch (Exception err) {
                MessageBox.Show(err.Message);
            }
        }

       
        private DataTable SortAccordingToLength(DataTable tb, int idx)
        {
            DataTable table = tb.Copy();
            table.Columns.Add(identify, typeof(int));
            string processStr = null;
            //删除无效行
            for (int i = table.Rows.Count - 1; i >= 0; i--)
            {
                DataRow dr = table.Rows[i];
                if (dr[idx] == DBNull.Value)
                {
                    dr.Delete();
                    continue;
                }
                processStr = processStr = StringTool.preprocess(dr[idx].ToString());
                if (processStr.Length == 0)
                {
                    dr.Delete();
                    continue;
                }
                dr[idx] = processStr;
                dr[identify] = processStr.Length;
            }
           
            table.AcceptChanges();
            table.DefaultView.Sort = identify + " asc";
            return table.DefaultView.ToTable();
        }

        private DataRow CreateAndAssignNewRow(DataTable dt, DataRow dr)
        {
            DataRow dataRow = dt.NewRow();
            for(int i=0;i<dt.Columns.Count;i++)
            {
                dataRow[i] = dr[i];
            }
            if (dataRow[2] == DBNull.Value)
            {
                if (dict != null)
                {
                    dataRow[2] = DoTranslation(dataRow[1].ToString());
                }
            }
            dt.Rows.Add(dataRow);
            return dataRow;
        }

        public string DoTranslation(string sent)
        {
            foreach (string key in dict.Keys)
            {
                if (sent.Contains(key))
                    sent = sent.Replace(key, dict[key]);
            }
            return sent;
        }

        //提取独字符串，把相似的字符串的id放入字典
        public Dictionary<string,string> SingleExtraction(DataTable org,DataTable output)
        {
            Dictionary<string,string> labelMap = new Dictionary<string, string>();
            HashSet<int> cmplist = new HashSet<int>();//存储已经比较过的index
            int sourceCol = 1;
            for (int i = 0; i < org.Rows.Count; i++)
            {
                if (cmplist.Contains(i))
                    continue;
                CreateAndAssignNewRow(output, org.Rows[i]);//组长和单条都进表，其他组员不进表。只进idxMap
                for (int j = i + 1; j < org.Rows.Count; j++)
                {
                    if (cmplist.Contains(j))
                        continue;
                    if (Levenshtein.Jaro_Winkler((string)org.Rows[i][sourceCol], (string)org.Rows[j][sourceCol]) > simRate)
                    {
                        cmplist.Add(j);
                        if (labelMap.ContainsKey(org.Rows[j][0].ToString()))
                            continue;
                        //labelMap.Add(Convert.ToInt32(org.Rows[j][0]), Convert.ToInt32(org.Rows[i][0]));
                        labelMap.Add(org.Rows[j][0].ToString(), org.Rows[i][0].ToString());
                    }
                    else
                    {
                        if (Convert.ToSingle(org.Rows[i][identify]) / Convert.ToSingle(org.Rows[j][identify]) < lenRate)
                            break;
                    }
                }
            }
            return labelMap;
        }
        private void SaveSplitTables(object sender, RoutedEventArgs e)
        {
            textBox.Text += "Saving index file...\n";
            CommonTools.SaveDataTable2Excel(idxMap, "保存索引表");
            textBox.Text += "Index file saved\n";
            textBox.Text += "Saving simple table...\n";
            CommonTools.SaveDataTable2Excel(single,"保存输出简表");
            textBox.Text += "All Saved\n";
            Close();
        }

        private void LoadDict(object sender, RoutedEventArgs e)
        {
            var filePath = CommonTools.OpenOneFile(excelSuffix);
            if (filePath != null)
            {
                string[] tableTags = { "待翻译文本", "目标文本" };
                TagChoose tagChoose = new TagChoose(tableTags);
                tagChoose.sendTable += new TagChoose.Tabledelegate(tagChooseDict_clicked);
                textBox.Text = "Loading " + filePath + "\n";
                tagChoose.SetData(filePath);
                tagChoose.Show();

            }
        }
        private void LoadTransText(object sender, RoutedEventArgs e)
        {
            var filePath = CommonTools.OpenOneFile(excelSuffix);
            if (filePath != null)
            {
                string[] tableTags = { "UniqueId", "Source", "Target" };
                TagChoose tagChoose = new TagChoose(tableTags);
                tagChoose.sendTable += new TagChoose.Tabledelegate(tagChooseText_clicked);
                textBox.Text += "Loading " + filePath + "\n";
                tagChoose.SetData(filePath);
                tagChoose.Show();
            }
        }


        public static DataTable CreatEmptyTableColumns(DataColumnCollection columns)
        {
            DataTable table = new DataTable();
            foreach (DataColumn dc in columns)
            {
                DataColumn col = new DataColumn();
                col.DataType = dc.DataType;
                col.ColumnName = dc.ColumnName;
            }
            return table;
        }
    }
}
