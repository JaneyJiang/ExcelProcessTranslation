using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Win32;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
//using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Threading;

namespace UI_Design
{
    /// <summary>
    /// Merge.xaml 的交互逻辑
    /// </summary>
    public partial class Merge : Window
    {
        private Dictionary<string,string> labelMap;
        private Dictionary<string, string> singleDict;
        private DataTable simALL;
        private string excelSuffix;
        public Merge()
        {
            InitializeComponent();
            InitData();
        }
        public void InitData()
        {
            labelMap = null;
            singleDict = null;
            simALL = null;
            excelSuffix = "xlsx";
        }

        private void tagChooseOrg_clicked(DataTable tb)
        {
            textBox.Text += "File loaded\n";
            simALL = tb.Copy();
        }
        private void tagChooseIndex_clicked(DataTable dt)
        {
            textBox.Text += "File loaded\n";
            labelMap = CommonTools.DataTableToDict(dt);
        }
        private void tagChooseSingle_clicked(DataTable tb)
        {
            textBox.Text += "File loaded\n";
            singleDict = new Dictionary<string, string>();
            for (int i = 1; i < tb.Rows.Count; i++)
            {
                string id = tb.Rows[i][0].ToString();
                if (!singleDict.ContainsKey(id))
                    singleDict.Add(id, tb.Rows[i][1].ToString());
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
                    case InputType.IndexFile:
                        tagChoose.sendTable += new TagChoose.Tabledelegate(tagChooseIndex_clicked);
                        break;
                    case InputType.TranslationFile:
                        tagChoose.sendTable += new TagChoose.Tabledelegate(tagChooseOrg_clicked);
                        break;
                    case InputType.MappingFile:
                        tagChoose.sendTable += new TagChoose.Tabledelegate(tagChooseSingle_clicked);
                        break;
                    default:
                        tagChoose.sendTable += new TagChoose.Tabledelegate(tagChooseOrg_clicked);
                        break;
                }
                textBox.Text += "Loading " + filePath + "...\n";
                tagChoose.SetData(filePath);
                tagChoose.Show();
            }
        }

        private void LoadIndexFile(object sender, RoutedEventArgs e)
        {
            string[] tableTags = { "Src", "Target" };
            LoadFile(tableTags, InputType.IndexFile);
        }

        private void LoadOrgFile(object sender, RoutedEventArgs e)
        {
            string[] tableTags = { "UniqueId", "Src", "Target" };
            LoadFile(tableTags, InputType.TranslationFile);
        }

        private void LoadSingleFile(object sender, RoutedEventArgs e)
        {
            string[] tableTags = { "UniqueId", "Target" };
            LoadFile(tableTags, InputType.MappingFile);
        }
        /*
        private void LoadIndexFile(object sender, RoutedEventArgs e)
        {
            var filePath = CommonTools.OpenOneFile(excelSuffix);
            if (filePath != null)
            {
                string[] tableTags = { "Src", "Target" };
                TagChoose tagChoose = new TagChoose(tableTags);
                tagChoose.sendTable += new TagChoose.Tabledelegate(tagChooseIndex_clicked);
                textBox.Text += "Loading " + filePath + "...\n";
                tagChoose.SetData(filePath);
                tagChoose.Show();
            }
        }
        
        private void LoadOrgFile(object sender, RoutedEventArgs e)
        {
            var filePath = CommonTools.OpenOneFile(excelSuffix);
            if (filePath != null)
            {
                string[] tableTags = { "UniqueId", "Src","Target" };
                TagChoose tagChoose = new TagChoose(tableTags);
                tagChoose.sendTable += new TagChoose.Tabledelegate(tagChooseOrg_clicked);
                textBox.Text += "Loading " + filePath + "...\n";
                tagChoose.SetData(filePath);
                tagChoose.Show();
            }
        }
        private void LoadSingleFile(object sender, RoutedEventArgs e)
        {
            var filePath = CommonTools.OpenOneFile(excelSuffix);
            if (filePath != null)
            {
                string[] tableTags = { "UniqueId", "Target" };
                TagChoose tagChoose = new TagChoose(tableTags);
                tagChoose.sendTable += new TagChoose.Tabledelegate(tagChooseSingle_clicked);
                textBox.Text += "Loading "+ filePath + "...\n";
                tagChoose.SetData(filePath);
                tagChoose.Show();
            }
        }
        */
        private void DoMerge(object sender, RoutedEventArgs e)
        {
            if (simALL == null)
            {
                MessageBox.Show("please load origin file!");
                return;
            }
            if (labelMap == null)
            {
                MessageBox.Show("please load index file!");
                return;
            }
            if (singleDict == null)
            {
                MessageBox.Show("please load simple file!");
                return;
            }

            textBox.Text += "Processing... \n";
            for (int i = 0; i < simALL.Rows.Count; i++)
            {
                string id = simALL.Rows[i][0].ToString();
                if (labelMap.ContainsKey(id))
                    id = labelMap[id];
                if (singleDict.ContainsKey(id))
                {
                    if (simALL.Rows[i][2] == DBNull.Value)
                    {
                        simALL.Rows[i][2] = singleDict[id];
                    }
                }
            }            
            CommonTools.SaveDataTable2Excel(simALL,"保存输出表");
            textBox.Text += "Finished!! \n";
            //Thread.Sleep(5000);
            Close();
        }
        
    }
}
