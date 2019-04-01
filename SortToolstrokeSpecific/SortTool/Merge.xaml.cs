using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
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
    /// Merge.xaml 的交互逻辑
    /// </summary>
    /// 
    enum InputType
    {
        ToSortedFile,
        SortedFile
    }
public partial class Merge : Window
    {
        private string file_suffix = "xlsx";
        private DataTable org;
        private Dictionary<string, string> sortedMap;
        public Merge()
        {
            InitializeComponent();
        }
        private void tagChooseOrg_clicked(DataTable tb)
        {
            textBox.Text += "File loaded! ";
            org = tb;
            textBox.Text += "input row lines:" + tb.Rows.Count.ToString() + "\n";
        }
        private void tagChooseSorted_clicked(DataTable tb)
        {
            textBox.Text += "File loaded! ";
            sortedMap = CommonTools.DataTableToDict(tb);
            textBox.Text += "input row lines:" + tb.Rows.Count.ToString() + "\n";
        }
        private void LoadFile(string[] tableTags, InputType type)
        {
            var filePath = CommonTools.OpenOneFile(file_suffix);
            if (filePath != null)
            {
                ExcelTagChoose tagChoose = new ExcelTagChoose(tableTags);
                switch (type)
                {
                    case InputType.ToSortedFile:
                        tagChoose.sendTable += new ExcelTagChoose.Tabledelegate(tagChooseOrg_clicked);
                        break;
                    case InputType.SortedFile:
                        tagChoose.sendTable += new ExcelTagChoose.Tabledelegate(tagChooseSorted_clicked);
                        break;
                    default:
                        break;
                     
                }
                textBox.Text += "Loading " + filePath + "\n";
                tagChoose.SetData(filePath);
                tagChoose.Show();
            }
        }
        private void LoadOrgFile(object sender, RoutedEventArgs e)
        {
            string[] tableTags = { "Label", "toSortNo" };
            try
            {
                LoadFile(tableTags,InputType.ToSortedFile);
                //textBox.Text += "载入字典" + this.dict.Count.ToString() + "条";
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
            }

        }

        private void LoadSortedFile(object sender, RoutedEventArgs e)
        {
            string[] tableTags = { "Label", "SortedNo" };
            try
            {
                LoadFile(tableTags, InputType.SortedFile);
                //textBox.Text += "载入字典" + this.dict.Count.ToString() + "条";
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
            }

        }
        private void MergeAndOutput(object sender, RoutedEventArgs e)
        {
            if (org == null)
            {
                MessageBox.Show("please load total file!");
                return;
            }
            if (sortedMap == null)
            {
                MessageBox.Show("please load sorted file!");
                return;
            }
            
            textBox.Text += "Processing... \n";
            string id;
            string sortedNo;
            for (int i = 0; i < org.Rows.Count; i++)
            {
                id = org.Rows[i][0].ToString();
                if (sortedMap.ContainsKey(id))
                {
                    sortedNo = sortedMap[id];
                    if (org.Rows[i][1] == DBNull.Value)
                    {
                        org.Rows[i][1] = sortedNo;
                    }
                }

            }
            CommonTools.SaveDataTable2Excel(org, "保存输出表");

        }
    }
}
