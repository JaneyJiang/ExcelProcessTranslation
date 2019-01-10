using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using FuncLibrary;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace UI_Design
{
    /// <summary>
    /// TagChoose.xaml 的交互逻辑
    /// </summary>
    public partial class TagChoose : Window
    {
        private ConnectExcel connExcel;
        private List<System.Data.DataTable> tablelist = null;
       // private string[] labelNames;
        private string filepath;
        //private List<string> excludeSheet;
        private DataTable dt;
        public TagChoose(string[] names)
        {
            InitializeComponent();
            foreach (string name in names)
            {
                CustomControl cc = new CustomControl(name);
                stackPanel.Children.Add(cc);
            }
            //labelNames = names;
        }
        public void SetData(string filepath)
        {
            this.filepath = filepath;
            connExcel = new ConnectExcel(filepath);
            string[] sheetnames = connExcel.GetSheetsNames();
            tablelist = connExcel.GetTableList();
            List<string> colNames = new List<string>();
            foreach (DataColumn dc in tablelist[0].Columns)
            {
                colNames.Add(dc.ColumnName);
            }
            foreach (CustomControl cc in stackPanel.Children.OfType<CustomControl>())
            {
                cc.Add(colNames.ToArray());
            }
        }
        public delegate void Tabledelegate(DataTable table);//定义一个委托
        public event Tabledelegate sendTable;//实例化对象（二）
        private void selectedBtn_Click(object sender, RoutedEventArgs e)
        {
            List<string> selectedColumns = new List<string>();
            foreach(CustomControl cc in stackPanel.Children.OfType<CustomControl>())
            {
                selectedColumns.Add(cc.labelSelect);
            }
            dt = CopyTableColumns(tablelist, selectedColumns.ToArray());
            //dt = CreateTableWithSpecifiedColumns(tablelist[0].Columns,labelNames);
            //dt = FillTableData(dt, tablelist,labelNames);
            //ExportExcel(dt, System.IO.Path.GetDirectoryName(filepath) + "\\test.xlsx");
            
            if (dt!=null)
            {
                sendTable(dt);
            }
            //MessageBox.Show("Finished");
            Close();
        }

       

        private DataTable CopyTableColumns(List<DataTable> tables, string[] colNames)
        {
            DataTable all = new DataTable();
            List<string> excludeSheet = new List<string>();
            for (int j = 0; j < tablelist.Count; j++)
            {
                System.Data.DataTable table = tables[j];
                if (!isContains(table, colNames))
                {
                    excludeSheet.Add(table.TableName);
                    continue;
                }
                DataTable dat = table.DefaultView.ToTable(false, colNames);
                //DataTable dat = table.DefaultView.ToTable();
                all.Merge(dat);               
            }
            return all;

        }
        private DataTable CreateTableWithSpecifiedColumns(DataColumnCollection columns,string[]names)
        {
            DataTable dt = new DataTable();
            foreach (DataColumn dc in columns)
            {
                if (names.Contains(dc.ColumnName))
                {
                    DataColumn col = new DataColumn();
                    col.DataType = dc.DataType;
                    col.ColumnName = dc.ColumnName;
                }
            }
            return dt;
        }
        private bool isContains(DataTable tb,string[] names)
        {
            foreach (string name in names)
            {
                if (!tb.Columns.Contains(name))
                    return false;
            }
            return true;
        }
        private bool isNullExist(DataRow tr, string[] names)
        {
            foreach (string name in names)
            {
                if (tr[name] == DBNull.Value)
                    return true;
            }
            return false;
        }
        private DataRow fillRow(DataRow src, DataRow tgt,string[] names)
        {
            foreach (string name in names)
            {
                tgt[name] = src[name];
            }
            return tgt;
        }
        private DataRow preProcess(DataRow dr, string colName)
        {
            dr[colName] = StringTool.preprocess((string)dr[colName]);
            return dr;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

    }
}
