using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows;


namespace SortTool
{
    /// <summary>
    /// ExcelTagChoose.xaml 的交互逻辑
    /// </summary>
    public partial class ExcelTagChoose : Window
    {
        private ConnectExcel connExcel;
        private List<System.Data.DataTable> tablelist = null;
        // private string[] labelNames;
        private string filepath;
        //private List<string> excludeSheet;
        private DataTable dt;
        public ExcelTagChoose()
        {
            InitializeComponent();
        }

        public ExcelTagChoose(string[] names)
        {
            InitializeComponent();
            foreach (string name in names)
            {
                UserControlExcelTag cc = new UserControlExcelTag(name);
                stackPanel.Children.Add(cc);
            }
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
            foreach (UserControlExcelTag cc in stackPanel.Children.OfType<UserControlExcelTag>())
            {
                cc.Add(colNames.ToArray());
            }
        }
        public delegate void Tabledelegate(DataTable table);//定义一个委托
        public event Tabledelegate sendTable;//实例化对象（二）
        private void SelecetedOK_Click(object sender, RoutedEventArgs e)
        {
            List<string> selectedColumns = new List<string>();
            foreach (UserControlExcelTag cc in stackPanel.Children.OfType<UserControlExcelTag>())
            {
                selectedColumns.Add(cc.labelSelect);
            }
            dt = CopyTableColumns(tablelist, selectedColumns.ToArray());
            //dt = CreateTableWithSpecifiedColumns(tablelist[0].Columns,labelNames);
            //dt = FillTableData(dt, tablelist,labelNames);
            //ExportExcel(dt, System.IO.Path.GetDirectoryName(filepath) + "\\test.xlsx");

            if (dt != null)
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

        private bool isContains(DataTable tb, string[] names)
        {
            foreach (string name in names)
            {
                if (!tb.Columns.Contains(name))
                    return false;
            }
            return true;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

    }
}
