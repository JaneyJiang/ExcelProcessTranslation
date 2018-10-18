//using Microsoft.Office.Interop.Excel;
using FuncLibrary;
using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Text;
using System.Windows;


namespace ExcelProcess
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        private List<System.Data.DataTable> tablelist = null;
        private string connStr = null;
        // private string[] filepaths;//当需要打开多个文件获得多个路径的时候可以用这个
        private string filepath;
        private string[] sheetnames;
        public MainWindow()
        {
            InitializeComponent();
            //connStr = "Provider=Microsoft.ACE.OLEDB.12.0; Extended Properties = 'Excel 12.0 Xml;HDR=YES;IMEX=1;MAXSCANROWS=0'";
            tablelist = new List<System.Data.DataTable>();
        }

        private void btnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Multiselect = false;
            fileDialog.Filter = "Excel Files|*.xlsx;";
            fileDialog.DefaultExt = ".xlsx";
            Nullable<bool> dialogOk = fileDialog.ShowDialog();
            if (dialogOk == true)
            {
                filepath = fileDialog.FileName;
                txbText.Text = filepath;
            }
            else
            {
                filepath = "";
                txbText.Text = "open no file";
            }
            //获取文件扩展名来获得不同的连接设置
            string strExtension = System.IO.Path.GetExtension(filepath);
            //Excel与Excel的连接
            //HDR=Yes,这代表第一行是标题，不作为数据使用
            //IMEx0:写入，1：读取，2：读取 写入
            switch (strExtension)
            {
                case ".xlsx":
                    connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1;MAXSCANROWS=0';";
                    break;
                default:
                    break;
            }
            if (connStr == null)
                throw new Exception("input file type error");
            else
                connStr += "Data Source=" + filepath;

            sheetnames = GetSchemaTable(connStr);
            cbSelect_set(sender, e);
        }
        private void cbSelect_set(object sender, EventArgs e)
        {
            foreach (DataColumn dc in tablelist[0].Columns)
            {
                cbSelect.Items.Add(dc.ColumnName);
                cbSelectLabel.Items.Add(dc.ColumnName);
            }
            cbSelectLabel.Items.Add(null);
            cbSelect.SelectedIndex = -1;
            cbSelectLabel.Items.Add(null);
            cbSelectLabel.SelectedIndex = -1;
        }

        private void btnComfirm_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("选定label和source", "提示", MessageBoxButton.OKCancel) == MessageBoxResult.OK)
            {
                //这儿就是点击确定 
                txbText.Text += "\n Processing " + sheetnames[0] + "...";
                try
                {
                    List<string> delet = MakeTable((string)cbSelectLabel.SelectedValue, (string)cbSelect.SelectedValue);
                    foreach (string name in delet)
                    {
                        txbText.Text += "   exclued file: " + name;
                    }
                    txbText.Text += "   Finished! ";
                }
                catch (Exception err)
                {
                    MessageBox.Show(err.Message);
                }
            }

        }

        private string[] GetSchemaTable(string connectionString)
        {
    
            using (OleDbConnection connection = new
                       OleDbConnection(connectionString))
            {
                connection.Open();
                System.Data.DataTable schemaTable = connection.GetOleDbSchemaTable(
                    OleDbSchemaGuid.Tables,
                    new object[] { null, null, null, "TABLE" });
                if (schemaTable == null)
                    return null;
                string[] excelSheets = new string[schemaTable.Rows.Count];

                int i = 0;
                foreach (System.Data.DataRow dr in schemaTable.Rows)
                {
                    string sheetname = dr["TABLE_NAME"].ToString().Trim();
                    excelSheets[i]=sheetname;               
                    string sql = "SELECT * FROM [" + sheetname + "]";
                    System.Data.DataSet ds = new DataSet();
                    ds.Clear();
                    OleDbDataAdapter data = new OleDbDataAdapter(sql, connStr);
                    data.Fill(ds);
                    tablelist.Add(ds.Tables[0]);
                    tablelist[i].TableName = sheetname;
                    Console.WriteLine(tablelist[i].TableName);
                    Console.WriteLine(i);
                    i += 1;
                }             
                return excelSheets; ;
            }
        }

        //CreatEmptyTable:根据给定的表名和列名，创建空的Table。给定的名字中，第一个是表格名，第二个开始就是列名
        private DataTable CreatEmptyTable(string[] names, string[] types)
        {
            int i = 0;
            DataTable tb = new DataTable((string)names[i]);
            for (i=1; i < names.Length-1; i++)
            {
                tb.Columns.Add(newCol(types[i-1], (string)names[i]));
            }
            tb.Columns.Add(newCol(types[i-1], (string)names[i]));
            return tb;
        }
        /*private DataTable CreatEmptyTable(string name, string labelCol, string srcCol)
        {
            //创建一个空表格结构，其中，表格内包括 id， labelCol，srcCol，strLen,其中id是对记个数的从1开始，strLen是对srcCol的长度来记长度的。
            DataTable tb = new DataTable(name);
            //define the new DataTable's column names.
            List<String> colnames = new List<String>();
            colnames.Add("id");
            colnames.Add(labelCol);
            colnames.Add(srcCol);
            colnames.Add("srcLen");

            //Create DataColumns and set various Properties. 
            DataColumn col;
            col = newCol("System.Int32", colnames[0]);
            tb.Columns.Add(col);
            tb.PrimaryKey = new DataColumn[] { col };//find 方法需要主键，如果不适用find方法可以不用主键。

            if (labelCol != null)
            {
                tb.Columns.Add(newCol("System.String", colnames[1]));
            }
            tb.Columns.Add(newCol("System.String", colnames[2]));
            tb.Columns.Add(newCol("System.Int32", colnames[3]));

            return tb;
        }*/
        System.Data.DataColumn newCol(string type, string colname)
        {
            System.Data.DataColumn col = new System.Data.DataColumn();
            col.DataType = System.Type.GetType(type);
            col.ColumnName = colname;
            return col;
        }
        private List<string> MakeTable(string labelCol, string sourceCol)
        {
            List<string> exclueSheet = new List<string>();
            //create a DataTable
            string[] names = { "all", "id", labelCol, sourceCol, "srcLen","tranTo" };
            string[] types = { "System.Int32", "System.String", "System.String", "System.Int32", "System.String"};
            System.Data.DataTable tb = CreatEmptyTable(names, types);
            DataRow row;
            string preProcess = null;
            int t = 1;//if you have only one sheet than it matches your index.
            //create rows and set the values.for several tables
            for (int j=0;j < tablelist.Count; j++)
            {
                System.Data.DataTable table = tablelist[j];
                if (table.Columns.Contains(labelCol) && table.Columns.Contains(sourceCol))
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        if (table.Rows[i][labelCol] == DBNull.Value || table.Rows[i][sourceCol] == DBNull.Value)
                        {
                            t++;
                            continue;                          
                        }
                        preProcess = StringTool.preprocess((string)table.Rows[i][sourceCol]);
                        if (preProcess.Length ==0)
                        {
                            t++;
                            continue;
                        }
                        row = tb.NewRow();
                        row[tb.Columns[0]] = t + 1;
                        row[tb.Columns[1]] = table.Rows[i][labelCol];
                        
                        row[tb.Columns[2]] = preProcess; 
                        row[tb.Columns[3]] = preProcess.Length;
                        tb.Rows.Add(row);
                        t++;
                    }
                }
                else
                {
                    exclueSheet.Add(table.TableName);
                }
                
            }
            tb.DefaultView.Sort = tb.Columns[3] + " ASC";
            tb = tb.DefaultView.ToTable();
            DataTabletoCsv(tb, "./"+sheetnames[0]+"origin.csv");

            //according to the whole table and string lenthg information to generate simiality table.
            DataTable simTb = tb.Clone();//复制原表的结构
            //DataTable simTb = CreatEmptyTable("simData", "id", labelCol, sourceCol, "srcLen");
            Hashtable simpair = new Hashtable();
            HashSet<int> cmplist = new HashSet<int>();
            int count = 0;
            for (int i = 0;i<tb.Rows.Count;i++)
            {
                int groupRow = 0;
                int groupNum = 0;
                if (cmplist.Contains(i))
                    continue;
                cmplist.Add(i);
                DataRow dataRow = tb.Rows[i];
                dataRow[sourceCol] = Translate((string)dataRow[sourceCol]);
                simTb.ImportRow(dataRow);
                groupRow = count;
                count += 1;
                for (int j = i + 1; j < tb.Rows.Count; j++)
                {
                    if (cmplist.Contains(j))
                        continue;
                    if (Levenshtein.Jaro_Winkler((string)tb.Rows[i][sourceCol], (string)tb.Rows[j][sourceCol]) > 0.8)
                    {
                        cmplist.Add(j);
                        tb.Rows[j]["srcLen"] = -1;
                        simTb.ImportRow(tb.Rows[j]);
                        groupNum += 1;
                        count += 1;
                    }
                    else {
                        if (Convert.ToSingle(tb.Rows[i]["srcLen"]) / Convert.ToSingle(tb.Rows[j]["srcLen"]) < 0.8)
                            break;
                    }
                }
                simTb.Rows[groupRow]["srcLen"] = groupNum;
                

            }
            DataTabletoCsv(simTb, "./" + sheetnames[0] + "_sim.csv");

            return exclueSheet;
        }

        private void DataTabletoCsv(System.Data.DataTable dt, string path)

        {
            StreamWriter sw = null;

            //这里的文件打开异常不需要捕获太早，当然可以通过返回bool型来对异常进行判断处理。
            sw = new StreamWriter(path, false, Encoding.GetEncoding("utf-8"));


            StringBuilder sb = new StringBuilder();

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                sb.Append(dt.Columns[i].ColumnName.ToString() + ",");
            }

            sb.Append(Environment.NewLine);

            for (int m = 0; m < dt.Rows.Count; m++)
            {
                //System.Windows.Forms.Application.DoEvents();
                for (int n = 0; n < dt.Columns.Count; n++)
                {
                    sb.Append(dt.Rows[m][n].ToString() + ",");
                }

                sb.Append(Environment.NewLine);

            }
            sw.Write(sb.ToString());
            sw.Flush();
            sw.Close();

        }

        private System.Data.DataTable CsvToDataTable(string path)
        {
            DataTable dt = new DataTable();
            FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read);
            StreamReader sr = new StreamReader(fs, System.Text.Encoding.Default);
            //记录每次读取的一行记录
            string strLine = "";
            //记录每行记录中的各字段内容
            string[] aryLine;
            //标示列数
            int columnCount = 0;
            //标示是否是读取的第一行
            bool IsFirst = true;
            //逐行读取CSV中的数据
            while ((strLine = sr.ReadLine()) != null)
            {
                aryLine = strLine.Split(',');
                if (IsFirst == true)
                {
                    IsFirst = false;
                    columnCount = aryLine.Length;
                    for (int i = 0; i < columnCount; i++)
                    {
                        DataColumn dc = new DataColumn(aryLine[i]);
                        dt.Columns.Add(dc);
                    }
                }
                else
                {
                    DataRow dr = dt.NewRow();
                    for (int j = 0; j < columnCount; j++)
                    {
                        dr[j] = aryLine[j];
                    }
                    dt.Rows.Add(dr);
                }
            }
            sr.Close();
            fs.Close();
            return dt;
        }
        private string Translate(string sent)
        {
            return sent;
        }



    }
}
