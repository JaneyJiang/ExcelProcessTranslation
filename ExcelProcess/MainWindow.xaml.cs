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
        // private string[] filepaths;//当需要打开多个文件获得多个路径的时候可以用这个
        private string filepath;
        private string[] sheetnames;
        public MainWindow()
        {
            InitializeComponent();
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
                filepath = null;
                txbText.Text = "open no file";
            }
            if (filepath != null)
            {
                ConnectExcel connExcel = new ConnectExcel(filepath);
                sheetnames = connExcel.GetSheetsNames();
                tablelist = connExcel.GetTableList();
            }
           
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

        //CreatEmptyTable:根据给定的表名和列名，创建空的Table。给定的名字中，第一个是表格名，第二个开始就是列名
        private DataTable CreatEmptyTable(string[] names, string[] types)
        {
            int i = 0;
            DataTable tb = new DataTable((string)names[i]);
            for (i = 1; i < names.Length; i++)
            {
                DataColumn col = newCol(types[i - 1], (string)names[i]);
                tb.Columns.Add(col);
                //添加主键id
                if ((string)names[i] == "id")
                {
                    tb.PrimaryKey = new DataColumn[] { col };
                }
            }
            return tb;
        }
        System.Data.DataColumn newCol(string type, string colname)
        {
            System.Data.DataColumn col = new System.Data.DataColumn();
            col.DataType = System.Type.GetType(type);
            col.ColumnName = colname;
            return col;
        }

        //创建一个空表格结构，其中，表格内包括 id， labelCol，srcCol，strLen,其中id是对记个数的从1开始，strLen是对srcCol的长度来记长度的。
        private List<string> MakeTable(string labelCol, string sourceCol)
        {
            List<string> exclueSheet = new List<string>();
            //create a DataTable
            string[] tb_names = { "all", "id", labelCol, sourceCol, "srcLen" };
            string[] tb_types = { "System.Int32", "System.String", "System.String", "System.Int32" };
            System.Data.DataTable tb = CreatEmptyTable(tb_names, tb_types);
            DataRow row;
            string preProcess = null;
            int t = 1;//if you have only one sheet than it matches your index.
            //create rows and set the values.for several tables
            for (int j = 0; j < tablelist.Count; j++)
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
                        if (preProcess.Length == 0)
                        {
                            t++;
                            continue;
                        }
                        RowDataClass rdc = new RowDataClass(t + 1, (string)table.Rows[i][labelCol], preProcess);
                        row = rdc.getRow(tb.NewRow(), RowFormat.ORIGIN);
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
            CsvTable.DataTableToCsv(tb, "./" + sheetnames[0] + "origin.csv");

            //according to the whole table and string lenthg information to generate simiality table.
            //DataTable simTb = tb.Clone();//复制原表的结构
            string[] sim_names = { "simData", "id", labelCol, sourceCol, "srcLen", "group", "transTo" };
            string[] sim_types = { "System.Int32", "System.String", "System.String", "System.Int32", "System.Int32", "System.String" };
            DataTable simTb = CreatEmptyTable(sim_names, sim_types);
            Hashtable simpair = new Hashtable();
            HashSet<int> cmplist = new HashSet<int>();
            int count = 0;
            for (int i = 0; i < tb.Rows.Count; i++)
            {
                int groupRow = 0;
                int groupNum = 0;
                if (cmplist.Contains(i))
                    continue;
                cmplist.Add(i);
                DataRow dataRow;
                RowDataClass rdc = new RowDataClass(tb.Rows[i], 0);
                dataRow = rdc.getRow(simTb.NewRow(), RowFormat.SIMALL);
                simTb.Rows.Add(dataRow);
                groupRow = count;
                count += 1;
                for (int j = i + 1; j < tb.Rows.Count; j++)
                {
                    if (cmplist.Contains(j))
                        continue;
                    if (Levenshtein.Jaro_Winkler((string)tb.Rows[i][sourceCol], (string)tb.Rows[j][sourceCol]) > 0.8)
                    {
                        cmplist.Add(j);
                        RowDataClass rd = new RowDataClass(tb.Rows[j], -1);
                        dataRow = rd.getRow(simTb.NewRow(), RowFormat.SIMALL);
                        simTb.Rows.Add(dataRow);
                        groupNum += 1;
                        count += 1;
                    }
                    else {
                        if (Convert.ToSingle(tb.Rows[i]["srcLen"]) / Convert.ToSingle(tb.Rows[j]["srcLen"]) < 0.65)
                            break;
                    }
                }
                simTb.Rows[groupRow]["group"] = groupNum;


            }
            CsvTable.DataTableToCsv(simTb, "./" + sheetnames[0] + "_sim.csv");

            SplitSimTb(simTb);

            DataTable group = CsvToDataTable(sheetnames[0] + "_group.csv");
            DataTable single = CsvToDataTable(sheetnames[0] + "_single.csv");
            simTb= GetTogether(simTb, group, single);
            CsvTable.DataTableToCsv(simTb, "./" + sheetnames[0] + "_sim.csv");
            simTb = FollowIndex(simTb);
            CsvTable.DataTableToCsv(simTb, "./" + sheetnames[0] + "_trans.csv");

            return exclueSheet;
        }
        //把simTb表按照id不能少的重新建立一个表，不存在的id就用空行填充。
        private DataTable FollowIndex(DataTable simTb)
        {
            simTb.DefaultView.Sort = simTb.Columns[0] + " ASC";
            simTb = simTb.DefaultView.ToTable();
            DataTable tb = simTb.Clone();
            int simIdx = 0;
            int startIdx = 2;
            while(simIdx < simTb.Rows.Count)
            {
                if ((int)simTb.Rows[simIdx][0] == startIdx)
                {
                    tb.ImportRow(simTb.Rows[simIdx]);
                    simIdx++;
                }
                else {
                    DataRow dr = tb.NewRow();
                    dr[0] = startIdx;
                    tb.Rows.Add(dr);   
                }
                startIdx++;

            }
                return tb;
        }

        //把翻译好的Group和single整合到simTb表中。
        private DataTable GetTogether(DataTable simTb, DataTable group, DataTable single)
        {
            DataRowCollection groupRows = group.Rows;
            DataRowCollection singleRows = single.Rows;
            DataRowCollection simRows = simTb.Rows;
            int singleIdx = 0;
            int groupIdx = 0;
            int groupId = 0;
            int singleId = 0;
            int groupcount = -1;
            foreach(DataRow simRow in simRows)
            {
                int id = (int)simRow[0];
                if(groupIdx < group.Rows.Count)
                {
                    groupId = (int)groupRows[groupIdx][0];
                }
                if (singleIdx < single.Rows.Count)
                {
                    singleId = (int)singleRows[singleIdx][0];
                }
                if (id == singleId)
                {
                    /*simRow.BeginEdit();
                    simRow["transTo"] = singleRows[singleIdx]["transTo"];
                    simRow.EndEdit();*/
                    CopyCell(simRow, "transTo", singleRows[singleIdx]["transTo"]);

                    singleIdx++;
                }
                else {
                    if (id == groupId)
                    {
                        /*simRow.BeginEdit();
                         simRow["transTo"] = groupRows[groupIdx]["transTo"];
                        simRow.EndEdit();*/
                        CopyCell(simRow, "transTo", groupRows[groupIdx]["transTo"]);
                        groupcount = (int)groupRows[groupIdx]["group"];
                    }
                    else {
                        if (groupcount > 0)
                        {
                            /*simRow.BeginEdit();
                            simRow["transTo"] = groupRows[groupIdx]["transTo"];
                            simRow.EndEdit();*/
                            CopyCell(simRow, "transTo", groupRows[groupIdx]["transTo"]);
                            groupcount--;
                        }
                        if (groupcount == 0)
                        {
                            groupIdx++;
                            groupcount--;
                        }
                    }
                }

            }
            return simTb;
        }

        public static void CopyCell(DataRow dr, string colName, object cell)
        {
            dr.BeginEdit();
            //simRow["transTo"] = singleRows[singleRowCount]["transTo"];
            dr[colName] = cell;
            dr.EndEdit();
        }
        private void SplitSimTb(DataTable simTb)//把sim表拆分成single和group两张表，供翻译补齐翻译内容
        {
            string[] sig_names = { "single","id", "label", "src", "transTo" };
            string[] group_names = { "group","id", "label", "src", "group","transTo" };
            string[] sig_types = { "System.Int32", "System.String", "System.String" ,"System.String" };
            string[] group_types = { "System.Int32", "System.String", "System.String", "System.Int32","System.String" };
            DataTable single = CreatEmptyTable(sig_names, sig_types);
            DataTable group = CreatEmptyTable(group_names,group_types);
            foreach (DataRow dr in simTb.Rows)
            {
                DataRow row;
                int tag = (int)dr["group"];
                if ( tag >=0 )
                {
                    if (tag == 0)
                    {
                        RowDataClass rdc = new RowDataClass(dr);
                        row = rdc.getRow(single.NewRow(), RowFormat.SINGLE);
                        single.Rows.Add(row);
                    }
                    else
                    {
                        RowDataClass rdc = new RowDataClass(dr);
                        row = rdc.getRow(group.NewRow(), RowFormat.GROUP);
                        group.Rows.Add(row);
                    }
                }
               
            }
            CsvTable.DataTableToCsv(group, "./" + sheetnames[0] + "_group.csv");
            CsvTable.DataTableToCsv(single, "./" + sheetnames[0] + "_single.csv");
        }
 
        private System.Data.DataTable CsvToDataTable(string path)
        {
            Dictionary<string, string> typeDict = new Dictionary<string, string>{ { "id","System.Int32"},{ "strLen","System.Int32"},{"group","System.Int32"} };
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
                        DataColumn dc;
                        if (typeDict.ContainsKey(aryLine[i]))
                        {
                            dc = newCol(typeDict[aryLine[i]], aryLine[i]);
                        }
                        else {
                            dc = newCol("System.String", aryLine[i]);
                        }
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
