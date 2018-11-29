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
        //第一块功能区的私有变量
        private List<System.Data.DataTable> tablelist = null;
        // private string[] filepaths;//当需要打开多个文件获得多个路径的时候可以用这个
        private string filepath;
        private string[] sheetnames;
        private double simRate = 0.85;//相似度大小
        private double lenRate = 0.65;//句子长度比
        private Dictionary<string, string> dict = Translation.TranslationDict();

        //第二块功能区的私有变量
        private string[] mergeFiles;


        public MainWindow()
        {
            InitializeComponent();
            tablelist = new List<System.Data.DataTable>();
        }

        private void btnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            /*OpenFileDialog fileDialog = new OpenFileDialog();
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
            }*/
            bool multiselect = false;
            string filter = "Excel Files|*.xlsx;";
            string defaultExt = ".xlsx";
            filepath = OpenFile(multiselect, filter, defaultExt);
            if (filepath == null)
            {
                txbText.Text = "open no file";
            }
            else {
                txbText.Text = filepath;
            }
            if (filepath != null)
            {
                ConnectExcel connExcel = new ConnectExcel(filepath);
                sheetnames = connExcel.GetSheetsNames();
                tablelist = connExcel.GetTableList();
                cbSelect_set(sender, e);
            }
           
        }
        private string OpenFile(bool multiselect, string filter, string defaultExt)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Multiselect = multiselect;
            fileDialog.Filter = filter;
            fileDialog.DefaultExt = defaultExt;
            Nullable<bool> dialogOk = fileDialog.ShowDialog();
            if(dialogOk== true)
            {
                return fileDialog.FileName;
            }
            else
            {
                return null;
            }
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
                    List<string> delet = MakeTables((string)cbSelectLabel.SelectedValue, (string)cbSelect.SelectedValue);
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

        //创建一个空表格结构，其中，表格内包括 id， labelCol，srcCol，strLen,其中id是对记个数的从1开始，strLen是对srcCol的长度来记长度的。
        private List<string> MakeTables(string labelCol, string sourceCol)
        {
            List<string> exclueSheet = new List<string>();
            //create a DataTable
            string[] tb_cols = { "id", labelCol, sourceCol, "srcLen" };
            string[] tb_types = { "System.Int32", "System.String", "System.String", "System.Int32" };
            System.Data.DataTable tb = CsvTable.CreatEmptyTable("all",tb_cols, tb_types);
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
            string[] sim_cols = { "id", labelCol, sourceCol, "srcLen", "group", "transTo" };
            string[] sim_types = { "System.Int32", "System.String", "System.String", "System.Int32", "System.Int32", "System.String" };
            DataTable simTb = CsvTable.CreatEmptyTable("simData",sim_cols, sim_types);
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
                    if (Levenshtein.Jaro_Winkler((string)tb.Rows[i][sourceCol], (string)tb.Rows[j][sourceCol]) > simRate)
                    {
                        cmplist.Add(j);
                        RowDataClass rd = new RowDataClass(tb.Rows[j], -1);
                        dataRow = rd.getRow(simTb.NewRow(), RowFormat.SIMALL);
                        simTb.Rows.Add(dataRow);
                        groupNum += 1;
                        count += 1;
                    }
                    else {
                        if (Convert.ToSingle(tb.Rows[i]["srcLen"]) / Convert.ToSingle(tb.Rows[j]["srcLen"]) < lenRate)
                            break;
                    }
                }
                simTb.Rows[groupRow]["group"] = groupNum;


            }
            CsvTable.DataTableToCsv(simTb, "./" + sheetnames[0] + "_sim.csv");

            SplitSimTb(simTb);

            /*DataTable group = CsvToDataTable(sheetnames[0] + "_group.csv");
            DataTable single = CsvToDataTable(sheetnames[0] + "_single.csv");
           // simTb= GetTogether(simTb, group, single);
            simTb = CsvTable.MergeTable(group, single, simTb, "transTo");
            CsvTable.DataTableToCsv(simTb, "./" + sheetnames[0] + "_sim.csv");
            simTb = FollowIndex(simTb);
            CsvTable.DataTableToCsv(simTb, "./" + sheetnames[0] + "_trans.csv");*/

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

        private void SplitSimTb(DataTable simTb)//把sim表拆分成single和group两张表，供翻译补齐翻译内容
        {
            string[] sig_cols = { "id", "label", "src", "transTo" };
            string[] group_cols = { "id", "label", "src", "group","transTo" };
            string[] sig_types = { "System.Int32", "System.String", "System.String" ,"System.String" };
            string[] group_types = { "System.Int32", "System.String", "System.String", "System.Int32","System.String" };
            DataTable single = CsvTable.CreatEmptyTable("single",sig_cols, sig_types);
            DataTable group = CsvTable.CreatEmptyTable("group",group_cols,group_types);
            foreach (DataRow dr in simTb.Rows)
            {
                DataRow row;
                int tag = (int)dr["group"];
                if ( tag >=0 )
                {
                    if (tag == 0)
                    {
                        RowDataClass rdc = new RowDataClass(dr);
                        row = rdc.getRow(single.NewRow(), RowFormat.SINGLE,dict);
                        single.Rows.Add(row);
                    }
                    else
                    {
                        RowDataClass rdc = new RowDataClass(dr);
                        row = rdc.getRow(group.NewRow(), RowFormat.GROUP,dict);
                        group.Rows.Add(row);
                    }
                }
               
            }
            CsvTable.DataTableToCsv(group, "./" + sheetnames[0] + "_group.csv");
            CsvTable.DataTableToCsv(single, "./" + sheetnames[0] + "_single.csv");
        }
 
        private string Translate(string sent)
        {
            return sent;
        }

        private void btnOpenMergeFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Multiselect = true;
            fileDialog.Filter = "Comma-Separated File|*.csv;";
            fileDialog.DefaultExt = ".csv";
            Nullable<bool> dialogOk = fileDialog.ShowDialog();
            if (dialogOk == true)
            {
                mergeFileNames.Text = null;
                //filePath = fileDialog.FileNames;
                foreach (string name in fileDialog.FileNames)
                    mergeFileNames.Text += name + "\n";
                mergeFiles = fileDialog.FileNames;
            }
            else
            {
                mergeFiles = null;
                txbText.Text = "open no file";
            }
            if (mergeFiles!=null && mergeFiles.Length < 2)
            {
                throw new Exception("merge need at least two files");
            }
        }

        private void btnMergeFile_Click(object sender, RoutedEventArgs e)
        {
            DataTable group=null;
            DataTable single=null;
            DataTable similarity=null;
            //会有两种情况，一种延用当前的表格，一种新输入文件
            if (mergeFiles == null)
            {//如果不额外输入文件，即延用当前产生的表格
                if (sheetnames != null)
                {
                    //MergeAll(sheetnames[0] + "_group.csv", sheetnames[0] + "_single.csv", sheetnames[0] + "_sim.csv");
                    group = CsvTable.CsvToDataTable(sheetnames[0] + "_group.csv");
                    single = CsvTable.CsvToDataTable(sheetnames[0] + "_single.csv");
                    similarity = CsvTable.CsvToDataTable(sheetnames[0] + "_sim.csv");
                    // simTb= GetTogether(simTb, group, single);
                    similarity = CsvTable.MergeTable(group, single, similarity, "transTo");
                    CsvTable.DataTableToCsv(similarity, "./" + sheetnames[0] + "_sim.csv");
                    mergeFileNames.Text +=  "Merge Finished!\n";
                }
                else {
                    mergeFileNames.Text += "\n" + "Merge Nothing!";
                }
            }
            else
            {//如果额外输入文件
                string simName = null;
                foreach (string name in mergeFiles)
                {
                    string[] sArray = name.Split('_');
                    string tailname = sArray[sArray.Length - 1];
                    switch (tailname)
                    {
                        case "group.csv":
                            group = CsvTable.CsvToDataTable(name);
                            break;
                        case "single.csv":
                            single = CsvTable.CsvToDataTable(name);
                            break;
                        case "sim.csv":
                            similarity = CsvTable.CsvToDataTable(name);
                            simName = name;
                            break;
                        default:
                            break;
                    }
                }
                try
                {
                    if (mergeFiles.Length == 3)
                    {
                        CsvTable.MergeTable(group, single, similarity, "transTo");
                    }
                    else
                    {
                        if (group == null)
                            CsvTable.MergeSingleTable(single, similarity, "transTo");
                        else
                            CsvTable.MergeGroupTable(group, similarity, "transTo");
                    }
                }
                catch (Exception exception)
                {
                    MessageBox.Show("MergeTable: Check your merge files' format " + exception.Message);
                    return;
                }
                CsvTable.DataTableToCsv(similarity, simName);
                mergeFileNames.Text += "\n"+ "Merge Finished!";
            }

        }

        private void btnOutputFull_Click(object sender, RoutedEventArgs e)
        {
            DataTable similarity;
            DataTable full;
            if (mergeFiles == null)
            {
                if (sheetnames[0] != null)
                {
                    similarity = CsvTable.CsvToDataTable(sheetnames[0] + "_sim.csv");
                    full = FollowIndex(similarity);
                    CsvTable.DataTableToCsv(full, "./" + sheetnames[0] + "_trans.csv");
                    mergeFileNames.Text += "Output Finished!";
                }
                else
                {
                    mergeFileNames.Text += "Output Nothing!";
                }
            }
            else
            {
                foreach (string name in mergeFiles)
                {
                    string[] sArray = name.Split('_');
                    string tailname = sArray[sArray.Length - 1];
                    if (tailname == "sim.csv")
                    {
                        similarity = CsvTable.CsvToDataTable(name);
                        full = FollowIndex(similarity);
                        CsvTable.DataTableToCsv(full, name.Replace("_sim.csv", "_trans.csv"));
                        mergeFileNames.Text += "Output Finished!";
                        return;
                    }

                }
                mergeFileNames.Text += "Output Nothing!";
            }          

        }
    }
}
