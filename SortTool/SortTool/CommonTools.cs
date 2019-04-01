using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Runtime.InteropServices;
//using System.Windows.Forms;
using Microsoft.Win32;
using System.Windows;

namespace SortTool
{
    static class CommonTools
    {
        
        static public DataTable DictToDataTable<T>(Dictionary<T, T> dict)
        {
            DataTable dt = new DataTable();
            DataColumn dc1 = new DataColumn("from");
            DataColumn dc2 = new DataColumn("to");
            dt.Columns.Add(dc1);
            dt.Columns.Add(dc2);
            foreach (KeyValuePair<T, T> item in dict)
            {
                dt.Rows.Add(item.Key, item.Value);
            }
            return dt;
        }

        //取表的前两列为字典的key，value,所以table的列不能小于2.
        static public Dictionary<string,string> DataTableToDict(DataTable dt)
        {
            if (dt.Columns.Count < 2)
            {
                MessageBox.Show("字典异常");
                return null;
            }
            //无法处理具有相同键的问题，一旦出现就会报错
            //return dt.Rows.Cast<DataRow>().ToDictionary(x => x[0].ToString(), x => x[1].ToString());
            return dt.Rows.Cast<DataRow>().ToLookup(x => x[0].ToString(), x => x[1].ToString())
                .ToDictionary(t => t.Key, t => t.First());//选取重复键的第一个,也可以t.Last()表示选的是最后一个。
        }

       /* static class ToDictionaryExtentions
        {
            public static IDictionary<TKey, TValue> ToDictionaryEx<TElement, TKey, TValue>(
                this IEnumerable<TElement> source,
                Func<TElement, TKey> keyGetter,
                Func<TElement, TValue> valueGetter)
            {
                IDictionary<TKey, TValue> dict = new Dictionary<TKey, TValue>();
                foreach (var e in source)
                {
                    var key = keyGetter(e);
                    if (dict.ContainsKey(key))
                    {
                        continue;
                    }

                    dict.Add(key, valueGetter(e));
                }
                return dict;
            }
        }*/

        static string g_WorkingFolder = "";
        static public Encoding gb = Encoding.GetEncoding("gb2312");

        static public void SaveDataTable2Excel(DataTable dt, string fileHint)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.InitialDirectory = g_WorkingFolder;// Environment.GetFolderPath(Environment.SpecialFolder.Desktop);// "d:\\";
            saveFileDialog1.Filter = "Excel(*.xlsx) | *.xlsx";//"CSV files (*.csv)|*.csv";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.RestoreDirectory = true;
            saveFileDialog1.Title = fileHint;
            Nullable<bool> dialogOk = saveFileDialog1.ShowDialog();
            if (dialogOk == true && saveFileDialog1.FileName.Length > 0)
            {

                ExportExcel(dt, saveFileDialog1.FileName);
                MessageBox.Show("Save Succeed！", "Export to Excel");
            }

        }
        static public void ExportExcel(DataTable dt, string savePath)
        {
            if (dt != null)
            {
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

                if (excel == null)
                {
                    return;
                }
                excel.Visible = false;
                Microsoft.Office.Interop.Excel.Workbooks workbooks = excel.Workbooks;

                Microsoft.Office.Interop.Excel.Workbook workbook = workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
                Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];

//                worksheet.Name = SheetName;
                try
                {
                    Microsoft.Office.Interop.Excel.Range range;

                    int rowIndex = 1;
                    int colIndex = 1;

                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        worksheet.Cells[rowIndex, colIndex + i] = dt.Columns[i].ColumnName;

                        range = worksheet.Cells[rowIndex, colIndex + i];
                        range.Interior.ColorIndex = 33;
                        //range.ColumnWidth = 36;//设置列宽
                        range.Font.Bold = true;
                        range.Font.Color = 0;
                        range.Font.Name = "Arial";
                        range.Font.Size = 12;
                        range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        range.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    }

                    rowIndex++;
                    int row_offset = -1;
                    int col_offset = -1;
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        int colorIdx = 0;
                        if (dt.Rows[i][0].ToString() == "[U]") colorIdx = 22;
                        else if (dt.Rows[i][0].ToString() == "[M]") colorIdx = 4;

                        if (colorIdx == 0)
                        {
                            row_offset++;
                            col_offset = 0;
                        }
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            if (j == 0 && colorIdx != 0) continue;
                            range = worksheet.Cells[rowIndex + row_offset, colIndex + col_offset];
                            range.Interior.ColorIndex = colorIdx;
                            worksheet.Cells[rowIndex + row_offset, colIndex + col_offset] = dt.Rows[i][j].ToString();
                            col_offset++;

                        }
                    }

                    //设置列宽自动匹配和WrapText 的前后顺序会影响输出表格列宽，
                    //WrapText =True，在前，输出表格以之前的字符的宽度为宽度，
                    //WrapText =True，在后，输出表格会以最宽的宽度为宽度。
                    worksheet.Cells.WrapText = true;
                    worksheet.Cells.Columns.AutoFit();//列宽自动匹配
                    
                    excel.DisplayAlerts = false;
                    workbook.Saved = true;
                    FileStream file = new FileStream(savePath, FileMode.Create);
                    file.Close();
                    file.Dispose();
                    workbook.SaveCopyAs(savePath);

                }
                catch (Exception err)
                {
                    MessageBox.Show(err.Message);
                }
                finally
                {
                    workbook.Close(false, Type.Missing, Type.Missing);
                    workbooks.Close();

                    excel.Quit();

                    Marshal.ReleaseComObject(worksheet);
                    Marshal.ReleaseComObject(workbook);
                    Marshal.ReleaseComObject(workbooks);
                    Marshal.ReleaseComObject(excel);

                    worksheet = null;
                    workbook = null;
                    workbooks = null;
                    excel = null;

                    GC.Collect();
                }
            }
        }
        static public void ExportCSV(DataTable dt, string savePath)
        {

            FileStream file = new FileStream(savePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);

            StreamWriter sw = new StreamWriter(file, Encoding.GetEncoding("UTF-8"));

            StringBuilder strbu = new StringBuilder();

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                strbu.Append(dt.Columns[i].ColumnName.ToString() + ",");
            }
            strbu.Append(Environment.NewLine);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    strbu.Append(dt.Rows[i][j].ToString() + ",");
                }
                strbu.Append(Environment.NewLine);
            }

            sw.Write(strbu.ToString());
            sw.Flush();
            file.Flush();

            sw.Close();
            sw.Dispose();

            file.Close();
            file.Dispose();
        }
        static public string[] GetExcelTableList(OleDbConnection myConn)
        {
            DataTable schemaTable = myConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            List<string> tablenamelist = new List<string>();
            for ( int i = 0; i < schemaTable.Rows.Count; i++)
            {
                string sheetnametemp = schemaTable.Rows[i].ItemArray[2].ToString();
                string lastchar = sheetnametemp.Substring(sheetnametemp.Length - 1, 1);
                string lasttwochar = sheetnametemp.Substring(sheetnametemp.Length - 2, 2);
                if (lastchar == "$" || lasttwochar=="$'")
//                if (!sheetnametemp.Contains("_FilterDatabase") && !sheetnametemp.Contains("_xlnm"))
                    tablenamelist.Add(sheetnametemp);
            }
            return tablenamelist.ToArray();
        }

        static public string[] GetFilterList(DataTable dt)
        {
            List<string> filterlist = new List<string>();
            DataRow[] tableRows = dt.Select();
            for (int i = 0; i < tableRows.Count(); i++)
            {
                string tempstr = tableRows[i]["Filter"].ToString();
                bool need_append = true;
                foreach (string filterstr in filterlist)
                    if (filterstr == tempstr)
                        need_append = false;
                if (need_append)
                    filterlist.Add(tempstr);
            }
            return filterlist.ToArray();
        }

        /*static public string OpenExcelFile()
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Filter = "Excel(*.xlsx)|*.xlsx|Excel(*.xls)|*.xls";
            openFile.InitialDirectory = g_WorkingFolder;// Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            openFile.Multiselect = false;
            string filename = null;

            if (openFile.ShowDialog() == DialogResult.OK)
                filename = openFile.FileName;
            g_WorkingFolder = System.IO.Path.GetDirectoryName(filename);
            return filename;

        }*/
        static public string OpenOneFile(string fileSuffix="")
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            switch (fileSuffix)
            {
                case "xlsx":
                    fileDialog.Filter = "Excel Files|*.xlsx;";
                    fileDialog.DefaultExt = ".xlsx";
                    break;
                case "txt":
                    fileDialog.Filter = "txt Files|*.txt;";
                    fileDialog.DefaultExt = ".txt";
                    break;
                default:
                    break;
            }
            fileDialog.Multiselect = false;
            fileDialog.FilterIndex = 2; //保存对话框是否记忆上次打开的目录
            fileDialog.RestoreDirectory = true; //点了保存按钮进入

            Nullable<bool> dialogOk = fileDialog.ShowDialog();
            if (dialogOk == true)
            {
                // g_WorkingFolder = System.IO.Path.GetDirectoryName(fileDialog.FileName);
                return fileDialog.FileName;
            }
            else
            {
                return null;
            }

        }
        static public string OpenExcelFile()
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Multiselect = false;
            fileDialog.Filter = "Excel Files|*.xlsx;";
            fileDialog.FilterIndex = 2; //保存对话框是否记忆上次打开的目录
            fileDialog.RestoreDirectory = true; //点了保存按钮进入
            fileDialog.DefaultExt = ".xlsx";
            Nullable<bool> dialogOk = fileDialog.ShowDialog();
            if (dialogOk == true)
            {
                // g_WorkingFolder = System.IO.Path.GetDirectoryName(fileDialog.FileName);
                return fileDialog.FileName;
            }
            else
            {
                return null;
            }
        }
        static public  string OpenTxtFile()
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Multiselect = false;
            fileDialog.Filter = "txt Files|*.txt;";
            fileDialog.DefaultExt = ".txt";
            fileDialog.FilterIndex = 2; //保存对话框是否记忆上次打开的目录
            fileDialog.RestoreDirectory = true; //点了保存按钮进入
            Nullable<bool> dialogOk = fileDialog.ShowDialog();
            if (dialogOk == true)
            {
                // g_WorkingFolder = System.IO.Path.GetDirectoryName(fileDialog.FileName);
                return fileDialog.FileName;
            }
            else
            {
                return null;
            }
        }
        static public void SaveDataTable2TXT(Dictionary<int, int> mydic)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.InitialDirectory = g_WorkingFolder;// Environment.GetFolderPath(Environment.SpecialFolder.Desktop);// "d:\\";
            saveFileDialog1.Filter = "TEXT(*.txt) | *.txt";//"CSV files (*.csv)|*.csv";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.RestoreDirectory = true;
            Nullable<bool> dialogOk = saveFileDialog1.ShowDialog();
            if (dialogOk == true && saveFileDialog1.FileName.Length > 0)
            {

                WriteTXT(saveFileDialog1.FileName, mydic);
                MessageBox.Show("Save Succeed！", "Export to TXT");
            }

        }
        static public void WriteTXT(string path, Dictionary<int, int> mydic) //将字典写入txt
        {
            FileStream fs = new FileStream(path, FileMode.Create);
            StreamWriter sw = new StreamWriter(fs);
            //开始写入
            foreach (var d in mydic)
            {
                sw.Write(d.Key + "\t" + d.Value); //键值对写入，用逗号隔开
            }
            //清空缓冲区
            sw.Flush();
            //关闭流
            sw.Close();
            fs.Close();
        }
        public static Dictionary<int, int> ReadDict(string path) //读txt文件 返回字典
        {
            StreamReader sr = new StreamReader(path, Encoding.Default);
            String line;
            var dic = new Dictionary<int, int>();
            try
            {
                while ((line = sr.ReadLine()) != null)
                {
                    var li = line.ToString().Split(new Char[] { '\t' }, StringSplitOptions.RemoveEmptyEntries); ; //将一行用,分开成键值对
                    if (dic.ContainsKey(Convert.ToInt32(li[0])))
                    {
                        dic[Convert.ToInt32(li[0])] = Convert.ToInt32(li[1]);
                    }
                    else
                    {
                        dic.Add(Convert.ToInt32(li[0]), Convert.ToInt32(li[1]));
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return dic;
        }
    }
}
