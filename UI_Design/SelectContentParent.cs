using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace UI_Design
{
    public class SelectContentParentForm : Form
    {
        public const int MAX_KEY_NUM = 10;
        protected DataTable SelectedTable = null;
        protected int KeyNumber;
//        protected int[] KeyIdx;
        protected List<int[]> KeyIdxList;//= new List<int[]>();
        protected int CurrentKeySet;
        protected string[] CurrentKeyTag;
        protected string[] TableColumnTag;
        protected bool need_updateFilter = false;
        protected int IgnoreStrID;

        public void SetSheetChoice(int num, int[] idx)
        {
            for (int i = 0; i < num; i++)
                KeyIdxList[CurrentKeySet][i] = idx[i];
            //    KeyIdx[i] = idx[i];
        }
        public void update_fileter(bool update, int filterstrID)
        {
            need_updateFilter = update;
            IgnoreStrID = filterstrID;

        }

        public void initContent(int key_num, string[] columnTag)
        {
            KeyNumber = key_num;
            KeyIdxList = new List<int[]>();
            SelectedTable = new DataTable();
            SelectedTable.Columns.Add("id", Type.GetType("System.Int32"));
            SelectedTable.Columns["id"].AutoIncrement = true;
            CurrentKeySet = 0;
            int[] keyidx = new int[KeyNumber];
            KeyIdxList.Add(keyidx);
            for (int i = 0; i < KeyNumber; i++)
            {
//                KeyIdx[i] = 0;
                KeyIdxList[CurrentKeySet][i] = 0;
                CurrentKeyTag[i] = "[null]";
                TableColumnTag[i] = columnTag[i];
                SelectedTable.Columns.Add(columnTag[i], Type.GetType("System.String"));
            }
            SelectedTable.Columns.Add("Source", Type.GetType("System.String"));

        }

        public string SelectContent(string excelfilename)
        {
            string debugText = "";
            string fileType = System.IO.Path.GetExtension(excelfilename);
            if (!string.IsNullOrEmpty(fileType))
            {
                bool hasTitle = false;
                using (DataSet ds = new DataSet())
                {
                    string strCon = string.Format("Provider=Microsoft.{0}.OLEDB.{1}.0;" +
                                    "Extended Properties=\"Excel {2}.0;HDR={3};IMEX=1;\";" +
                                    "data source={4};", (fileType == ".xls" ? "JET" : "ACE"),
                                    (fileType == ".xls" ? 4 : 12), (fileType == ".xls" ? 8 : 12), (hasTitle ? "Yes" : "NO"), excelfilename);

                    using (OleDbConnection myConn = new OleDbConnection(strCon))
                    {
                        myConn.Open();
                        string[] Tablename = { "" };//CommonFunction.GetExcelTableList(myConn);
                        int TableCount = Tablename.Count();
                        for (int i = 0; i < TableCount; i++)
                        {
                            ds.Clear();
                            string strCom = " SELECT * FROM [" + Tablename[i] + "]";
                            using (OleDbDataAdapter myCommand = new OleDbDataAdapter(strCom, myConn))
                            {
                                myCommand.Fill(ds);
                                if (ds != null && ds.Tables.Count > 0)
                                {
                                    DataTable TableTemp = ds.Tables[0]; ;
                                    DataRow[] Rows = TableTemp.Select();
                                    int ColumnCount = TableTemp.Columns.Count;
                                    string[] Columnname = new string[ColumnCount];
                                    string[] ColumnCap = new string[ColumnCount];

                                    for (int j = 0; j < ColumnCount; j++)
                                    {
                                        Columnname[j] = TableTemp.Columns[j].ColumnName;
                                        ColumnCap[j] = Rows[0][Columnname[j]].ToString();
                                    }
                                    string[] CurrentTabelTag = new string[KeyNumber];
                                    bool b_NeedNewSelection = false;
                                    for (int k=0; k< KeyIdxList.Count; k++)
                                    {
                                        b_NeedNewSelection = false;
                                        for (int m = 0; m < KeyNumber; m++)
                                        {
                                            CurrentTabelTag[m] = ColumnCap[KeyIdxList[k][m]];   //KeyIdx
                                            if (CurrentTabelTag[m] != CurrentKeyTag[m])
                                                b_NeedNewSelection = true;
                                        }
                                        if (!b_NeedNewSelection) {
                                            CurrentKeySet = k;
                                            break;
                                        }
                                    }
                                    if (b_NeedNewSelection)
                                    {
                                        int[] keyidx = new int[KeyNumber];
                                        KeyIdxList.Add(keyidx);
                                        CurrentKeySet = KeyIdxList.Count-1;

                                        SelectContent TerminForm = new SelectContent(KeyNumber, TableColumnTag);
                                        TerminForm.Text = "Sheet: " + Tablename[i];
                                        TerminForm.SetContent(ColumnCap);

                                        TerminForm.ShowDialog(this);
/*                                        debugText += "Sheet " + Tablename[i] + ":";// Original Tag:" + KeyIdx[0].ToString() + " " + CurrentKeyTag[0] + " Target Tag:" + KeyIdx[1].ToString() + " " + CurrentKeyTag[1] + "\r\n";
                                        for (int m = 0; m < KeyNumber; m++)
                                        { 
                                            CurrentKeyTag[m] = ColumnCap[KeyIdxList[CurrentKeySet][m]];
                                            debugText += "Column "+ KeyIdxList[CurrentKeySet][m].ToString() + ":" + CurrentKeyTag[m]+"\t";
                                        }
                                        debugText += "\r\n";
 */                                   }
//                                    else
//                                    {
//                                        debugText += "Sheet " + Tablename[i] + "have the same structure, ignore selection!\r\n";
//                                    }
                                    debugText += "Sheet " + Tablename[i] + ":";// Original Tag:" + KeyIdx[0].ToString() + " " + CurrentKeyTag[0] + " Target Tag:" + KeyIdx[1].ToString() + " " + CurrentKeyTag[1] + "\r\n";
                                    for (int m = 0; m < KeyNumber; m++)
                                    {
                                        CurrentKeyTag[m] = ColumnCap[KeyIdxList[CurrentKeySet][m]];
                                        debugText += "Column " + KeyIdxList[CurrentKeySet][m].ToString() + ":" + CurrentKeyTag[m] + "\t";
                                    }
                                    debugText += "\r\n";
                                    for (int j = 1; j < Rows.Count(); j++)
                                    {
                                        bool need_append = false;
                                        string[] tempVal = new string[KeyNumber];
                                        for (int m = 0; m < KeyNumber; m++)
                                        {
                                            tempVal[m] = Rows[j][Columnname[KeyIdxList[CurrentKeySet][m]]].ToString();
                                            if (tempVal[m] != "")
                                                need_append = true;
                                        }
                                        if (need_append)
                                        {
                                            DataRow dw = SelectedTable.NewRow();
                                            for (int m = 0; m < KeyNumber; m++)
                                                dw[TableColumnTag[m]] = tempVal[m];
                                            dw["Source"] = Tablename[i];
                                            SelectedTable.Rows.Add(dw);

                                        }

                                    }
                                }


                            }

                        }
                        myConn.Close();

                    }
                }
            }
            KeyIdxList.Clear();
            return debugText;
        }

    }
}
