using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ExcelProcess
{
    public class CsvTable
    {
        private static DataColumn newCol(string type, string colname)
        {
            System.Data.DataColumn col = new System.Data.DataColumn();
            col.DataType = System.Type.GetType(type);
            col.ColumnName = colname;
            return col;
        }
        //CreatEmptyTable:根据给定的表名和列名，创建空的Table。给定的名字中，第一个是表格名，第二个开始就是列名
        public static DataTable CreatEmptyTable(string tableName, string[] names, string[] types)
        {
            DataTable tb = new DataTable(tableName);
            for (int i = 0; i < names.Length; i++)
            {
                DataColumn col = newCol(types[i], (string)names[i]);
                tb.Columns.Add(col);
                //添加主键id
                if ((string)names[i] == "id")
                {
                    tb.PrimaryKey = new DataColumn[] { col };
                }
            }
            return tb;
        }
        public static DataTable CsvToDataTable(string path)
        {
            Dictionary<string, string> typeDict = new Dictionary<string, string> { { "id", "System.Int32" }, { "strLen", "System.Int32" }, { "group", "System.Int32" } };
            DataTable dt = new DataTable();
            FileStream fs = null;
            StreamReader sr = null;
            try
            {
                fs = new FileStream(path, FileMode.Open, FileAccess.Read);
                sr = new StreamReader(fs, System.Text.Encoding.Default);
            }
            catch (Exception e)
            {
                MessageBox.Show("CsvToDataTable: " + path + " " + e.Message);
                return null;
            }

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
                if (strLine[strLine.Length - 1] == ',')
                    strLine = strLine.Substring(0, strLine.Length - 1);
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
                        else
                        {
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
        public static void DataTableToCsv(DataTable inTable, string fileName)
        {
            StreamWriter sw = null;
            try
            {
                 sw = new StreamWriter(fileName, false, Encoding.GetEncoding("utf-8"));
            }
            catch (Exception e)
            {
                MessageBox.Show("DataTableToCsv: "+ fileName+ " "+e.Message);
            }

            StringBuilder sb = new StringBuilder();
            foreach (DataColumn col in inTable.Columns)
            {
                sb.Append(col.ColumnName + ",");
            }

            sb.Append(Environment.NewLine);

            foreach (DataRow dr in inTable.Rows)
            {
                foreach (DataColumn dc in inTable.Columns)
                {
                    sb.Append(dr[dc.ColumnName].ToString() + ",");
                }
                sb.Append(Environment.NewLine);
            }

            sw.Write(sb.ToString());
            sw.Flush();
            sw.Close();
        }

        //根据group和single两张表中，有组别的主句和没有组别的单句的翻译内容，整合到原来的相似度sim表中带翻译项, mergeCol表示需要被填入的列名。
        public static DataTable MergeTable(DataTable group, DataTable single, DataTable sim, string mergeCol)
        {
            DataRowCollection groupRows = group.Rows;
            DataRowCollection singleRows = single.Rows;
            DataRowCollection simRows = sim.Rows;
            int singleRowCount = 0;
            int groupRowCount = 0;
            int groupRowValueId = 0;
            int singleRowValueId = 0;
            int groupCount = -1;
            bool setSingleValue = true;
            bool setGroupValue = true;
            foreach (DataRow simRow in simRows)
            {
                int simRowId = (int)simRow[0];
                if (groupRowCount < group.Rows.Count && setGroupValue)
                {
                    groupRowValueId = (int)groupRows[groupRowCount][0];
                    setGroupValue = false;
                }
                if (singleRowCount < single.Rows.Count && setSingleValue)
                {
                    singleRowValueId = (int)singleRows[singleRowCount][0];
                    setSingleValue = false;
                }
                if (simRowId == singleRowValueId)
                {
                    CopyCell(simRow, mergeCol, singleRows[singleRowCount][mergeCol]);

                    singleRowCount++;
                    setSingleValue = true;
                }
                else
                {
                    if (simRowId == groupRowValueId || groupCount > 0)
                    {
                        CopyCell(simRow, mergeCol, groupRows[groupRowCount][mergeCol]);
                        if (groupCount < 0)
                        {
                            groupCount = (int)groupRows[groupRowCount]["group"];
                        }
                        else
                        {
                            groupCount--;
         
                        }
                        if (groupCount == 0)
                        {
                            groupRowCount++;
                            groupCount--;
                            setGroupValue = true;
                        }
                    }
                }
            }
            return sim;
        }
        public static DataTable MergeGroupTable(DataTable group,DataTable sim, string mergeCol)
        {
            DataRowCollection groupRows = group.Rows;
            DataRowCollection simRows = sim.Rows;
            int groupRowCount = 0;
            int groupRowValueId = 0;
            int groupCount = -1;
            bool setNewValue = true;
            foreach (DataRow simRow in simRows)
            {
                int simRowId = (int)simRow[0];
                if (groupRowCount < group.Rows.Count && setNewValue)
                {
                    groupRowValueId = (int)groupRows[groupRowCount][0];
                    setNewValue = false;
                }

                if (simRowId == groupRowValueId || groupCount > 0)
                {
                    CopyCell(simRow, mergeCol, groupRows[groupRowCount][mergeCol]);

                    if (groupCount < 0)
                    {
                        groupCount = (int)groupRows[groupRowCount]["group"];
                    }
                    else
                    {
                        groupCount--;
                    }
                }

                if (groupCount == 0)
                {
                    groupRowCount++;
                    groupCount--;
                    setNewValue = true;
                }

            }
            return sim;
        }
        public static DataTable MergeSingleTable(DataTable single, DataTable sim, string mergeCol)
        {
            DataRowCollection singleRows = single.Rows;
            DataRowCollection simRows = sim.Rows;
            int singleRowCount = 0;
            int singleRowValueId = 0;
            bool setNewValue = true;
            foreach (DataRow simRow in simRows)
            {
                int simRowId = (int)simRow[0];
                if (singleRowCount < single.Rows.Count && setNewValue)
                {
                    singleRowValueId = (int)singleRows[singleRowCount][0];
                    setNewValue = false;
                }
                if (simRowId == singleRowValueId)
                {
                    CopyCell(simRow, mergeCol, singleRows[singleRowCount][mergeCol]);
                    singleRowCount++;
                    if (singleRowCount >= single.Rows.Count)
                        break;
                    else
                    {
                        setNewValue = true;
                    }       

                }
                
            }
            return sim;
        }
        public static void CopyCell(DataRow dr, string colName, object cell)
        {
            dr.BeginEdit();
            //simRow["transTo"] = singleRows[singleRowCount]["transTo"];
            dr[colName] = cell;
            dr.EndEdit();
        }
    }
}
