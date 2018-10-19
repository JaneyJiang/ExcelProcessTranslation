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
        public static void DataTableToCsv(DataTable inTable, string fileName)
        {
            StreamWriter sw = null;
            try
            {
                //这里的文件打开异常不需要捕获太早，当然可以通过返回bool型来对异常进行判断处理。
                sw = new StreamWriter(fileName, false, Encoding.GetEncoding("utf-8"));
            }
            catch (Exception e)
            {
                MessageBox.Show("DataTableToCsv: "+e.Message);
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
            foreach (DataRow simRow in simRows)
            {
                int simRowId = (int)simRow[0];
                if (groupRowCount < group.Rows.Count)
                {
                    groupRowValueId = (int)groupRows[groupRowCount][0];
                }
                if (singleRowCount < single.Rows.Count)
                {
                    singleRowValueId = (int)singleRows[singleRowCount][0];
                }
                if (simRowId == singleRowValueId)
                {
                    simRow.BeginEdit();
                    simRow["transTo"] = singleRows[singleRowCount]["transTo"];
                    simRow.EndEdit();

                    singleRowCount++;
                }
                else
                {

                    if (simRowId == groupRowValueId || groupCount > 0)
                    {
                        simRow.BeginEdit();
                        simRow["transTo"] = groupRows[groupRowCount]["transTo"];
                        simRow.EndEdit();
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
