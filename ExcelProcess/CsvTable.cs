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
    }
}
