using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SortTool
{
    class ConnectExcel//using OLEDB to read xlsx data using database
    {
        private List<DataTable> _tablelist = null;

        private readonly string _connParser = "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1;MAXSCANROWS=0';";
        private string _connStr = null;
        private string _excelFilePath = null;
        private List<string> _sheetNames = null;

        public ConnectExcel(string excelFilePath)
        {
            _tablelist = new List<DataTable>();
            _excelFilePath = excelFilePath;
            _sheetNames = new List<string>();
            SetConnStr();
            ProcessSchemaTable();
        }
        public void SetConnStr()
        {
            //获取文件扩展名来获得不同的连接设置
            string strExtension = System.IO.Path.GetExtension(_excelFilePath);
            //Excel与Excel的连接
            //HDR=Yes,这代表第一行是标题，不作为数据使用
            //IMEx0:写入，1：读取，2：读取 写入
            switch (strExtension)
            {
                case ".xlsx":
                    _connStr = _connParser;
                    break;
                default:
                    break;
            }
            if (_connStr == null)
                throw new Exception("input file type error, using *.xlsx");
            else
                _connStr += "Data Source=" + _excelFilePath;
        }
        public string[] GetSheetsNames()
        {
            return _sheetNames.ToArray() ;
        }

        public List<DataTable> GetTableList()
        {
            return _tablelist;
        }
        private void ProcessSchemaTable()
        {
            using (OleDbConnection connection = new
                       OleDbConnection(_connStr))
            {
                connection.Open();
                DataTable schemaTable = connection.GetOleDbSchemaTable(
                    OleDbSchemaGuid.Tables,
                    new object[] { null, null, null, "TABLE" });
                if (schemaTable == null)
                    throw new Exception("Excel is empty");

                foreach (System.Data.DataRow dr in schemaTable.Rows)
                {
                    string sheetname = dr["TABLE_NAME"].ToString();//.Trim();                   
                    string sql = "SELECT * FROM [" + sheetname + "]";
                    System.Data.DataSet ds = new DataSet();
                    ds.Clear();
                    OleDbDataAdapter data = new OleDbDataAdapter(sql, _connStr);
                    data.Fill(ds);
                    ds.Tables[0].TableName = sheetname;
                    _tablelist.Add(ds.Tables[0]);
                    _sheetNames.Add(sheetname);
                    //Console.WriteLine(tablelist[i].TableName);
                    // Console.WriteLine(i);
                }
            }
        }
    }
}
