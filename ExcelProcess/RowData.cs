using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelProcess
{
    public enum RowFormat
    {
        ORIGIN,
        SIMALL,
        SINGLE,
        GROUP,
        TRANS,

    }
    //为不同的输出数据提供格式
    public class RowDataClass
    {
        private int _id;
        private string _label;
        private string _src;
        private int _srcLen;
        private int _group;
        private string _tranTo;
        public RowDataClass(int id, string label, string src)
        {
            _id = id;
            _label = label;
            _src = src;
            _srcLen = src.Length;
        }
        public RowDataClass(int id, string label, string src, int group)
        {
            _id = id;
            _label = label;
            _src = src;
            _srcLen = src.Length;
            _group = group;
        }
        public RowDataClass(DataRow dr, int group)
        {
            object[] rowArray = dr.ItemArray;
            _id = (int)rowArray[0];
            _label = (string)rowArray[1];
            _src = (string)rowArray[2];
            _srcLen = _src.Length;
            _group = group;
        }
        public RowDataClass(DataRow dr)
        {
            object[] rowArray = dr.ItemArray;
            _id = (int)rowArray[0];
            _label = (string)rowArray[1];
            _src = (string)rowArray[2];
            if (rowArray.Length > 4)
            {
                _group = (int)rowArray[4];
            }
        }
        public void setValue(int id, string label, string src, string trans)
        {
            _id = id;
            _label = label;
            _src = src;
            _tranTo = trans;
        }


        //place to add translation
        public DataRow getRow(DataRow newRow,RowFormat k)
        {
            DataRow row;
            row = newRow;
            row[0] = _id;
            row[1] = _label;
            row[2] = _src;
            switch (k)
            {
                case RowFormat.ORIGIN://这是一开始获得表格中数据的元素
                    row[3] = _srcLen;
                    break;
                case RowFormat.SINGLE://这是保存没有相似组的字符串
                    row[3] = _src;////如果有可以翻译的软件，可以直接把这个trans(src)
                    break;
                case RowFormat.GROUP://这是保存有相似度的主字符串
                    row[3] = _group;
                    row[4] = _src;//如果有可以翻译的软件，可以直接把这个trans(src)
                    break;
                case RowFormat.SIMALL://这是对原表格中数据元素进行相似度计算和匹配。
                    row[3] = _srcLen;
                    row[4] = _group;
                    break;
                case RowFormat.TRANS://这是最后输出的有数据的单项的表格。
                    row[3] = _tranTo;
                    break;
                default:
                    break;                  
            }
            return row;
        }
    }
}
