using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace SortTool
{
    /// <summary>
    /// UserControlExcelTag.xaml 的交互逻辑
    /// </summary>
    public partial class UserControlExcelTag : UserControl
    {
        public UserControlExcelTag()
        {
            InitializeComponent();
        }
        public UserControlExcelTag(string name)
        {
            InitializeComponent();
            labelName = name;
        }
        public String labelName
        {
            get { return LabelName.Content.ToString(); }
            set { LabelName.Content = value; }
        }
        public String labelSelect
        {
            get { return LabelSelect.SelectedValue.ToString(); }
            set { LabelSelect.SelectedIndex = 0; }
        }
        public void Add(string[] names)
        {
            foreach (string name in names)
            {
                LabelSelect.Items.Add(name);
            }

        }
    }
}
