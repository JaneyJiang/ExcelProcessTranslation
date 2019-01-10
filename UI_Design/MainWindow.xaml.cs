using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace UI_Design
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void SimilarityWindow(object sender, RoutedEventArgs e)
        {
            Window similarity = new Similarity();
            //similarity.Show();
            this.Hide();
            Nullable<bool>  d = similarity.ShowDialog();
            if (d==false)
                this.Show();
        }

        private void MergeWindow(object sender, RoutedEventArgs e)
        {
            Window merge = new Merge();
            //merge.Show();
            this.Hide();
            Nullable<bool> d = merge.ShowDialog();
            if (d == false)
                this.Show();
        }
    }
}
