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
using System.Windows.Shapes;

namespace tests
{
    /// <summary>
    /// configurationAleart.xaml の相互作用ロジック
    /// </summary>
    public partial class configurationAleart : Window
    {
        List<string[]> ConfChangeList { get; set; }
        public configurationAleart()
        {
            InitializeComponent();
            ConfChangeList = new List<string[]>();
            ConfChangeList.Add(new string[] { "DocNum", "ChangeNum", "->", "af_DocNum", "af_ChangeNum" });
            ConfChangeList.Add(new string[] { "DocNum", "ChangeNum", "->", "af_DocNum", "af_ChangeNum" });
            ConfChangeList.Add(new string[] { "DocNum", "ChangeNum", "->", "af_DocNum", "af_ChangeNum" });
            ConfChangeGrid.ItemsSource = (System.Collections.IEnumerable)ConfChangeList;

        }
    }
}
