using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace tests.OneclickDatagridEdit
{
    /// <summary>
    /// OneclickDatagridEdit.xaml の相互作用ロジック
    /// </summary>
    public partial class OneclickDatagridEdit_Window : Window
    {
        ObservableCollection<Item> items = new ObservableCollection<Item>();
        public OneclickDatagridEdit_Window()
        {
            InitializeComponent();
            items = new ObservableCollection<Item>
            {
                new Item { data1 = "test1", data2 = "test2", data3 = "test3" },
                new Item { data1 = "test4", data2 = "test5", data3 = "test6" },
                new Item { data1 = "test7", data2 = "test8", data3 = "test9" }
            };
            oneclickdatagrid.ItemsSource = items;
        }


        private void DataGridCell_GotFocus(object sender, System.Windows.RoutedEventArgs e)
        {

            DataGridCell cell = sender as DataGridCell;
            if (cell != null && !cell.IsEditing)
            {
                // セルが編集可能で、かつ既に編集モードではない場合のみ編集を開始
                if (cell.Column is DataGridBoundColumn column)
                {
                    // カラムがDataGridBoundColumnの場合にのみ編集を開始
                    // DataGridTemplateColumnなどの場合は、編集テンプレートのコントロールに
                    // フォーカスを移すなどの追加処理が必要になることがあります。
                    oneclickdatagrid.BeginEdit(e); // DataGrid の編集を開始
                }
            }
        }
    }
    public class Item
    {
        public string? data1 { get; set; }
        public string? data2 { get; set; }
        public string? data3 { get; set; }
    }
}
