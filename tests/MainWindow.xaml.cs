using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

// Window namespaces
using TimerFunc.tests;
using tests.ExcelEditor;
using tests.OneclickDatagridEdit;

namespace tests
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

        }

        private void Button_Click_TimerWindow(object sender, RoutedEventArgs e)
        {
            Window timerWindow = new TimerWindow();
            timerWindow.Show();
        }

        private void Button_Click_ExcelEditor(object sender, RoutedEventArgs e)
        {
            ExcelEditor.ExcelEditor excelEditor = new ExcelEditor.ExcelEditor();
        }
        private void Button_Click_OneclickDatagridEdit(object sender, RoutedEventArgs e)
        {
            Window newwindow = new OneclickDatagridEdit_Window();
            newwindow.Show();
        }

        private void Button_Click_Logging(object sender, RoutedEventArgs e)
        {

        }
    }
}