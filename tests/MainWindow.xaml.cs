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
using tests.Logger;

// ★ILogger を使うために必要
using Microsoft.Extensions.Logging; 
using System.Collections.ObjectModel;
using System.Threading.Tasks;
using Serilog.Core;
using Microsoft.Extensions.DependencyInjection;
namespace tests
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly IServiceProvider _serviceProvider;

        //ロガー用
        private readonly ILogger<MainWindow> _logger; // このクラスのロガー
        private readonly DataCacheService _dataCacheService; // 前回の回答で作成したサービス
        public ObservableCollection<MyDataItem> DisplayedData { get; set; }



        public MainWindow(ILogger<MainWindow> logger, DataCacheService dataCacheService)
        {
            InitializeComponent();
            //ロガー用
            _logger = logger;
            _dataCacheService = dataCacheService;

            _logger.LogInformation("MainWindowが初期化されました。"); // Informationレベルのログ

            DisplayedData = new ObservableCollection<MyDataItem>();//表示するデータ
            this.DataContext = this;

            _ = LoadAndDisplayDataAsync(); // データ読み込みを非同期で開始
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
            _logger.LogTrace("ボタンがクリックされました。"); // Traceレベルのログ（非常に詳細）
            //Window newwindow = new Logger.LoggerWindow();
            Window newwindow = App.ServiceProvider.GetService<Logger.LoggerWindow>();
            newwindow.Show();
        }
        private async Task LoadAndDisplayDataAsync()
        {
            _logger.LogInformation("データのロードと表示を開始します。");

            try
            {
                // キャッシュからデータを読み込む
                var cachedData = await _dataCacheService.LoadDataCacheAsync();

                if (cachedData != null && cachedData.Count > 0)
                {
                    _logger.LogInformation("キャッシュからデータをロードしました。");
                    // UIスレッドでコレクションをクリアし、データを追加
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        DisplayedData.Clear();
                        foreach (var item in cachedData)
                        {
                            DisplayedData.Add(item);
                        }
                    });
                }
                else
                {
                    _logger.LogWarning("キャッシュが見つからないか、古いため、新しいデータを取得します。");
                    // ここでデータベースなどから実際のデータを取得する処理を呼び出す
                    var newData = await FetchDataFromDatabaseAsync(); // データベースからのデータ取得をシミュレート
                    if (newData != null)
                    {
                        _logger.LogInformation("新しいデータをデータベースから取得しました。");
                        // UIスレッドでコレクションをクリアし、データを追加
                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            DisplayedData.Clear();
                            foreach (var item in newData)
                            {
                                DisplayedData.Add(item);
                            }
                        });
                        // 取得した新しいデータをキャッシュに保存
                        await _dataCacheService.SaveDataCacheAsync(newData);
                        _logger.LogInformation("新しいデータをキャッシュに保存しました。");
                    }
                    else
                    {
                        _logger.LogError("データベースからのデータ取得に失敗しました。");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "データのロード中に予期せぬエラーが発生しました。");
                MessageBox.Show($"データのロード中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task<List<MyDataItem>> FetchDataFromDatabaseAsync()
        {
            _logger.LogDebug("データベースからデータを取得中...");
            await Task.Delay(2000); // データベースアクセスをシミュレートするために2秒待機

            // ダミーデータを返す
            return new List<MyDataItem>
            {
                new MyDataItem { Id = 1, Name = "Product A", Value = 100, LastUpdated = DateTime.Now },
                new MyDataItem { Id = 2, Name = "Product B", Value = 250, LastUpdated = DateTime.Now.AddDays(-1) },
                new MyDataItem { Id = 3, Name = "Product C", Value = 300, LastUpdated = DateTime.Now.AddHours(-5) }
            };
        }
    }
}