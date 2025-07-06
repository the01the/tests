using System;
using System.Collections.Generic;
using System.ComponentModel;
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
using System.Windows.Threading;
using TimerFunc.tests;



//ログ用
using System.Windows;
using Microsoft.Extensions.Logging;
using System.Collections.ObjectModel; // MyDataItem のコレクション用
using System.Threading.Tasks; // 非同期処理用


namespace tests.Logger
{
    /// <summary>
    /// LoggerWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class LoggerWindow : Window
    {
        private readonly ILogger<LoggerWindow> _logger;
        private readonly DataCacheService _dataCacheService;
        public ObservableCollection<MyDataItem> DisplayedData { get; set; }

        public LoggerWindow(ILogger<LoggerWindow> logger, DataCacheService dataCacheService)
        {
            InitializeComponent();
            SetupTimer();// タイマーセットアップ
            // ロガーとデータキャッシュサービスのインスタンスを取得
            _logger = logger;
            _dataCacheService = dataCacheService;
            _logger.LogInformation("MainWindowが初期化されました。");
            // データグリッドにバインドするコレクションを初期化
            DisplayedData = new ObservableCollection<MyDataItem>();
            this.DataContext = this;
            // キャッシュデータとUIの更新を非同期で行う
            _ = LoadAndDisplayDataAsync();
        }
        private async Task LoadAndDisplayDataAsync()
        {
            _logger.LogInformation("データのロードと表示を開始します。");

            // キャッシュからデータを読み込む
            var cachedData = await _dataCacheService.LoadDataCacheAsync();

            if (cachedData != null && cachedData.Count > 0)
            {
                _logger.LogInformation("キャッシュからデータをロードしました。");
                DisplayedData.Clear();
                foreach (var item in cachedData)
                {
                    DisplayedData.Add(item);
                }
            }
            else
            {
                _logger.LogInformation("キャッシュが見つからないか、古いため、新しいデータを取得します。");
                // ここでデータベースなどから実際のデータを取得する処理を呼び出す
                var newData = await FetchDataFromDatabaseAsync(); // 仮のメソッド
                if (newData != null)
                {
                    DisplayedData.Clear();
                    foreach (var item in newData)
                    {
                        DisplayedData.Add(item);
                    }
                    // 新しいデータをキャッシュに保存
                    await _dataCacheService.SaveDataCacheAsync(newData);
                    _logger.LogInformation("新しいデータを取得し、キャッシュに保存しました。");
                }
                else
                {
                    _logger.LogWarning("データベースからのデータ取得に失敗しました。");
                }
            }
        }

        // 仮のデータベースからデータを取得するメソッド
        private async Task<List<MyDataItem>> FetchDataFromDatabaseAsync()
        {
            _logger.LogInformation("データベースからデータを取得中...");
            await Task.Delay(2000); // データベースアクセスをシミュレート
            return new List<MyDataItem>
            {
                new MyDataItem { Id = 1, Name = "Database Item 1", Value = 100, LastUpdated = DateTime.Now },
                new MyDataItem { Id = 2, Name = "Database Item 2", Value = 200, LastUpdated = DateTime.Now.AddDays(-1) },
                new MyDataItem { Id = 3, Name = "Database Item 3", Value = 300, LastUpdated = DateTime.Now.AddHours(-5) }
            };
        }
        //ロギングの場所：ボタンクリック、タイマー
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            _logger.LogDebug("ボタン1がクリックされました。");
            try
            {
                // 何かの処理
                int result = 10 / 1; // 例外を発生させる例
            }
            catch (DivideByZeroException ex)
            {
                _logger.LogError(ex, "ゼロ除算エラーが発生しました。");
            }
        }
        private void Button_Click2(object sender, RoutedEventArgs e)
        {
            _logger.LogDebug("ボタン2がクリックされました。");
            try
            {
                // 何かの処理
                int result = 10 / 1; // 例外を発生させる例
            }
            catch (DivideByZeroException ex)
            {
                _logger.LogError(ex, "ゼロ除算エラーが発生しました。");
            }
        }
        private void Button_Click3(object sender, RoutedEventArgs e)
        {
            _logger.LogDebug("ボタン3がクリックされました。");
            try
            {
                // 何かの処理
                int result = 10 / 1; // 例外を発生させる例
            }
            catch (DivideByZeroException ex)
            {
                _logger.LogError(ex, "ゼロ除算エラーが発生しました。");
            }
        }
        
        string eventtext = "";
        // タイマのインスタンス
        private DispatcherTimer _timer;
        // タイマメソッド
        private void MyTimerMethod(object sender, EventArgs e)
        {
            this.TextBlock1.Text = $"{DateTime.Now.ToString("HH:mm:ss")}{eventtext}";

            if (DateTime.Now.Second % 10 == 0) // 10秒ごとにメッセージを表示
            {
                eventtext = " - 10秒経過しました";
                Timerevent_func timerevent_func = new Timerevent_func();
                timerevent_func.TimerTarget_SearchFunc();

            }
            else
            {
                eventtext = "";
            }
        }
        // タイマを設定する
        private void SetupTimer()
        {
            // タイマのインスタンスを生成
            _timer = new DispatcherTimer(); // 優先度はDispatcherPriority.Background
            // インターバルを設定
            _timer.Interval = new TimeSpan(0, 0, 1); //hh, mm, ss
            // タイマメソッドを設定
            _timer.Tick += new EventHandler(MyTimerMethod);
            // タイマを開始
            _timer.Start();

            // 画面が閉じられるときに、タイマを停止
            this.Closing += new CancelEventHandler(StopTimer);
        }
        // タイマを停止
        private void StopTimer(object sender, CancelEventArgs e)
        {
            _timer.Stop();
        }
    }
}
