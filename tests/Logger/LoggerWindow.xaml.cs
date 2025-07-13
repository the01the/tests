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
using Microsoft.Extensions.Logging; //別のウィンドウ側でも必要
using System.Collections.ObjectModel; // MyDataItem のコレクション用
using System.Threading.Tasks; // 非同期処理用


namespace tests.Logger
{
    /// <summary>
    /// LoggerWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class LoggerWindow : Window
    {
        private readonly ILogger<LoggerWindow> _logger; // このクラスのロガー

        // コンストラクタでILogger<SubWindow>を受け取る

        public LoggerWindow(ILogger<LoggerWindow> logger)
        //public LoggerWindow()
        {
            InitializeComponent();
            SetupTimer();// タイマーセットアップ
            // ロガーとデータキャッシュサービスのインスタンスを取得
            _logger = logger;

            _logger.LogInformation("SubWindow が初期化されました。");
            this.Title = "サブウィンドウ"; // ウィンドウのタイトルを設定
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
                _logger.LogDebug(eventtext);

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
