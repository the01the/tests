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

//timer
using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Threading;

using TimerFunc.tests; 

namespace TimerFunc.tests
{
    /// <summary>
    /// TimerWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class TimerWindow : Window
    {
        public TimerWindow()
        {
            InitializeComponent();
            SetupTimer();// タイマーセットアップ
        }
        string eventtext = "";

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

        // タイマのインスタンス
        private DispatcherTimer _timer;

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
