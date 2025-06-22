using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.InteropServices.JavaScript;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Documents;

namespace TimerFunc.tests
{
    class Timerevent_func
    {
        public void TimerTarget_SearchFunc()
        {
            // タイマメソッド
            // this.TextBlock1.Text = $"{DateTime.Now.ToString("HH:mm:ss")}{eventtext}";
            Console.WriteLine($"{DateTime.Now.ToString("HH:mm:ss")}");

            //検索処理
            Console.WriteLine("検索処理を実行中...");

            Window configurationAleart = new ConfigurationAleart();
            configurationAleart.Show();
        }
    }
}
