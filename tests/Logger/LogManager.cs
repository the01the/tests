using System;
using System.IO;
using System.Reflection;
namespace tests.Logger
{
    public static class LogManager
    {
        //使用法
        //LogManager.Write("アプリケーションが起動しました。");


        // ログファイルのパスをアプリケーションの実行ディレクトリに設定
        //private static readonly string LogFilePath = Path.Combine(
        //    Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
        //    "application.log"
        //);

        private static readonly string LogFilePath = Path.Combine(
            @$"E:\\",
            "application.log"
        );

        // スレッドセーフな書き込みのためのロックオブジェクト
        private static readonly object _lockObject = new object();

        /// <summary>
        /// 指定されたテキストをログファイルに追記します。
        /// </summary>
        /// <param name="text">書き込むテキスト</param>
        public static void Write(string text)
        {
            // 排他制御
            lock (_lockObject)
            {
                try
                {
                    // 日時, 書き込みたいテキスト,
                    string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}, {text},\n";

                    // File.AppendText はファイルが存在しない場合は新規作成、存在する場合は追記します
                    File.AppendAllText(LogFilePath, logEntry);
                }
                catch (Exception ex)
                {
                    // ログファイルへの書き込みに失敗した場合の処理
                    // デバッグウィンドウにエラーを出力するなど
                    System.Diagnostics.Debug.WriteLine($"ログ書き込みエラー: {ex.Message}");
                }
            }
        }
    }
}
