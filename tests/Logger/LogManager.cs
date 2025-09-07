using System;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
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

        /// <summary>
        /// ログファイルから「中断」のログを読み込み、指定されたIDに一致する変数名とデータを取得します。
        /// </summary>
        /// <param name="id">検索するID</param>
        /// <returns>変数名とデータのペアのリスト。見つからない場合は空のリストを返します。</returns>
        public static List<Tuple<string, string>> ReadInterruptedLogs(string id)
        {
            var resultList = new List<Tuple<string, string>>();

            if (!File.Exists(LogFilePath))
            {
                return resultList;
            }

            // 排他制御
            lock (_lockObject)
            {
                try
                {
                    // ファイルを読み込み、1行ずつ処理
                    foreach (string line in File.ReadLines(LogFilePath))
                    {
                        // 正規表現を使用してログエントリを解析
                        var match = Regex.Match(line, @"^[^,]+,\s*中断,\s*([^,]+),\s*([^,]+),\s*(.*),?$");

                        // マッチした場合、かつIDが一致した場合
                        if (match.Success && match.Groups[1].Value.Trim() == id)
                        {
                            string variableName = match.Groups[2].Value.Trim();
                            string data = match.Groups[3].Value.Trim();
                            resultList.Add(new Tuple<string, string>(variableName, data));
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"ログ読み込みエラー: {ex.Message}");
                }
            }

            return resultList;
        }

    }
}
