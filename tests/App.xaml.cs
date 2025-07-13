using System.Configuration;
using System.Data;
using System.Windows;

//ロガー用
// App.xaml.cs
using System.IO;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Serilog; // Serilog.Log を使うため


using Serilog.Debugging; // ★これを追加
using System.Text; // ★これを追加
namespace tests
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public static IServiceProvider ServiceProvider { get; private set; }

        protected override void OnStartup(StartupEventArgs e)
        {
            // Serilogの内部エラーをデバッグ出力に表示するように設定
            SelfLog.Enable(msg => System.Diagnostics.Debug.WriteLine($"[Serilog SelfLog] {msg}")); // ★これを追加


            base.OnStartup(e);

            // 1. 設定の読み込み (appsettings.json)
            var configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .Build();

            // Serilog の設定をコードでより詳細に制御
            var loggerConfiguration = new LoggerConfiguration()
                .ReadFrom.Configuration(configuration); // appsettings.jsonの大部分の設定を読み込む

            // appsettings.json の "WriteTo" セクションを完全に削除した場合のコード:
            loggerConfiguration
                .WriteTo.File(
                    path: Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                        "YourAppName",
                        "logs",
                        "log-{Date}.txt"
                    ),
                    rollingInterval: RollingInterval.Day,
                    retainedFileCountLimit: 7,
                    outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] {Message:lj}{NewLine}{Exception}",
                    encoding: Encoding.UTF8 // ★ここで Encoding.UTF8 を直接指定★
                )
                .WriteTo.Debug(); // デバッグシンクもコードで追加

            Log.Logger = loggerConfiguration.CreateLogger();
            // 2. Serilogの初期化とMicrosoft.Extensions.Loggingへの統合
            //Log.Logger = new LoggerConfiguration()
            //    .ReadFrom.Configuration(configuration) // appsettings.jsonから設定を読み込む
            //    .Enrich.FromLogContext() // ログコンテキストからプロパティをエンリッチ
            //    .WriteTo.Debug() // デバッグ出力ウィンドウにもログを出す例
            //    .CreateLogger();

            // 3. DIコンテナの設定
            var serviceCollection = new ServiceCollection();
            ConfigureServices(serviceCollection, configuration);
            ServiceProvider = serviceCollection.BuildServiceProvider();

            // アプリケーション起動ログ
            // ILogger<App> を取得してログを記録
            ServiceProvider.GetService<ILogger<App>>()?.LogInformation("アプリケーションが起動しました。");

            // メインウィンドウの表示 (DIコンテナから取得)
            var mainWindow = ServiceProvider.GetService<MainWindow>();
            mainWindow.Show();

            // デバッグ用: 実際のLOCALAPPDATAパスを確認
            string localAppDataPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            System.Diagnostics.Debug.WriteLine($"LOCALAPPDATA Path: {localAppDataPath}");

            // デバッグ用: Serilogが最終的に解決するはずのパスを推測
            string expectedLogPath = Path.Combine(localAppDataPath, "YourAppName", "logs", "log-YYYYMMDD.txt"); // YYYYMMDDは自動で入る
            System.Diagnostics.Debug.WriteLine($"Expected Log Path Pattern: {expectedLogPath}");

        }

        private void ConfigureServices(IServiceCollection services, IConfiguration configuration)
        {
            // ロギングサービスを追加
            services.AddLogging(builder =>
            {
                // SerilogをMicrosoft.Extensions.Loggingに統合
                // App.xaml.cs の Serilog.Log.Logger を使用する
                builder.AddSerilog(Log.Logger, dispose: true); // ★重要: ここで Serilog.Log.Logger を渡す
                                                               // または、設定ファイルから直接Serilogを設定させる場合 (より一般的なパターン)
                                                               // builder.AddSerilog(dispose: true);

                // コンソール出力やデバッグ出力など、他の標準プロバイダーを追加することも可能
                // builder.AddConsole();
                // builder.AddDebug();

                // ログレベルのフィルタリング（appsettings.jsonで設定済みなら不要だが、コードで設定も可能）
                // builder.SetMinimumLevel(LogLevel.Information);
            });

            // MainWindowをDIに登録
            // DIによってILogger<MainWindow>などがコンストラクタに自動で注入される
            services.AddSingleton<MainWindow>();
            services.AddTransient<Logger.LoggerWindow>(); // 要求されるたびに新しいSubWindowのインスタンスを生成

            // 他のサービスも登録 (例: DataCacheService)
            services.AddSingleton<Logger.DataCacheService>();
        }

        protected override void OnExit(ExitEventArgs e)
        {
            // ログのフラッシュ
            ServiceProvider.GetService<ILogger<App>>()?.LogInformation("アプリケーションが終了します。");
            Log.CloseAndFlush(); // Serilog のログバッファをフラッシュ

            // DIコンテナのリソース解放
            (ServiceProvider as System.IDisposable)?.Dispose();

            base.OnExit(e);
        }
    }
}