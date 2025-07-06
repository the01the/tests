using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

//nuget追加
//
//

using System.IO;
using System.Windows;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Serilog; // Serilog.Log を使うため


namespace tests.Logger
{

    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class Logger : Application
    {
        public static IServiceProvider ServiceProvider { get; private set; } // DIコンテナ

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // 1. 設定の読み込み (appsettings.json)
            IConfigurationRoot configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .Build();
            

            // 2. Serilogの初期化
            // Microsoft.Extensions.Logging と統合するために、LoggerConfigurationからILoggerを構築
            Log.Logger = new LoggerConfiguration()
                .ReadFrom.Configuration(configuration) // appsettings.jsonから設定を読み込む
                .CreateLogger();



            // 3. DIコンテナの設定
            // LoggerFactoryを通じてMicrosoft.Extensions.Logging.ILoggerを提供する
            var serviceCollection = new ServiceCollection();
            ConfigureServices(serviceCollection, configuration);
            ServiceProvider = serviceCollection.BuildServiceProvider();

            // ログの開始メッセージ
            ServiceProvider.GetService<ILogger<App>>()?.LogInformation("アプリケーションが起動しました。");

            // メインウィンドウの表示
            var mainWindow = ServiceProvider.GetService<MainWindow>();
            mainWindow.Show();
        }

        private void ConfigureServices(IServiceCollection services, IConfiguration configuration)
        {
            // ロギングサービスを追加 (Serilogを内部的に使用)
            services.AddLogging(builder =>
            {
                builder.AddSerilog(dispose: true); // SerilogをMicrosoft.Extensions.Loggingに統合
                // builder.AddConfiguration(configuration.GetSection("Logging")); // 必要に応じて標準ロギングの設定も追加
            });

            // メインウィンドウをDIに登録
            services.AddSingleton<MainWindow>();

            // その他のサービス（例: データサービス、リポジトリなど）
            services.AddSingleton<DataCacheService>(); // データキャッシュサービスを登録
        }

        protected override void OnExit(ExitEventArgs e)
        {
            // ログバッファをフラッシュして、未書き込みのログを確実にファイルに保存
            Log.Information("アプリケーションが終了します。");
            Log.CloseAndFlush();

            // DIコンテナのリソースを解放
            (ServiceProvider as System.IDisposable)?.Dispose();

            base.OnExit(e);
        }
    }
}
