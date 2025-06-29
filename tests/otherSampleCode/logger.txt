using NPOI.HPSF;
using NPOI.SS.Formula.Atp;
using NPOI.SS.Formula.Functions;
using NPOI.XSSF.Streaming.Values;
using Org.BouncyCastle.Asn1.X509;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using tests.otherSampleCode;

namespace tests.otherSampleCode
{
    internal class logger
    {
    }
}

/*
はい、NLogを使ってソフトウェアの動作状態ログと構造化された実績データログを別々のファイルに記録し、それぞれ異なる保存期間（1日と7日）で管理する方法ですね。ユーザーディレクトリに.UserSwLogという隠しフォルダを作成して使うように設定します。

NLogのターゲット (Targets) と ルール (Rules) を適切に組み合わせることで、この要件を実現できます。特に、ログファイルの場所指定にはNLogのレイアウトレンダラーを活用し、保存期間の管理にはファイルターゲットのアーカイブ機能を使います。

1. NLog NuGetパッケージのインストール
まず、プロジェクトにNLogのコアパッケージをインストールします。

Bash

Install-Package NLog
2. NLog.config ファイルの作成と設定
プロジェクトのルートディレクトリに**NLog.configという名前のXMLファイルを作成し、以下の内容を記述します。
ビルドアクションを「コンテンツ」に、出力ディレクトリにコピーを「新しい場合はコピーする」または「常にコピーする」**に設定することを忘れないでください。

NLog.config の例:
 */
<? xml version = "1.0" encoding = "utf-8" ?>
< nlog xmlns = "http://www.nlog-project.org/schemas/NLog.xsd"
      xmlns: xsi = "http://www.w3.org/2001/XMLSchema-instance"
      autoReload = "true"
      throwExceptions = "false" >

  < variable name = "logBaseDir" value = "${specialfolder:folder=LocalApplicationData}/.UserSwLog" />

  < targets >
    < target xsi: type = "File"
            name = "appStatusLog"
            fileName = "${logBaseDir}/app_status_${shortdate}.log"
            layout = "${longdate}|${level:uppercase=true}|${logger}|${message} ${exception:format=tostring}"
            archiveEvery = "Day"
            archiveFileName = "${logBaseDir}/archives/app_status_${archiveDate:format=yyyyMMdd}_{#}.log"
            archiveNumbering = "DateAndSequence"
            maxArchiveFiles = "7"          keepFileOpen = "false"
            encoding = "utf-8"
            concurrentWrites = "true"
            />

    < target xsi: type = "File"
            name = "structuredDataLog"
            fileName = "${logBaseDir}/structured_data_${shortdate}.log"
            layout = "${message}"         archiveEvery = "Day"
            archiveFileName = "${logBaseDir}/archives/structured_data_${archiveDate:format=yyyyMMdd}_{#}.log"
            archiveNumbering = "DateAndSequence"
            maxArchiveFiles = "1"          keepFileOpen = "false"
            encoding = "utf-8"
            concurrentWrites = "true"
            />

    < target xsi: type = "Console"
            name = "console"
            layout = "${longdate}|${level:uppercase=true}|${logger}|${message} ${exception:format=tostring}" />
  </ targets >

  < rules >
    < logger name = "StructuredDataLogger" minlevel = "Info" writeTo = "structuredDataLog" final = "true" />


    < logger name = "*" minlevel = "Debug" writeTo = "appStatusLog" />

    < logger name = "*" minlevel = "Info" writeTo = "console" />
  </ rules >
</ nlog >
/*
3. C# コードでの利用例
C#コードでは、ログの種類に応じて異なるロガーインスタンスを使用し、構造化データは指定されたフォーマットで文字列を生成してログに渡します。
 */
using System;
using NLog;
using System.IO; // Directory.CreateDirectory 用

public class DualLoggingWithRetentionExample
{
    // ソフトウェアの動作状態ログ用ロガー
    // 通常のクラス名ベースのロガーを使用します。
    private static readonly ILogger AppStatusLogger = LogManager.GetCurrentClassLogger();

    // 構造化された実績データログ用ロガー
    // NLog.configで定義した "StructuredDataLogger" という名前のロガーを取得します。
    private static readonly ILogger StructuredDataLogger = LogManager.GetLogger("StructuredDataLogger");

    public void SimulateApplicationProcess()
    {
        // ログフォルダが存在しない場合は作成する（NLogが自動で作成することもありますが、明示的に行うと確実です）
        string logBaseDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), ".UserSwLog");
        if (!Directory.Exists(logBaseDir))
        {
            Directory.CreateDirectory(logBaseDir);
            // オプション: フォルダを隠し属性にする（Windowsの場合）
            // new DirectoryInfo(logBaseDir).Attributes |= FileAttributes.Hidden;
        }

        AppStatusLogger.Info("アプリケーションが起動しました。");
        AppStatusLogger.Debug("初期設定を読み込み中...");

        try
        {
            // ある処理を実行し、構造化された実績データを生成
            string dataAContent = "ユーザーID: 54321";
            string dataBContent = "処理結果: 正常完了";
            int processingTimeMs = 85;

            // 構造化された実績データを指定のテキストフォーマットでログに出力
            // NLog.configのlayout="${message}"に合わせて、ここでフォーマット済みの文字列を生成します。
            string structuredLogMessage = $"[データA:{dataAContent},データB:{dataBContent},処理時間:{processingTimeMs}ms]";
            StructuredDataLogger.Info(structuredLogMessage);

            AppStatusLogger.Info($"処理が完了しました。実績データをログに記録しました。");

            // 例外が発生した場合のログ
            // throw new InvalidOperationException("データベース接続に失敗しました。");
        }
        catch (Exception ex)
        {
            AppStatusLogger.Error(ex, "アプリケーション処理中に予期せぬエラーが発生しました。");
        }

        AppStatusLogger.Info("アプリケーションが終了します。");
    }

    public static void Main(string[] args)
    {
        Console.WriteLine("NLogによる2種類のログ出力テストを開始します。");
        Console.WriteLine("ログファイルはユーザーのAppData\\Local\\.UserSwLog フォルダに出力されます。");
        Console.WriteLine("  - 動作状態ログ: app_status_YYYYMMDD.log (7日分保持)");
        Console.WriteLine("  - 構造化データログ: structured_data_YYYYMMDD.log (1日分保持)");

        DualLoggingWithRetentionExample app = new DualLoggingWithRetentionExample();
        app.SimulateApplicationProcess();

        // アプリケーション終了時にNLogをシャットダウンすることで、
        // バッファされたログが確実にファイルに書き込まれます。
        LogManager.Shutdown();

        Console.WriteLine("\nログ出力が完了しました。ファイルパスを確認してください。");
        Console.WriteLine("何かキーを押して終了します...");
        Console.ReadKey();
    }
}
/*
解説とポイント
ログ格納場所 (${specialfolder:folder=LocalApplicationData}/.UserSwLog):

logBaseDir 変数を定義し、fileName で使用しています。

${specialfolder:folder=LocalApplicationData} は、Windowsの場合、C:\Users\<ユーザー名>\AppData\Local を指します。これは、アプリケーションがユーザーごとにデータを保存するための推奨される場所です。

.UserSwLog というフォルダ名を指定することで、Windowsではデフォルトで隠しフォルダとして扱われます（ファイルエクスプローラーで「隠しファイル」を表示する設定でないと見えません）。

C#コードで明示的なフォルダ作成: Directory.CreateDirectory(logBaseDir) を追加しました。NLogはログファイル書き込み時にフォルダを自動作成しますが、事前に作成しておくことでパス解決の問題などを避けられます。

動作状態ログ (appStatusLog):

fileName="${logBaseDir}/app_status_${shortdate}.log": app_status_YYYYMMDD.log という形式で、日ごとにファイルが作成されます。

maxArchiveFiles="7": アーカイブされたログファイルのうち、直近の7日分のみを保持し、それより古いものは自動的に削除されます。

構造化データログ (structuredDataLog):

fileName="${logBaseDir}/structured_data_${shortdate}.log": structured_data_YYYYMMDD.log という形式で、日ごとにファイルが作成されます。

layout="${message}": ここが重要です。NLogの標準的な情報（タイムスタンプ、レベル、ロガー名など）は含まず、C#コードでStructuredDataLogger.Info()に渡した文字列がそのままファイルに書き込まれます。これにより、[データA:中身,データB:中身]という純粋な構造化データのみを記録できます。

maxArchiveFiles="1": アーカイブされたファイルの最大数が1に設定されています。archiveEvery="Day" と組み合わせることで、日付が変わると前日のファイルはアーカイブフォルダに移動されますが、そのアーカイブファイルはすぐに削除されるため、実質的に常に今日のファイルのみが存在し、1日を過ぎると削除される挙動になります。

ロガーの分離とルーティング:

AppStatusLogger = LogManager.GetCurrentClassLogger();: 動作ログは通常のロガーで出力。NLog.configのname="*"ルールにマッチし、appStatusLogターゲットに送られます。

StructuredDataLogger = LogManager.GetLogger("StructuredDataLogger");: 構造化データログは、NLog.configで明示的に定義した"StructuredDataLogger"という名前のロガーで出力。このロガーからのログは、final="true"が設定されたルールによってstructuredDataLogターゲットにのみ送られ、他のログファイルには重複して出力されません。

この設定と実装により、異なる種類のログをそれぞれのファイルに、指定された保存期間で自動的に管理できるようになります。
 */