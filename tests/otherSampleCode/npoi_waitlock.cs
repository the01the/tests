using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace tests.otherSampleCode
{
    internal class npoi_waitlock
    {
    }
}

/*
 はい、おっしゃる通りです。別のプロセス（特にExcelアプリケーション自体）でファイルが開かれている場合、そのファイルにはロックがかかり、プログラムからの書き込みがブロックされてしまう可能性が非常に高いです。これはNPOIやFileStreamの問題ではなく、OSレベルのファイルロックの挙動ですね。

この問題を解決するには、ファイルがロックされていないかを確認し、ロックされていたら一定時間待機して再試行する、という処理を実装するのが一般的です。

ファイルロック解除待ち処理の実装
ファイルロックの解除を待つには、基本的に以下のロジックを繰り返します。

ファイルを指定のアクセスモードで開けるか試す。

開けた場合： ロックが解除されたと判断し、処理を続行する。

開けなかった場合（IOExceptionが発生した場合）： ロックされていると判断し、

少し待機する。

設定された最大試行回数に達していないか確認する。

最大試行回数に達していなければ、再度試行する。

最大試行回数に達したら、エラーとして処理を終了する。

このロジックを、既存のExcelファイルを開く FileStream の部分に組み込みます。

修正案のコード例
前回の InsertImageIntoExistingSheet メソッドを修正し、ファイルロック解除待ちのロジックを追加します。



変更点の解説
FileStream の確立ループ:

for (int attempt = 0; attempt < MAX_RETRY_ATTEMPTS; attempt++) ループを追加しました。

このループ内で new FileStream(...) を試行します。

try-catch (IOException ex): FileStream のオープンが IOException で失敗した場合（ファイルロックが主な原因）、これはファイルが利用中であることを意味します。

再試行と待機: attempt < MAX_RETRY_ATTEMPTS - 1 で最大試行回数に達していないかを確認し、達していなければ Console.WriteLine でメッセージを表示し、Thread.Sleep(RETRY_DELAY_MS) で指定されたミリ秒数だけ処理を一時停止します。

ループの終了: FileStream が正常に開けたら break でループを抜けます。最大試行回数に達しても開けない場合は、エラーメッセージを表示してメソッドを終了します。

FileStream の using ブロック: ファイルストリームが確立された後（ループを抜けた後）、その FileStream オブジェクトを using (fileStream) ブロックで囲むことで、残りの処理が行われた後、確実にファイルが閉じられるようにしています。

定数 MAX_RETRY_ATTEMPTS と RETRY_DELAY_MS:

試行回数と待機時間をカスタマイズできるように、定数として定義しました。これらの値を調整することで、ロック解除を待つ戦略を変更できます。

MAX_RETRY_ATTEMPTS = 10 と RETRY_DELAY_MS = 500 の場合、最大で 10 * 0.5秒 = 5秒 の待機時間となります。

この修正により、Excelファイルが一時的にロックされていても、プログラムがすぐにエラーで終了せず、ロックが解除されるのを待って処理を続行できるようになります。ぜひお試しください。
 */

using System;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel; // .xlsx ファイルの場合
using NPOI.HSSF.UserModel; // .xls ファイルの場合
using System.Threading; // Thread.Sleep を使うため

public class ExcelImageModifier
{
    public static void InsertImageIntoExistingSheet(string excelFilePath, string imageFilePath, string targetSheetName, int col1, int row1, int col2, int row2)
    {
        IWorkbook workbook = null;
        const int MAX_RETRY_ATTEMPTS = 10; // 最大再試行回数
        const int RETRY_DELAY_MS = 500;   // 再試行間隔 (ミリ秒)

        // ファイルストリームを確立するループ
        FileStream fileStream = null;
        for (int attempt = 0; attempt < MAX_RETRY_ATTEMPTS; attempt++)
        {
            try
            {
                // ファイルを開く試行
                fileStream = new FileStream(excelFilePath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
                Console.WriteLine($"Excelファイル '{Path.GetFileName(excelFilePath)}' を開きました。");
                break; // 成功したらループを抜ける
            }
            catch (IOException ex)
            {
                if (attempt < MAX_RETRY_ATTEMPTS - 1)
                {
                    Console.WriteLine($"ファイルロックを検出しました。再試行します... ({attempt + 1}/{MAX_RETRY_ATTEMPTS}) - {ex.Message}");
                    Thread.Sleep(RETRY_DELAY_MS); // 少し待機
                }
                else
                {
                    Console.WriteLine($"エラー: Excelファイル '{Path.GetFileName(excelFilePath)}' へのアクセスに失敗しました。ファイルがロックされている可能性があります。");
                    Console.WriteLine($"詳細: {ex.Message}");
                    return; // 最大試行回数に達したら終了
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"予期せぬエラーが発生しました: {ex.Message}");
                return;
            }
        }

        // fileStreamがnullの場合（開けなかった場合）はここで終了
        if (fileStream == null)
        {
            return;
        }

        // usingブロックでFileStreamを確実に閉じる
        using (fileStream)
        {
            try
            {
                // ファイル形式に応じて適切なWorkbookクラスを選択
                if (excelFilePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    workbook = new XSSFWorkbook(fileStream);
                }
                else if (excelFilePath.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                {
                    workbook = new HSSFWorkbook(fileStream);
                }
                else
                {
                    Console.WriteLine("エラー: サポートされていないExcelファイル形式です。(.xlsx または .xls)");
                    return;
                }

                // ターゲットシートの取得、または新規作成
                ISheet sheet = workbook.GetSheet(targetSheetName);
                if (sheet == null)
                {
                    sheet = workbook.CreateSheet(targetSheetName);
                    Console.WriteLine($"シート '{targetSheetName}' が見つからなかったため、新規作成しました。");
                }

                // 画像ファイルをバイト配列として読み込む
                byte[] imageBytes = File.ReadAllBytes(imageFilePath);

                // Workbookに画像をPicturとして追加し、インデックスを取得
                int pictureIdx = workbook.AddPicture(imageBytes, GetPictureType(imageFilePath));

                // シートに描画オブジェクトを作成
                IDrawing drawing = sheet.CreateDrawingPatriarch();

                // 画像を配置するアンカー（位置とサイズ）を定義
                IClientAnchor anchor;
                if (workbook is XSSFWorkbook)
                {
                    anchor = new XSSFClientAnchor(0, 0, 0, 0, col1, row1, col2, row2);
                }
                else
                {
                    anchor = new HSSFClientAnchor(0, 0, 0, 0, (short)col1, row1, (short)col2, row2);
                }

                // 描画オブジェクトとアンカーを使って画像を作成
                IPicture pict = drawing.CreatePicture(anchor, pictureIdx);

                // 画像のサイズをアンカーに合わせるか、元の比率を保つかを選択
                // pict.Resize(); 

                // 変更内容をファイルストリームに書き込む
                fileStream.Position = 0;
                workbook.Write(fileStream);
                Console.WriteLine($"画像を '{excelFilePath}' の '{targetSheetName}' シートの [{col1},{row1}]～[{col2},{row2}] に貼り付けました。");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"処理中に予期せぬエラーが発生しました: {ex.Message}");
            }
        } // usingブロックを抜ける際にfileStreamは自動的に閉じられる
    }

    // 画像の拡張子からPictureTypeを取得するヘルパー関数 (変更なし)
    private static PictureType GetPictureType(string imageFilePath)
    {
        string extension = Path.GetExtension(imageFilePath)?.ToLowerInvariant();
        switch (extension)
        {
            case ".jpg":
            case ".jpeg":
                return PictureType.JPEG;
            case ".png":
                return PictureType.PNG;
            case ".gif":
                return PictureType.GIF;
            case ".bmp":
                return PictureType.BMP;
            default:
                throw new NotSupportedException($"画像形式 '{extension}' はサポートされていません。");
        }
    }

    public static void Main(string[] args)
    {
        // --- 使用例 (変更なし) ---
        string currentDirectory = AppDomain.CurrentDomain.BaseDirectory;
        string excelFileName = "MyExistingWorkbook.xlsx";
        string imageFileName = "sample_image.png";

        string excelPath = Path.Combine(currentDirectory, excelFileName);
        string imagePath = Path.Combine(currentDirectory, imageFileName);

        // テスト用のExcelファイルがなければ作成（既存ファイルがある場合はこの部分は不要）
        if (!File.Exists(excelPath))
        {
            Console.WriteLine($"'{excelFileName}' が存在しないため、テスト用ファイルを作成します。");
            using (var newWorkbook = new XSSFWorkbook())
            {
                newWorkbook.CreateSheet("Sheet1");
                using (var newFs = new FileStream(excelPath, FileMode.Create, FileAccess.Write))
                {
                    newWorkbook.Write(newFs);
                }
            }
        }
        else
        {
            Console.WriteLine($"既存のExcelファイル '{excelFileName}' を使用します。");
        }

        // テスト用の画像ファイルが存在するか確認
        if (!File.Exists(imagePath))
        {
            Console.WriteLine($"エラー: 画像ファイル '{imageFileName}' が見つかりません。");
            Console.WriteLine($"'{imageFileName}' を実行ファイルと同じディレクトリに配置してください。");
            return;
        }

        // 画像貼り付け処理を実行
        InsertImageIntoExistingSheet(excelPath, imagePath, "Sheet1", 1, 1, 3, 4);

        Console.WriteLine("\n処理が完了しました。何かキーを押して終了します...");
        Console.ReadKey();
    }
}