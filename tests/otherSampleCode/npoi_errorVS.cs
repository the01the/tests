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
    internal class npoi_errorVS
    {
    }
}
/*
 NPOIを使って既存のExcelシートに画像を貼り付けようとしていて、FileMode.OpenOrCreate を使っているのに workbook.Write(fs) でエラーが出るのですね。

このエラーは、いくつかの原因が考えられます。

FileMode.OpenOrCreate の問題: OpenOrCreate はファイルが存在しない場合に新しく作成し、存在する場合は開きます。しかし、既存のファイルを「上書き」する際には、ファイルのポインタ位置やストリームの扱いに注意が必要です。特に、ファイルを読み込んだ直後はポインタが終端にあるため、書き込む前に先頭に戻す必要があります。

ファイルがロックされている: Excelファイルが他のアプリケーション（Excelソフト自体など）で開かれている場合、FileStream がファイルへの書き込みをブロックされることがあります。FileShare.ReadWrite を指定しても、排他的なロックがかかっている場合は回避できません。

NPOIのバージョンやファイル形式の不整合: NPOIのバージョンが古い、または.xls（HSSF）と.xlsx（XSSF）で処理を間違えている、あるいはファイルの破損などが原因で内部的にエラーが発生している可能性もあります。

既存シートへの貼り付け方法とエラー対策
既存のシートに画像を貼り付けるための修正案と、エラーに対する具体的な対策を盛り込んだコードを示します。

修正案のコード例

NPOIを使って既存のExcelシートに画像を貼り付けようとしていて、FileMode.OpenOrCreate を使っているのに workbook.Write(fs) でエラーが出るのですね。

このエラーは、いくつかの原因が考えられます。

FileMode.OpenOrCreate の問題: OpenOrCreate はファイルが存在しない場合に新しく作成し、存在する場合は開きます。しかし、既存のファイルを「上書き」する際には、ファイルのポインタ位置やストリームの扱いに注意が必要です。特に、ファイルを読み込んだ直後はポインタが終端にあるため、書き込む前に先頭に戻す必要があります。

ファイルがロックされている: Excelファイルが他のアプリケーション（Excelソフト自体など）で開かれている場合、FileStream がファイルへの書き込みをブロックされることがあります。FileShare.ReadWrite を指定しても、排他的なロックがかかっている場合は回避できません。

NPOIのバージョンやファイル形式の不整合: NPOIのバージョンが古い、または.xls（HSSF）と.xlsx（XSSF）で処理を間違えている、あるいはファイルの破損などが原因で内部的にエラーが発生している可能性もあります。

既存シートへの貼り付け方法とエラー対策
既存のシートに画像を貼り付けるための修正案と、エラーに対する具体的な対策を盛り込んだコードを示します。

修正案のコード例
C#

using System;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel; // .xlsx ファイルの場合
using NPOI.HSSF.UserModel; // .xls ファイルの場合

public class ExcelImageModifier
{
    public static void InsertImageIntoExistingSheet(string excelFilePath, string imageFilePath, string targetSheetName, int col1, int row1, int col2, int row2)
    {
        IWorkbook workbook = null; // workbookをnullで初期化

        // 1. 既存のExcelファイルを読み込むためのFileStream
        //    FileMode.Open を使うことで「既存ファイルを開く」意図を明確にする
        //    FileAccess.ReadWrite で読み書き両方の権限を確保
        //    FileShare.ReadWrite で他のプロセスからの読み書きを許可し、ロックを軽減
        try
        {
            using (FileStream fileStream = new FileStream(excelFilePath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
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

                // 2. ターゲットシートの取得、または新規作成
                ISheet sheet = workbook.GetSheet(targetSheetName);
                if (sheet == null)
                {
                    sheet = workbook.CreateSheet(targetSheetName);
                    Console.WriteLine($"シート '{targetSheetName}' が見つからなかったため、新規作成しました。");
                }

                // 3. 画像ファイルをバイト配列として読み込む
                byte[] imageBytes = File.ReadAllBytes(imageFilePath);

                // 4. Workbookに画像をPicturとして追加し、インデックスを取得
                //    画像タイプは実際のファイルに合わせて変更 (PictureType.PNG, PictureType.JPEG, など)
                int pictureIdx = workbook.AddPicture(imageBytes, GetPictureType(imageFilePath)); 

                // 5. シートに描画オブジェクトを作成
                IDrawing drawing = sheet.CreateDrawingPatriarch();

                // 6. 画像を配置するアンカー（位置とサイズ）を定義
                //    クライアントアンカーは通常、セル内でのオフセット（dx1, dy1, dx2, dy2）は0に設定し、
                //    開始・終了の列と行で位置とサイズを決定します。
                //    XSSFFactory (XSSF) または HSSFFactory (HSSF) から適切なClientAnchorを作成
                IClientAnchor anchor;
                if (workbook is XSSFWorkbook)
                {
                    anchor = new XSSFClientAnchor(0, 0, 0, 0, col1, row1, col2, row2);
                }
                else // HSSFWorkbookの場合
                {
                    anchor = new HSSFClientAnchor(0, 0, 0, 0, (short)col1, row1, (short)col2, row2);
                }
                
                // 画像がアンカーの範囲内に収まるように設定（またはResize()で自動調整）
                // anchor.PreferTwoColumns = true; // 必要であれば設定

                // 7. 描画オブジェクトとアンカーを使って画像を作成
                IPicture pict = drawing.CreatePicture(anchor, pictureIdx);

                // 8. 画像のサイズをアンカーに合わせるか、元の比率を保つかを選択
                //    コメントアウトすると指定したアンカーのサイズに拡大縮小される
                //    pict.Resize(); // これを有効にすると、画像を元の比率で挿入し、アンカーは自動的に調整される

                // 9. 変更内容をファイルストリームに書き込む
                //    重要: FileStreamの書き込みポインタをファイルの先頭に戻す
                fileStream.Position = 0; // これがないと、ファイル末尾に追記されてしまうかエラーになる
                workbook.Write(fileStream); // FileStreamに書き込む
            } // usingブロックを抜ける際にfileStreamは自動的に閉じられる
            
            Console.WriteLine($"画像を '{excelFilePath}' の '{targetSheetName}' シートの [{col1},{row1}]～[{col2},{row2}] に貼り付けました。");
        }
        catch (IOException ex)
        {
            Console.WriteLine($"ファイルアクセスエラーが発生しました: {ex.Message}");
            Console.WriteLine("ファイルが他のアプリケーションで開かれているか、パスが不正な可能性があります。");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"予期せぬエラーが発生しました: {ex.Message}");
        }
    }

    // 画像の拡張子からPictureTypeを取得するヘルパー関数
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
            case ".bmp": // NPOIではサポートされていない場合があるため注意
                return PictureType.BMP; 
            default:
                throw new NotSupportedException($"画像形式 '{extension}' はサポートされていません。");
        }
    }

    public static void Main(string[] args)
    {
        // --- 使用例 ---
        string currentDirectory = AppDomain.CurrentDomain.BaseDirectory; // 実行ファイルのあるディレクトリ
        string excelFileName = "MyExistingWorkbook.xlsx"; // 既存のExcelファイル名
        string imageFileName = "sample_image.png";       // 貼り付けたい画像ファイル名

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
        // "Sheet1" シートの B2セル (col1=1, row1=1) から D5セル (col2=3, row2=4) に画像を挿入
        InsertImageIntoExistingSheet(excelPath, imagePath, "Sheet1", 1, 1, 3, 4); 

        Console.WriteLine("\n処理が完了しました。何かキーを押して終了します...");
        Console.ReadKey();
    }
}
エラーの原因と対策の詳細
1. FileMode.OpenOrCreate と workbook.Write(fs) の問題
エラーの原因:

FileMode.OpenOrCreate でファイルを開くと、NPOIはファイル全体を読み込みます。この時、FileStream の内部ポインタはファイルの終端に移動しています。

この状態で workbook.Write(fs) を呼び出すと、NPOIはストリームの現在のポインタ位置から書き込みを開始しようとします。つまり、既存の内容を上書きするのではなく、ファイルに追記しようとしたり、ファイル終端での書き込みが許されない状況になったりしてエラーが発生します。

対策:

fileStream.Position = 0; の追加: workbook.Write(fs) を呼び出す直前に、FileStream の Position プロパティを 0 に設定して、ストリームの書き込み位置をファイルの先頭に戻します。これにより、NPOIはファイルの内容を先頭から完全に上書きすることができます。

FileMode.Open を使用: 既存のファイルを扱う場合は、FileMode.Open を使うことで、「ファイルが存在しない場合はエラー」という挙動になり、意図がより明確になります。OpenOrCreateは通常、存在しない場合に新しく作成し、存在する場合は開く、という場合に使うことが多いです。

2. ファイルのロックと IOException
エラーの原因:

System.IO.IOException: The process cannot access the file '...' because it is being used by another process. （別のプロセスによって使用されているため、ファイルにアクセスできません）

これは、ExcelファイルがすでにExcelアプリケーションなどで開かれている場合や、以前のプログラム実行で FileStream が適切に閉じられなかった場合に発生します。

対策:

using ステートメントの徹底: using (FileStream fileStream = new FileStream(...)) のように using ブロックを使用することで、FileStream がスコープを抜ける際に確実にクローズされ、ファイルハンドルが解放されます。これは非常に重要です。

FileShare.ReadWrite の指定: FileShare.ReadWrite は、他のプロセスがそのファイルにアクセスすることを許可する共有モードです。これにより、Excelソフトなどで開かれている場合でも、読み書きアクセスを試みることが可能になります（ただし、完全な保証はありません。Excelソフト側が排他ロックをかけている場合はアクセスできません）。

ユーザーへの通知と再試行: 上記のコード例のように try-catch ブロックで IOException を捕捉し、「ファイルが他のアプリケーションで開かれている可能性があります」といったメッセージをユーザーに表示することで、問題を特定しやすくなります。

3. NPOIのバージョンやファイル形式の不整合
エラーの原因:

.xlsx ファイルを HSSFWorkbook で開こうとしたり、その逆を行ったりすると、InvalidFormatException や NullReferenceException など、様々なエラーが発生します。

NPOIの古いバージョンでは、特定の機能が実装されていなかったり、バグがあったりする可能性があります。

対策:

ファイル拡張子による自動判別: 上記コード例のように、excelFilePath.EndsWith(".xlsx", ...) で拡張子をチェックし、XSSFWorkbook または HSSFWorkbook を動的に選択します。

最新版のNPOIを使用: NuGetパッケージマネージャーで、NPOIの最新の安定版 (NPOI, NPOI.XSSF, NPOI.HSSF) をインストールまたは更新してください。

これらの修正と確認事項で、workbook.Write(fs) でのエラーが解決し、既存のExcelシートに画像を貼り付けられるようになるはずです。

何か不明な点があれば、またお気軽にお尋ねください！

*/

using System;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel; // .xlsx ファイルの場合
using NPOI.HSSF.UserModel; // .xls ファイルの場合

public class ExcelImageModifier
{
    public static void InsertImageIntoExistingSheet(string excelFilePath, string imageFilePath, string targetSheetName, int col1, int row1, int col2, int row2)
    {
        IWorkbook workbook = null; // workbookをnullで初期化

        // 1. 既存のExcelファイルを読み込むためのFileStream
        //    FileMode.Open を使うことで「既存ファイルを開く」意図を明確にする
        //    FileAccess.ReadWrite で読み書き両方の権限を確保
        //    FileShare.ReadWrite で他のプロセスからの読み書きを許可し、ロックを軽減
        try
        {
            using (FileStream fileStream = new FileStream(excelFilePath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
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

                // 2. ターゲットシートの取得、または新規作成
                ISheet sheet = workbook.GetSheet(targetSheetName);
                if (sheet == null)
                {
                    sheet = workbook.CreateSheet(targetSheetName);
                    Console.WriteLine($"シート '{targetSheetName}' が見つからなかったため、新規作成しました。");
                }

                // 3. 画像ファイルをバイト配列として読み込む
                byte[] imageBytes = File.ReadAllBytes(imageFilePath);

                // 4. Workbookに画像をPicturとして追加し、インデックスを取得
                //    画像タイプは実際のファイルに合わせて変更 (PictureType.PNG, PictureType.JPEG, など)
                int pictureIdx = workbook.AddPicture(imageBytes, GetPictureType(imageFilePath));

                // 5. シートに描画オブジェクトを作成
                IDrawing drawing = sheet.CreateDrawingPatriarch();

                // 6. 画像を配置するアンカー（位置とサイズ）を定義
                //    クライアントアンカーは通常、セル内でのオフセット（dx1, dy1, dx2, dy2）は0に設定し、
                //    開始・終了の列と行で位置とサイズを決定します。
                //    XSSFFactory (XSSF) または HSSFFactory (HSSF) から適切なClientAnchorを作成
                IClientAnchor anchor;
                if (workbook is XSSFWorkbook)
                {
                    anchor = new XSSFClientAnchor(0, 0, 0, 0, col1, row1, col2, row2);
                }
                else // HSSFWorkbookの場合
                {
                    anchor = new HSSFClientAnchor(0, 0, 0, 0, (short)col1, row1, (short)col2, row2);
                }

                // 画像がアンカーの範囲内に収まるように設定（またはResize()で自動調整）
                // anchor.PreferTwoColumns = true; // 必要であれば設定

                // 7. 描画オブジェクトとアンカーを使って画像を作成
                IPicture pict = drawing.CreatePicture(anchor, pictureIdx);

                // 8. 画像のサイズをアンカーに合わせるか、元の比率を保つかを選択
                //    コメントアウトすると指定したアンカーのサイズに拡大縮小される
                //    pict.Resize(); // これを有効にすると、画像を元の比率で挿入し、アンカーは自動的に調整される

                // 9. 変更内容をファイルストリームに書き込む
                //    重要: FileStreamの書き込みポインタをファイルの先頭に戻す
                fileStream.Position = 0; // これがないと、ファイル末尾に追記されてしまうかエラーになる
                workbook.Write(fileStream); // FileStreamに書き込む
            } // usingブロックを抜ける際にfileStreamは自動的に閉じられる

            Console.WriteLine($"画像を '{excelFilePath}' の '{targetSheetName}' シートの [{col1},{row1}]～[{col2},{row2}] に貼り付けました。");
        }
        catch (IOException ex)
        {
            Console.WriteLine($"ファイルアクセスエラーが発生しました: {ex.Message}");
            Console.WriteLine("ファイルが他のアプリケーションで開かれているか、パスが不正な可能性があります。");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"予期せぬエラーが発生しました: {ex.Message}");
        }
    }

    // 画像の拡張子からPictureTypeを取得するヘルパー関数
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
            case ".bmp": // NPOIではサポートされていない場合があるため注意
                return PictureType.BMP;
            default:
                throw new NotSupportedException($"画像形式 '{extension}' はサポートされていません。");
        }
    }

    public static void Main(string[] args)
    {
        // --- 使用例 ---
        string currentDirectory = AppDomain.CurrentDomain.BaseDirectory; // 実行ファイルのあるディレクトリ
        string excelFileName = "MyExistingWorkbook.xlsx"; // 既存のExcelファイル名
        string imageFileName = "sample_image.png";       // 貼り付けたい画像ファイル名

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
        // "Sheet1" シートの B2セル (col1=1, row1=1) から D5セル (col2=3, row2=4) に画像を挿入
        InsertImageIntoExistingSheet(excelPath, imagePath, "Sheet1", 1, 1, 3, 4);

        Console.WriteLine("\n処理が完了しました。何かキーを押して終了します...");
        Console.ReadKey();
    }
}