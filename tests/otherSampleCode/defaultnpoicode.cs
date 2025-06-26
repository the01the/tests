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
    internal class defaultnpoicode
    {
    }
}
/*
 NPOIで既存のExcelファイルに画像を貼り付ける際に must write のようなエラーが出るのですね。そのエラーメッセージは、おそらくファイルストリームのアクセス権限や使用方法に関するものだと思われます。具体的には、ファイルを読み込み専用で開いている、または同時に書き込みができないモードで開いていることが原因で発生している可能性が高いです。

C#とNPOIで既存のExcelファイルに画像を貼り付ける場合、以下の点に注意して FileStream を使用する必要があります。

FileStream のアクセス権限と共有モード
Excelファイルを開く際に、FileStream のコンストラクタで指定する FileMode と FileAccess、そして FileShare が非常に重要です。

FileMode: ファイルを開く方法を指定します。既存ファイルへの書き込みでは FileMode.Open または FileMode.OpenOrCreate を使います。

FileAccess: ファイルへのアクセス権限を指定します。読み書き両方を行うため FileAccess.ReadWrite が必須です。

FileShare: 他のプロセスとファイルを共有する方法を指定します。これにより、ファイルが開かれている間に他のプロセスがファイルにアクセスできるかを制御します。

FileShare.None: ファイルを排他的に開きます。他のプロセスはファイルを開けません。

FileShare.Read: 他のプロセスが読み取り専用でファイルを開くことを許可します。

FileShare.Write: 他のプロセスが書き込み専用でファイルを開くことを許可します。

FileShare.ReadWrite: 他のプロセスが読み書き両方でファイルを開くことを許可します。

既存のファイルを読み込み、変更を加えて保存する場合、FileAccess.ReadWrite を指定するのが一般的です。また、他のアプリケーションが同時にそのファイルを開いている可能性がある場合は、FileShare.Read や FileShare.ReadWrite を指定することで競合を避けることができます。



エラーの具体的な原因と対処法
上記のコード例で修正されている主なポイントは以下の通りです。

FileAccess.ReadWrite の指定:

new FileStream(excelFilePath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite)

これでファイルを読み込みと書き込みの両方で開きます。FileAccess.Read のみだと、読み込んだ後に workbook.Write(fs) で書き込もうとした際に must write または類似のエラーが発生します。

FileShare.ReadWrite の指定:

FileShare.ReadWrite を指定することで、Excelファイルがすでに他のアプリケーション（例: Excelソフト自体）で開かれている場合でも、読み書きアクセスを許可し、共有違反によるエラーを避ける可能性が高まります。ただし、完全な競合回避は保証できません。もしファイルがロックされている場合は、それでもエラーになることがあります。

fs.Position = 0;:

ファイルを読み込んだ後、ストリームの現在位置はファイルの終端にあります。workbook.Write(fs) を呼び出す前に、fs.Position = 0; でストリームの書き込み位置をファイルの先頭に戻すことで、既存のファイル内容を上書きすることができます。これを忘れると、内容が追記される形になるか、書き込みに失敗することがあります。

using ステートメントの利用:

using (FileStream fs = ...) ブロックを使用することで、FileStream オブジェクトがスコープを抜ける際に自動的に Dispose() メソッドが呼び出され、ファイルハンドルが適切に閉じられることを保証します。これを怠ると、ファイルがロックされたままになり、後続の処理やExcelソフトからのアクセスができなくなることがあります。

解決しない場合の確認事項
ファイルが他のプロセスで開かれていないか？

処理を実行する前に、対象のExcelファイルがExcelアプリケーションなどで開かれていないか確認してください。開かれている場合は、一度閉じてから実行してください。

ファイルパスは正しいか？

指定したExcelファイルや画像ファイルのパスが正しいか、スペルミスやディレクトリの誤りがないか再確認してください。

デバッグ時に、excelPath や imagePath 変数の中身を確認するのが有効です。

管理者権限の問題か？

ファイルを操作しようとしているディレクトリが、書き込み権限のないシステムディレクトリなどではないか確認してください。もしそうなら、デスクトップやドキュメントフォルダなど、書き込み権限のある場所にファイルを置いて試してみてください。

NPOIのバージョン:

使用しているNPOIのバージョンが古い場合、不具合があるかもしれません。NuGetパッケージマネージャーで最新版にアップデートしてみてください。

上記の修正と確認事項で、問題が解決することを願っています。
 
 */
using System;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel; // .xlsx ファイルの場合
// using NPOI.HSSF.UserModel; // .xls ファイルの場合

public class ExcelImageWriter
{
    public static void InsertImageIntoExistingExcel(string excelFilePath, string imageFilePath, string sheetName, int col1, int row1, int col2, int row2)
    {
        IWorkbook workbook;

        // 既存のExcelファイルを読み込む
        // usingステートメントでFileStreamを適切に閉じることが重要
        using (FileStream fs = new FileStream(excelFilePath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
        {
            if (excelFilePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                workbook = new XSSFWorkbook(fs);
            }
            else if (excelFilePath.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
            {
                // .xls ファイルの場合は HSSFWorkbook を使用
                // workbook = new HSSFWorkbook(fs);
                Console.WriteLine("警告: .xls ファイルは HSSFWorkbook で処理する必要があります。このサンプルコードは .xlsx 向けです。");
                return;
            }
            else
            {
                Console.WriteLine("エラー: サポートされていないファイル形式です。(.xlsx または .xls)");
                return;
            }

            ISheet sheet = workbook.GetSheet(sheetName);
            if (sheet == null)
            {
                // シートが存在しない場合は新規作成
                sheet = workbook.CreateSheet(sheetName);
                Console.WriteLine($"シート '{sheetName}' が見つからなかったため、新規作成しました。");
            }

            // 画像ファイルをバイト配列として読み込む
            byte[] bytes = File.ReadAllBytes(imageFilePath);
            int pictureIdx = workbook.AddPicture(bytes, PictureType.PNG); // 画像タイプは適宜変更 (PNG, JPEGなど)

            IDrawing drawing = sheet.CreateDrawingPatriarch();
            XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, col1, row1, col2, row2); // x1, y1, x2, y2 は通常0で、col/rowで位置指定
            IPicture pict = drawing.CreatePicture(anchor, pictureIdx);

            // 画像のサイズをアンカーに合わせて調整しない（画像を元のサイズで挿入したい場合）
            // pict.Resize(); // これをコメントアウトすることで、指定したアンカーの範囲内に収まるように自動調整される

            // 変更内容を元のファイルに上書き保存する
            // FileStreamが終了する前に、書き込みを行う
            fs.Position = 0; // ストリームの現在位置をファイルの先頭に戻す
            workbook.Write(fs); // ストリームに書き込む
        } // usingブロックを抜ける際にfsは自動的に閉じられる

        Console.WriteLine($"画像を '{excelFilePath}' の '{sheetName}' シートに貼り付けました。");
    }

    public static void Main(string[] args)
    {
        // 使用例
        string currentDirectory = Directory.GetCurrentDirectory();
        string excelFileName = "test_excel.xlsx";
        string imageFileName = "sample_image.png"; // 実際の画像ファイル名に置き換える

        string excelPath = Path.Combine(currentDirectory, excelFileName);
        string imagePath = Path.Combine(currentDirectory, imageFileName);

        // テスト用のExcelファイルが存在しない場合は作成
        if (!File.Exists(excelPath))
        {
            using (var newWorkbook = new XSSFWorkbook())
            {
                newWorkbook.CreateSheet("Sheet1");
                using (var newFs = new FileStream(excelPath, FileMode.Create, FileAccess.Write))
                {
                    newWorkbook.Write(newFs);
                }
            }
            Console.WriteLine($"テスト用のExcelファイル '{excelPath}' を作成しました。");
        }

        // テスト用の画像ファイルが存在するか確認
        if (!File.Exists(imagePath))
        {
            Console.WriteLine($"エラー: テスト用の画像ファイル '{imagePath}' が見つかりません。");
            Console.WriteLine("実行ディレクトリに 'sample_image.png' を配置してください。");
            return;
        }

        // 画像貼り付け処理を実行
        // シート名 "Sheet1" の B2セルからD5セルの範囲に画像を挿入する例
        // col1, row1, col2, row2 は0始まりのインデックス
        // B2 (1,1) から D5 (3,4) に画像を配置
        InsertImageIntoExistingExcel(excelPath, imagePath, "Sheet1", 1, 1, 3, 4);

        Console.WriteLine("処理が完了しました。");
        Console.ReadKey();
    }
}