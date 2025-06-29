using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace tests.otherSampleCode
{
    internal class npoi_wordtemplate_write
    {
    }
}


/*
 承知いたしました。C# と NPOI を使用して、既存の Word ファイル（.docx 形式が前提となります）をテンプレートとして、特定の箇所に文字列を書き込む処理を実装する方法について解説します。

NPOIはWord文書の操作もサポートしており、特にXSSF (Excel) と並んでXWPF (Word) がその役割を担います。

NPOIによるWord文書操作の基本
XWPFDocument: Word文書全体を表すクラスです。

XWPFParagraph: 段落を表します。

XWPFRun: 段落内のスタイルが同じテキストのまとまりを表します。

XWPFTable: 表を表します。

プレースホルダーの考え方: テンプレートでは、書き換えたい場所に{{PlaceholderText}}や_KEY_のような独自の文字列（プレースホルダー）を埋め込んでおき、NPOIを使ってその文字列を検索し、置換する手法が一般的です。

実装手順とコード例
以下の手順で実装を進めます。

NPOIパッケージのインストール:
プロジェクトにNPOIの関連パッケージをNuGetで追加します。

NPOI

NPOI.OOXML (これは.docx形式を扱うために必要です)

NPOI.OpenXml4Net (通常 NPOI.OOXML と共に自動的にインストールされますが、念のため)

テンプレートWordファイルの準備:
置換したい箇所にプレースホルダーを記述したWordファイル（例: template.docx）を作成します。

例: "契約日: {{contractDate}}" や "氏名: _NAME_"

C# コードの実装:
既存のWordファイルを開き、内容を読み込み、指定のプレースホルダーを検索して置換するロジックを記述します。

テンプレートの準備例
template.docx を以下の内容で作成し、プログラムの実行ファイルと同じディレクトリに配置してください。

契約書

拝啓

平素は格別のご高配を賜り、厚く御礼申し上げます。
下記の通り、契約を締結いたします。

契約日: {{contractDate}}

契約者: {{clientName}} 様

本契約に関する詳細につきましては、別途書面にてご確認ください。

敬具
コードの解説とポイント
必要なNuGetパッケージ:

NPOI

NPOI.OOXML (.docx形式を扱うために必須)

NPOI.OpenXml4Net (通常、上記と共に自動インストール)

テンプレートファイルの読み込み:

FileStream で FileMode.Open と FileAccess.Read を使ってテンプレートを読み込みます。

FileShare.ReadWrite を指定することで、Wordファイルが他のアプリケーションで開かれている場合でも読み込みが可能になります。

XWPFDocument document = new XWPFDocument(fs); でWordドキュメントオブジェクトを作成します。

プレースホルダーの置換ロジック:

段落 (Paragraphs) の走査:
Word文書のテキストは、基本的に段落 (XWPFParagraph) の中にあります。

Run (XWPFRun) の走査:
段落はさらに Run という単位に分割されます。Runは同じ書式設定を持つテキストの連続です。NPOIでテキストを置換する場合、通常はRun内のテキストを操作します。
例えば、"契約日: {{contractDate}}" という文字列があった場合、これが一つのRunとして認識されていれば、run.Text.Replace("{{contractDate}}", "...") で置換できます。
注意点: もし「{{ と contractDate と }}」がWordの書式設定（例: 太字の一部、別のフォントなど）によって複数のRunに分かれてしまっている場合、単純なrun.Text.Containsでは検出できません。この場合、複数のRunを結合して置換するなど、より複雑なロジックが必要になります。上記のコードは、プレースホルダーが単一のRunに含まれることを前提としています。

テーブル (Tables) の走査:
テーブル内のセルにもプレースホルダーがある場合を考慮し、document.Tables を走査して同様の置換処理を行います。

変更後のドキュメントの保存:

FileStream outputFs = new FileStream(outputFilePath, FileMode.Create, FileAccess.Write): 新しいファイルを作成（または既存ファイルを上書き）し、書き込みアクセスで開きます。

document.Write(outputFs);: 変更を加えた XWPFDocument の内容を、この FileStream に書き込みます。

エラーハンドリング:
ファイル読み書き時の IOException やその他の例外を try-catch で捕捉し、ユーザーに分かりやすいメッセージを表示するようにしています。

このコードを参考に、NPOIを使ったWordテンプレートからの書き込み処理を実装してみてください。特に、テンプレートのプレースホルダーが単一のRunとして認識されるように、Wordで作成する際に書式設定をシンプルにしておくことが重要です。
 
 */

using System;
using System.IO;
using NPOI.XWPF.usermodel; // Word文書を扱うための名前空間
using System.Collections.Generic;

public class WordTemplateProcessor
{
    /// <summary>
    /// 既存のWordテンプレートファイル(.docx)を読み込み、プレースホルダーを置換して新しいファイルとして保存します。
    /// </summary>
    /// <param name="templateFilePath">テンプレートとなるWordファイルのパス。</param>
    /// <param name="outputFilePath">置換結果を保存する新しいWordファイルのパス。</param>
    /// <param name="replacements">置換する文字列の辞書。キーがプレースホルダー、値が置換後の文字列。</param>
    public static void ProcessWordTemplate(string templateFilePath, string outputFilePath, Dictionary<string, string> replacements)
    {
        if (!File.Exists(templateFilePath))
        {
            Console.WriteLine($"エラー: テンプレートファイル '{templateFilePath}' が見つかりません。");
            return;
        }

        XWPFDocument document = null;

        // テンプレートファイルを読み込みます
        try
        {
            // FileShare.ReadWrite を指定し、他のプロセスがファイルを開いていても読み込めるようにします。
            using (FileStream fs = new FileStream(templateFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                document = new XWPFDocument(fs);
            }
        }
        catch (IOException ex)
        {
            Console.WriteLine($"エラー: テンプレートファイル '{templateFilePath}' の読み込み中にファイルアクセスエラーが発生しました。");
            Console.WriteLine($"詳細: {ex.Message}");
            return;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"エラー: テンプレートファイル '{templateFilePath}' の読み込み中に予期せぬエラーが発生しました。");
            Console.WriteLine($"詳細: {ex.Message}");
            return;
        }

        if (document == null) return; // 読み込みに失敗したら処理を終了

        // --- ドキュメント内の段落を走査してプレースホルダーを置換 ---
        foreach (var paragraph in document.Paragraphs)
        {
            // 段落内の全てのRun（スタイルが同じテキストのまとまり）を走査
            foreach (var run in paragraph.Runs)
            {
                string text = run.Text;
                if (!string.IsNullOrEmpty(text))
                {
                    foreach (var kvp in replacements)
                    {
                        // プレースホルダーが存在するかチェックし、置換
                        // 例: {{contractDate}} を置換
                        if (text.Contains(kvp.Key))
                        {
                            text = text.Replace(kvp.Key, kvp.Value);
                        }
                    }
                    // 置換後のテキストをRunに設定
                    run.SetText(text, 0); // 0はRun内のテキストのインデックス
                }
            }
        }

        // --- ドキュメント内のテーブルを走査してプレースホルダーを置換（オプション） ---
        // テーブル内のセルにもプレースホルダーがある場合
        foreach (var table in document.Tables)
        {
            foreach (var row in table.Rows)
            {
                foreach (var cell in row.GetTableCells())
                {
                    foreach (var paragraph in cell.Paragraphs)
                    {
                        foreach (var run in paragraph.Runs)
                        {
                            string text = run.Text;
                            if (!string.IsNullOrEmpty(text))
                            {
                                foreach (var kvp in replacements)
                                {
                                    if (text.Contains(kvp.Key))
                                    {
                                        text = text.Replace(kvp.Key, kvp.Value);
                                    }
                                }
                                run.SetText(text, 0);
                            }
                        }
                    }
                }
            }
        }

        // --- 最終的なドキュメントを新しいファイルとして保存 ---
        try
        {
            // 新しいファイルを作成し、XWPFDocumentの内容を書き込む
            // FileMode.Create を使うことで、既存ファイルがあれば上書きされる
            using (FileStream outputFs = new FileStream(outputFilePath, FileMode.Create, FileAccess.Write))
            {
                document.Write(outputFs);
            }
            Console.WriteLine($"テンプレート処理が完了し、ファイル '{outputFilePath}' に保存されました。");
        }
        catch (IOException ex)
        {
            Console.WriteLine($"エラー: 出力ファイル '{outputFilePath}' の書き込み中にファイルアクセスエラーが発生しました。");
            Console.WriteLine($"詳細: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"エラー: 出力ファイル '{outputFilePath}' の保存中に予期せぬエラーが発生しました。");
            Console.WriteLine($"詳細: {ex.Message}");
        }
    }

    public static void Main(string[] args)
    {
        string currentDirectory = AppDomain.CurrentDomain.BaseDirectory;
        string templateFileName = "template.docx";
        string outputFileName = "output_document.docx";

        string templatePath = Path.Combine(currentDirectory, templateFileName);
        string outputPath = Path.Combine(currentDirectory, outputFileName);

        // テンプレートファイルが存在しない場合は、ユーザーに作成を促す
        if (!File.Exists(templatePath))
        {
            Console.WriteLine($"エラー: テンプレートファイル '{templateFileName}' が見つかりません。");
            Console.WriteLine("以下の内容で '{templateFileName}' を作成し、実行ファイルと同じディレクトリに配置してください:");
            Console.WriteLine("---------------------------------------------");
            Console.WriteLine("件名：契約書");
            Console.WriteLine("契約日：{{contractDate}}");
            Console.WriteLine("契約者：{{clientName}} 様");
            Console.WriteLine("---------------------------------------------");
            Console.ReadKey();
            return;
        }

        // 置換するデータ
        Dictionary<string, string> replacements = new Dictionary<string, string>
        {
            { "{{contractDate}}", DateTime.Now.ToString("yyyy年MM月dd日") },
            { "{{clientName}}", "株式会社ABC" }
        };

        // テンプレート処理を実行
        ProcessWordTemplate(templatePath, outputPath, replacements);

        Console.WriteLine("\n処理が完了しました。何かキーを押して終了します...");
        Console.ReadKey();
    }
}