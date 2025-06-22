using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using NPOI;

using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using System.IO;
using System.Windows;

namespace tests.ExcelEditor
{
    public class ExcelEditor
    {
        string filePath = "sample.xlsx";
        string sheetName = "newSheet";
        //
        public ExcelEditor()
        {
            // コンストラクタ
            // Excelファイルの読み込みや初期化処理をここに記述
            createExcelFile();
            writeExcelFile();
            writeExcelFile_changeFont(filePath);
        }

        void createExcelFile()
        {
            // Excelファイルを作成するメソッド
            // NPOIライブラリを使用してExcelファイルを生成
            try
            {
                //ブック作成
                var book = CreateNewBook(filePath);

                //シート無しのexcelファイルは保存は出来るが、開くとエラーが発生する
                book.CreateSheet(sheetName);

                //ブックを保存
                using (var fs = new FileStream(filePath, FileMode.Create))
                {
                    book.Write(fs);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            Console.WriteLine("Excelファイルを作成しました。");
        }
        //ブック作成
        static IWorkbook CreateNewBook(string filePath)
        {
            IWorkbook book;
            var extension = Path.GetExtension(filePath);

            // HSSF => Microsoft Excel(xls形式)(excel 97-2003)
            // XSSF => Office Open XML Workbook形式(xlsx形式)(excel 2007以降)
            if (extension == ".xls")
            {
                book = new HSSFWorkbook();
            }
            else if (extension == ".xlsx")
            {
                book = new XSSFWorkbook();
            }
            else
            {
                throw new ApplicationException("CreateNewBook: invalid extension");
            }

            return book;
        }
        void writeExcelFile()
        {
            // Excelファイルにデータを書き込むメソッド
            // NPOIライブラリを使用してExcelファイルにデータを追加
            try
            {
                //ブック読み込み
                var book = WorkbookFactory.Create(filePath);

                //シート名からシート取得
                var sheet = book.GetSheet(sheetName);

                //セルに設定
                WriteCell(sheet, 0, 0, "0-0");
                WriteCell(sheet, 1, 1, "1-1");
                WriteCell(sheet, 0, 3, 100);
                WriteCell(sheet, 0, 4, DateTime.Today);


                //ブックを保存
                using (var fs = new FileStream("sample2.xlsx", FileMode.Create))
                {
                    book.Write(fs);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            MessageBox.Show("Excelファイルにデータを書き込みました。");
        }
        void writeExcelFile_changeFont(string filePath)
        {
            // Excelファイルにデータを書き込むメソッド
            // NPOIライブラリを使用してExcelファイルにデータを追加
            try
            {
                //ブック読み込み
                var book = WorkbookFactory.Create(filePath);

                //シート名からシート取得
                var sheet = book.GetSheet(sheetName);

                //セルに設定
                WriteCell(sheet, 0, 4, DateTime.Today);

                //日付表示するために書式変更
                var style = book.CreateCellStyle();
                style.DataFormat = book.CreateDataFormat().GetFormat("yyyy/mm/dd");
                WriteStyle(sheet, 0, 4, style);

                //フォントサイズ変更
                IFont font = book.CreateFont();

                font.FontHeightInPoints = 24;

                font.FontName = "Courier New";

                font.IsItalic = true;

                font.IsStrikeout = true;

                //スタイル変更

                ICellStyle style2 = book.CreateCellStyle();

                style2.SetFont(font);

                WriteCell(sheet, 0, 5, "フォント変更データ");
                WriteStyle(sheet, 0, 5, style2);

                //ブックを保存
                using (var fs = new FileStream("sample3.xlsx", FileMode.Create))
                {
                    book.Write(fs);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            MessageBox.Show($"Excelファイル {filePath} にデータを書き込みました。");
        }


        //セルに値を設定するメソッド
        //セル設定(文字列用)
        public static void WriteCell(ISheet sheet, int columnIndex, int rowIndex, string value)
        {
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellValue(value);
        }

        //セル設定(数値用)
        public static void WriteCell(ISheet sheet, int columnIndex, int rowIndex, double value)
        {
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellValue(value);
        }

        //セル設定(日付用)
        public static void WriteCell(ISheet sheet, int columnIndex, int rowIndex, DateTime value)
        {
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellValue(value);
        }

        //書式変更
        public static void WriteStyle(ISheet sheet, int columnIndex, int rowIndex, ICellStyle style)
        {
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.CellStyle = style;
        }



    }
}
