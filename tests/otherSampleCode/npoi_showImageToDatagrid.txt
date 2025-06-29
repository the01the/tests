using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection.Metadata;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Input;

namespace tests.otherSampleCode
{
    internal class npoi_showImageToDatagrid
    {
    }
}
/*
 C#でNPOIを使ってExcelデータをWPFのDataGridにMVVMパターンで表示する際に、画像やセルの結合も反映させたいのですね。これは、NPOIとWPF DataGridの標準機能だけでは直接的に対応できない部分が多く、カスタムロジックやコントロールの活用が必要になります。

DataGridは基本的に表形式のテキストデータを表示するのに特化しており、Excelの画像や複雑なセル結合をそのまま再現するのは困難です。しかし、いくつかの工夫で「それらしく」見せることは可能です。

1. Excelデータの読み込み (NPOI)
まず、Excelファイルからデータを読み込む部分です。テキストデータは通常の読み込みで問題ありませんが、画像や結合セルの情報は別途取得する必要があります。

データのモデリング
DataGridに表示するためのデータモデルは、セルの値を保持するだけでなく、画像情報や結合セルの情報も保持できるように拡張します。
 */

// ViewModel層に置くデータモデルの例
public class ExcelCellData : INotifyPropertyChanged
{
    public event PropertyChangedEventHandler PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string name = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public int RowIndex { get; set; }
    public int ColumnIndex { get; set; }
    public object Value { get; set; } // セルの値 (文字列、数値など)
    public bool IsMergedCell { get; set; } // このセルが結合セルの一部か
    public bool IsMergedCellMaster { get; set; } // 結合セルの左上隅のセルか
    public int MergedCellRowSpan { get; set; } = 1; // 結合セルの行スパン
    public int MergedCellColumnSpan { get; set; } = 1; // 結合セルの列スパン

    // 画像情報 (後述のImageDataクラスを使用)
    public ImageData CellImage { get; set; }

    // 背景色、文字色など、必要に応じて追加
    // public string BackgroundColor { get; set; } 
}

public class ImageData : INotifyPropertyChanged
{
    public event PropertyChangedEventHandler PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string name = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public int PictureIndex { get; set; } // NPOIのPictureインデックス
    public byte[] ImageBytes { get; set; } // 画像のバイナリデータ
    public PictureType ImageType { get; set; } // 画像のタイプ (PNG, JPEGなど)
    public int TopLeftRow { get; set; }
    public int TopLeftCol { get; set; }
    public int BottomRightRow { get; set; }
    public int BottomRightCol { get; set; }

    // WPFで表示しやすいようにBitmapImageに変換するプロパティ
    public BitmapImage ImageSource
    {
        get
        {
            if (ImageBytes == null || ImageBytes.Length == 0) return null;
            using (var ms = new MemoryStream(ImageBytes))
            {
                var image = new BitmapImage();
                image.BeginInit();
                image.CacheOption = BitmapCacheOption.OnLoad;
                image.StreamSource = ms;
                image.EndInit();
                image.Freeze(); // スレッド間で共有可能にする
                return image;
            }
        }
    }
}
/*
 Excelデータの読み込みロジック
NPOIでExcelファイルからデータを読み込み、上記モデルにマッピングします。
 */
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.Collections.ObjectModel; // ObservableCollectionを使用
using System.Windows.Media.Imaging; // BitmapImageを使用

public class ExcelDataReaderService
{
    public ObservableCollection<ObservableCollection<ExcelCellData>> ReadExcelData(string filePath, string sheetName)
    {
        ObservableCollection<ObservableCollection<ExcelCellData>> gridData = new ObservableCollection<ObservableCollection<ExcelCellData>>();
        IWorkbook workbook;

        using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        {
            if (filePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                workbook = new XSSFWorkbook(fs);
            }
            else
            {
                workbook = new HSSFWorkbook(fs);
            }

            ISheet sheet = workbook.GetSheet(sheetName);
            if (sheet == null)
            {
                // シートが存在しない場合のエラー処理
                return gridData;
            }

            // --- セルの結合情報を先に取得する ---
            // DataGridに反映させるために、結合されたセルの情報を保持するDictionary
            // キー: MergedRegion.FirstRow + "_" + MergedRegion.FirstColumn (例: "1_0" for B2)
            // 値: MergedRegion
            var mergedRegions = new Dictionary<string, CellRangeAddress>();
            for (int i = 0; i < sheet.NumMergedRegions; i++)
            {
                CellRangeAddress mergedRegion = sheet.GetMergedRegion(i);
                mergedRegions[$"{mergedRegion.FirstRow}_{mergedRegion.FirstColumn}"] = mergedRegion;
            }

            // --- 画像情報を取得する ---
            // NPOIで画像はIDrawingオブジェクトから取得できます。
            // 画像はセルに紐づいていない場合があるため、別途リストで管理し、後でセルにマッピングします。
            var sheetPictures = new List<ImageData>();
            if (sheet.DrawingPatriarch != null)
            {
                foreach (IShape shape in sheet.DrawingPatriarch.Shapes)
                {
                    if (shape is IPicture picture)
                    {
                        var anchor = picture.GetClientAnchor();
                        var imgData = new ImageData
                        {
                            PictureIndex = picture.PictureData.PictureType,
                            ImageBytes = picture.PictureData.Data,
                            ImageType = picture.PictureData.PictureType,
                            TopLeftRow = anchor.Row1,
                            TopLeftCol = anchor.Col1,
                            BottomRightRow = anchor.Row2,
                            BottomRightCol = anchor.Col2
                        };
                        sheetPictures.Add(imgData);
                    }
                }
            }


            // --- セルの値を読み込む ---
            // 最大行数と最大列数を推定（空白行や列が多いと非効率になる可能性あり）
            int firstRow = sheet.FirstRowNum;
            int lastRow = sheet.LastRowNum;
            int maxCol = 0;
            for (int r = firstRow; r <= lastRow; r++)
            {
                IRow row = sheet.GetRow(r);
                if (row != null)
                {
                    if (row.LastCellNum > maxCol) maxCol = row.LastCellNum;
                }
            }

            for (int r = firstRow; r <= lastRow; r++)
            {
                ObservableCollection<ExcelCellData> rowData = new ObservableCollection<ExcelCellData>();
                IRow excelRow = sheet.GetRow(r);

                for (int c = 0; c < maxCol; c++) // maxColまでループすることで、列数にばらつきがあっても対応
                {
                    ExcelCellData cellData = new ExcelCellData { RowIndex = r, ColumnIndex = c };
                    ICell cell = excelRow?.GetCell(c);

                    // セルの値を取得
                    if (cell != null)
                    {
                        switch (cell.CellType)
                        {
                            case CellType.String:
                                cellData.Value = cell.StringCellValue;
                                break;
                            case CellType.Numeric:
                                if (DateUtil.IsCellDateFormatted(cell))
                                {
                                    cellData.Value = cell.DateCellValue;
                                }
                                else
                                {
                                    cellData.Value = cell.NumericCellValue;
                                }
                                break;
                            case CellType.Boolean:
                                cellData.Value = cell.BooleanCellValue;
                                break;
                            case CellType.Formula:
                                // 計算結果を表示する場合
                                try
                                {
                                    cellData.Value = cell.CachedFormulaResultType switch
                                    {
                                        CellType.String => cell.StringCellValue,
                                        CellType.Numeric => DateUtil.IsCellDateFormatted(cell) ? (object)cell.DateCellValue : cell.NumericCellValue,
                                        CellType.Boolean => cell.BooleanCellValue,
                                        _ => null
                                    };
                                }
                                catch (Exception)
                                {
                                    cellData.Value = cell.CellFormula; // エラーの場合は式を表示
                                }
                                break;
                            case CellType.Blank:
                            case CellType.Error:
                            default:
                                cellData.Value = null;
                                break;
                        }
                    }

                    // 結合セルの情報をセット
                    string key = $"{r}_{c}";
                    if (mergedRegions.ContainsKey(key))
                    {
                        var region = mergedRegions[key];
                        cellData.IsMergedCell = true;
                        cellData.IsMergedCellMaster = true; // 左上隅のセル
                        cellData.MergedCellRowSpan = region.LastRow - region.FirstRow + 1;
                        cellData.MergedCellColumnSpan = region.LastColumn - region.FirstColumn + 1;
                    }
                    else
                    {
                        // 結合セルの一部かどうかを判断 (マスターセル以外)
                        // これは非常にシンプルな判定方法。厳密には全ての結合領域をチェックする必要がある
                        foreach (var entry in mergedRegions)
                        {
                            var region = entry.Value;
                            if (r >= region.FirstRow && r <= region.LastRow &&
                                c >= region.FirstColumn && c <= region.LastColumn &&
                                (r != region.FirstRow || c != region.FirstColumn)) // マスターセルではない
                            {
                                cellData.IsMergedCell = true;
                                cellData.IsMergedCellMaster = false;
                                break;
                            }
                        }
                    }

                    // セルに対応する画像を探してセット
                    // 画像のアンカーがセルの範囲内に完全に収まっている場合を想定
                    // Excelの画像配置は複雑なため、これは簡易的なマッピングです
                    foreach (var img in sheetPictures)
                    {
                        if (img.TopLeftRow == r && img.TopLeftCol == c) // 左上隅のセルに画像を紐付ける
                        {
                            cellData.CellImage = img;
                            break;
                        }
                    }

                    rowData.Add(cellData);
                }
                gridData.Add(rowData);
            }
        }
        return gridData;
    }
}
/*
 2. ViewModel層の準備
DataGridにバインドするためのViewModelを定義します。
 */

using System.ComponentModel;
using System.Collections.ObjectModel;
using System.Runtime.CompilerServices;

public class MainViewModel : INotifyPropertyChanged
{
    public event PropertyChangedEventHandler PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string name = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    private ObservableCollection<ObservableCollection<ExcelCellData>> _excelData;
    public ObservableCollection<ObservableCollection<ExcelCellData>> ExcelData
    {
        get { return _excelData; }
        set
        {
            _excelData = value;
            OnPropertyChanged();
        }
    }

    public MainViewModel()
    {
        // ここでデータを読み込む
        // 例えば、ボタンクリックで読み込む場合はコマンドにする
        LoadExcelCommand = new RelayCommand(LoadExcelData);
    }

    public RelayCommand LoadExcelCommand { get; set; }

    private void LoadExcelData(object parameter)
    {
        // 実際のファイルパスに置き換える
        string filePath = "path/to/your/excel_file.xlsx";
        string sheetName = "Sheet1";

        ExcelDataReaderService reader = new ExcelDataReaderService();
        ExcelData = reader.ReadExcelData(filePath, sheetName);
    }
}

/*
 3. View (XAML) の実装
WPF DataGridで画像とセルの結合を「再現」するには、DataGridのカスタマイズが必須です。

画像の表示
DataGridCell内にImageコントロールを配置し、ExcelCellData.CellImage.ImageSourceにバインドします。

セルの結合の再現
WPF DataGrid自体にはExcelのようなネイティブなセル結合機能はありません。これを再現するには、いくつかの方法があります。

RowSpan/ColumnSpanを適用したカスタムコントロール: DataGridCellをテンプレート化し、セルのIsMergedCellMasterがtrueの場合に、そのセルのGrid.RowSpanやGrid.ColumnSpanを設定するカスタムコントロールやConverterを作成する方法。これは非常に複雑で、DataGridの仮想化などと競合する可能性があります。

空白セルを非表示にする（最も現実的）:
最も簡単な方法は、結合セルのマスターセル（左上隅のセル）に値や画像を表示し、結合された残りのセル（IsMergedCellがtrueでIsMergedCellMasterがfalseのセル）は非表示にすることで、見た目上、結合されているように見せる方法です。ただし、これは厳密なセル結合ではなく、見た目のごまかしです。

XAMLの例 (上記2番目の「空白セル非表示」アプローチ)
DataGridはItemsSourceにObservableCollection<ObservableCollection<ExcelCellData>>のような入れ子になったコレクションを受け入れることができません。DataGridは通常、フラットなデータソース（List<MyObject>, DataTableなど）を想定します。

そのため、DataGridの各列をDataGridTemplateColumnとして定義し、手動でバインドする必要があります。これは列数が可変の場合には管理が複雑になります。

より現実的なアプローチとしては、データを List<List<ExcelCellData>> または List<List<object>> として読み込み、DataGrid.AutoGenerateColumns = true を使うのではなく、動的に DataGridTextColumn や DataGridTemplateColumn を生成する方法が考えられます。

しかし、最も簡単なMVVMでの実装は、DataTable を使うか、または列を固定してそれぞれの ExcelCellData をバインドする方法です。

ここでは、固定列数を前提とした、最も基本的なXAMLの構造を示します。
 */


< Window x: Class = "ExcelDisplayApp.MainWindow"
        xmlns = "http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns: x = "http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns: d = "http://schemas.microsoft.com/expression/blend/2008"
        xmlns: mc = "http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns: local = "clr-namespace:ExcelDisplayApp"
        mc: Ignorable = "d"
        Title = "Excel Data Display" Height = "450" Width = "800" >
    < Window.DataContext >
        < local:MainViewModel />
    </ Window.DataContext >
    < Grid >
        < Grid.RowDefinitions >
            < RowDefinition Height = "Auto" />
            < RowDefinition Height = "*" />
        </ Grid.RowDefinitions >

        < Button Grid.Row = "0" Command = "{Binding LoadExcelCommand}" Content = "Load Excel Data" Margin = "10" />

        < DataGrid Grid.Row = "1" ItemsSource = "{Binding ExcelData}"
                  AutoGenerateColumns = "False" CanUserAddRows = "False" CanUserDeleteRows = "False"
                  HeadersVisibility = "Column" SelectionMode = "Single"
                  AlternatingRowBackground = "LightCyan" RowBackground = "White" >


            < DataGrid.Columns >
                < DataGridTemplateColumn Header = "Column A" >
                    < DataGridTemplateColumn.CellTemplate >
                        < DataTemplate >
                            < Grid >
                                < TextBlock Text = "{Binding [0].Value}"
                                           HorizontalAlignment = "Left" VerticalAlignment = "Center"
                                           Margin = "2" />
                                < Image Source = "{Binding [0].CellImage.ImageSource}"
                                       Width = "30" Height = "30" Stretch = "Uniform"
                                       Visibility = "{Binding [0].CellImage, Converter={StaticResource NullToCollapsedConverter}}" />


                                < Border Background = "White"
                                        Visibility = "{Binding [0].IsMergedCell, Converter={StaticResource BooleanToCollapsedConverter}}" />
                            </ Grid >
                        </ DataTemplate >
                    </ DataGridTemplateColumn.CellTemplate >
                </ DataGridTemplateColumn >

                < DataGridTemplateColumn Header = "Column B" >
                    < DataGridTemplateColumn.CellTemplate >
                        < DataTemplate >
                            < Grid >
                                < TextBlock Text = "{Binding [1].Value}"
                                           HorizontalAlignment = "Left" VerticalAlignment = "Center"
                                           Margin = "2" />
                                < Image Source = "{Binding [1].CellImage.ImageSource}"
                                       Width = "30" Height = "30" Stretch = "Uniform"
                                       Visibility = "{Binding [1].CellImage, Converter={StaticResource NullToCollapsedConverter}}" />

                                < Border Background = "White"
                                        Visibility = "{Binding [1].IsMergedCell, Converter={StaticResource BooleanToCollapsedConverter}}" />
                            </ Grid >
                        </ DataTemplate >
                    </ DataGridTemplateColumn.CellTemplate >
                </ DataGridTemplateColumn >


                </ DataGrid.Columns >
        </ DataGrid >
    </ Grid >
</ Window >
/*
 Converterの追加 (App.xaml.cs または 任意のConverter.cs)
XAMLで Visibility を制御するためのConverterが必要です。
 */
// NullToCollapsedConverter.cs
using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

public class NullToCollapsedConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value == null ? Visibility.Collapsed : Visibility.Visible;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}

// BooleanToCollapsedConverter.cs
using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

public class BooleanToCollapsedConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is bool boolValue && boolValue)
        {
            return Visibility.Collapsed; // trueの場合、非表示
        }
        return Visibility.Visible; // falseの場合、表示
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}
/*
 App.xaml にリソースとして追加:
 */
< Application x: Class = "ExcelDisplayApp.App"
             xmlns = "http://schemas.microsoft.com/winfx/2006/xaml/presentation"
             xmlns: x = "http://schemas.microsoft.com/winfx/2006/xaml"
             xmlns: local = "clr-namespace:ExcelDisplayApp"
             StartupUri = "MainWindow.xaml" >
    < Application.Resources >
        < local:NullToCollapsedConverter x:Key = "NullToCollapsedConverter" />
        < local:BooleanToCollapsedConverter x:Key = "BooleanToCollapsedConverter" />
    </ Application.Resources >
</ Application >
/*
 考慮事項と限界
DataGridの動的列生成:
上記のXAMLは列数を固定して書く例です。Excelの列数が可変の場合、DataGrid.AutoGenerateColumns = true は使わず、ViewModelまたはコードビハインドでDataGridTextColumnやDataGridTemplateColumnを動的に生成し、DataGrid.Columnsコレクションに追加するロジックが必要になります。
その場合、ItemsSourceはList<object[]>のように、各行がオブジェクトの配列であるようなシンプルな構造にする方が扱いやすいかもしれません。その代わり、ExcelCellDataのようなリッチな情報を直接バインドできなくなります。

画像の正確な配置:
Excelの画像は特定のセルに紐づくとは限りません。セルをまたがる場合や、セルとは独立して配置される場合もあります。上記のコードは、画像の左上隅が特定のセルと一致する場合のみを想定した簡易的な紐付けです。Excelのアンカー情報を正確に再現するには、DataGridのセル上に絶対座標で画像を配置するような、DataGridではないカスタムコントロール（例えば、ItemsControlとCanvasの組み合わせ）を検討する必要があります。

セルの結合の再現:

DataGridの制約: WPFのDataGridは、行と列が厳密にグリッド構造になっているため、Excelのような「複数のセルが1つに結合される」機能をネイティブでは持ちません。

見た目のごまかし: 上記XAMLで示した「結合セル（マスター以外）を非表示にする」方法は、あくまで見た目のごまかしです。選択時に結合されたセルが個別に選択されてしまったり、データのコピー＆ペーストがExcelのように機能しなかったりする限界があります。

本格的な再現には:

UniformGridやGridを組み合わせたカスタムコントロール: DataGridを使わず、Excelのレイアウトを模倣したカスタムのグリッドコントロールを自作する方法です。これは最も柔軟ですが、開発コストが非常に高いです。

サードパーティ製グリッドコントロール: Infragistics、Telerik、DevExpressなどの高機能なUIライブラリが提供するグリッドコントロールには、Excelライクなセル結合機能をサポートしているものがあります。しかし、これらは通常、有償です。

パフォーマンス:
非常に大きなExcelファイルを読み込む場合、全てのセルデータ（特に画像データ）をObservableCollectionに保持すると、メモリ使用量が増大し、パフォーマンスが低下する可能性があります。その場合、必要な範囲だけを読み込む、仮想化を実装するなどの工夫が必要です。

まとめ
ExcelのテキストデータをWPF DataGridに表示するのはNPOIで比較的容易ですが、画像とセルの結合を忠実に再現するには、DataGridの制約により非常に高いハードルがあります。

画像表示: DataGridTemplateColumn を使用し、Image コントロールをバインドすることで、セル内に画像を表示することは可能です。ただし、画像のアンカー位置の再現には限界があります。

セルの結合: DataGridでExcelのようなセル結合を再現することは、標準機能ではほぼ不可能です。最も簡単な回避策は、結合されたセルのうち、マスターセル以外を非表示にする、または空白にする方法ですが、これは見た目のごまかしであり、厳密な結合ではありません。本格的に再現したい場合は、カスタムコントロールの自作か、有償のサードパーティ製UIコンポーネントの導入が必要になります。

まずは、上記コードを参考にテキストデータと画像の基本的な表示、そしてセルの結合部分の非表示化から試してみて、要件と開発コストのバランスを考慮して次のステップを検討することをお勧めします。
 */
