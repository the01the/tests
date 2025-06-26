using Microsoft.Win32;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection.Metadata;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using System.Windows.Media;
using System.Windows;

namespace tests.otherSampleCode
{
    internal class npoi_showImage
    {
    }
}
/*
承知しました。NPOIを使ってExcelから画像を読み込み、WPFのDataGridに表示する部分に焦点を絞って、シンプルなパターンでコードを再構成します。セルの結合に関する複雑な扱いは除外し、画像が特定のセルに紐づいていると仮定して表示する形に絞り込みますね。

WPFのDataGridはあくまで表形式のデータを扱うためのもので、Excelのようにセルに「描画された」画像をそのままインライン表示する機能は持っていません。そのため、ここでは画像の左上隅がExcelの特定のセルと一致すると仮定し、そのセルに画像を表示するというアプローチを取ります。

1. Excelデータモデルの定義
DataGridの各セルに表示するデータと、関連する画像情報を保持するためのモデルです。

C#
*/
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.IO;
using System.Windows.Media.Imaging; // BitmapImage用

// INotifyPropertyChangedはWPFのデータバインディングに必要です
public class ExcelCellDisplayData : INotifyPropertyChanged
{
    public event PropertyChangedEventHandler PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string name = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    // Excelセルの値
    private object _value;
    public object Value
    {
        get { return _value; }
        set
        {
            _value = value;
            OnPropertyChanged();
        }
    }

    // セルに紐づく画像データ（BitmapImage形式で直接バインド可能にする）
    private BitmapImage _cellImageSource;
    public BitmapImage CellImageSource
    {
        get { return _cellImageSource; }
        set
        {
            _cellImageSource = value;
            OnPropertyChanged();
        }
    }

    // 内部的な画像バイトデータ（NPOIから読み込む際に一時的に保持）
    // DataGridには直接バインドしないが、CellImageSource生成時に利用
    public byte[] ImageBytes { get; set; }
}
/*
2.Excelデータ読み込みサービス(NPOI使用)
Excelファイルからセルの値と画像データを読み込み、上記 ExcelCellDisplayData モデルにマッピングします。

C#
*/
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq; // ToList() のために必要

public class ExcelDataReadService
{
    /// <summary>
    /// Excelファイルからシートのデータを読み込み、DataGrid表示用のコレクションを生成します。
    /// </summary>
    /// <param name="filePath">Excelファイルのパス。</param>
    /// <param name="sheetName">読み込むシート名。</param>
    /// <returns>各行がExcelCellDisplayDataのリストであるコレクション。</returns>
    public ObservableCollection<ObservableCollection<ExcelCellDisplayData>> ReadExcelWithImages(string filePath, string sheetName)
    {
        var gridData = new ObservableCollection<ObservableCollection<ExcelCellDisplayData>>();
        IWorkbook workbook;

        // FileStreamを確実に閉じるためにusingを使用
        using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        {
            if (filePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                workbook = new XSSFWorkbook(fs);
            }
            else if (filePath.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
            {
                workbook = new HSSFWorkbook(fs);
            }
            else
            {
                Console.WriteLine("エラー: サポートされていないファイル形式です。(.xlsx または .xls)");
                return gridData;
            }

            ISheet sheet = workbook.GetSheet(sheetName);
            if (sheet == null)
            {
                Console.WriteLine($"エラー: シート '{sheetName}' が見つかりませんでした。");
                return gridData;
            }

            // --- シート内の画像情報を事前に取得する ---
            // NPOIで画像はIDrawingオブジェクトから取得できます。
            // 画像のアンカー情報（位置）を使って、どのセルに関連するかを後で判断します。
            var sheetPictures = new List<Tuple<byte[], IClientAnchor>>(); // byte[], アンカー情報
            if (sheet.DrawingPatriarch != null)
            {
                foreach (IShape shape in sheet.DrawingPatriarch.Shapes)
                {
                    if (shape is IPicture picture)
                    {
                        sheetPictures.Add(new Tuple<byte[], IClientAnchor>(picture.PictureData.Data, picture.GetClientAnchor()));
                    }
                }
            }

            // --- 最大行数と最大列数を取得 (正確なGridのサイズを決定するため) ---
            int firstRow = sheet.FirstRowNum;
            int lastRow = sheet.LastRowNum;
            int maxCol = 0;
            for (int r = firstRow; r <= lastRow; r++)
            {
                IRow row = sheet.GetRow(r);
                if (row != null)
                {
                    // NPOIのLastCellNumは1始まりで、最後のセル+1を返すため、
                    // その行の最後のセルがある列のインデックスは LastCellNum - 1 です。
                    // そのため maxCol は最大のLastCellNumを保持するようにします。
                    if (row.LastCellNum > maxCol) maxCol = row.LastCellNum;
                }
            }

            // 最終的な表示のために、一番右のセル番号が0の場合もあるので、+1して最低1列は確保
            if (maxCol == 0 && sheet.GetRow(firstRow)?.FirstCellNum >= 0)
            { // 最低でもセルが一つある行の場合
                maxCol = sheet.GetRow(firstRow).LastCellNum;
            }
            if (maxCol == 0) maxCol = 10; // もしデータがない場合でも最低限の列数を表示

            // --- 各セルの値と画像情報を読み込む ---
            for (int r = firstRow; r <= lastRow; r++)
            {
                var rowData = new ObservableCollection<ExcelCellDisplayData>();
                IRow excelRow = sheet.GetRow(r);

                for (int c = 0; c < maxCol; c++)
                {
                    var cellData = new ExcelCellDisplayData();
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
                                    cellData.Value = $"= {cell.CellFormula}"; // エラーの場合は式を表示
                                }
                                break;
                            case CellType.Blank:
                            case CellType.Error:
                            default:
                                cellData.Value = null;
                                break;
                        }
                    }

                    // 画像をセルに紐付ける（画像がセルの左上隅にアンカーされていると仮定）
                    var imageInfo = sheetPictures.FirstOrDefault(img =>
                        img.Item2.Row1 == r && img.Item2.Col1 == c);

                    if (imageInfo != null)
                    {
                        cellData.ImageBytes = imageInfo.Item1;
                        // BitmapImageへの変換はExcelCellDisplayDataプロパティゲッターで
                        // CellImageSourceプロパティを初期化
                        cellData.CellImageSource = new BitmapImage();
                        using (var ms = new MemoryStream(cellData.ImageBytes))
                        {
                            cellData.CellImageSource.BeginInit();
                            cellData.CellImageSource.CacheOption = BitmapCacheOption.OnLoad;
                            cellData.CellImageSource.StreamSource = ms;
                            cellData.CellImageSource.EndInit();
                            cellData.CellImageSource.Freeze(); // UIスレッド以外でもアクセス可能にする
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
3.ViewModelの準備
DataGridにバインドするデータを保持し、読み込みロジックを呼び出すViewModelです。

C#
*/
using System.ComponentModel;
using System.Collections.ObjectModel;
using System.Runtime.CompilerServices;
using System.Windows.Input; // ICommand用
using Microsoft.Win32; // OpenFileDialog用

public class MainViewModel : INotifyPropertyChanged
{
    public event PropertyChangedEventHandler PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string name = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    private ObservableCollection<ObservableCollection<ExcelCellDisplayData>> _excelData;
    public ObservableCollection<ObservableCollection<ExcelCellDisplayData>> ExcelData
    {
        get { return _excelData; }
        set
        {
            _excelData = value;
            OnPropertyChanged();
        }
    }

    public ICommand LoadExcelCommand { get; set; }

    public MainViewModel()
    {
        LoadExcelCommand = new RelayCommand(LoadExcelData);
    }

    private void LoadExcelData(object parameter)
    {
        OpenFileDialog openFileDialog = new OpenFileDialog();
        openFileDialog.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*";
        if (openFileDialog.ShowDialog() == true)
        {
            string filePath = openFileDialog.FileName;
            // 実際にはユーザーにシート名を選択させるか、最初のシートをデフォルトにする
            string sheetName = "Sheet1"; // 例として固定

            ExcelDataReadService reader = new ExcelDataReadService();
            ExcelData = reader.ReadExcelWithImages(filePath, sheetName);
        }
    }
}

// 簡単なRelayCommandの実装（MVVMパターンでコマンドをバインドするために必要）
public class RelayCommand : ICommand
{
    private readonly Action<object> _execute;
    private readonly Func<object, bool> _canExecute;

    public event EventHandler CanExecuteChanged
    {
        add { CommandManager.RequerySuggested += value; }
        remove { CommandManager.RequerySuggested -= value; }
    }

    public RelayCommand(Action<object> execute, Func<object, bool> canExecute = null)
    {
        _execute = execute ?? throw new ArgumentNullException(nameof(execute));
        _canExecute = canExecute;
    }

    public bool CanExecute(object parameter)
    {
        return _canExecute == null || _canExecute(parameter);
    }

    public void Execute(object parameter)
    {
        _execute(parameter);
    }
}
/*
4.View(XAML) の実装
WPFのDataGridは、ItemsSourceにIEnumerable<IEnumerable<T>>のような入れ子になったコレクションを直接列として自動生成する機能がありません。そのため、DataGridの各列を**DataGridTemplateColumn**として定義し、その中に画像とテキストを表示するコントロールを配置します。

重要な注意点: DataGridの列を動的に生成する方が柔軟ですが、ここではExcelの列数を事前に把握しているか、最大列数を仮定して静的にXAMLに定義する形となります。もしExcelの列数が完全に可変で、しかも多数の列があり得る場合は、コードビハインドで DataGrid.Columns コレクションを動的に構築する方が現実的です。

XML
*/
<Window x:Class = "ExcelDisplayApp.MainWindow"
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

        < Button Grid.Row = "0" Command = "{Binding LoadExcelCommand}" Content = "Load Excel File" Margin = "10" />

        < DataGrid Grid.Row = "1" ItemsSource = "{Binding ExcelData}"
                  AutoGenerateColumns = "False" CanUserAddRows = "False" CanUserDeleteRows = "False"
                  HeadersVisibility = "Column" SelectionMode = "Single"
                  AlternatingRowBackground = "LightCyan" RowBackground = "White" >


            < DataGrid.Resources >
                < local:NullToCollapsedConverter x:Key = "NullToCollapsedConverter" />
            </ DataGrid.Resources >

            < DataGrid.Columns >
                < DataGridTemplateColumn Header = "Column A" >
                    < DataGridTemplateColumn.CellTemplate >
                        < DataTemplate >
                            < Grid >
                                < Grid.RowDefinitions >
                                    < RowDefinition Height = "Auto" />
                                    < RowDefinition Height = "*" />
                                </ Grid.RowDefinitions >


                                < TextBlock Grid.Row = "0" Text = "{Binding [0].Value}"
                                           HorizontalAlignment = "Left" VerticalAlignment = "Center"
                                           Margin = "2" />


                                < Image Grid.Row = "1" Source = "{Binding [0].CellImageSource}"
                                       Width = "50" Height = "50" Stretch = "Uniform"
                                       Visibility = "{Binding [0].CellImageSource, Converter={StaticResource NullToCollapsedConverter}}"
                                       Margin = "2" />
                            </ Grid >
                        </ DataTemplate >
                    </ DataGridTemplateColumn.CellTemplate >
                </ DataGridTemplateColumn >

                < DataGridTemplateColumn Header = "Column B" >
                    < DataGridTemplateColumn.CellTemplate >
                        < DataTemplate >
                             < Grid >
                                < Grid.RowDefinitions >
                                    < RowDefinition Height = "Auto" />
                                    < RowDefinition Height = "*" />
                                </ Grid.RowDefinitions >
                                < TextBlock Grid.Row = "0" Text = "{Binding [1].Value}"
                                           HorizontalAlignment = "Left" VerticalAlignment = "Center"
                                           Margin = "2" />
                                < Image Grid.Row = "1" Source = "{Binding [1].CellImageSource}"
                                       Width = "50" Height = "50" Stretch = "Uniform"
                                       Visibility = "{Binding [1].CellImageSource, Converter={StaticResource NullToCollapsedConverter}}"
                                       Margin = "2" />
                            </ Grid >
                        </ DataTemplate >
                    </ DataGridTemplateColumn.CellTemplate >
                </ DataGridTemplateColumn >

                </ DataGrid.Columns >
        </ DataGrid >
    </ Grid >
</ Window >
/*
Converterの追加
NullToCollapsedConverter.cs をプロジェクトに追加します。

C#
*/
using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

public class NullToCollapsedConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        // valueがnullの場合（画像がない場合）、Visibility.Collapsedを返す（非表示）
        // それ以外の場合（画像がある場合）、Visibility.Visibleを返す（表示）
        return value == null ? Visibility.Collapsed : Visibility.Visible;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}
/*
そして、App.xaml の Application.Resources にこのConverterを定義します。

XML
*/
< Application x:Class = "ExcelDisplayApp.App"
             xmlns = "http://schemas.microsoft.com/winfx/2006/xaml/presentation"
             xmlns: x = "http://schemas.microsoft.com/winfx/2006/xaml"
             xmlns: local = "clr-namespace:ExcelDisplayApp"
             StartupUri = "MainWindow.xaml" >
    < Application.Resources >
        < local:NullToCollapsedConverter x:Key = "NullToCollapsedConverter" />
    </ Application.Resources >
</ Application >
/*
動作確認のための準備
Excelファイル: 任意のExcelファイル（.xlsxまたは.xls）を用意し、画像を貼り付けてください。画像は特定のセルの左上隅にアンカーされているように配置すると、このサンプルコードで表示されやすくなります。
例: Sheet1のA1セルにテキスト「Hello」と画像、B1セルにテキスト「World」と画像、というように。

プロジェクト構成:

WPFプロジェクトを作成。

NuGetパッケージ NPOI, NPOI.XSSF, NPOI.HSSF をインストール。

上記の ExcelCellDisplayData.cs, ExcelDataReadService.cs, MainViewModel.cs, RelayCommand.cs, NullToCollapsedConverter.cs をプロジェクトに追加。

MainWindow.xaml と MainWindow.xaml.cs を上記XAMLと一致するように修正。

App.xaml にConverterリソースを追加。

限界と追加の考慮事項
画像のアンカーとサイズ:
Excelでは画像をセル内に厳密に配置するだけでなく、セルをまたいで配置したり、浮動して配置したりできます。このサンプルコードは、画像の左上隅がExcelシートの特定のセルの左上隅と一致する場合にのみ、そのセルに画像を関連付けて表示します。より複雑な画像の配置（複数セルにまたがる、セル境界に位置する等）を再現するには、DataGridの各セルをさらに細かく分割したり、DataGrid以外のカスタムコントロール（例: Canvasを使った自由な配置）を検討する必要があります。

DataGridの動的な列数:
上記のXAMLは列数が固定されています。Excelの列数が常に変動する場合、XAMLで静的に列を定義するのではなく、コードビハインドでDataGrid.ColumnsコレクションにDataGridTemplateColumnを動的に追加するロジックを実装する必要があります。これは少し複雑になります。

パフォーマンス:
巨大なExcelファイルを読み込む場合、特に多数の画像を読み込むとメモリ消費が増大する可能性があります。BitmapImageの生成はメモリを消費するため、必要に応じて画像を遅延読み込みしたり、表示領域外の画像を解放したりする最適化が必要になるかもしれません。

このパターンは、Excelの画像表示をDataGridで「それらしく」見せるための、シンプルで比較的実装しやすい方法です。これで基本的な要件は満たせるかと思います。
*/