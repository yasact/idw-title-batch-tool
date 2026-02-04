using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using IdwTitleBatchTool.ViewModels;

namespace IdwTitleBatchTool.Views;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
    }

    private DateTime _lastClickTime = DateTime.MinValue;
    private DataGridCell? _lastClickedCell = null;

    private void DataGrid_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
    {
        // F2キーやテキスト入力の場合は許可
        if (e.EditingEventArgs is KeyEventArgs)
        {
            return;
        }

        // マウスクリックの場合、ダブルクリック相当の時間内に同じセルをクリックした場合のみ許可
        var now = DateTime.Now;
        var cell = e.Column.GetCellContent(e.Row)?.Parent as DataGridCell;

        if (cell != null && cell == _lastClickedCell &&
            (now - _lastClickTime).TotalMilliseconds < 500)
        {
            // ダブルクリック相当 - 編集を許可
            _lastClickedCell = null;
            return;
        }

        // シングルクリック - 編集をキャンセルして時間を記録
        _lastClickedCell = cell;
        _lastClickTime = now;
        e.Cancel = true;
    }

    private void DataGrid_Copy(object sender, ExecutedRoutedEventArgs e)
    {
        var selectedCells = dataGrid.SelectedCells;
        if (selectedCells.Count == 0) return;

        // 最初の選択セルの値をコピー
        var firstCell = selectedCells[0];
        if (firstCell.Column is DataGridBoundColumn boundColumn &&
            boundColumn.Binding is System.Windows.Data.Binding binding)
        {
            var item = firstCell.Item;
            var propertyName = binding.Path.Path;
            var property = item.GetType().GetProperty(propertyName);
            var value = property?.GetValue(item);

            string textValue;
            if (value is DateTime dateValue)
            {
                textValue = dateValue.ToString("yyyy/MM/dd");
            }
            else
            {
                textValue = value?.ToString() ?? "";
            }

            Clipboard.SetText(textValue);
            e.Handled = true;
        }
    }

    private void DataGrid_Paste(object sender, ExecutedRoutedEventArgs e)
    {
        if (!Clipboard.ContainsText()) return;

        // 編集モードを終了
        dataGrid.CommitEdit(DataGridEditingUnit.Cell, true);
        dataGrid.CommitEdit(DataGridEditingUnit.Row, true);

        var selectedCells = dataGrid.SelectedCells;
        if (selectedCells.Count >= 1)
        {
            var text = Clipboard.GetText().Trim();
            foreach (var cellInfo in selectedCells)
            {
                if (cellInfo.Column.IsReadOnly) continue;

                if (cellInfo.Column is DataGridBoundColumn boundColumn &&
                    boundColumn.Binding is System.Windows.Data.Binding binding)
                {
                    var item = cellInfo.Item;
                    var propertyName = binding.Path.Path;
                    var property = item.GetType().GetProperty(propertyName);
                    if (property != null)
                    {
                        if (property.PropertyType == typeof(string))
                        {
                            property.SetValue(item, text);
                        }
                        else if (property.PropertyType == typeof(DateTime))
                        {
                            if (DateTime.TryParse(text, out var dateValue))
                            {
                                property.SetValue(item, dateValue);
                            }
                        }
                    }
                }
            }
            dataGrid.Items.Refresh();
            e.Handled = true;
        }
    }

    protected override void OnClosed(EventArgs e)
    {
        if (DataContext is IDisposable disposable)
        {
            disposable.Dispose();
        }
        base.OnClosed(e);
    }
}
