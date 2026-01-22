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

    private void DataGrid_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.V && Keyboard.Modifiers == ModifierKeys.Control)
        {
            var dataGrid = sender as DataGrid;
            if (dataGrid?.SelectedCells.Count > 1 && Clipboard.ContainsText())
            {
                var text = Clipboard.GetText().Trim();
                foreach (var cellInfo in dataGrid.SelectedCells)
                {
                    if (cellInfo.Column is DataGridBoundColumn boundColumn &&
                        boundColumn.Binding is System.Windows.Data.Binding binding &&
                        !cellInfo.Column.IsReadOnly)
                    {
                        var item = cellInfo.Item;
                        var propertyName = binding.Path.Path;
                        var property = item.GetType().GetProperty(propertyName);
                        property?.SetValue(item, text);
                    }
                }
                dataGrid.Items.Refresh();
                e.Handled = true;
            }
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
