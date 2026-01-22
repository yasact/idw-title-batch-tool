using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using IdwTitleBatchTool.Models;
using IdwTitleBatchTool.Services;
using Microsoft.Win32;

namespace IdwTitleBatchTool.ViewModels;

public partial class MainWindowViewModel : ObservableObject, IDisposable
{
    private readonly InventorService _inventorService;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(ReadCommand))]
    private string _folderPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

    [ObservableProperty]
    private string _statusText = "準備完了";

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(ReadCommand))]
    [NotifyCanExecuteChangedFor(nameof(WriteCommand))]
    [NotifyCanExecuteChangedFor(nameof(ApplyFileNameCommand))]
    private bool _isBusy;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(WriteCommand))]
    [NotifyCanExecuteChangedFor(nameof(ApplyFileNameCommand))]
    private ObservableCollection<TitleProperty> _properties = [];

    public MainWindowViewModel()
    {
        _inventorService = new InventorService();
        _inventorService.StatusChanged += OnStatusChanged;
    }

    private void OnStatusChanged(string status)
    {
        Application.Current.Dispatcher.Invoke(() => StatusText = status);
    }

    [RelayCommand]
    private void SelectFolder()
    {
        var dialog = new OpenFolderDialog
        {
            Title = "IDWファイルが含まれるフォルダを選択してください",
            InitialDirectory = FolderPath
        };

        if (dialog.ShowDialog() == true)
        {
            FolderPath = dialog.FolderName;
        }
    }

    [RelayCommand(CanExecute = nameof(CanRead))]
    private async Task ReadAsync()
    {
        if (string.IsNullOrWhiteSpace(FolderPath) || !Directory.Exists(FolderPath))
        {
            MessageBox.Show("有効なフォルダパスを選択してください。", "エラー",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var result = MessageBox.Show(
            $"{FolderPath}\nの.idwを読み込みます。",
            "読み込み確認",
            MessageBoxButton.OKCancel,
            MessageBoxImage.Question);

        if (result != MessageBoxResult.OK) return;

        IsBusy = true;

        try
        {
            var props = await Task.Run(() => _inventorService.ReadTitleProperties(FolderPath));
            Properties = new ObservableCollection<TitleProperty>(props);

            MessageBox.Show($"{Properties.Count} 件のファイルを読み込みました。", "完了",
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"エラーが発生しました:\n{ex.Message}", "エラー",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            IsBusy = false;
            StatusText = "準備完了";
        }
    }

    private bool CanRead() => !IsBusy && !string.IsNullOrWhiteSpace(FolderPath);

    [RelayCommand(CanExecute = nameof(CanWrite))]
    private async Task WriteAsync()
    {
        var result = MessageBox.Show(
            $"{FolderPath}\nの.idwに書き込みます。",
            "書き込み確認",
            MessageBoxButton.OKCancel,
            MessageBoxImage.Question);

        if (result != MessageBoxResult.OK) return;

        IsBusy = true;

        try
        {
            await Task.Run(() => _inventorService.WriteTitleProperties([.. Properties]));

            MessageBox.Show($"{Properties.Count} 件のファイルに書き込みました。", "完了",
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"エラーが発生しました:\n{ex.Message}", "エラー",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            IsBusy = false;
            StatusText = "準備完了";
        }
    }

    private bool CanWrite() => !IsBusy && Properties.Count > 0;

    [RelayCommand(CanExecute = nameof(CanApplyFileName))]
    private void ApplyFileName()
    {
        foreach (var prop in Properties)
        {
            var fileName = Path.GetFileNameWithoutExtension(prop.FileName);
            var parts = fileName.Split('_');

            if (parts.Length >= 1)
            {
                prop.DrawingNumber = parts[0];
            }

            if (parts.Length >= 4)
            {
                prop.Title2 = string.Join("_", parts.Skip(3));
            }
        }

        // Refresh the collection to update UI
        var temp = Properties;
        Properties = [];
        Properties = temp;

        MessageBox.Show("ファイル名から図番・名称2を取得しました。", "完了",
            MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private bool CanApplyFileName() => !IsBusy && Properties.Count > 0;

    [RelayCommand]
    private void Close()
    {
        Application.Current.MainWindow?.Close();
    }

    public void Dispose()
    {
        _inventorService.Dispose();
        GC.SuppressFinalize(this);
    }
}
