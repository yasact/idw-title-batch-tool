using IdwTitleBatchTool.Models;
using IdwTitleBatchTool.Services;

namespace IdwTitleBatchTool.Forms;

public partial class MainForm : Form
{
    private readonly InventorService _inventorService;
    private BindingSource _bindingSource;
    private List<TitleProperty> _properties;

    public MainForm()
    {
        InitializeComponent();
        _inventorService = new InventorService();
        _inventorService.StatusChanged += OnStatusChanged;
        _bindingSource = new BindingSource();
        _properties = new List<TitleProperty>();

        // Set initial folder path
        txtFolderPath.Text = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
    }

    private void OnStatusChanged(string status)
    {
        if (InvokeRequired)
        {
            Invoke(() => OnStatusChanged(status));
            return;
        }

        lblStatus.Text = status;
        Application.DoEvents();
    }

    private void btnSelectFolder_Click(object sender, EventArgs e)
    {
        using var dialog = new FolderBrowserDialog
        {
            Description = "IDWファイルが含まれるフォルダを選択してください",
            SelectedPath = txtFolderPath.Text,
            ShowNewFolderButton = false
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            txtFolderPath.Text = dialog.SelectedPath;
        }
    }

    private async void btnRead_Click(object sender, EventArgs e)
    {
        if (string.IsNullOrWhiteSpace(txtFolderPath.Text) || !Directory.Exists(txtFolderPath.Text))
        {
            MessageBox.Show("有効なフォルダパスを選択してください。", "エラー",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var result = MessageBox.Show(
            $"{txtFolderPath.Text}\nの.idwを読み込みます。",
            "読み込み確認",
            MessageBoxButtons.OKCancel,
            MessageBoxIcon.Question);

        if (result != DialogResult.OK) return;

        SetButtonsEnabled(false);

        try
        {
            _properties = await Task.Run(() => _inventorService.ReadTitleProperties(txtFolderPath.Text));
            _bindingSource.DataSource = _properties;
            dataGridView.DataSource = _bindingSource;

            ConfigureDataGridViewColumns();

            MessageBox.Show($"{_properties.Count} 件のファイルを読み込みました。", "完了",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"エラーが発生しました:\n{ex.Message}", "エラー",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            SetButtonsEnabled(true);
            lblStatus.Text = "準備完了";
        }
    }

    private async void btnWrite_Click(object sender, EventArgs e)
    {
        if (_properties.Count == 0)
        {
            MessageBox.Show("書き込むデータがありません。先にファイルを読み込んでください。", "エラー",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var result = MessageBox.Show(
            $"{txtFolderPath.Text}\nの.idwに書き込みます。",
            "書き込み確認",
            MessageBoxButtons.OKCancel,
            MessageBoxIcon.Question);

        if (result != DialogResult.OK) return;

        SetButtonsEnabled(false);

        try
        {
            // Commit any pending edits
            dataGridView.EndEdit();
            _bindingSource.EndEdit();

            await Task.Run(() => _inventorService.WriteTitleProperties(_properties));

            MessageBox.Show($"{_properties.Count} 件のファイルに書き込みました。", "完了",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"エラーが発生しました:\n{ex.Message}", "エラー",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            SetButtonsEnabled(true);
            lblStatus.Text = "準備完了";
        }
    }

    private void btnClose_Click(object sender, EventArgs e)
    {
        Close();
    }

    private void SetButtonsEnabled(bool enabled)
    {
        btnRead.Enabled = enabled;
        btnWrite.Enabled = enabled;
        btnClose.Enabled = enabled;
        btnSelectFolder.Enabled = enabled;
        lblStatus.Text = enabled ? "準備完了" : "処理中...";
    }

    private void ConfigureDataGridViewColumns()
    {
        if (dataGridView.Columns.Count == 0) return;

        // Set column headers in Japanese
        var columnSettings = new Dictionary<string, (string Header, int Width, bool ReadOnly)>
        {
            ["FilePath"] = ("ファイルパス", 200, true),
            ["FileName"] = ("ファイル名", 120, true),
            ["CompanyName1"] = ("会社名1", 100, false),
            ["CompanyName2"] = ("会社名2", 100, false),
            ["Title1"] = ("名称1", 120, false),
            ["Title2"] = ("名称2", 120, false),
            ["DrawingNumber"] = ("図番", 100, false),
            ["DecisionNo"] = ("決定No", 80, false),
            ["DrawnBy"] = ("製図", 80, false),
            ["DesignedBy"] = ("設計", 80, false),
            ["CheckedBy"] = ("検図", 80, false),
            ["ApprovedBy"] = ("承認", 80, false),
            ["CreationDate"] = ("作成日", 100, false),
        };

        foreach (DataGridViewColumn column in dataGridView.Columns)
        {
            if (columnSettings.TryGetValue(column.Name, out var settings))
            {
                column.HeaderText = settings.Header;
                column.Width = settings.Width;
                column.ReadOnly = settings.ReadOnly;
            }
        }

        // Hide file path by default (too long)
        if (dataGridView.Columns["FilePath"] != null)
        {
            dataGridView.Columns["FilePath"].Visible = false;
        }
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        _inventorService.Dispose();
        base.OnFormClosing(e);
    }
}
