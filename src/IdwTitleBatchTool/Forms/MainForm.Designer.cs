namespace IdwTitleBatchTool.Forms;

partial class MainForm
{
    private System.ComponentModel.IContainer components = null;

    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }
        base.Dispose(disposing);
    }

    #region Windows Form Designer generated code

    private void InitializeComponent()
    {
        this.components = new System.ComponentModel.Container();

        // Controls
        this.lblFolder = new Label();
        this.txtFolderPath = new TextBox();
        this.btnSelectFolder = new Button();
        this.dataGridView = new DataGridView();
        this.btnRead = new Button();
        this.btnWrite = new Button();
        this.btnApplyFileName = new Button();
        this.btnClose = new Button();
        this.lblStatus = new Label();
        this.panelTop = new Panel();
        this.panelBottom = new Panel();

        ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
        this.panelTop.SuspendLayout();
        this.panelBottom.SuspendLayout();
        this.SuspendLayout();

        // panelTop
        this.panelTop.Controls.Add(this.lblFolder);
        this.panelTop.Controls.Add(this.txtFolderPath);
        this.panelTop.Controls.Add(this.btnSelectFolder);
        this.panelTop.Dock = DockStyle.Top;
        this.panelTop.Height = 50;
        this.panelTop.Padding = new Padding(10);

        // lblFolder
        this.lblFolder.Text = "フォルダ:";
        this.lblFolder.AutoSize = true;
        this.lblFolder.Location = new Point(10, 17);

        // txtFolderPath
        this.txtFolderPath.Location = new Point(70, 14);
        this.txtFolderPath.Size = new Size(500, 23);
        this.txtFolderPath.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

        // btnSelectFolder
        this.btnSelectFolder.Text = "参照...";
        this.btnSelectFolder.Location = new Point(580, 13);
        this.btnSelectFolder.Size = new Size(80, 25);
        this.btnSelectFolder.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        this.btnSelectFolder.Click += btnSelectFolder_Click;

        // dataGridView
        this.dataGridView.Dock = DockStyle.Fill;
        this.dataGridView.AllowUserToAddRows = false;
        this.dataGridView.AllowUserToDeleteRows = false;
        this.dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
        this.dataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        this.dataGridView.SelectionMode = DataGridViewSelectionMode.CellSelect;
        this.dataGridView.BackgroundColor = SystemColors.Window;
        this.dataGridView.BorderStyle = BorderStyle.Fixed3D;
        this.dataGridView.EditMode = DataGridViewEditMode.EditOnEnter;

        // panelBottom
        this.panelBottom.Controls.Add(this.btnRead);
        this.panelBottom.Controls.Add(this.btnWrite);
        this.panelBottom.Controls.Add(this.btnApplyFileName);
        this.panelBottom.Controls.Add(this.btnClose);
        this.panelBottom.Controls.Add(this.lblStatus);
        this.panelBottom.Dock = DockStyle.Bottom;
        this.panelBottom.Height = 60;
        this.panelBottom.Padding = new Padding(10);

        // btnRead
        this.btnRead.Text = "読込";
        this.btnRead.Size = new Size(100, 35);
        this.btnRead.Location = new Point(10, 12);
        this.btnRead.Click += btnRead_Click;

        // btnWrite
        this.btnWrite.Text = "書込";
        this.btnWrite.Size = new Size(100, 35);
        this.btnWrite.Location = new Point(120, 12);
        this.btnWrite.Click += btnWrite_Click;

        // btnApplyFileName
        this.btnApplyFileName.Text = "ファイル名から取得";
        this.btnApplyFileName.Size = new Size(130, 35);
        this.btnApplyFileName.Location = new Point(230, 12);
        this.btnApplyFileName.Click += btnApplyFileName_Click;

        // btnClose
        this.btnClose.Text = "閉じる";
        this.btnClose.Size = new Size(100, 35);
        this.btnClose.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        this.btnClose.Location = new Point(560, 12);
        this.btnClose.Click += btnClose_Click;

        // lblStatus
        this.lblStatus.Text = "準備完了";
        this.lblStatus.AutoSize = true;
        this.lblStatus.Location = new Point(370, 20);
        this.lblStatus.ForeColor = Color.DarkBlue;

        // MainForm
        this.AutoScaleDimensions = new SizeF(7F, 15F);
        this.AutoScaleMode = AutoScaleMode.Font;
        this.ClientSize = new Size(2000, 600);
        this.Controls.Add(this.dataGridView);
        this.Controls.Add(this.panelTop);
        this.Controls.Add(this.panelBottom);
        this.MinimumSize = new Size(700, 400);
        this.StartPosition = FormStartPosition.CenterScreen;
        this.Text = "IDW表題欄一括変更ツール";

        ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
        this.panelTop.ResumeLayout(false);
        this.panelTop.PerformLayout();
        this.panelBottom.ResumeLayout(false);
        this.panelBottom.PerformLayout();
        this.ResumeLayout(false);
    }

    #endregion

    private Label lblFolder;
    private TextBox txtFolderPath;
    private Button btnSelectFolder;
    private DataGridView dataGridView;
    private Button btnRead;
    private Button btnWrite;
    private Button btnClose;
    private Label lblStatus;
    private Panel panelTop;
    private Panel panelBottom;
    private Button btnApplyFileName;
}
