namespace IdwTitleBatchTool.Models;

/// <summary>
/// IDW図面の表題欄プロパティを表すモデル
/// </summary>
public class TitleProperty
{
    /// <summary>ファイルのフルパス</summary>
    public string FilePath { get; set; } = string.Empty;

    /// <summary>ファイル名（拡張子なし）</summary>
    public string FileName { get; set; } = string.Empty;

    /// <summary>会社名1（客先名1）</summary>
    public string CompanyName1 { get; set; } = string.Empty;

    /// <summary>会社名2（客先名2）</summary>
    public string CompanyName2 { get; set; } = string.Empty;

    /// <summary>名称1</summary>
    public string Title1 { get; set; } = string.Empty;

    /// <summary>名称2</summary>
    public string Title2 { get; set; } = string.Empty;

    /// <summary>図番</summary>
    public string DrawingNumber { get; set; } = string.Empty;

    /// <summary>決定No</summary>
    public string DecisionNo { get; set; } = string.Empty;

    /// <summary>製図</summary>
    public string DrawnBy { get; set; } = string.Empty;

    /// <summary>設計</summary>
    public string DesignedBy { get; set; } = string.Empty;

    /// <summary>検図</summary>
    public string CheckedBy { get; set; } = string.Empty;

    /// <summary>承認</summary>
    public string ApprovedBy { get; set; } = string.Empty;

    /// <summary>作成日</summary>
    public DateTime CreationDate { get; set; }
}
