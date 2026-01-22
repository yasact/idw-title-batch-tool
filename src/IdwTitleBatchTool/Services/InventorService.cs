using System.IO;
using System.Runtime.InteropServices;
using IdwTitleBatchTool.Models;

namespace IdwTitleBatchTool.Services;

/// <summary>
/// Inventor Apprentice を使用してIDWファイルの表題欄プロパティを読み書きするサービス
/// </summary>
public class InventorService : IDisposable
{
    private dynamic? _apprenticeApp;
    private bool _disposed;

    // PropertySet GUIDs
    private const string UserDefinedPropertySetGuid = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
    private const string DesignTrackingPropertySetGuid = "{32853F0F-3444-11d1-9E93-0060B03C1CA6}";

    // Property IDs for Design Tracking (from PropertiesForDesignTrackingPropertiesEnum)
    // kAuthorDesignTrackingProperties = 1
    // kTitleDesignTrackingProperties = 2
    // kSubjectDesignTrackingProperties = 3
    // kCreationDateDesignTrackingProperties = 4
    private const int CreationDatePropertyId = 4;

    public event Action<string>? StatusChanged;

    public void Initialize()
    {
        if (_apprenticeApp != null) return;

        ReportStatus("Inventorを起動中...");

        // Try multiple ProgIDs for different Inventor versions
        var progIds = new[] { "Inventor.ApprenticeServer", "Inventor.ApprenticeServer.1" };
        Type? apprenticeType = null;
        string? foundProgId = null;

        foreach (var progId in progIds)
        {
            apprenticeType = Type.GetTypeFromProgID(progId);
            if (apprenticeType != null)
            {
                foundProgId = progId;
                break;
            }
        }

        if (apprenticeType == null)
        {
            throw new InvalidOperationException(
                "Inventor Apprentice が見つかりません。Autodesk Inventor がインストールされていることを確認してください。");
        }

        try
        {
            ReportStatus($"ProgID: {foundProgId} でインスタンス作成中...");
            _apprenticeApp = Activator.CreateInstance(apprenticeType);
            ReportStatus("Inventor Apprentice 起動完了");
        }
        catch (COMException comEx)
        {
            throw new InvalidOperationException(
                $"Inventor Apprentice の起動に失敗しました。\n" +
                $"ProgID: {foundProgId}\n" +
                $"HRESULT: 0x{comEx.HResult:X8}\n" +
                $"詳細: {comEx.Message}", comEx);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                $"Inventor Apprentice の起動に失敗しました。\n" +
                $"ProgID: {foundProgId}\n" +
                $"エラー: {ex.GetType().Name}\n" +
                $"詳細: {ex.Message}", ex);
        }
    }

    public List<TitleProperty> ReadTitleProperties(string folderPath)
    {
        Initialize();

        var results = new List<TitleProperty>();
        var idwFiles = Directory.GetFiles(folderPath, "*.idw", SearchOption.TopDirectoryOnly);

        if (idwFiles.Length == 0)
        {
            throw new FileNotFoundException($"フォルダ内に.idwファイルが見つかりません: {folderPath}");
        }

        for (int i = 0; i < idwFiles.Length; i++)
        {
            var filePath = idwFiles[i];
            var fileName = Path.GetFileNameWithoutExtension(filePath);

            ReportStatus($"({i + 1}/{idwFiles.Length}) {fileName}.idw を読み込み中...");

            try
            {
                var property = ReadSingleFile(filePath);
                results.Add(property);
            }
            catch (Exception ex)
            {
                // エラーが発生してもスキップして続行
                System.Diagnostics.Debug.WriteLine($"Error reading {filePath}: {ex.Message}");
                results.Add(new TitleProperty
                {
                    FilePath = filePath,
                    FileName = fileName
                });
            }
        }

        ReportStatus("読み込み完了");
        return results;
    }

    private TitleProperty ReadSingleFile(string filePath)
    {
        dynamic doc = _apprenticeApp!.Open(filePath);
        try
        {
            dynamic userProps = doc.PropertySets[UserDefinedPropertySetGuid];
            dynamic designProps = doc.PropertySets[DesignTrackingPropertySetGuid];

            var property = new TitleProperty
            {
                FilePath = filePath,
                FileName = Path.GetFileNameWithoutExtension(filePath),
                CompanyName1 = GetPropertyValue(userProps, "客先名1"),
                CompanyName2 = GetPropertyValue(userProps, "客先名2"),
                Title1 = GetPropertyValue(userProps, "名称1"),
                Title2 = GetPropertyValue(userProps, "名称2"),
                DrawingNumber = GetPropertyValue(userProps, "図番"),
                DecisionNo = GetPropertyValue(userProps, "決定No"),
                DrawnBy = GetPropertyValue(userProps, "製図"),
                DesignedBy = GetPropertyValue(userProps, "設計"),
                CheckedBy = GetPropertyValue(userProps, "検図"),
                ApprovedBy = GetPropertyValue(userProps, "承認"),
                CreationDate = GetDatePropertyValue(designProps, CreationDatePropertyId)
            };

            return property;
        }
        finally
        {
            doc.Close();
        }
    }

    public void WriteTitleProperties(List<TitleProperty> properties)
    {
        Initialize();

        for (int i = 0; i < properties.Count; i++)
        {
            var property = properties[i];
            ReportStatus($"({i + 1}/{properties.Count}) {property.FileName}.idw に書き込み中...");

            try
            {
                WriteSingleFile(property);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    $"ファイル '{property.FileName}.idw' への書き込み中にエラーが発生しました: {ex.Message}", ex);
            }
        }

        ReportStatus("書き込み完了");
    }

    private void WriteSingleFile(TitleProperty property)
    {
        dynamic doc = _apprenticeApp!.Open(property.FilePath);
        try
        {
            // Check if migration is needed
            if (doc.NeedsMigrating)
            {
                throw new InvalidOperationException(
                    $"ファイル '{property.FileName}' はマイグレーションが必要なため保存できません。");
            }

            dynamic userProps = doc.PropertySets[UserDefinedPropertySetGuid];
            dynamic designProps = doc.PropertySets[DesignTrackingPropertySetGuid];

            SetPropertyValue(userProps, "客先名1", property.CompanyName1);
            SetPropertyValue(userProps, "客先名2", property.CompanyName2);
            SetPropertyValue(userProps, "名称1", property.Title1);
            SetPropertyValue(userProps, "名称2", property.Title2);
            SetPropertyValue(userProps, "図番", property.DrawingNumber);
            SetPropertyValue(userProps, "決定No", property.DecisionNo);
            SetPropertyValue(userProps, "製図", property.DrawnBy);
            SetPropertyValue(userProps, "設計", property.DesignedBy);
            SetPropertyValue(userProps, "検図", property.CheckedBy);
            SetPropertyValue(userProps, "承認", property.ApprovedBy);
            SetDatePropertyValue(designProps, CreationDatePropertyId, property.CreationDate);

            // Save using FileSaveAs
            ReportStatus($"{property.FileName}.idw を保存中...");
            dynamic fileSaveAs = _apprenticeApp!.FileSaveAs;
            fileSaveAs.AddFileToSave(doc, doc.FullFileName);
            fileSaveAs.ExecuteSave();
        }
        finally
        {
            doc.Close();
        }
    }

    private static string GetPropertyValue(dynamic propertySet, string propertyName)
    {
        try
        {
            return propertySet[propertyName].Value?.ToString() ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static DateTime GetDatePropertyValue(dynamic propertySet, int propertyId)
    {
        try
        {
            // ItemByPropId is a method, not an indexer
            return (DateTime)propertySet.ItemByPropId(propertyId).Value;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"GetDatePropertyValue error: {ex.Message}");
            return DateTime.MinValue;
        }
    }

    private static void SetPropertyValue(dynamic propertySet, string propertyName, string value)
    {
        try
        {
            propertySet[propertyName].Value = value ?? string.Empty;
        }
        catch
        {
            // Property might not exist, ignore
        }
    }

    private static void SetDatePropertyValue(dynamic propertySet, int propertyId, DateTime value)
    {
        try
        {
            // ItemByPropId is a method, not an indexer
            propertySet.ItemByPropId(propertyId).Value = value;
        }
        catch
        {
            // Property might not exist, ignore
        }
    }

    private void ReportStatus(string status)
    {
        StatusChanged?.Invoke(status);
    }

    public void Dispose()
    {
        if (_disposed) return;

        if (_apprenticeApp != null)
        {
            try
            {
                Marshal.ReleaseComObject(_apprenticeApp);
            }
            catch
            {
                // Ignore cleanup errors
            }
            _apprenticeApp = null;
        }

        _disposed = true;
        GC.SuppressFinalize(this);
    }

    ~InventorService()
    {
        Dispose();
    }
}
