Option Explicit
'----------------------------------------------------------------
' idw表題欄一括変更ツール
' YasAct
' ---------------------------------------------------------------
' バージョン履歴
' 0.1 | リリース
' 0.2 | Mainプログラムを標準モジュールに移動
' 0.3 | InventorApplicationをApprenticeを使用するように変更
'--------------------------------------------------------------
'バージョン番号
Public Const SOFTWAREVERSION As String = "0.3"
' ヘッダー行サイズ
Public Const HEADERROWSIZE As Long = 5
'読み込むファイル拡張子
Public Const SEARCHEXTENSION As String = "idw"

Public Sub GetTitleInfo()
    On Error GoTo ErrorHandler
    Dim hRs As Long                              ' HeaderRowSize
    hRs = HEADERROWSIZE
    'SheetをActivate
    ThisWorkbook.Sheets("iProperty一覧").Activate

    InitializeTitle (hRs)

    'ProcessingLabel-----------------------------
    Application.StatusBar = "Inventorを起動中"
    FORM_titleBatchTool.processingLabel.Caption = "Inventorを起動中"
    '--------------------------------------------
    ' Apprenticeの作成
    Dim oInvApp As ApprenticeServerComponent
    Set oInvApp = New ApprenticeServerComponent  'Apprentice対応

    Dim f As Variant
    Dim filePaths() As String
    Dim baseNames() As String
    Dim companyName1() As String
    Dim companyName2() As String
    Dim drawingTitle1() As String
    Dim drawingTitle2() As String
    Dim drawingNumber() As String
    Dim No() As String
    Dim Drawer() As String
    Dim Designer() As String
    Dim Check() As String
    Dim Approval() As String
    Dim drawingDate() As Date

    Dim cnt As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim msg As String

    Dim fso As FileSystemObject
    Set fso = New FileSystemObject

    Dim myPath As String
    myPath = FORM_titleBatchTool.dirPathLabel.Caption

    Dim searchExt As String
    searchExt = SEARCHEXTENSION

    ReDim filePaths(fso.GetFolder(myPath).Files.Count)
    ReDim baseNames(fso.GetFolder(myPath).Files.Count)
    ReDim companyName1(fso.GetFolder(myPath).Files.Count)
    ReDim companyName2(fso.GetFolder(myPath).Files.Count)
    ReDim drawingTitle1(fso.GetFolder(myPath).Files.Count)
    ReDim drawingTitle2(fso.GetFolder(myPath).Files.Count)
    ReDim drawingNumber(fso.GetFolder(myPath).Files.Count)
    ReDim No(fso.GetFolder(myPath).Files.Count)
    ReDim Drawer(fso.GetFolder(myPath).Files.Count)
    ReDim Designer(fso.GetFolder(myPath).Files.Count)
    ReDim Check(fso.GetFolder(myPath).Files.Count)
    ReDim Approval(fso.GetFolder(myPath).Files.Count)
    ReDim drawingDate(fso.GetFolder(myPath).Files.Count)

    ' ApprenticeDocumentの宣言
    Dim oDoc As ApprenticeServerDocument

    Dim oUDProps As PropertySet
    Dim oDTProps As PropertySet
    Dim oProps As Property

    For Each f In fso.GetFolder(myPath).Files
        If LCase(fso.GetExtensionName(f.Name)) = searchExt Then
            cnt = cnt + 1

            filePaths(cnt) = myPath & "\" & fso.GetFileName(f.Name) 'ファイルのフルパスを取得して配列に格納
            baseNames(cnt) = fso.GetBaseName(f.Name) 'ファイルのBaseNameを取得して配列に格納
            'processingLabel----------------------------
            Application.StatusBar = cnt & "個の." & searchExt & "がみつかりました   " & baseNames(cnt) & "." & searchExt & "を発見"
            FORM_titleBatchTool.processingLabel.Caption = baseNames(cnt) & "." & searchExt & "を発見" & vbCrLf & cnt _
                                                                       & "個の." & searchExt & "がみつかりました"
            '--------------------------------------------
            Set oDoc = oInvApp.Open(filePaths(cnt))
            Set oUDProps = oDoc.PropertySets.Item("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
            Set oDTProps = oDoc.PropertySets.Item("{32853F0F-3444-11d1-9E93-0060B03C1CA6}")
            companyName1(cnt) = oUDProps.Item("客先名1").Value
            companyName2(cnt) = oUDProps.Item("客先名2").Value
            drawingTitle1(cnt) = oUDProps.Item("名称1").Value
            drawingTitle2(cnt) = oUDProps.Item("名称2").Value
            drawingNumber(cnt) = oUDProps.Item("図番").Value
            No(cnt) = oUDProps.Item("決定No").Value
            Drawer(cnt) = oUDProps.Item("製図").Value
            Designer(cnt) = oUDProps.Item("設計").Value
            Check(cnt) = oUDProps.Item("検図").Value
            Approval(cnt) = oUDProps.Item("承認").Value
            drawingDate(cnt) = oDTProps.ItemByPropId(PropertiesForDesignTrackingPropertiesEnum.kCreationDateDesignTrackingProperties).Value

            oDoc.Close
        End If
    Next f
    If cnt = 0 Then
        MsgBox "." & searchExt & "ファイルはありません", vbExclamation
    Else
        'ProcessingLabel-------------------
        Application.StatusBar = "Excelに出力中"
        FORM_titleBatchTool.processingLabel.Caption = "Excelに出力中"
        '--------------------------------------------
        For i = 1 To cnt
            k = i + hRs + 1                      'hRsに項目名を表示したのでk=i+ hRs + 1
            With Worksheets("iProperty一覧")
                .Cells(k, 1) = filePaths(i)
                .Cells(k, 2) = baseNames(i)
                .Cells(k, 3) = companyName1(i)
                .Cells(k, 4) = companyName2(i)
                .Cells(k, 5) = drawingTitle1(i)
                .Cells(k, 6) = drawingTitle2(i)
                .Cells(k, 7) = drawingNumber(i)
                .Cells(k, 8) = No(i)
                .Cells(k, 9) = Drawer(i)
                .Cells(k, 10) = Designer(i)
                .Cells(k, 11) = Check(i)
                .Cells(k, 12) = Approval(i)
                .Cells(k, 13) = drawingDate(i)
            End With
        Next i
    End If
    '後始末--------
    Set fso = Nothing
    If Not oInvApp Is Nothing Then
        Set oInvApp = Nothing
    Else
    End If
    Application.StatusBar = False

    GoTo finally

ErrorHandler:
    Debug.Print Err.Description
    '後始末--------
    If Not oDoc Is Nothing Then
        oDoc.Close
    Else
    End If
    Set fso = Nothing
    Set oInvApp = Nothing
    msg = Application.StatusBar & "にエラーが発生しました( " _
        & Err.Number & "): "
    Application.StatusBar = False

finally:
    MsgBox "完了しました"
End Sub

Public Sub SetTitleInfo()
    'エクセルの内容を書き込む。
    On Error GoTo ErrorHandler
    Dim hRs As Long                              ' HeaderRowSize
    hRs = HEADERROWSIZE
    ThisWorkbook.Sheets("iProperty一覧").Activate

    '行がなにもなかったら終了
    If Cells(hRs + 2, 1).Value = "" Then
        MsgBox "データがありません"
        Exit Sub
    Else

        'ProcessingLabel-------------------
        Application.StatusBar = "Inventorを起動中"
        FORM_titleBatchTool.processingLabel.Caption = "Inventorを起動中"
        '--------------------------------------------
        ' Apprenticeの作成
        Dim oInvApp As ApprenticeServerComponent
        Set oInvApp = New ApprenticeServerComponent

        Dim f As Variant
        Dim filePaths() As String
        Dim baseNames() As String
        Dim companyName1() As String
        Dim companyName2() As String
        Dim drawingTitle1() As String
        Dim drawingTitle2() As String
        Dim drawingNumber() As String
        Dim No() As String
        Dim Drawer() As String
        Dim Designer() As String
        Dim Check() As String
        Dim Approval() As String
        Dim drawingDate() As Date

        Dim cnt As Long
        Dim i As Long
        Dim j As Long
        Dim k As Long
        Dim msg As String

        Dim fso As FileSystemObject
        Set fso = New FileSystemObject

        Dim myPath As String
        myPath = FORM_titleBatchTool.dirPathLabel.Caption

        Dim searchExt As String
        searchExt = SEARCHEXTENSION

        ' 配列のサイズを変更
        ReDim filePaths(fso.GetFolder(myPath).Files.Count)
        ReDim baseNames(fso.GetFolder(myPath).Files.Count)
        ReDim companyName1(fso.GetFolder(myPath).Files.Count)
        ReDim companyName2(fso.GetFolder(myPath).Files.Count)
        ReDim drawingTitle1(fso.GetFolder(myPath).Files.Count)
        ReDim drawingTitle2(fso.GetFolder(myPath).Files.Count)
        ReDim drawingNumber(fso.GetFolder(myPath).Files.Count)
        ReDim No(fso.GetFolder(myPath).Files.Count)
        ReDim Drawer(fso.GetFolder(myPath).Files.Count)
        ReDim Designer(fso.GetFolder(myPath).Files.Count)
        ReDim Check(fso.GetFolder(myPath).Files.Count)
        ReDim Approval(fso.GetFolder(myPath).Files.Count)
        ReDim drawingDate(fso.GetFolder(myPath).Files.Count)

        Dim oDoc As ApprenticeServerDocument
        Dim oUDProps As PropertySet
        Dim oDTProps As PropertySet
        Dim oProps As Property
        Dim oFileSaveAs As FileSaveAs
        'ProcessingLabel-------------------
        Application.StatusBar = "ファイルの数を数えています。"
        FORM_titleBatchTool.processingLabel.Caption = "ファイルの数を数えています。"
        '--------------------------------------------
        For Each f In fso.GetFolder(myPath).Files
            If LCase(fso.GetExtensionName(f.Name)) = searchExt Then
                cnt = cnt + 1
            End If
        Next f

        For i = 1 To cnt
            k = i + hRs + 1                      'hRs + 1 に項目名を表示したのでk=i+ hRs + 2

            filePaths(i) = Cells(k, 1).Value
            baseNames(i) = Cells(k, 2).Value
            companyName1(i) = Cells(k, 3).Value
            companyName2(i) = Cells(k, 4).Value
            drawingTitle1(i) = Cells(k, 5).Value
            drawingTitle2(i) = Cells(k, 6).Value
            drawingNumber(i) = Cells(k, 7).Value
            No(i) = Cells(k, 8).Value
            Drawer(i) = Cells(k, 9).Value
            Designer(i) = Cells(k, 10).Value
            Check(i) = Cells(k, 11).Value
            Approval(i) = Cells(k, 12).Value
            drawingDate(i) = Cells(k, 13).Value

        Next i

        For j = 1 To cnt

            'ProcessingLabel-------------------
            'Application.StatusBar = "(" & j & "/" & cnt & ")" & baseNames(j) & "." & searchExt & "に書き込み中"
            FORM_titleBatchTool.processingLabel.Caption = "(" & j & "/" & cnt & ")" & baseNames(j) & "." & searchExt & "に書き込み中"
            '--------------------------------------------
            Set oDoc = oInvApp.Open(filePaths(j))
            Set oUDProps = oDoc.PropertySets.Item("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
            Set oDTProps = oDoc.PropertySets.Item("{32853F0F-3444-11d1-9E93-0060B03C1CA6}")
            oUDProps.Item("客先名1").Value = companyName1(j)
            oUDProps.Item("客先名2").Value = companyName2(j)
            oUDProps.Item("名称1").Value = drawingTitle1(j)
            oUDProps.Item("名称2").Value = drawingTitle2(j)
            oUDProps.Item("図番").Value = drawingNumber(j)
            oUDProps.Item("決定No").Value = No(j)
            oUDProps.Item("製図").Value = Drawer(j)
            oUDProps.Item("設計").Value = Designer(j)
            oUDProps.Item("検図").Value = Check(j)
            oUDProps.Item("承認").Value = Approval(j)
            oDTProps.ItemByPropId(PropertiesForDesignTrackingPropertiesEnum.kCreationDateDesignTrackingProperties).Value = drawingDate(j)

            'ProcessingLabel-------------------
            Application.StatusBar = "(" & j & "/" & cnt & ")" & baseNames(j) & "." & searchExt & "を保存中"
            FORM_titleBatchTool.processingLabel.Caption = "(" & j & "/" & cnt & ")" & baseNames(j) & "." & searchExt & "を保存中"
            '--------------------------------------------

            Debug.Print ("これは実行された")
            'Dim oFileSaveAs As Object
            Dim oIsNeedMigrating As Boolean
            oIsNeedMigrating = oDoc.NeedsMigrating
            If oIsNeedMigrating Then
                MsgBox "Migrationが必要なため保存できません"
                 Debug.Print ("Migrationが必要です")
                 Err.Raise vbObjectError + 513, "SetTitleInfo", "Migrationが必要なため保存できません。"

                 Exit Sub
            End If




            Set oFileSaveAs = oInvApp.FileSaveAs
            Debug.Print ("ここ")
           Call oFileSaveAs.AddFileToSave(oDoc, oDoc.FullFileName)
            Debug.Print (oDoc)
            Debug.Print (oDoc.FullFileName)
'            oDoc.PropertySets.FlushToFile
'            oInvApp.FileSaveAs.AddFileToSave

            oFileSaveAs.ExecuteSave

            'ProcessingLabel-------------------
            Application.StatusBar = "(" & j & "/" & cnt & ")" & baseNames(j) & "." & searchExt & "をClose中"
            FORM_titleBatchTool.processingLabel.Caption = "(" & j & "/" & cnt & ")" & baseNames(j) & "." & searchExt & "をClose中"
            '--------------------------------------------
            oDoc.Close
            Set oDoc = Nothing

        Next j

        '後始末--------
        Set fso = Nothing
        Set oInvApp = Nothing
        Application.StatusBar = False
        GoTo finally
    End If

ErrorHandler:
    Debug.Print Err.Description
    '後始末--------
    oDoc.Close
    Set fso = Nothing
    Set oInvApp = Nothing
    msg = Application.StatusBar & "にエラーが発生しました( " & Err.Number & "): "
    Application.StatusBar = False

finally:
    'MsgBox "完了しました"
End Sub

Sub toolOpen()
    FORM_titleBatchTool.Show
End Sub
