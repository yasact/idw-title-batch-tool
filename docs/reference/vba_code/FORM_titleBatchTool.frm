Option Explicit

Private Sub closeButton_Click()
    Unload FORM_titleBatchTool
End Sub

Private Sub dirSelectButton_Click()
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = ThisWorkbook.Path
        Debug.Print (ThisWorkbook.Path)
        If .Show = True Then
            dirPathLabel.Caption = .SelectedItems(1)
            Debug.Print (dirPathLabel.Caption)
        End If
    End With
End Sub

Private Sub getterButton_Click()
    Dim ans As Long
    Dim myPath As String
    ' myPath = ThisWorkbook.Path
    myPath = dirPathLabel.Caption

    ans = MsgBox(myPath & vbCrLf & "の." & A_Main.SEARCHEXTENSION & "を読み込みます。", vbOKCancel, "読み込み確認")

    If ans = vbOK Then
        '処理中はボタン類を隠す-----
        getterButton.Visible = False
        setterButton.Visible = False
        closeButton.Visible = False
        processingLabel.Visible = True

        Call A_Main.GetTitleInfo

        '処理終えたのでボタン類再表示
        getterButton.Visible = True
        setterButton.Visible = True
        closeButton.Visible = True
        processingLabel.Visible = False

    ElseIf ans = vbCancel Then
        Exit Sub
    End If

End Sub



Private Sub setterButton_Click()

    Dim ans As Long
    Dim myPath As String
    'myPath = ThisWorkbook.Path
    myPath = dirPathLabel.Caption
    ans = MsgBox(myPath & vbCrLf & "の." & A_Main.SEARCHEXTENSION & "に書き込みます。", vbOKCancel, "書き込み確認")

    If ans = vbOK Then
        '処理中はボタン類を隠す-----
        getterButton.Visible = False
        setterButton.Visible = False
        closeButton.Visible = False
        processingLabel.Visible = True

        Call A_Main.SetTitleInfo

        '処理終えたのでボタン類再表示
        getterButton.Visible = True
        setterButton.Visible = True
        closeButton.Visible = True
        processingLabel.Visible = False

    ElseIf ans = vbCancel Then
        Exit Sub

    End If

End Sub

Private Sub UserForm_Activate()
    Dim myPath As String
    myPath = ThisWorkbook.Path
    FORM_titleBatchTool.dirPathLabel.Caption = myPath
End Sub
