Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim hRs As Long                              ' HeaderRowSize
    hRs = A_Main.HEADERROWSIZE
    InitializeTitle (hRs)
    ThisWorkbook.Save
End Sub

Private Sub Workbook_Open()
    FORM_titleBatchTool.Show
End Sub
