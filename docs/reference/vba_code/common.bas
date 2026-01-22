Option Explicit

Public Sub InitializeTitle(hRs As Long)
    'Worksheetの初期化----------------------
    'Worksheetのセルの値をすべてクリア
    Worksheets("iProperty一覧").Cells.Clear
    'Workshettにタイトルを記入
    Cells(hRs + 1, 1) = "ファイルパス"
    Cells(hRs + 1, 2) = "ファイル名"
    Cells(hRs + 1, 3) = "会社名1"
    Cells(hRs + 1, 4) = "会社名2"
    Cells(hRs + 1, 5) = "名称1"
    Cells(hRs + 1, 6) = "名称2"
    Cells(hRs + 1, 7) = "図番"
    Cells(hRs + 1, 8) = "決定No"
    Cells(hRs + 1, 9) = "製図"
    Cells(hRs + 1, 10) = "設計"
    Cells(hRs + 1, 11) = "検図"
    Cells(hRs + 1, 12) = "承認"
    Cells(hRs + 1, 13) = "作成日"
    'タイトル行を色付
    Dim i As Long
    For i = 1 To 13
        'Cells(hRs + 1, i).Interior.Color = RGB(252, 228, 214)
        Cells(hRs + 1, i).Interior.Color = &HD0CECE
    Next i
    '-------------------------------------------
    With Cells(1, 2)
        .Value = "." & A_Main.SEARCHEXTENSION & "表題欄一括変更ツール"
    End With
    With Cells(1, 3)
        .Value = "ver."
        .HorizontalAlignment = xlRight
    End With
    With Cells(1, 4)
        .Value = SOFTWAREVERSION
        .HorizontalAlignment = xlLeft
    End With

    'ウインドウ枠の固定---------
    ActiveWindow.FreezePanes = False
    Range("B7").Select
    ActiveWindow.FreezePanes = True
End Sub
