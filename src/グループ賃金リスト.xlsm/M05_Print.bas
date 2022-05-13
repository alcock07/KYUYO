Attribute VB_Name = "M05_Print"
Option Explicit

'********************** シートの印刷 *********************************

Sub 個別印刷A4()

Dim lngP   As Long    'ページ数

    lngP = InputBox("印刷枚数を入力して下さい", "", "1")
    If lngP = 0 Then Exit Sub
    Call Print_OK(lngP, False, "A4")
    
End Sub

Sub 個別印刷B4()

Dim lngP   As Long    'ページ数

    lngP = InputBox("印刷枚数を入力して下さい", "", "1")
    If lngP = 0 Then Exit Sub
    Call Print_OK(lngP, False, "B4")
    
End Sub

Sub 個別印刷年齢付()

Dim lngP   As Long    'ページ数

    lngP = InputBox("印刷枚数を入力して下さい", "", "1")
    If lngP = 0 Then Exit Sub
    
    Columns("Y:AD").Select
    Selection.ColumnWidth = 0
    
    Call Print_OK(lngP, True, "B4")
    
    Columns("Y:AD").Select
    Selection.ColumnWidth = 9
    Range("A1").Select
    
End Sub

Sub Print_OK(lngP As Long, boolA As Boolean, strPS As String)
Attribute Print_OK.VB_ProcData.VB_Invoke_Func = " \n14"

    '１頁目印刷
    Range("A6:Z54").Interior.ColorIndex = xlNone
    If boolA Then
        ActiveSheet.PageSetup.PrintArea = "$B$3:$AE$56"
    Else
        ActiveSheet.PageSetup.PrintArea = "$B$3:$Z$56"
    End If
    If strPS = "A4" Then
        With ActiveSheet.PageSetup
            .Zoom = 70
            .PaperSize = xlPaperA4
        End With
    Else
        With ActiveSheet.PageSetup
            .Zoom = 88
            .PaperSize = xlPaperB4
        End With
    End If
    ActiveWindow.SelectedSheets.PrintOut Copies:=lngP, Collate:=True
    
    '２頁目印刷
    If Range("B67").Value <> "" Then
        If boolA Then
            ActiveSheet.PageSetup.PrintArea = "$B$63:$AE$115"
        Else
            ActiveSheet.PageSetup.PrintArea = "$B$63:$Z$115"
        End If
       If strPS = "A4" Then
        With ActiveSheet.PageSetup
            .Zoom = 70
            .PaperSize = xlPaperA4
        End With
    Else
        With ActiveSheet.PageSetup
            .Zoom = 88
            .PaperSize = xlPaperB4
        End With
    End If
        ActiveWindow.SelectedSheets.PrintOut Copies:=lngP, Collate:=True
    End If

End Sub

'年齢付与
Sub Age_Add()

Dim strB  As String
Dim strD  As String
Dim lngR  As Long
    

    strD = Format(Date, "yyyymmdd")
    strD = InputBox("年齢計算の基準にする日を指定してください", "基準日入力", strD)
    For lngR = 7 To 54
        strB = Cells(lngR, 7)
        If strB = "" Then
            Cells(lngR, 31) = ""
        Else
            Cells(lngR, 31) = AgeCal(Format(strB, "yyyymmdd"), strD)
        End If
    Next lngR
    For lngR = 67 To 113
        strB = Cells(lngR, 7)
        If strB = "" Then
            Cells(lngR, 31) = ""
        Else
            Cells(lngR, 31) = AgeCal(Format(strB, "yyyymmdd"), strD)
        End If
    Next lngR
    
End Sub

Public Function AgeCal(strBirth As String, strBuff As String) As String
    
'strBirth --- 生年月日(yyyymmdd)
'strBuff ---- 基準日(yyyymmdd)

Dim intAgeYear      As Integer
Dim intAgeMonth     As Integer

    intAgeYear = CInt(Mid(strBuff, 1, 4)) - CInt(Mid(strBirth, 1, 4))
    '誕生月が過ぎているか判定（今月－誕生月＝０以上）
    If (CInt(Mid(strBuff, 5, 2)) - CInt(Mid(strBirth, 5, 2))) >= 0 Then
        '今月－誕生月
        intAgeMonth = CInt(Mid(strBuff, 5, 2)) - CInt(Mid(strBirth, 5, 2))
    Else
        '今月－誕生月＋１２、年－１
        intAgeMonth = CInt(Mid(strBuff, 5, 2)) - CInt(Mid(strBirth, 5, 2)) + 12
        intAgeYear = intAgeYear - 1
    End If
    If (CInt(Mid(strBuff, 5, 2)) - CInt(Mid(strBirth, 5, 2))) = 0 Then
        If (CInt(Mid(strBuff, 7, 2)) - CInt(Mid(strBirth, 7, 2))) < 0 Then
            intAgeMonth = intAgeMonth - 1
            If (intAgeMonth - 1) < 0 Then
                intAgeYear = intAgeYear - 1
                intAgeMonth = 11
            End If
        End If
    End If
    AgeCal = CStr(intAgeYear) & "歳" & CStr(intAgeMonth) & "ヶ月"
End Function
