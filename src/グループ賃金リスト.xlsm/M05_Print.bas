Attribute VB_Name = "M05_Print"
Option Explicit

'********************** �V�[�g�̈�� *********************************

Sub �ʈ��A4()

Dim lngP   As Long    '�y�[�W��

    lngP = InputBox("�����������͂��ĉ�����", "", "1")
    If lngP = 0 Then Exit Sub
    Call Print_OK(lngP, False, "A4")
    
End Sub

Sub �ʈ��B4()

Dim lngP   As Long    '�y�[�W��

    lngP = InputBox("�����������͂��ĉ�����", "", "1")
    If lngP = 0 Then Exit Sub
    Call Print_OK(lngP, False, "B4")
    
End Sub

Sub �ʈ���N��t()

Dim lngP   As Long    '�y�[�W��

    lngP = InputBox("�����������͂��ĉ�����", "", "1")
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

    '�P�Ŗڈ��
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
    
    '�Q�Ŗڈ��
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

'�N��t�^
Sub Age_Add()

Dim strB  As String
Dim strD  As String
Dim lngR  As Long
    

    strD = Format(Date, "yyyymmdd")
    strD = InputBox("�N��v�Z�̊�ɂ�������w�肵�Ă�������", "�������", strD)
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
    
'strBirth --- ���N����(yyyymmdd)
'strBuff ---- ���(yyyymmdd)

Dim intAgeYear      As Integer
Dim intAgeMonth     As Integer

    intAgeYear = CInt(Mid(strBuff, 1, 4)) - CInt(Mid(strBirth, 1, 4))
    '�a�������߂��Ă��邩����i�����|�a�������O�ȏ�j
    If (CInt(Mid(strBuff, 5, 2)) - CInt(Mid(strBirth, 5, 2))) >= 0 Then
        '�����|�a����
        intAgeMonth = CInt(Mid(strBuff, 5, 2)) - CInt(Mid(strBirth, 5, 2))
    Else
        '�����|�a�����{�P�Q�A�N�|�P
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
    AgeCal = CStr(intAgeYear) & "��" & CStr(intAgeMonth) & "����"
End Function
