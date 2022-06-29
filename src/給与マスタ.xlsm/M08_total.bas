Attribute VB_Name = "M08_total"
Option Explicit

Sub Get_Total()

Dim cnA As New ADODB.Connection
Dim rsA As New ADODB.Recordset
Dim strSQL As String
Dim strNT  As String
Dim strR   As String
Dim DateA  As Date
Dim DateB  As Date
Dim strEJ  As String

    Range("D3:E12").ClearContents
    Range("G3:G12").ClearContents

    strNT = "Initial Catalog=KYUYO;"
    cnA.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD 'SQLServer
    cnA.Open
    
    '�O���[�v�Ј��}�X�^�[�Ǎ���
    strR = Range("M1")
    If strR = "ALL" Then
         strSQL = ""
         strSQL = strSQL & "SELECT SEX,"
         strSQL = strSQL & "       SKBN,"
         strSQL = strSQL & "       MGR,"
         strSQL = strSQL & "       BMNNM,"
         strSQL = strSQL & "       DATE1,"
         strSQL = strSQL & "       DATE2"
         strSQL = strSQL & "    FROM KYUMTA"
         strSQL = strSQL & "         WHERE Left(KBN,1) = 'R'"
         strSQL = strSQL & "         AND MGR <> '����'"
    Else
        strSQL = ""
         strSQL = strSQL & "SELECT SEX,"
         strSQL = strSQL & "       SKBN,"
         strSQL = strSQL & "       MGR,"
         strSQL = strSQL & "       BMNNM,"
         strSQL = strSQL & "       DATE1,"
         strSQL = strSQL & "       DATE2"
         strSQL = strSQL & "    FROM KYUMTA"
         strSQL = strSQL & "         WHERE KBN = '" & strR & "'"
    End If
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly

    If rsA.EOF = False Then rsA.MoveFirst
    Do Until rsA.EOF
        If InStr(1, rsA.Fields("BMNNM"), "�c��") <> 0 Then
            strEJ = "E"
        ElseIf InStr(1, rsA.Fields("BMNNM"), "�J��") <> 0 Then
            strEJ = "E"
         ElseIf InStr(1, rsA.Fields("BMNNM"), "�f��") <> 0 Then
            strEJ = "E"
        ElseIf InStr(1, rsA.Fields("BMNNM"), "���H") <> 0 Then
            strEJ = "K"
        Else
            strEJ = "J"
        End If
        If rsA.Fields("SKBN") = "A" Or rsA.Fields("SKBN") = "B" Then  '���Ј�
            If strEJ = "E" Then
                Range("D3") = Range("D3") + 1
                Range("E3") = Range("E3") + GetAge(rsA.Fields(4), Date)
                Range("G3") = Range("G3") + GetAge(rsA.Fields(5), Date)
            ElseIf strEJ = "J" Then
                If rsA.Fields("SEX") = "M" Then
                    Range("D4") = Range("D4") + 1
                    Range("E4") = Range("E4") + GetAge(rsA.Fields(4), Date)
                    Range("G4") = Range("G4") + GetAge(rsA.Fields(5), Date)
                Else
                    Range("D5") = Range("D5") + 1
                    Range("E5") = Range("E5") + GetAge(rsA.Fields(4), Date)
                    Range("G5") = Range("G5") + GetAge(rsA.Fields(5), Date)
                End If
            ElseIf strEJ = "K" Then
                If rsA.Fields("SEX") = "M" Then
                    Range("D6") = Range("D6") + 1
                    Range("E6") = Range("E6") + GetAge(rsA.Fields(4), Date)
                    Range("G6") = Range("G6") + GetAge(rsA.Fields(5), Date)
                Else
                    Range("D7") = Range("D7") + 1
                    Range("E7") = Range("E7") + GetAge(rsA.Fields(4), Date)
                    Range("G7") = Range("G7") + GetAge(rsA.Fields(5), Date)
                End If
            End If
        Else
            If strEJ = "E" Then
                Range("D8") = Range("D8") + 1
                Range("E8") = Range("E8") + GetAge(rsA.Fields(4), Date)
                Range("G8") = Range("G8") + GetAge(rsA.Fields(5), Date)
            ElseIf strEJ = "J" Then
                If rsA.Fields("SEX") = "M" Then
                    Range("D9") = Range("D9") + 1
                    Range("E9") = Range("E9") + GetAge(rsA.Fields(4), Date)
                    Range("G9") = Range("G9") + GetAge(rsA.Fields(5), Date)
                Else
                    Range("D10") = Range("D10") + 1
                    Range("E10") = Range("E10") + GetAge(rsA.Fields(4), Date)
                    Range("G10") = Range("G10") + GetAge(rsA.Fields(5), Date)
                End If
            ElseIf strEJ = "K" Then
                If rsA.Fields("SEX") = "M" Then
                    Range("D11") = Range("D11") + 1
                    Range("E11") = Range("E11") + GetAge(rsA.Fields(4), Date)
                    Range("G11") = Range("G11") + GetAge(rsA.Fields(5), Date)
                Else
                    Range("D12") = Range("D12") + 1
                    Range("E12") = Range("E12") + GetAge(rsA.Fields(4), Date)
                    Range("G12") = Range("G12") + GetAge(rsA.Fields(5), Date)
                End If
            End If
        End If

        rsA.MoveNext
    Loop

Exit_DB:

   If Not rsA Is Nothing Then
        If rsA.State = adStateOpen Then rsA.Close
        Set rsA = Nothing
    End If
    If Not cnA Is Nothing Then
        If cnA.State = adStateOpen Then cnA.Close
        Set cnA = Nothing
    End If

End Sub


Sub Get_���ݒn()

Dim cnA As New ADODB.Connection
Dim rsA As New ADODB.Recordset
Dim strSQL As String
Dim strNT  As String
Dim DateA  As Date
Dim DateB  As Date
Dim strEJ  As String '�E�픻��
Dim lngSZ  As Long   '���ݒn����

    Range("I6:P11").ClearContents

    strNT = "Initial Catalog=KYUYO;"
    cnA.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD 'SQLServer
    cnA.Open

    '�O���[�v�Ј��}�X�^�[�Ǎ���
    strSQL = ""
    strSQL = strSQL & "SELECT MGR,"
    strSQL = strSQL & "       BMNNM,"
    strSQL = strSQL & "       SKBN,"
    strSQL = strSQL & "       SEX,"
    strSQL = strSQL & "       SNAME"
    strSQL = strSQL & "    FROM KYUMTA"
    strSQL = strSQL & "        WHERE Left(KBN,1) = 'R'"
    strSQL = strSQL & "        AND MGR <> '����'"
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly

    If rsA.EOF = False Then rsA.MoveFirst
    Do Until rsA.EOF
        '�E�픻��
        If InStr(1, rsA.Fields(1), "�c��") <> 0 Then
            strEJ = "E"
        ElseIf InStr(1, rsA.Fields(1), "�J��") <> 0 Then
            strEJ = "E"
         ElseIf InStr(1, rsA.Fields(1), "�f��") <> 0 Then
            strEJ = "E"
        ElseIf InStr(1, rsA.Fields(1), "���H") <> 0 Then
            strEJ = "K"
        Else
            strEJ = "J"
        End If
        '���ݒn����
        Select Case Left(rsA.Fields(1), 2)
            Case "����"
                lngSZ = 7
            Case "����"
                lngSZ = 8
            Case "����"
                lngSZ = 9
            Case "���"
                lngSZ = 10
            Case "���"
                lngSZ = 11
            Case Else
                lngSZ = 6
        End Select
        '�Z���ɃZ�b�g
        If rsA.Fields(2) = "A" Or rsA.Fields(2) = "B" Then
            If strEJ = "E" Then
                Cells(lngSZ, 9) = Cells(lngSZ, 9) + 1
            ElseIf strEJ = "J" Then
                If rsA.Fields(3) = "M" Then
                    Cells(lngSZ, 10) = Cells(lngSZ, 10) + 1
                Else
                    Cells(lngSZ, 11) = Cells(lngSZ, 11) + 1
                End If
            ElseIf strEJ = "K" Then
                If rsA.Fields(3) = "M" Then
                    Cells(lngSZ, 12) = Cells(lngSZ, 12) + 1
                Else
                    MsgBox "���H�ۂ̐��Ј��ŏ���������悤�ł��B�\���g�����ĉ������B"
                End If
            End If
        Else
            If strEJ = "E" Then
                Cells(lngSZ, 9) = Cells(lngSZ, 9) + 1
            ElseIf strEJ = "J" Then
                If rsA.Fields(3) = "M" Then
                    Cells(lngSZ, 13) = Cells(lngSZ, 13) + 1
                Else
                    Cells(lngSZ, 14) = Cells(lngSZ, 14) + 1
                End If
            ElseIf strEJ = "K" Then
                If rsA.Fields(3) = "M" Then
                    Cells(lngSZ, 15) = Cells(lngSZ, 15) + 1
                Else
                    Cells(lngSZ, 16) = Cells(lngSZ, 16) + 1
                End If
            End If
        End If
        rsA.MoveNext
    Loop

Exit_DB:

    If Not rsA Is Nothing Then
        If rsA.State = adStateOpen Then rsA.Close
        Set rsA = Nothing
    End If
    If Not cnA Is Nothing Then
        If cnA.State = adStateOpen Then cnA.Close
        Set cnA = Nothing
    End If

End Sub

Sub NenreiSample()

    Dim dBirthday As Date
    Dim lAge      As Long


    lAge = GetAge(dBirthday, Date)


End Sub

Public Function GetAge(Birthday As Date, BaseDate As Date) As Long

    Dim lAge As Long

    '����
    '  Birthday : �a����(���t/�����^)
    '  DateNow  : ���(���t/�����^)
    '�߂�l
    '  �N��(�������^)

    lAge = DateDiff("yyyy", Birthday, BaseDate) + (Format(Birthday, "mmdd") > Format(BaseDate, "mmdd"))

    GetAge = lAge

End Function

