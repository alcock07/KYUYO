Attribute VB_Name = "M03_Update"
Option Explicit

Sub �Ј��X�V()

'Const SQL1 = "SELECT * FROM �O���[�v�Ј��}�X�^�[ WHERE (((���Ə��敪)='"
'Const SQL2 = "') AND ((�Ј��R�[�h)='"
'Const SQL3 = "'))"
'
'Const SQL4 = "SELECT * FROM �����{���\ WHERE (((����)="
'Const SQL5 = ") AND ((����)="

Dim cnA      As New ADODB.Connection
Dim rsA      As New ADODB.Recordset
Dim Cmd      As New ADODB.Command
Dim rsT      As New ADODB.Recordset
Dim strSQL   As String
Dim strNT    As String
Dim lngR     As Long   '�s����
Dim lngC     As Long   '����
Dim strKey1  As String '���Ə��敪
Dim strKey2  As String '�Ј��R�[�h
Dim strDel   As String '�폜��
Dim lngTKY   As Long   '����
Dim lngGSU   As Long   '����

    '�Ј����X�V
    strNT = "Initial Catalog=KYUYO;"
    cnA.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD 'SQLServer
    cnA.Open
    Set Cmd.ActiveConnection = cnA
    lngR = 7
    Do
        If Cells(lngR, 2) = "" Then Exit Do
        strDel = StrConv(Left(Cells(lngR, 27), 1), vbUpperCase)  '�폜��
        strKey1 = StrConv(Left(Cells(lngR, 2), 2), vbUpperCase)  '���Ə��敪
        strKey2 = Format(Cells(lngR, 3), "00000")                '�Ј�����
        lngTKY = Cells(lngR, 10)                                 '����
        lngGSU = Cells(lngR, 12)                                 '����
        '�{���\�ƃ`�F�b�N
        If Cells(lngR, 11) = "A" Then
            strSQL = ""
            strSQL = strSQL & "SELECT * "
            strSQL = strSQL & "       FROM KYUHYO"
            strSQL = strSQL & "            WHERE CLASS = '" & lngTKY & "'"
            strSQL = strSQL & "            AND ISSUE = '" & lngGSU & "'"
            rsT.Open strSQL, cnA, adOpenStatic, adLockPessimistic
            If Cells(lngR, 17) = rsT.Fields("PAY1") And Cells(lngR, 18) = rsT.Fields("PAY2") Then
            Else
                MsgBox "�{�����邢�͉������Ԉ���Ă��܂��I" & vbCrLf & "������ " & Cells(lngR, 4) & " ������", vbCritical
            End If
            If rsT.State = adStateOpen Then rsT.Close
        End If
        
        strSQL = ""
        strSQL = strSQL & "SELECT * "
        strSQL = strSQL & "       FROM KYUMTA"
        strSQL = strSQL & "            WHERE KBN = '" & strKey1 & "'"
        strSQL = strSQL & "            AND SCODE = '" & strKey2 & "'"
        rsA.Open strSQL, cnA, adOpenStatic, adLockPessimistic
        If strDel = "D" Then '�폜����
            If rsA.EOF Then
                strDel = ""
            Else
                rsA.Fields(0) = "1"
            End If
        Else
            If rsA.EOF Then
                '�}�X�^�ɖ�����Βǉ�
                rsA.AddNew
                rsA.Fields("DATKB") = "1"
                rsA.Fields("KBN") = strKey1
                rsA.Fields("SCODE") = strKey2
                rsA.Fields("HOUR") = 0
            End If
            If rsA.EOF Then
                rsA.AddNew
                rsA.Fields("KBN") = Cells(lngR, 2)
                rsA.Fields("SCODE") = Cells(lngR, 3)
            Else
                rsA.MoveFirst
            End If
            rsA.Fields("SNAME") = Cells(lngR, 4)
            rsA.Fields("SEX") = Cells(lngR, 5)
            rsA.Fields("DATE1") = CDate(Cells(lngR, 7))
            rsA.Fields("DATE2") = CDate(Cells(lngR, 8))
            rsA.Fields("SKBN") = Cells(lngR, 9)
            rsA.Fields("CLASS") = Cells(lngR, 10)
            rsA.Fields("ISSUE") = Cells(lngR, 12)
            rsA.Fields("MGR") = Cells(lngR, 14)
            For lngC = 10 To 16
                If Cells(lngR, lngC + 5) = "" Then '�{��->�蓖
                    rsA.Fields(lngC) = 0
                Else
                    rsA.Fields(lngC) = Cells(lngR, lngC + 5)
                End If
            Next lngC
            rsA.Fields("PRN") = Cells(lngR, 23)
            rsA.Fields("OFFICE") = Cells(lngR, 24)
            rsA.Fields("JIKYU") = 0
            rsA.Fields("HOUR") = 0
            If StrConv(rsA![SKBN], vbUpperCase) = "P" Then
                rsA.Fields("JIKYU") = Cells(lngR, 15) / Cells(lngR, 27) '�{�� / �߰ď��莞��
                rsA.Fields("HOUR") = Cells(lngR, 27) '�߰ď��莞��
            End If
            rsA.Update
        End If
        rsA.Close
        lngR = lngR + 1
        If lngR = 54 Then lngR = 66
        If lngR > 112 Then Exit Do
    Loop
    
    MsgBox "�X�V���܂����I(*'��'*)", vbInformation, "�}�X�^�X�V"
    
Exit_DB:

    If Not rsT Is Nothing Then
        If rsT.State = adStateOpen Then rsT.Close
        Set rsT = Nothing
    End If
    If Not rsA Is Nothing Then
        If rsA.State = adStateOpen Then rsA.Close
        Set rsA = Nothing
    End If
    If Not cnA Is Nothing Then
        If cnA.State = adStateOpen Then cnA.Close
        Set cnA = Nothing
    End If
    
End Sub
