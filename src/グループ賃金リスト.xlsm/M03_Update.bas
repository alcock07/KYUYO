Attribute VB_Name = "M03_Update"
Option Explicit

Sub �Ј��X�V()

Const SQL1 = "SELECT * FROM �O���[�v�Ј��}�X�^�[ WHERE (((���Ə��敪)='"
Const SQL2 = "') AND ((�Ј��R�[�h)='"
Const SQL3 = "'))"

Const SQL4 = "SELECT * FROM �����{���\ WHERE (((����)="
Const SQL5 = ") AND ((����)="

Dim cnA      As New ADODB.Connection
Dim rsA      As New ADODB.Recordset
Dim rsT      As New ADODB.Recordset
Dim strSQL   As String
Dim strUNM   As String
Dim strKBN   As String
Dim strDB    As String
Dim lngR     As Long   '�s����
Dim lngC     As Long   '����
Dim strKey1  As String '���Ə��敪
Dim strKey2  As String '�Ј��R�[�h
Dim strDel   As String '�폜��
Dim strDAT1  As String '���N����
Dim strDAT2  As String '���ДN����
Dim DateS    As Date   '���N����2
Dim DateN    As Date   '���ДN����2
Dim DateA    As Date   '��Ɨp�ϐ�
Dim lngTKY   As Long
Dim lngGSU   As Long

    strKBN = Sheets("Menu").Range("AI5")

    '�Ј����X�V
    strUNM = Strings.UCase(GetUserNameString)
    If strUNM = "SCOTT" Or strUNM = "TAKA" Or strUNM = "SIMO" Then
        If strKBN = "TA" Or strKBN = "KA" Then
            strDB = dbT
        Else
            strDB = dbK
        End If
    Else
        Call Back_Menu
        GoTo Exit_DB
    End If
    cnA.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & strDB
    cnA.Open
    lngR = 7
    Do
        If Cells(lngR, 2) = "" Then Exit Do
        strDel = StrConv(Left(Cells(lngR, 27), 1), vbUpperCase)  '�폜��
        strKey1 = StrConv(Left(Cells(lngR, 2), 2), vbUpperCase)  '���Ə��敪
        strKey2 = Format(Cells(lngR, 3), "00000")                '�Ј�����
        lngTKY = Cells(lngR, 12)                                 '����
        lngGSU = Cells(lngR, 14)                                 '����
        '�{���\�ƃ`�F�b�N
        If Cells(lngR, 11) = "A" Then
            strSQL = SQL4 & lngTKY & SQL5 & lngGSU & "))"
            rsT.Open strSQL, cnA, adOpenStatic, adLockPessimistic
            If Cells(lngR, 17) = rsT.Fields("�{��") And Cells(lngR, 18) = rsT.Fields("����") Then
            Else
                MsgBox "�{�����邢�͉������Ԉ���Ă��܂��I" & vbCrLf & "������ " & Cells(lngR, 4) & " ������", vbCritical
            End If
            rsT.Close
        End If
        
        strSQL = SQL1 & strKey1 & SQL2 & strKey2 & SQL3
        rsA.Open strSQL, cnA, adOpenStatic, adLockPessimistic
        If strDel = "D" Then '�폜����
            If rsA.EOF Then
                strDel = ""
            Else
                rsA.Delete
            End If
        Else
            '���N����
            If Cells(lngR, 7) = "" Then
                strDAT1 = ""
            Else
                DateS = Cells(lngR, 7)
                strDAT1 = Format(DateA, "yyyymmdd")
                
            End If
            '���ДN����
            If Cells(lngR, 10) = "" Then
                strDAT1 = ""
            Else
                DateN = Cells(lngR, 10)
                strDAT2 = Format(DateA, "yyyymmdd")
            End If
            If rsA.EOF Then
                '�}�X�^�ɖ�����Βǉ�
                rsA.AddNew
                rsA.Fields(0) = strKey1
                rsA.Fields(1) = strKey2
                rsA.Fields(20) = 0
            End If
            If rsA.EOF Then
                rsA.AddNew
                rsA.Fields(0) = Cells(lngR, 2)
                rsA.Fields(1) = Cells(lngR, 3)
            Else
                rsA.MoveFirst
            End If
            rsA.Fields("�Ј���") = Cells(lngR, 4)
            rsA.Fields("����") = Cells(lngR, 6)
            rsA.Fields("���N����") = DateS
            rsA.Fields("���ДN����") = DateN
            rsA.Fields("�Ј����") = Cells(lngR, 11)
            rsA.Fields("����") = Cells(lngR, 12)
            rsA.Fields("����") = Cells(lngR, 14)
            rsA.Fields("�Ǘ��E��") = Cells(lngR, 16)
            For lngC = 9 To 15
                If Cells(lngR, lngC + 8) = "" Then '�{��->�蓖
                    rsA.Fields(lngC) = 0
                Else
                    rsA.Fields(lngC) = Cells(lngR, lngC + 8)
                End If
            Next lngC
            rsA.Fields("�������") = Cells(lngR, 25)
            rsA.Fields("�������Ə�") = Cells(lngR, 26)
            rsA.Fields("�����A�C��") = ""
            rsA.Fields("�p�[�g���ԋ�") = 0
            rsA.Fields("�p�[�g���莞�Ԑ�") = 0
            If StrConv(rsA![�Ј����], vbUpperCase) = "P" Then
                rsA.Fields("�p�[�g���ԋ�") = Cells(lngR, 17) / Cells(lngR, 29) '�{�� / �߰ď��莞��
                rsA.Fields("�p�[�g���莞�Ԑ�") = Cells(lngR, 29) '�߰ď��莞��
            End If
            rsA.Update
        End If
        rsA.Close
        lngR = lngR + 1
        If lngR = 55 Then lngR = 67
        If lngR > 113 Then Exit Do
    Loop
    
    MsgBox "�X�V���܂����I(*'��'*)", vbInformation, "�}�X�^�X�V"
    
Exit_DB:
    cnA.Close
    Set rsA = Nothing
    Set cnA = Nothing
    
End Sub
