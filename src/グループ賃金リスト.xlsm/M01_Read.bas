Attribute VB_Name = "M01_Read"
Option Explicit

Sub Select_STN()

'=======================================
'���j���[��ʂŎ��Ə���I���������̏���
'=======================================
    
    Sheets("List").Select 'List�V�[�g�ֈړ�
    Range("B3").Select
    
    Call Get_Data '�f�[�^�ǂݍ���
    
End Sub

Sub Get_Data()

'=======================================
'�����f�[�^�ǂݍ���
'=======================================

    Dim strSTN As String

    strSTN = Sheets("Menu").Range("AI5") '���_�敪�擾(RH,RO,RT,TA,KA)
     
    Call �����Ǎ�(strSTN)
    
End Sub


Sub �����Ǎ�(strKBN As String)

Dim cnA    As New ADODB.Connection
Dim rsA    As New ADODB.Recordset
Dim Cmd    As New ADODB.Command
Dim strSQL As String
Dim strNT As String
Dim lngR   As Long
    
    '�Ј�������
    Call CLR_CELL          '�ް���ĸر
    
    strNT = "Initial Catalog=KYUYO;"
    cnA.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD 'SQLServer
    cnA.Open
    Set Cmd.ActiveConnection = cnA
    strSQL = ""
    strSQL = strSQL & "SELECT *"
    strSQL = strSQL & "     FROM KYUMTA"
    strSQL = strSQL & "        WHERE KBN ='" & strKBN & "'"
    strSQL = strSQL & "        AND DATKB ='1'"
    strSQL = strSQL & "     ORDER BY CLASS DESC,"
    strSQL = strSQL & "              ISSUE DESC,"
    strSQL = strSQL & "              SKBN,"
    strSQL = strSQL & "              SCODE"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    If rsA.EOF = False Then rsA.MoveFirst
    '��Ă��ް��\��t��
    lngR = 7
    Do Until rsA.EOF
        If Trim(rsA![MGR] & "") <> "����" Then '��ʎЈ�
            Cells(lngR, 2) = rsA.Fields("KBN")
            Cells(lngR, 3) = rsA.Fields("SCODE")
            Cells(lngR, 4) = rsA.Fields("SNAME")
            Cells(lngR, 5) = Trim(rsA.Fields("SEX"))
            Cells(lngR, 7) = Format(rsA.Fields("DATE1"), "ggge�Nm��d��")
            Cells(lngR, 8) = Format(rsA.Fields("DATE2"), "ggge�Nm��d��")
            Cells(lngR, 9) = rsA.Fields("SKBN")
            Cells(lngR, 10) = rsA.Fields("CLASS")
            Cells(lngR, 12) = rsA.Fields("ISSUE")
            Cells(lngR, 13) = �Ǘ��E��T��(Trim(rsA.Fields("MGR")) & "")
            Cells(lngR, 15) = rsA.Fields("PAY1") '�{��
            Cells(lngR, 16) = rsA.Fields("PAY2") '����
            Cells(lngR, 17) = rsA.Fields("OPT1")
            Cells(lngR, 18) = rsA.Fields("OPT2")
            Cells(lngR, 19) = rsA.Fields("OPT3")
            Cells(lngR, 20) = rsA.Fields("OPT4") '�Ɛю蓖
            Cells(lngR, 21) = rsA.Fields("OPT5")
            Cells(lngR, 22) = "=SUM(RC[-7]:RC[-1])"
            Cells(lngR, 23) = rsA.Fields("PRN")
            Cells(lngR, 24) = rsA.Fields("OFFICE")
            Cells(lngR, 27) = rsA.Fields("HOUR")
        End If
        rsA.MoveNext
        lngR = lngR + 1
        If lngR = 54 Then lngR = 66
        If lngR > 112 Then Exit Do
    Loop
       
    Range("A2").Select

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

Sub CLR_CELL()
    '�P���ڃN���A
    Range("B7:E53,G7:J53,L7:M53,O7:U53,W7:Y53,AA7:AA53").Select
    Selection.ClearContents
    '�Q���ڃN���A
    Range("B66:E112,G66:J112,L66:M112,O66:U112,W66:Y112,AA66:AA112").Select
    Selection.ClearContents
    Range("A1").Select
End Sub

Function �Ǘ��E��T��(strK As String)

    Select Case strK
        Case "����"
            �Ǘ��E��T�� = "YY"
        Case "�x�X��"
            �Ǘ��E��T�� = "SS"
        Case "����"
            �Ǘ��E��T�� = "BB"
        Case "����"
            �Ǘ��E��T�� = "JJ"
        Case "�ے�"
            �Ǘ��E��T�� = "KK"
        Case "��C"
            �Ǘ��E��T�� = "KS"
        Case "�ے��㗝"
            �Ǘ��E��T�� = "HD"
        Case "�W��"
            �Ǘ��E��T�� = "HK"
        Case "�ǒ�"
            �Ǘ��E��T�� = "HH"
        Case Else
            �Ǘ��E��T�� = ""
    End Select
    
End Function

