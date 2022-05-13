Attribute VB_Name = "M01_Read"
Option Explicit

Sub Select_STN()
    
    Sheets("List").Select
    Range("B3").Select
    
    Call Get_Data
    
End Sub

Sub Get_Data()

Dim strSTN As String
Dim strSNM As String

    strSTN = Sheets("Menu").Range("AI5")
     
    Call �Ј��Ǎ�(strSTN)
    
End Sub


Sub �Ј��Ǎ�(strKBN As String)

Const SQL1 = "SELECT * FROM �O���[�v�Ј��}�X�^�[ WHERE (((���Ə��敪)='"
Const SQL2 = "')) ORDER BY ���� DESC, �Ј����, �Ј��R�[�h"
Const SQL2T = "')) ORDER BY ���� DESC, ���� DESC, �Ј��R�[�h"

Dim cnA    As New ADODB.Connection
Dim rsA    As New ADODB.Recordset
Dim Cmd    As New ADODB.Command

Dim strSQL As String
Dim strUNM As String
Dim strDB  As String
Dim lngR   As Long
Dim lngC   As Long
Dim P_Hant As String
    
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
    Set Cmd.ActiveConnection = cnA
    
    '�Ј�������
    Call CLR_CELL          '�ް���ĸر
        
    strSQL = SQL1 & strKBN & SQL2
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    If rsA.EOF = False Then rsA.MoveFirst
    lngR = 7
    Do Until rsA.EOF
        If Trim(rsA![�Ǘ��E��] & "") <> "����" Then '��ʎЈ�
            Cells(lngR, 2) = rsA.Fields("���Ə��敪")
            Cells(lngR, 3) = rsA.Fields("�Ј��R�[�h")
            Cells(lngR, 4) = rsA.Fields("�Ј���")
            If rsA.Fields("����") = "�j" Then
                Cells(lngR, 5) = "M"
            Else
                Cells(lngR, 5) = "W"
            End If
            Cells(lngR, 7) = rsA.Fields("���N����")
            Cells(lngR, 10) = rsA.Fields("���ДN����")
            Cells(lngR, 11) = rsA.Fields("�Ј����")
            Cells(lngR, 12) = rsA.Fields("����")
            Cells(lngR, 14) = rsA.Fields("����")
            Cells(lngR, 15) = �Ǘ��E��T��(rsA.Fields("�Ǘ��E��") & "")
            Cells(lngR, 17) = rsA.Fields("��{���P") '�{��
            Cells(lngR, 18) = rsA.Fields("��{���Q") '����
            Cells(lngR, 19) = rsA.Fields("�Ǘ��E�蓖")
            Cells(lngR, 20) = rsA.Fields("�Ƒ��蓖")
            Cells(lngR, 21) = rsA.Fields("��s�s�Ζ��蓖")
            Cells(lngR, 22) = rsA.Fields("�����蓖") '�Ɛю蓖
            Cells(lngR, 23) = rsA.Fields("�����Ǝ蓖")
            Cells(lngR, 24) = "=SUM(RC[-7]:RC[-1])"
            Cells(lngR, 25) = rsA.Fields("�������")
            Cells(lngR, 26) = rsA.Fields("�������Ə�")
            Cells(lngR, 29) = rsA.Fields("�p�[�g���莞�Ԑ�")
        End If
        rsA.MoveNext
        lngR = lngR + 1
        If lngR = 55 Then lngR = 67
        If lngR > 113 Then Exit Do
    Loop
       
    Range("A2").Select

    '�ڑ��̃N���[�Y
    rsA.Close
    cnA.Close

Exit_DB:

    '�I�u�W�F�N�g�̔j��
    Set rsA = Nothing
    Set cnA = Nothing
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
    Range("B7:E53,G7:H53,J7:L53,N7:O53,Q7:W53,Y7:AA53,AC7:AC53").Select
    Selection.ClearContents
    Range("AG7:AR44").Select
    Selection.ClearContents
    '�Q���ڃN���A
    Range("B67:E113,G67:H113,J67:L113,N67:O113,Q67:W113,Y67:AA113,AC67:AC113").Select
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

