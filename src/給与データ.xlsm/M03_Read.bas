Attribute VB_Name = "M03_Read"
Option Explicit

'=== PCA�ް���蒊�o ����===

Sub �f�[�^�Ǎ�()

    '[�~�σf�[�^]����N�����̃f�[�^�𒊏o
    '���喈�ɕ����ăV�[�g�ɓ\��t����
    
    Dim cnA            As New ADODB.Connection
    Dim rsS            As ADODB.Recordset
    Dim strYM          As String '�N��
    Dim strKS          As String
    Dim lngC           As Long   'ٰ�߶���
    Dim lngR           As Long   '  �V
    Dim lngKIN(17, 21) As Long   '���z�i�[�z��
    Dim K_cell         As Range  '���
    
    '������
    Erase lngKIN
    Sheets("����").Select
    Range("C11:S30").ClearContents
    Range("C37:S37").ClearContents
    Range("L2") = 0
    strDate = Range("B4").Value  '�x�����擾
    strYM = Strings.Format(strDate, "yyyy/mm")
    
    '[�~�σf�[�^]�����ް��擾���Ĕz��Ɋi�[�i�x�X���ƂɏW�v�j
    cnA.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & dbW '���^�f�[�^
    cnA.Open
    Set rsS = New ADODB.Recordset
    If Range("C4") = "����" Then
        strKS = "K"
    Else
        strKS = "S"
    End If
    strSQL = ""
    strSQL = strSQL & "SELECT *"
    strSQL = strSQL & "         FROM �~�σf�[�^"
    strSQL = strSQL & "                        WHERE �x���N�� = '" & strYM & "'"
    strSQL = strSQL & "                        And ���^�敪 = '" & strKS & "'"
    strSQL = strSQL & "         ORDER BY ����敪"
    rsS.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsS.EOF = False Then rsS.MoveFirst
    Do Until rsS.EOF
        If rsS!�����x���z <> 0 Then
            '�񔻒�
            lngC = ���唻��(rsS!�Ј��R�[�h)
            If lngC = 0 Then
                Exit Sub
            End If
            '�ް��\��t��
            lngKIN(lngC, 1) = lngKIN(lngC, 1) + 1  '�l��
            lngKIN(lngC, 2) = lngKIN(lngC, 2) + (rsS!�����x���z - rsS!��ېŌ�ʔ�)  '�U���z
            lngKIN(lngC, 4) = lngKIN(lngC, 4) + rsS!��ېŌ�ʔ�  '��ېŌ�ʔ�
            lngKIN(lngC, 5) = lngKIN(lngC, 5) + rsS!���N�ی���
            lngKIN(lngC, 6) = lngKIN(lngC, 6) + rsS!���ی���
            lngKIN(lngC, 7) = lngKIN(lngC, 7) + rsS!�����N���ی���
            lngKIN(lngC, 8) = lngKIN(lngC, 8) + rsS!�m�苒�o�N��
            lngKIN(lngC, 9) = lngKIN(lngC, 9) + rsS!�ٗp�ی���
            lngKIN(lngC, 10) = lngKIN(lngC, 10) + rsS!���򏊓���
            If Range("C4") = "����" Then
                lngKIN(lngC, 11) = lngKIN(lngC, 11) + rsS!�����s����
                lngKIN(lngC, 12) = lngKIN(lngC, 12) + rsS!�ݕt��       '�ݕt��
                lngKIN(lngC, 14) = lngKIN(lngC, 14) + rsS!�N�b�N��     '�N�b�N��
                lngKIN(lngC, 15) = lngKIN(lngC, 15) + rsS!���s�ϗ���   '���s�ϗ���
                lngKIN(lngC, 16) = lngKIN(lngC, 16) + rsS!���`���~     '���`���~
                lngKIN(lngC, 17) = lngKIN(lngC, 17) + rsS!�a���a����� '�ƒ���
                lngKIN(lngC, 18) = lngKIN(lngC, 18) + rsS!���̑��T����
                lngKIN(lngC, 19) = lngKIN(lngC, 19) + rsS!�H����a����
                lngKIN(lngC, 20) = lngKIN(lngC, 20) + rsS!���̑��a�����
            Else
                lngKIN(lngC, 12) = lngKIN(lngC, 12) + rsS!�ݕt��       '�ݕt��
                lngKIN(lngC, 16) = lngKIN(lngC, 16) + rsS!���`���~     '���`���~
            End If
            If rsS!�ٗp�ی��� <> 0 Then lngKIN(lngC, 21) = lngKIN(lngC, 21) + rsS!���x���z '�ٗp�ی��Ώێx���z
        End If
        rsS.MoveNext
    Loop
    
    If Not rsS Is Nothing Then
        If rsS.State = adStateOpen Then rsS.Close
        Set rsS = Nothing
    End If
    If Not cnA Is Nothing Then
        If cnA.State = adStateOpen Then cnA.Close
        Set cnA = Nothing
    End If
    
    '�V�[�g�ɓ\��t�����܂�
    Set K_cell = Sheets("����").Range("C10")
    For lngC = 1 To 17 Step 1
        If lngKIN(lngC, 1) > 0 Then
            For lngR = 1 To 20 Step 1
                If lngKIN(lngC, lngR) = 0 Then
                    K_cell.Offset(lngR, lngC - 1).Value = ""
                Else
                    K_cell.Offset(lngR, lngC - 1).Value = lngKIN(lngC, lngR)
                End If
            Next lngR
            K_cell.Offset(27, lngC - 1).Value = lngKIN(lngC, 21)
        End If
    Next lngC
    Range("C11").Select
    
End Sub

Function ���唻��(strSCD As String) As Long
    
    Dim cnA     As New ADODB.Connection
    Dim rsS     As New ADODB.Recordset
    Dim Cmd     As New ADODB.Command
    Dim strNT   As String
    Dim lngBMN  As Long
        
    strNT = "Initial Catalog=KYUYO;"
    cnA.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD 'SQLServer
    cnA.Open
    Set Cmd.ActiveConnection = cnA
    
    strSQL = ""
    strSQL = strSQL & "SELECT *"
    strSQL = strSQL & "�@�@�@�@FROM KYUMTA"
    strSQL = strSQL & "             WHERE SCODE Like '" & "%" & Right(strSCD, 3) & "'"
    strSQL = strSQL & "             And Left(KBN,1) = 'R'"
    Cmd.CommandText = strSQL
    Set rsS = Cmd.Execute
    If rsS.EOF = False Then
        rsS.MoveFirst
        If rsS.Fields("BMN3") = "" Or IsNull(rsS.Fields("BMN3")) Then
            MsgBox "�Ј��̕��傪�o�^����Ă��܂���I" & vbCrLf & "�o�^���Ă��珈������蒼���ĉ������B(ToT)/~~~" & _
            "�Ј��R�[�h=" & strSCD
            ���唻�� = 0
            GoTo Exit_DB
        End If
        Select Case rsS.Fields("KBN")
            Case "RH"
                If rsS.Fields("BMN3") = "41" Then '�f��
                    lngBMN = 3
                Else
                    lngBMN = 1
                End If
            Case "RO"
                If rsS.Fields("BMN3") = "19" Then     '�����c��
                    lngBMN = 5
                ElseIf rsS.Fields("BMN3") = "20" Then '�������
                        lngBMN = 5
                ElseIf rsS.Fields("BMN3") = "18" Then '�����H
                        lngBMN = 7
                ElseIf rsS.Fields("BMN3") = "21" Then '�������H
                        lngBMN = 8
                ElseIf rsS.Fields("BMN3") = "22" Then '���É��c��
                        lngBMN = 6
                ElseIf rsS.Fields("BMN3") = "23" Then '���É����
                        lngBMN = 6
                ElseIf rsS.Fields("BMN3") = "24" Then '���É����H
                        lngBMN = 9
                Else
                    lngBMN = 4
                End If
            Case "RT"
                If rsS.Fields("BMN3") = "31" Then '�������H
                    lngBMN = 11
                ElseIf rsS.Fields("BMN3") = "27" Then '��֓��c��
                    lngBMN = 12
                ElseIf rsS.Fields("BMN3") = "32" Then '��֓����
                    lngBMN = 12
                ElseIf rsS.Fields("BMN3") = "33" Then '��֓����H
                    lngBMN = 13
                ElseIf rsS.Fields("BMN3") = "28" Then '���c��
                    lngBMN = 14
                ElseIf rsS.Fields("BMN3") = "34" Then '�����
                    lngBMN = 14
                ElseIf rsS.Fields("BMN3") = "35" Then '�����H
                    lngBMN = 15
                ElseIf rsS.Fields("BMN3") = "29" Then '�k�֓��c��
                    lngBMN = 16
                ElseIf rsS.Fields("BMN3") = "36" Then '�k�֓����
                    lngBMN = 16
                ElseIf rsS.Fields("BMN3") = "37" Then '�k�֓����H
                    lngBMN = 17
                Else
                    lngBMN = 10
                End If
            Case "RX", "TA", "KA"
                If rsS.Fields("SCODE") = "00089" Or rsS.Fields("SCODE") = "00490" Or rsS.Fields("SCODE") = "00472" Or rsS.Fields("SCODE") = "00694" Then '�{��
                    lngBMN = 2
                ElseIf rsS.Fields("SCODE") = "00497" Then  '���x�X
                    lngBMN = 4
                ElseIf rsS.Fields("SCODE") = "00526" Then  '�����x�X
                    lngBMN = 10
                Else
                    MsgBox "�Ј��R�[�h���s���ł��I" & vbCrLf & "�v�����I�I(ToT)/~~~" & _
                    "�Ј��R�[�h=" & strSCD
                End If
        End Select
        ���唻�� = lngBMN
    Else
        MsgBox "�Ј��}�X�^�[�ɓo�^���Ȃ��f�[�^������܂��I" & vbCrLf & "�}�X�^�ɓo�^���Ă����蒼���ĉ�����" & _
        "�Ј��R�[�h=" & strSCD
        ���唻�� = 0
        GoTo Exit_DB
    End If

Exit_DB:

    If Not rsS Is Nothing Then
        If rsS.State = adStateOpen Then rsS.Close
        Set rsS = Nothing
    End If
    If Not cnA Is Nothing Then
        If cnA.State = adStateOpen Then cnA.Close
        Set cnA = Nothing
    End If
    
End Function
