Attribute VB_Name = "M02_Data"
Option Explicit

Private strBMN As String
Private strBKA As String
Private strSCD As String

Sub Proc_Get()
    
    '�e�L�X�g�f�[�^��ǂݍ����Data�V�[�g�֏�������
    '���^�f�[�^��[�C���|�[�g�f�[�^]�e�[�u���֓����
    '[�C���|�[�g�f�[�^]�̃f�[�^�𐸍����Ē~�σf�[�^�ֈڍs����
    
    If Range("S11") = 1 Then
        strKBN = "K"
    ElseIf Range("S11") = 2 Then
        strKBN = "S"
    Else
        strKBN = InputBox("�x���敪����͂��ĉ������B" & vbCrLf & "���^=K , �ܗ^=S", "�x���敪���", "K")
    End If
    
    '�e�L�X�g�f�[�^��ǂݍ����Data�V�[�g�֏�������
    READ_TextFile
    
    'Data�V�[�g���ް���DB�̃C���|�[�g�f�[�^(Tempð���)�֏�������
    If strKBN = "S" Then
        Call �ܗ^�f�[�^�ړ�
    Else
        Call ���^�f�[�^�ړ�
    End If
    
    '[�C���|�[�g�f�[�^]�̃f�[�^��~�σf�[�^�ֈڍs����
    Call �~�Ϗ���
    
     MsgBox "�f�[�^�ǂݍ��݂��������܂����B"
     
     Sheets("Menu").Select
    
End Sub

Sub READ_TextFile()

    Const cnsTITLE = "�e�L�X�g�t�@�C���ǂݍ��ݏ���"
    Const cnsFILTER = "�e�L�X�g�`���t�@�C�� (*.txt),*.txt,�S�Ẵt�@�C��(*.*),*.*"
    
    Dim xlAPP       As Application ' Application�I�u�W�F�N�g
    Dim intFF       As Integer     ' FreeFile�l
    Dim strFileName As String      ' OPEN����t�@�C����(�t���p�X)
    Dim vFileName   As Variant     ' �t�@�C��������p
    Dim X(1 To 54)  As Variant     ' �ǂݍ��񂾃��R�[�h���e
    Dim lngR        As Long        ' ���e����Z���̍s
    Dim lngCnt      As Long        ' ���R�[�h�����J�E���^
    
'    strKBN = Sheets("Wait").Range("S11")
    Sheets("Data").Select
    Range("A1:BB100").ClearContents
    'Application�I�u�W�F�N�g�擾
    Set xlAPP = Application
    '��t�@�C�����J����̃t�H�[���Ńt�@�C�����̎w����󂯂�
    xlAPP.StatusBar = "�ǂݍ��ރt�@�C�������w�肵�ĉ������B"
    ChDrive "K:"
    ChDir dtW
    vFileName = xlAPP.GetOpenFilename(cnsFILTER, 1, cnsTITLE, , False)
    '�L�����Z�����ꂽ�ꍇ��False���Ԃ�̂ňȍ~�̏����͍s�Ȃ�Ȃ�
    If VarType(vFileName) = vbBoolean Then Exit Sub
    strFileName = vFileName

    'FreeFile�l�̎擾(�ȍ~���̒l�œ��o�͂���)
    intFF = FreeFile
    '�w��t�@�C����OPEN(���̓��[�h)
    Open strFileName For Input As #intFF
    lngR = 0
    '�t�@�C����EOF�܂ŌJ��Ԃ�
    Do Until EOF(intFF)
        '���R�[�h�����J�E���^�̉��Z
        lngCnt = lngCnt + 1
        xlAPP.StatusBar = "�ǂݍ��ݒ��ł��D�D�D�D(" & lngCnt & "���R�[�h��)"
        '���R�[�h��ǂݍ���
        If strKBN = "K" Then
            Input #intFF, X(1), X(2), X(3), X(4), X(5), X(6), X(7), X(8), X(9), X(10), _
                          X(11), X(12), X(13), X(14), X(15), X(16), X(17), X(18), X(19), X(20), _
                          X(21), X(22), X(23), X(24), X(25), X(26), X(27), X(28), X(29), X(30), _
                          X(31), X(32), X(33), X(34), X(35), X(36), X(37), X(38), X(39), X(40), _
                          X(41), X(42), X(43), X(44), X(45), X(46), X(47), X(48), X(49), X(50), _
                          X(51), X(52), X(53), X(54)
        Else
            Input #intFF, X(1), X(2), X(3), X(4), X(5), X(6), X(7), X(8), X(9), X(10), _
                          X(11), X(12), X(13), X(14), X(15), X(16), X(17), X(18), X(19), X(20), _
                          X(21), X(22), X(23), X(24), X(25), X(26), X(27), X(28), X(29), X(30), _
                          X(31), X(32), X(33), X(34), X(35), X(36), X(37), X(38), X(39), X(40), _
                          X(41), X(42), X(43), X(44), X(45)
        End If
        '�s�����Z��A�`E��Ƀ��R�[�h���e��\��
        lngR = lngR + 1
        If strKBN = "K" Then
            Range(Cells(lngR, 1), Cells(lngR, 54)).Value = X   ' �z��n��
        Else
            Range(Cells(lngR, 1), Cells(lngR, 45)).Value = X
        End If
    Loop
    
    xlAPP.StatusBar = False
    
End Sub

Sub �ܗ^�f�[�^�ړ�()

'�ŏ��ɋ��^�f�[�^�̒��̃C���|�[�g�f�[�^���N���A����
'�V�[�gData�ɂ���f�[�^���C���|�[�g�f�[�^�ɓ����

Dim cnW As New ADODB.Connection
Dim rsK As New ADODB.Recordset
Dim lngR   As Long
Dim lngC   As Long

    cnW.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & dbW '���^�f�[�^
    cnW.Open
    
    '�C���|�[�g�f�[�^�N���A
    strSQL = "DELETE FROM �C���|�[�g�f�[�^"
    rsK.Open strSQL, cnW, adOpenStatic, adLockPessimistic
    
    '�C���|�[�g�f�[�^�I�[�v��
    rsK.Open "�C���|�[�g�f�[�^", cnW, adOpenStatic, adLockPessimistic
        
    Sheets("Data").Select
    lngR = 3
    Do
        rsK.AddNew
        strBMN = Strings.Format(Cells(lngR, 1), "000")   '���庰��
        strBKA = Strings.Format(Cells(lngR, 2), "000")   '���ۺ���
        strSCD = Strings.Format(Cells(lngR, 3), "00000") '�Ј�����
                
        '����敪�̐ݒ�(rsK.Fields(0))
        Select Case strBMN
            Case "100"
                rsK.Fields(0) = "000"
            Case "200"
                If strBKA = "010" Then
                    rsK.Fields(0) = "010"
                ElseIf strBKA = "020" Then
                    rsK.Fields(0) = "020"
                ElseIf strBKA = "030" Then
                    rsK.Fields(0) = "030"
                End If
            Case "300"
                If strBKA = "010" Then
                    rsK.Fields(0) = "040"
                ElseIf strBKA = "030" Then
                    rsK.Fields(0) = "050"
                ElseIf strBKA = "040" Then
                    rsK.Fields(0) = "060"
                ElseIf strBKA = "050" Then
                    rsK.Fields(0) = "070"
                End If
            Case "400"
                Select Case strSCD
                    Case "00089"
                        rsK.Fields(0) = "000"
                    Case "00472"
                        rsK.Fields(0) = "000"
                    Case "00490"
                        rsK.Fields(0) = "000"
                    Case "00491"
                        rsK.Fields(0) = "000"
                    Case "00694"
                        rsK.Fields(0) = "000"
                    Case "00497"
                        rsK.Fields(0) = "010"
                    Case "00215"
                        rsK.Fields(0) = "040"
                    Case "00526"
                        rsK.Fields(0) = "040"
                    Case Else
                        MsgBox "���ʂł��Ȃ������̃R�[�h������܂��I" & vbCrLf & strSCD, vbCritical
                        rsK.Fields(0) = "400"
                End Select
            Case Else
                MsgBox "���ʏo���Ȃ����傪����悤�ł��B" & vbCrLf & "�m�F���ĉ������B" & strBMN & " - " & strBKA
        End Select
        
        rsK.Fields(45) = strBKA '���ۺ���
        rsK.Fields(1) = strSCD  '�Ј�����
        '�Ј����`�x��15
        For lngC = 2 To 25
            rsK.Fields(lngC) = Cells(lngR, lngC + 2)
        Next lngC
        rsK.Fields(32) = Cells(lngR, 28) '�ݕt��
        rsK.Fields(35) = Cells(lngR, 30) '���`���~
        rsK.Fields(43) = Cells(lngR, 44) '�T�����v
        rsK.Fields(44) = Cells(lngR, 45) '�����x���z
        rsK.Update
        
        lngR = lngR + 1
        If Cells(lngR, 1) = "" And lngR > 2 Then Exit Do
    Loop
    
Exit_DB:

    If Not rsK Is Nothing Then
        If rsK.State = adStateOpen Then rsK.Close
        Set rsK = Nothing
    End If
    If Not cnW Is Nothing Then
        If cnW.State = adStateOpen Then cnW.Close
        Set cnW = Nothing
    End If

End Sub

Sub ���^�f�[�^�ړ�()

'�ŏ��ɋ��^�f�[�^�̒��̃C���|�[�g�f�[�^���N���A����
'�V�[�gData�ɂ���f�[�^���C���|�[�g�f�[�^�ɓ����

Dim cnW As New ADODB.Connection
Dim rsK As New ADODB.Recordset
Dim lngR   As Long
Dim lngC   As Long
    
    
    
    cnW.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & dbW '���^�f�[�^
    cnW.Open
    
    '�C���|�[�g�f�[�^�N���A
    strSQL = "DELETE FROM �C���|�[�g�f�[�^"
    rsK.Open strSQL, cnW, adOpenStatic, adLockPessimistic
    
    '�C���|�[�g�f�[�^�I�[�v��
    rsK.Open "�C���|�[�g�f�[�^", cnW, adOpenStatic, adLockPessimistic
    
    Sheets("Data").Select
    lngR = 3
    Do
        rsK.AddNew
        strBMN = Strings.Format(Cells(lngR, 1), "000")   '���庰��
        strBKA = Strings.Format(Cells(lngR, 2), "000")   '���ۺ���
        strSCD = Strings.Format(Cells(lngR, 3), "00000") '�Ј�����
        
        '����敪�̐ݒ�(rsK.Fields(0))
        Select Case strBMN
            Case "100"
                rsK.Fields(0) = "000"
            Case "200"
                If strBKA = "010" Then      '���x�X
                    rsK.Fields(0) = "010"
                ElseIf strBKA = "020" Then  '�����c�Ə�
                    rsK.Fields(0) = "020"
                ElseIf strBKA = "030" Then  '���É��c�Ə�
                    rsK.Fields(0) = "030"
                End If
            Case "300"
                If strBKA = "010" Then      '�����x�X
                    rsK.Fields(0) = "040"
                ElseIf strBKA = "030" Then  '��֓�
                    rsK.Fields(0) = "050"
                ElseIf strBKA = "040" Then  '���
                    rsK.Fields(0) = "060"
                ElseIf strBKA = "050" Then  '�k�֓�
                    rsK.Fields(0) = "070"
                End If
            Case "400"
                Select Case strSCD
                    Case "00089"
                        rsK.Fields(0) = "000" '��������
                    Case "00490"
                        rsK.Fields(0) = "000" '�����m�q
                    Case "00694"
                        rsK.Fields(0) = "000" '�K�싞�q
                    Case "00472"
                        rsK.Fields(0) = "000" '���V�O
                    Case "00497"
                        rsK.Fields(0) = "010" '�X�c�T�V
                    Case "00526"
                        rsK.Fields(0) = "040" '�����V��Y
                    Case Else
                        MsgBox "���ʂł��Ȃ������̃R�[�h������܂��I" & vbCrLf & strSCD, vbCritical
                        rsK.Fields(0) = "400"
                End Select
            Case Else
                MsgBox "���ʏo���Ȃ����傪����悤�ł��B" & vbCrLf & "�m�F���ĉ������B" & strBMN & " - " & strBKA
        End Select
        
        rsK.Fields(45) = strBKA '���ۺ���
        rsK.Fields(1) = strSCD  '�Ј�����
        '�����`�x��4�@--> �Ј����`�Ƒ��蓖
        For lngC = 2 To 7
            rsK.Fields(lngC) = Cells(lngR, lngC + 2)
        Next lngC
        '�c�ƌv
        rsK.Fields(8) = 0
        For lngC = 18 To 23
            rsK.Fields(8) = rsK.Fields(8) + Cells(lngR, lngC)
        Next lngC
        '�x������
        rsK.Fields(9) = Cells(lngR, 24)
        '���C�蓖
        rsK.Fields(10) = Cells(lngR, 12)
        '���ʎ蓖
        rsK.Fields(11) = Cells(lngR, 13)
        '����
        rsK.Fields(12) = Cells(lngR, 25)
        '�ېō��v�`�����x���z
        For lngC = 16 To 44
            rsK.Fields(lngC) = Cells(lngR, lngC + 10)
        Next lngC
        rsK.Update
        
        lngR = lngR + 1
        If Cells(lngR, 1) = "" And lngR > 2 Then Exit Do
    Loop

Exit_DB:

    If Not rsK Is Nothing Then
        If rsK.State = adStateOpen Then rsK.Close
        Set rsK = Nothing
    End If
    If Not cnW Is Nothing Then
        If cnW.State = adStateOpen Then cnW.Close
        Set cnW = Nothing
    End If

End Sub

Sub �~�Ϗ���()

'���߰��ް���~���ް��ֈڍs����
'�i�ް��Ȃ���Βǉ��A����΍X�V�j

Dim cnW     As New ADODB.Connection
Dim cnA     As New ADODB.Connection
Dim rsI     As New ADODB.Recordset '�~�σf�[�^
Dim rsK     As New ADODB.Recordset '�C���|�[�g�f�[�^
Dim rsM     As New ADODB.Recordset '���^Ͻ�(KYUMTA)
Dim Cmd     As New ADODB.Command
Dim strNT   As String
Dim lngR    As Long
Dim lngKIN  As Long
    
    DateA = Sheets("Menu").Range("F15")
    If DateA = "0�F00�F00" Then
        strDate = InputBox("�x��������͂��ĉ������B", "�x��������", Strings.Format(Date, "yyyy") & "/" & Strings.Format(Date, "mm"))
    Else
        strDate = Strings.Format(DateA, "yyyy") & "/" & Strings.Format(DateA, "mm")
    End If
    
    cnW.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & dbW '���^�f�[�^
    cnW.Open
    strNT = "Initial Catalog=KYUYO;"
    cnA.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD 'SQLServer
    cnA.Open
    Set Cmd.ActiveConnection = cnA
    
    '���^�f�[�^�̃C���|�[�g�f�[�^�Ǎ���
    rsK.Open "�C���|�[�g�f�[�^", cnW, adOpenStatic, adLockReadOnly
    If rsK.EOF = False Then rsK.MoveFirst
    Do Until rsK.EOF
        If rsK![�����x���z] <> 0 Then
           '���^Ͻ��Ƌ��^����v���Ă��邩�m�F����(�x���z����c�ƒx�����ތ��΂Ȃǂ�������������j
            lngKIN = rsK![�ېŎx���z�v] - rsK![�c�Ǝ蓖] + rsK![�x������] + rsK![����]
            strBMN = rsK![����敪]
            strSCD = rsK![�Ј��R�[�h]
            If rsK![���ۃR�[�h] <> "" And strKBN = "K" Then
                strSQL = ""
                strSQL = strSQL & "SELECT SCODE,"
                strSQL = strSQL & "       PAY1 + PAY2 + OPT1 + OPT2 + OPT3 + OPT4 + OPT5 as PAY,"
                strSQL = strSQL & "       SKBN"
                strSQL = strSQL & "  FROM KYUMTA"
                strSQL = strSQL & "       WHERE SCODE = '" & strSCD & "'"
                Cmd.CommandText = strSQL
                Set rsM = Cmd.Execute
                If rsM.Fields("SKBN") & "" <> "P" Then '�߰ĎЈ����O
                    If lngKIN <> rsM.Fields("PAY") Then
                        lngR = MsgBox("���z���Ⴂ�܂��I�I" & vbCrLf & "�v�`�F�b�N - " & strSCD & " " & rsK![�Ј���] & vbCrLf & "���s���܂����H", vbYesNo, "���^�`�F�b�N")
                        If lngR = vbNo Then
                            Exit Sub
                        End If
                    End If
                End If
                rsM.Close
            End If
           '�~�σf�[�^������
            strSQL = ""
            strSQL = strSQL & "SELECT *"
            strSQL = strSQL & "  FROM �~�σf�[�^"
            strSQL = strSQL & "       WHERE �x���N�� = '" & strDate & "'"
            strSQL = strSQL & "       AND   ���^�敪 = '" & strKBN & "'"
            strSQL = strSQL & "       AND   ����敪 = '" & strBMN & "'"
            strSQL = strSQL & "       AND   �Ј��R�[�h = '" & strSCD & "'"
            rsI.Open strSQL, cnW, adOpenStatic, adLockPessimistic
            If rsI.EOF Then
                rsI.AddNew '������Βǉ�
                rsI![�x���N��] = strDate
                rsI![���^�敪] = strKBN
                rsI![����敪] = strBMN
                rsI![�Ј��R�[�h] = strSCD
            End If
           '�ȉ��X�V
            rsI![�Ј���] = rsK![�Ј���]
            For lngR = 3 To 45 Step 1
                rsI.Fields(lngR + 2) = rsK.Fields(lngR)
            Next lngR
            rsI.Update
            rsI.Close
        End If
        rsK.MoveNext
    Loop

Exit_DB:
    
    If Not rsK Is Nothing Then
        If rsK.State = adStateOpen Then rsK.Close
        Set rsK = Nothing
    End If
    If Not rsI Is Nothing Then
        If rsI.State = adStateOpen Then rsI.Close
        Set rsI = Nothing
    End If
    If Not rsM Is Nothing Then
        If rsM.State = adStateOpen Then rsM.Close
        Set rsM = Nothing
    End If
    If Not cnW Is Nothing Then
        If cnW.State = adStateOpen Then cnW.Close
        Set cnW = Nothing
    End If
    If Not cnA Is Nothing Then
        If cnA.State = adStateOpen Then cnA.Close
        Set cnA = Nothing
    End If

End Sub
