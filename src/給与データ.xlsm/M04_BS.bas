Attribute VB_Name = "M04_BS"
Option Explicit

Const Fname = "Z:\��v�V�X�e��\�d��\�ܗ^�d��.txt"
Const Kname = "Z:\��v�V�X�e��\�d��\���^�d��.txt"

Private lngKIN(17, 6) As Long '���z�i�[�z��(���Ə�,����)
Private lngKYU       As Long '��V�����U�֊z
Private strKMC       As String  '�Ȗں���
Private strKMN       As String  '�Ȗږ�
Private strHJC       As String  '�⏕����
Private strHJN       As String  '�⏕��
Private strKMC2      As String  '�Ȗں���
Private strKMN2      As String  '�Ȗږ�
Private strHJC2      As String  '�⏕����
Private strHJN2      As String  '�⏕��
Private strBMC       As String  '���庰��
Private strBMC2      As String  '���庰��
Private strTXC       As String  '�ŋ敪
Private lngR         As Long    '�d��s����
Private lngCR        As Long    '�s����
Private lngCC        As Long    '���_����
Private lngNO        As Long    '�`�[��
Private lngGNO       As Long    '�s��
Private boolR        As Boolean

'=== �d����&�X�V ����===
Sub SEL_KS()
    boolR = False
    If Sheets("����").Range("C4") = "����" Then
        Call �d����("K")
    Else
        If Sheets("����").Range("C4") = "�Վ��ܗ^" Then
            boolR = True
            Call �d����("S")
        Else
            Call �d����("S")
        End If
    End If
    Range("B4").Select
End Sub

Sub �d����(strKBN As String)

Dim lngCock As Long

With Sheets("����")
.Select
DateA = Range("B4")
Erase lngKIN

'���z��z��ɾ�� =====
'�l���E�U���z�E��ʔ�E�ݕt���E�ٗp�ی�����Е��S��
For lngCC = 0 To 16
    lngKIN(lngCC, 0) = lngKIN(lngCC, 0) + Cells(11, lngCC + 3) '�l��
    lngKIN(lngCC, 1) = lngKIN(lngCC, 1) + Cells(12, lngCC + 3) '�U���z
    lngKIN(lngCC, 2) = lngKIN(lngCC, 2) + Cells(14, lngCC + 3) '��ʔ�
    lngKIN(lngCC, 4) = lngKIN(lngCC, 4) + Cells(22, lngCC + 3) '�ݕt��
    lngKIN(lngCC, 5) = lngKIN(lngCC, 5) + Cells(33, lngCC + 3) '�ٗp�ی�����Е��S��
    lngKIN(lngCC, 6) = lngKIN(lngCC, 6) + Cells(24, lngCC + 3) '�N�b�N��
Next lngCC
'�T���v
For lngCC = 0 To 16
    For lngCR = 0 To 15
        lngKIN(lngCC, 3) = lngKIN(lngCC, 3) + Cells(lngCR + 15, lngCC + 3)
    Next lngCR
Next lngCC

'�ݕt���̋��z������
If strKBN = "K" Then
    If (lngKIN(0, 4) + lngKIN(1, 4) + lngKIN(2, 4)) <> Range("AB19") Then
        MsgBox "�{���̑ݕt������v���܂���B�ݕt�����ׂ�ێ炵�Ă����蒼���ĉ������B", vbCritical
        Exit Sub
    ElseIf (lngKIN(3, 4) + lngKIN(4, 4) + lngKIN(5, 4) + lngKIN(6, 4) + lngKIN(7, 4) + lngKIN(8, 4)) <> Range("AE19") Then
        MsgBox "���̑ݕt������v���܂���B�ݕt�����ׂ�ێ炵�Ă����蒼���ĉ������B", vbCritical
        Exit Sub
    ElseIf (lngKIN(9, 4) + lngKIN(10, 4) + lngKIN(11, 4) + lngKIN(12, 4) + lngKIN(13, 4) + lngKIN(14, 4)) <> Range("AH19") Then
        MsgBox "�����̑ݕt������v���܂���B�ݕt�����ׂ�ێ炵�Ă����蒼���ĉ������B", vbCritical
        Exit Sub
    End If
    If Dir(Kname) <> "" Then Kill Kname
Else
    If (lngKIN(0, 4) + lngKIN(1, 4) + lngKIN(2, 4)) <> Range("AL19") Then
        MsgBox "�{���̑ݕt������v���܂���B�ݕt�����ׂ�ێ炵�ĉ������B", vbCritical
        Exit Sub
    ElseIf (lngKIN(3, 4) + lngKIN(4, 4) + lngKIN(5, 4) + lngKIN(6, 4) + lngKIN(7, 4) + lngKIN(8, 4)) <> Range("AO19") Then
        MsgBox "���̑ݕt������v���܂���B�ݕt�����ׂ�ێ炵�ĉ������B", vbCritical
        Exit Sub
    ElseIf (lngKIN(9, 4) + lngKIN(10, 4) + lngKIN(11, 4) + lngKIN(12, 4) + lngKIN(13, 4) + lngKIN(14, 4)) <> Range("AR19") Then
        MsgBox "�����̑ݕt������v���܂���B�ݕt�����ׂ�ێ炵�ĉ������B", vbCritical
        Exit Sub
    End If
    If Dir(Fname) <> "" Then Kill Fname
End If

Sheets("�d��").Select
Call CLS_�d��

If strKBN = "K" Then
    Range("B1") = "���^�d��"
    strKMC = .Range("U11")
    strKMN = .Range("W11")
    strHJC = ""
    strHJN = ""
    strBMC = .Range("C6")
    strTXC = "00"
Else
    Range("B1") = "�ܗ^�d��"
    If boolR Then
        strKMC = "713"
        strKMN = "�ܗ^"
        strHJC = ""
        strHJN = ""
        strBMC = .Range("C6")
    Else
        strKMC = "323"
        strKMN = "�����ܗ^"
        strHJC = "601"
        strHJN = "�{��"
        strBMC = ""
    End If
    strTXC = "00"
End If
lngR = 5  '�J�n�s
lngNO = 1 '�`�[��
lngGNO = 1

'=== ��U���z��{���Ōv�シ�� ===
'�U�����z�i�U���z+��ʔ�j
Cells(lngR, 1) = lngNO
Cells(lngR, 2) = strKMC
Cells(lngR, 3) = strKMN
If strHJC <> "" Then
    Cells(lngR + 1, 2) = strHJC
    Cells(lngR + 1, 3) = strHJN
End If
Cells(lngR, 4) = strTXC
Cells(lngR, 5) = Strings.Format(DateA, "ggge�Nm����") & .Range("C4")
Cells(lngR + 1, 5) = "�U�����z  " & .Range("T11") & "����"
Cells(lngR, 6) = .Range("T12") + .Range("T14")
Cells(lngR, 7) = .Range("U12") '�Ȗں���
Cells(lngR, 8) = .Range("W12") '�Ȗږ�
Cells(lngR, 10) = "00" '�ŋ敪
Cells(lngR + 1, 7) = .Range("V12") '�⏕����
Cells(lngR + 1, 8) = .Range("X12") '�⏕��
Cells(lngR + 1, 4) = strBMC
Cells(lngR + 1, 10) = ""
lngNO = lngNO + 1
lngR = lngR + 2

'�Љ�ی���
For lngCR = 15 To 30
    If .Cells(lngCR, 20) <> 0 And lngCR <> 22 Then
        Cells(lngR, 1) = lngNO
        Cells(lngR, 5) = Strings.Format(DateA, "ggge�Nm����") & .Range("C4")
        Cells(lngR + 1, 5) = .Cells(lngCR, 2) & " �a��"
        lngNO = lngNO + 1
        Cells(lngR, 6) = .Cells(lngCR, 20)
        Cells(lngR, 2) = strKMC
        Cells(lngR, 3) = strKMN
        Cells(lngR, 4) = strTXC
        Cells(lngR, 7) = .Cells(lngCR, 21)
        Cells(lngR, 8) = .Cells(lngCR, 23)
        Cells(lngR, 10) = "00"
        If strHJC <> "" Then
            Cells(lngR + 1, 2) = strHJC
            Cells(lngR + 1, 3) = strHJN
        End If
        Cells(lngR + 1, 4) = strBMC
        Cells(lngR + 1, 7) = .Cells(lngCR, 22)
        Cells(lngR + 1, 8) = .Cells(lngCR, 24)
        Cells(lngR + 1, 10) = ""
        lngR = lngR + 2
    End If
Next lngCR

Call Data_Export(DateA)
Call �V�[�g���
Call CLS_�d��
lngR = 5
lngNO = 1

'=== ��������U�� ==========
'�{����-����U��
If strKBN = "S" Then
    '�Վ��ܗ^�̏ꍇ�͒��ڏܗ^����ŏグ��
    If boolR Then
        strKMC = "713"
        strKMN = "�ܗ^"
    Else
        strKMC = "713"
        strKMN = "�ܗ^�����z"
    End If
End If

For lngCC = 0 To 2
     If strKBN = "K" Then
        If lngCC > 0 Then Call �����U��("101")
        Call ��ʔ�U��
    Else
        If lngCC > 0 Then Call �����U��("101")
    End If
    Call �ٗp�ی��U��
Next lngCC

'�ݕt���ʌv��
If strKBN = "K" And .Range("AB19") <> 0 Then  '�{��
    For lngCR = 13 To 18
        If .Cells(lngCR, 19) = "" Then Exit For
        Cells(lngR, 1) = lngNO
        Cells(lngR, 2) = strKMC
        Cells(lngR, 3) = strKMN
        Cells(lngR, 4) = "00"
        Cells(lngR, 5) = Strings.Format(DateA, "ggge�Nm����") & .Range("C4")
        Cells(lngR + 1, 5) = .Cells(lngCR, 27) & "�@�ݕt���v��"
        Cells(lngR, 6) = .Cells(lngCR, 28)
        Cells(lngR, 7) = .Range("U22")
        Cells(lngR, 8) = .Range("W22")
        Cells(lngR, 10) = "00"
        Cells(lngR + 1, 7) = .Cells(lngCR, 26) '�ݕt���⏕����
        Cells(lngR + 1, 8) = .Cells(lngCR, 27) '�ݕt���⏕��
        Cells(lngR + 1, 4) = strBMC
        Cells(lngR + 1, 10) = ""
        lngNO = lngNO + 1
        lngR = lngR + 2
    Next lngCR
ElseIf strKBN = "S" And .Range("AL19") <> 0 Then
    For lngCR = 13 To 18
        If .Cells(lngCR, 29) = "" Then Exit For
        Cells(lngR, 1) = lngNO
        Cells(lngR, 2) = "323"
        Cells(lngR, 3) = "�����ܗ^"
        Cells(lngR, 4) = "00"
        Cells(lngR + 1, 2) = "601"
        Cells(lngR + 1, 3) = "�{��"
        Cells(lngR, 5) = Strings.Format(DateA, "ggge�Nm����") & .Range("C4")
        Cells(lngR + 1, 5) = .Cells(lngCR, 37) & "�@�ݕt���v��"
        Cells(lngR, 6) = .Cells(lngCR, 38)
        Cells(lngR, 7) = .Range("U22")
        Cells(lngR, 8) = .Range("W22")
        Cells(lngR, 7) = .Cells(lngCR, 36)
        Cells(lngR, 8) = .Cells(lngCR, 37)
        Cells(lngR, 10) = "00"
        Cells(lngR + 1, 10) = ""
        lngNO = lngNO + 1
        lngR = lngR + 2
    Next lngCR
End If
Call Data_Export(DateA)
Call �V�[�g���
Call CLS_�d��
lngR = 5
lngNO = 1

'���x�X�U��
Range("B1") = "���U��"
If strKBN = "S" Then Call �ܗ^�U��("���")
For lngCC = 3 To 8
    strBMC = Cells(6, lngCC + 3)
    If strKBN = "K" Then
        Call �����U��("101")
        Call ��ʔ�U��
    Else
        If lngCC > 3 Then Call �c�Ə��U��("201")
    End If
    Call �ٗp�ی��U��
Next lngCC
'�ݕt���ʌv��
If strKBN = "K" And .Range("AE19") <> 0 Then
    For lngCR = 13 To 18
        If .Cells(lngCR, 29) = "" Then Exit For
        Cells(lngR, 1) = lngNO
        Cells(lngR, 2) = strKMC
        Cells(lngR, 3) = strKMN
        Cells(lngR, 4) = "00"
        Cells(lngR + 1, 4) = "101"
        Cells(lngR, 5) = Strings.Format(DateA, "ggge�Nm����") & .Range("C4")
        Cells(lngR + 1, 5) = .Cells(lngCR, 30) & "�@�ݕt���U��"
        Cells(lngR, 6) = .Cells(lngCR, 31)
        Cells(lngR, 7) = .Range("U22")
        Cells(lngR, 8) = .Range("W22")
        Cells(lngR + 1, 7) = .Cells(lngCR, 29)
        Cells(lngR + 1, 8) = .Cells(lngCR, 30)
        Cells(lngR, 10) = "00"
        Cells(lngR + 1, 10) = ""
        lngNO = lngNO + 1
        lngR = lngR + 2
        Next lngCR
ElseIf strKBN = "S" And .Range("AR19") <> 0 Then
    For lngCR = 13 To 18
        If .Cells(lngCR, 39) = "" Then Exit For
        Cells(lngR, 1) = lngNO
        If boolR Then
            Cells(lngR, 2) = "713"
            Cells(lngR, 3) = "�ܗ^"
            Cells(lngR + 1, 4) = "201"
        Else
            Cells(lngR, 2) = "323"
            Cells(lngR, 3) = "�����ܗ^"
            Cells(lngR + 1, 2) = "611"
            Cells(lngR + 1, 3) = "���"
        End If
        Cells(lngR, 5) = Strings.Format(DateA, "ggge�Nm����") & .Range("C4")
        Cells(lngR + 1, 5) = .Cells(lngCR, 38) & "�@�ݕt���v��"
        Cells(lngR, 6) = .Cells(lngCR, 41)
        Cells(lngR, 7) = .Range("U22")
        Cells(lngR, 8) = .Range("W22")
        Cells(lngR + 1, 7) = .Cells(lngCR, 39)
        Cells(lngR + 1, 8) = .Cells(lngCR, 40)
        Cells(lngR, 10) = "00"
        Cells(lngR + 1, 10) = ""
        lngNO = lngNO + 1
        lngR = lngR + 2
    Next lngCR
End If
Call Data_Export(DateA)
Call �V�[�g���
Call CLS_�d��
lngR = 5
lngNO = 1

'�����x�X�U��
Range("B1") = "�����U��"
If strKBN = "S" Then Call �ܗ^�U��("����")
For lngCC = 9 To 17
    If strKBN = "K" Then
        Call �����U��("101")
        Call ��ʔ�U��
        lngCock = lngCock + lngKIN(lngCC, 6)
    Else
         If lngCC > 9 Then Call �c�Ə��U��("301")
    End If
    Call �ٗp�ی��U��
Next lngCC

'�N�b�N��U��
If strKBN = "K" And lngCock <> 0 Then
    Cells(lngR, 1) = lngNO
    Cells(lngR, 2) = "326"
    Cells(lngR, 3) = "�a���"
    Cells(lngR + 1, 2) = "707"
    Cells(lngR + 1, 3) = "�N�b�N��"
    Cells(lngR, 4) = "00"
    Cells(lngR, 5) = Strings.Format(DateA, "ggge�Nm����") & .Range("C4")
    Cells(lngR + 1, 5) = "�������N�b�N���U��"
    Cells(lngR, 6) = lngCock
    Cells(lngR, 7) = "326"
    Cells(lngR, 8) = "�a���"
    Cells(lngR, 10) = "00"
    Cells(lngR + 1, 7) = "717"
    Cells(lngR + 1, 8) = "�N�b�N��-����"
    lngNO = lngNO + 1
    lngR = lngR + 2
End If

'�ݕt���ʌv��
If strKBN = "K" And .Range("AH19") <> 0 Then
    For lngCC = 13 To 18
        If .Cells(lngCC, 32) = "" Then Exit For
        Cells(lngR, 1) = lngNO
        Cells(lngR, 2) = strKMC
        Cells(lngR, 3) = strKMN
        Cells(lngR, 4) = "00"
        Cells(lngR + 1, 4) = "101"
        Cells(lngR, 5) = Strings.Format(DateA, "ggge�Nm����") & .Range("C4")
        Cells(lngR + 1, 5) = .Cells(lngCC, 33) & "�@�ݕt���U��"
        Cells(lngR, 6) = .Cells(lngCC, 34)
        Cells(lngR, 7) = .Range("U22")
        Cells(lngR, 8) = .Range("W22")
        Cells(lngR + 1, 7) = .Cells(lngCC, 32)
        Cells(lngR + 1, 8) = .Cells(lngCC, 33)
        Cells(lngR, 10) = "00"
        Cells(lngR + 1, 10) = ""
        lngNO = lngNO + 1
        lngR = lngR + 2
    Next lngCC
ElseIf strKBN = "S" And .Range("AR19") <> 0 Then
    For lngCR = 13 To 18
        If .Cells(lngCR, 44) = "" Then Exit For
        Cells(lngR, 1) = lngNO
        If boolR Then
            Cells(lngR, 2) = "713"
            Cells(lngR, 3) = "�ܗ^"
            Cells(lngR + 1, 4) = "301"
        Else
            Cells(lngR, 2) = "323"
            Cells(lngR, 3) = "�����ܗ^"
            Cells(lngR + 1, 2) = "631"
            Cells(lngR + 1, 3) = "����"
        End If
        Cells(lngR, 4) = "00"
        Cells(lngR, 5) = Strings.Format(DateA, "ggge�Nm����") & .Range("C4")
        Cells(lngR + 1, 5) = .Cells(lngCR, 43) & "�@�ݕt���v��"
        Cells(lngR, 6) = .Cells(lngCR, 44)
        Cells(lngR, 7) = .Range("U22")
        Cells(lngR, 8) = .Range("W22")
        Cells(lngR + 1, 7) = .Cells(lngCR, 42)
        Cells(lngR + 1, 8) = .Cells(lngCR, 43)
        Cells(lngR, 10) = "00"
        Cells(lngR + 1, 10) = ""
        lngNO = lngNO + 1
        lngR = lngR + 2
    Next lngCR
End If
Call Data_Export(DateA)
Call �V�[�g���

Sheets("�U��").Select
Call �V�[�g���

.Select
End With

End Sub

Sub �����U��(strRMC As String)

'��V�����Z�o�i�U���z+�T���v�j
lngKYU = 0
lngKYU = lngKYU + lngKIN(lngCC, 1) + lngKIN(lngCC, 3)

If lngKYU <> 0 Then
    Cells(lngR, 1) = lngNO
    Cells(lngR, 2) = strKMC
    Cells(lngR, 3) = strKMN
    Cells(lngR, 4) = "00"
    Cells(lngR, 5) = Strings.Format(DateA, "ggge�Nm����") & Sheets("����").Range("C4")
    Cells(lngR + 1, 5) = Sheets("����").Cells(5, lngCC + 3) & "���v��"
    Cells(lngR, 6) = lngKYU
    Cells(lngR, 7) = strKMC
    Cells(lngR, 8) = strKMN
    Cells(lngR, 10) = "00"
    Cells(lngR + 1, 4) = Sheets("����").Cells(6, lngCC + 3)
    Cells(lngR + 1, 10) = strRMC
    lngNO = lngNO + 1
    lngR = lngR + 2
    If lngR > 43 Then
        Call Data_Export(DateA)
        Call �V�[�g���
        Call CLS_�d��
        lngR = 5
    End If
End If
    
End Sub

Sub �c�Ə��U��(strRMC As String)

'��V�����Z�o�i�U���z+�T���v�j
lngKYU = 0
lngKYU = lngKYU + lngKIN(lngCC, 1) + lngKIN(lngCC, 3)

strKMC = "323"
strKMN = "�����ܗ^"
If strRMC = "201" Then
    strHJC = "611"
    strHJN = "���"
Else
    strHJC = "631"
    strHJN = "����"
End If

If lngKYU <> 0 Then
    Cells(lngR, 1) = lngNO
    Cells(lngR, 2) = strKMC
    Cells(lngR, 3) = strKMN
    Cells(lngR, 4) = "00"
    Cells(lngR, 5) = Strings.Format(DateA, "ggge�Nm����") & Sheets("����").Range("C4")
    Cells(lngR + 1, 5) = Sheets("����").Cells(5, lngCC + 3) & "���v��"
    Cells(lngR, 6) = lngKYU
    Cells(lngR, 7) = strKMC
    Cells(lngR, 8) = strKMN
    Cells(lngR, 10) = "00"
    Cells(lngR + 1, 2) = strHJC
    Cells(lngR + 1, 3) = strHJN
    Cells(lngR + 1, 7) = strHJC
    Cells(lngR + 1, 8) = strHJN
    lngNO = lngNO + 1
    lngR = lngR + 2
    If lngR > 43 Then
        Call Data_Export(DateA)
        Call �V�[�g���
        Call CLS_�d��
        lngR = 5
    End If
End If
    
End Sub

Sub �ܗ^�U��(strSTN As String)
'�ܗ^�Z�o�i�U���z+�T���v�j

Dim strKCD  As String
Dim strKNM  As String
Dim strKCD2 As String
Dim strKNM2 As String

lngKYU = 0
If strSTN = "���" Then
    For lngCC = 3 To 8
        lngKYU = lngKYU + lngKIN(lngCC, 1) + lngKIN(lngCC, 3)
    Next lngCC
    If boolR Then '�Վ��ܗ^����
        strKCD = "713"
        strKNM = "�ܗ^"
        strHJC = ""
        strHJN = ""
        strBMC = "201"
        strKCD2 = "713"
        strKNM2 = "�ܗ^"
        strHJC2 = ""
        strHJN2 = ""
        strBMC2 = "101"
    Else
        strKCD = "323"
        strKNM = "�����ܗ^"
        strHJC = "611"
        strHJN = "���"
        strBMC = ""
        strKCD2 = "323"
        strKNM2 = "�����ܗ^"
        strHJC2 = "601"
        strHJN2 = "�{��"
        strBMC2 = ""
    End If
Else
    For lngCC = 9 To 14
        lngKYU = lngKYU + lngKIN(lngCC, 1) + lngKIN(lngCC, 3)
    Next lngCC
    If boolR Then
        strKCD = "713"
        strKNM = "�ܗ^"
         strHJC = ""
        strHJN = ""
        strBMC = "301"
        strKCD2 = "713"
        strKNM2 = "�ܗ^"
        strHJC2 = ""
        strHJN2 = ""
        strBMC2 = "101"
    Else
        strKCD = "323"
        strKNM = "�����ܗ^"
        strHJC = "631"
        strHJN = "����"
        strBMC = ""
        strKCD2 = "323"
        strKNM2 = "�����ܗ^"
        strHJC2 = "601"
        strHJN2 = "�{��"
        strBMC2 = ""
    End If
End If
    
If lngKYU <> 0 Then
    Cells(lngR, 1) = lngNO
    Cells(lngR, 2) = strKCD
    Cells(lngR, 3) = strKNM
    Cells(lngR + 1, 2) = strHJC
    Cells(lngR + 1, 3) = strHJN
    Cells(lngR, 4) = "00"
    Cells(lngR + 1, 4) = strBMC
    Cells(lngR, 5) = Strings.Format(DateA, "ggge�Nm����") & Sheets("����").Range("C4")
    Cells(lngR + 1, 5) = strSTN & "���v��"
    Cells(lngR, 6) = lngKYU
    Cells(lngR, 7) = strKCD2
    Cells(lngR, 8) = strKNM2
    Cells(lngR + 1, 7) = strHJC2
    Cells(lngR + 1, 8) = strHJN2
    Cells(lngR, 10) = "00"
    Cells(lngR + 1, 10) = strBMC2
    lngNO = lngNO + 1
    lngR = lngR + 2
    If lngR > 43 Then
        Call Data_Export(DateA)
        Call �V�[�g���
        Call CLS_�d��
        lngR = 5
    End If
End If
    
End Sub

Sub ��ʔ�U��()

If lngKIN(lngCC, 2) <> 0 Then
    Cells(lngR, 1) = lngNO
    Cells(lngR, 2) = Sheets("����").Range("U14")
    Cells(lngR, 3) = Sheets("����").Range("W14")
    Cells(lngR, 4) = "Q5"
    Cells(lngR, 5) = Strings.Format(DateA, "ggge�Nm����") & Sheets("����").Range("C4")
    Cells(lngR + 1, 5) = "��ʔ�U�� " & " " & Sheets("����").Cells(5, lngCC + 3) & "���v��"
    Cells(lngR, 6) = lngKIN(lngCC, 2)
    Cells(lngR + 1, 6) = Round((lngKIN(lngCC, 2) / 110) * 10, 0)
    Cells(lngR, 7) = Sheets("����").Range("U11")
    Cells(lngR, 8) = Sheets("����").Range("W11")
    Cells(lngR, 10) = "00"
    Cells(lngR + 1, 4) = Sheets("����").Cells(6, lngCC + 3)
    Cells(lngR + 1, 10) = "101"
    lngNO = lngNO + 1
    lngR = lngR + 2
    If lngR > 43 Then
        Call Data_Export(DateA)
        Call �V�[�g���
        Call CLS_�d��
        lngR = 5
    End If
End If

End Sub

Sub �ٗp�ی��U��()

If lngKIN(lngCC, 5) <> 0 Then
    Cells(lngR, 1) = lngNO
    Cells(lngR, 2) = Sheets("����").Range("U33")
    Cells(lngR, 3) = Sheets("����").Range("W33")
    Cells(lngR, 4) = "P0"
    Cells(lngR, 5) = Strings.Format(DateA, "ggge�Nm����") & Sheets("����").Range("C4")
    Cells(lngR + 1, 5) = Sheets("����").Cells(33, 2) & " " & Sheets("����").Cells(5, lngCC + 3) & "���v��"
    
    Cells(lngR, 6) = lngKIN(lngCC, 5)
    Cells(lngR, 7) = Sheets("����").Range("U19")
    Cells(lngR, 8) = Sheets("����").Range("W19")
    Cells(lngR + 1, 7) = Sheets("����").Range("V19")
    Cells(lngR + 1, 8) = Sheets("����").Range("X19")
    Cells(lngR, 10) = "00"
    Cells(lngR + 1, 8) = Sheets("����").Range("X19")
    Cells(lngR + 1, 4) = Sheets("����").Cells(6, lngCC + 3)
    Cells(lngR + 1, 10) = ""
    lngNO = lngNO + 1
    lngR = lngR + 2
    If lngR > 43 Then
        Call Data_Export(DateA)
        Call �V�[�g���
        Call CLS_�d��
        lngR = 5
    End If
End If

End Sub

Sub CLS_�d��()
    Range("A5:J46") = ""
End Sub

Public Sub Data_Export(DateA As Date)

Dim boolT    As Boolean  '�߂�l
Dim strNO    As String   '�`�[�ԍ�
Dim strTEXT  As String   '�ް�÷��
Dim lngC     As Long     '����
Dim lngTaxL  As Long     '�ؕ������
Dim lngTaxR  As Long     '�ݕ������
Dim strBMNL  As String   '�ؕ�����
Dim strBMNR  As String   '�ݕ�����
Dim strKMKL  As String   '�ؕ��Ȗ�
Dim strKMKR  As String   '�ݕ��Ȗ�
Dim strHCDL  As String   '�ؕ�����溰��
Dim strHCDR  As String   '�ݕ�����溰��
Dim strTXBL  As String   '�ؕ��ŋ敪
Dim strTXBR  As String   '�ݕ��ŋ敪
Dim strTKYO  As String   '�E�v
Dim strKINL  As String   '�ؕ����z
Dim strKINR  As String   '�ݕ����z
Dim strTaxL  As String   '�ؕ����z
Dim strTaR   As String   '�ݕ����z

    '�`�[�ԍ�&�d��̧�ٖ��쐬
    strNO = "4" & Strings.Format(DateA, "mmdd")
    strDate = Format(DateA, "yyyymmdd")
    
    '�ėp�d��f�[�^����߰ď���
    For lngC = 5 To 45 Step 2
        If Cells(lngC, 3) = "" Or Cells(lngC, 6) = 0 Then
        Else
            strBMNL = Cells(lngC + 1, 4)
            strBMNR = Cells(lngC + 1, 10)
            strKMKL = Cells(lngC, 2)
            strKMKR = Cells(lngC, 7)
            strHCDL = Cells(lngC + 1, 2)
            strHCDR = Cells(lngC + 1, 7)
            strTXBL = Cells(lngC, 4)
            strTXBR = Cells(lngC, 10)
            strTKYO = Cells(lngC, 5) & " " & Cells(lngC + 1, 5)
            strKINL = Cells(lngC, 6)
            strKINR = Cells(lngC, 6)
            If Right(strTXBL, 1) = "0" Then
                lngTaxL = 0
            Else
                lngTaxL = Cells(lngC + 1, 6)
            End If
            If Right(strTXBR, 1) = "0" Then
                lngTaxR = 0
            Else
                lngTaxR = Cells(lngC + 1, 6)
            End If
                        
            strTEXT = strDate & "," & strNO & ",21,0,1," & _
            strBMNL & ",," & strKMKL & ",," & strHCDL & ",," & strTXBL & ",," & strKINL & "," & lngTaxL & ",1," & _
            strBMNR & ",," & strKMKR & ",," & strHCDR & ",," & strTXBR & ",," & strKINR & "," & lngTaxR & "," & _
            strTKYO & ",,,1,,,,,,,,,,,,,,,,,,,,,,,,,"
            
            If Sheets("����").Range("C4") = "����" Then
                boolT = AddText(Kname, strTEXT)
            Else
                boolT = AddText(Fname, strTEXT)
            End If
            lngGNO = lngGNO + 1
        End If
    Next lngC
    
End Sub

