Attribute VB_Name = "M01_Main"
Option Explicit

' �d��f�[�^�����}�N�� =========================
' ϸۋL�^�� : 2001/07/11  հ�ް�� : Shigeo ITOI
' �X�V�F2006/07/06  t_takazawa
' �X�V�F2006/11/30  t_takazawa
' �X�V�F2014/04/09  t_takazawa
' �X�V�F2022/06/23  t_takazawa
'===============================================

#If VBA7 Then
    Declare PtrSafe Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
#Else
    Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
#End If

Public Const MYPROVIDERE = "Provider=SQLOLEDB;"
Public Const MYSERVER9 = "Data Source=192.168.128.9\SQLEXPRESS;"
Public Const MYSERVER = "Data Source=HB14\SQLEXPRESS;"
Public Const USER = "User ID=sa;"
Public Const PSWD9 = "Password=ALCadmin!;"
Public Const PSWD = "Password=admin;"

Public Const MAX_COMPUTERNAME_LENGTH = 15
Public Const dtW As String = "\\192.168.128.4\hb\kyuyo\PCA���^�f�[�^�~��\"
Public Const dbW As String = "\\192.168.128.4\hb\kyuyo\���^�f�[�^.accdb"
Public Const dbM As String = "\\192.168.128.4\hb\kyuyo\�O���[�v����.accdb"

Public strSQL    As String
Public strTXT    As String '���߰�csv̧��
Public strKBN    As String '�x���敪(K or S)
Public DateA     As Date   '�x����
Public strDate   As String '�x����

Sub AP_END()

    Dim myBook As Workbook
    Dim strFN As String
    Dim boolB As Boolean
    
    Application.ReferenceStyle = xlA1
    
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    
    strFN = ThisWorkbook.Name '���̃u�b�N�̖��O
    boolB = False
    For Each myBook In Workbooks
        If myBook.Name <> strFN Then boolB = True
    Next
    If boolB Then
        ThisWorkbook.Close False  '�t�@�C�������
    Else
        Application.Quit  'Excell���I��
        ThisWorkbook.Saved = True
        ThisWorkbook.Close False
    End If
      
    
End Sub

Public Function CP_NAME() As String

    Const COMPUTERNAMBUFFER_LENGTH = MAX_COMPUTERNAME_LENGTH + 1
    Dim strComputerNameBuffer As String * COMPUTERNAMBUFFER_LENGTH
    Dim lngComputerNameLength As Long
    Dim lngWin32apiResultCode As Long
    
    ' �R���s���[�^�[���̒�����ݒ�
    lngComputerNameLength = Len(strComputerNameBuffer)
    ' �R���s���[�^�[�����擾
    lngWin32apiResultCode = GetComputerName(strComputerNameBuffer, lngComputerNameLength)
    ' �R���s���[�^�[����\��
    CP_NAME = Left(strComputerNameBuffer, InStr(strComputerNameBuffer, vbNullChar) - 1)

End Function

Function ���t����() As String
    
    Dim strDAY As String
    Dim DateA  As Date
    Dim lngM   As Long
    
    '�f�[�^���ͼ�Ăֶ���ް�̓��t���
    DateA = Now()
    DateA = Strings.Format(DateA, "yyyy/mm") & "/25"
    If Weekday(DateA) = vbSunday Then
        DateA = DateA - 2
    ElseIf Weekday(DateA) = vbSaturday Then
        DateA = DateA - 1
    End If
ReReRe:
    strDAY = InputBox("���^�x��������͂��ĉ�����", "�x��������", Strings.Format(DateA, "yyyy/mm/dd"))
    If IsError(CDate(strDAY)) Then
        lngM = MsgBox("���͂������t�̌`�����s���ł�", vbOKCancel, "���t�G���[")
        GoTo ReReRe
    End If
    
    ���t���� = strDAY
    
End Function

Sub �V�[�g���()
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
End Sub

Sub Go_Siwake()
    Sheets("����").Select
End Sub
