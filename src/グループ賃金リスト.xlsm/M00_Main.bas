Attribute VB_Name = "M00_Main"
Option Explicit

' �O���[�v�������X�g �}�N��
'2000/06/14 �쐬 : Shigeo ITOI
'2006/07/19 �X�V : takazawa
'2008/04/17 �X�V : takazawa
'2011/08/31 �X�V : takazawa
'2013/01/29 �X�V �Ftakazawa
'2021/03/04 �X�V �Ftakazawa
'2022/05/13 �X�V �Ftakazawa Git�o�^

Public Const MYPROVIDERE = "Provider=SQLOLEDB;"
'Public Const MYSERVER = "Data Source=192.168.128.9\SQLEXPRESS;"
Public Const MYSERVER = "Data Source=HB14\SQLEXPRESS;"
Public Const USER = "User ID=sa;"
'Public Const PSWD = "Password=ALCadmin!;"
Public Const PSWD = "Password=admin;"


'Public Const dbK = "\\192.168.128.4\hb\KYUYO\�O���[�v����.accdb"
'Public Const dbT = "\\192.168.128.4\hb\ta\���^�V�X�e��\�O���[�v����.accdb"

'=== API �֐��錾 ===
#If VBA7 Then
    Declare PtrSafe Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
#Else
    Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
#End If
Public Const MAX_COMPUTERNAME_LENGTH = 15

Public lngGN As Long
Public Enum PrBookOnApplication
    prBookOnApplicationOpened
    prBookOnApplicationNotOpened
    prBookOnApplicationNotExist
    prBookOnApplicationSameNameBookOpened
End Enum

'=== ���߭�����擾�֐� ===
Public Function CP_NAME() As String

    Const COMPUTERNAMBUFFER_LENGTH = MAX_COMPUTERNAME_LENGTH + 1
    Dim strComputerNameBuffer As String * COMPUTERNAMBUFFER_LENGTH
    Dim lngComputerNameLength As Long
    Dim lngWin32apiResultCode As Long
    
    ' �R���s���[�^�[���̒�����ݒ�
    lngComputerNameLength = Len(strComputerNameBuffer)
    ' �R���s���[�^�[�����擾
    lngWin32apiResultCode = GetComputerName(strComputerNameBuffer, _
                                            lngComputerNameLength)
    ' �R���s���[�^�[����\��
    CP_NAME = Strings.Left(strComputerNameBuffer, InStr(strComputerNameBuffer, vbNullChar) - 1)
    
End Function

'=== հ�ް���擾�֐� ===
Public Function GetUserNameString() As String
    Dim strNAME As String * 256     '������o�b�t�@
    Dim lngSize As Long             '������̒���
    Dim lngRet As Long              'API�֐��̖߂�l
    
    On Error GoTo ErrorHandle
    
    lngSize = Len(strNAME) - 1                      '������o�b�t�@�T�C�Y��ݒ�
    lngRet = GetUserName(strNAME, lngSize)          'API�֐��ɂ��R���s���[�^�����擾
    If lngRet = 0 Then
        '�G���[�����������ꍇ
        GetUserNameString = ""
    Else
        'API�֐�����I��
        GetUserNameString = Strings.Left(strNAME, lngSize - 1)  '�L��������݂̂�Ԃ�
        'GetUserName�͑�2�����Ƀk���������܂߂����������i�[����
    End If
    
    Exit Function
ErrorHandle:
    '�G���[�����������ꍇ
    GetUserNameString = ""
End Function

'=== EXCEL���J���Ă��邩�`�F�b�N ===
Public Function BookOnApplication(myBookFullName As String) As PrBookOnApplication
    Dim myAllFileAttr As VbFileAttribute
    Dim myBookName As String
    Dim myBook As Workbook
    myAllFileAttr = vbArchive Or vbHidden Or vbReadOnly Or vbSystem
    myBookName = Dir(myBookFullName, myAllFileAttr)
    If Len(myBookName) = 0 Then
        BookOnApplication = prBookOnApplicationNotExist
        Exit Function
    End If
    On Error Resume Next
    Set myBook = Workbooks(myBookName)
    On Error GoTo 0
    If myBook Is Nothing Then
        BookOnApplication = prBookOnApplicationNotOpened
    ElseIf UCase(myBook.FullName) = UCase(myBookFullName) Then
        BookOnApplication = prBookOnApplicationOpened
    Else
        BookOnApplication = prBookOnApplicationSameNameBookOpened
    End If
End Function

Sub Back_Menu()
    Sheets("Menu").Activate
    Range("A1").Select
End Sub

Sub Move_Masta()
    Sheets("Masta").Activate
    Range("A1").Select
End Sub

Sub �I������()
    CLR_CELL
    AP_END
End Sub

Sub AP_END()

    Dim myBook As Workbook
    Dim strFN As String
    Dim boolB As Boolean
    
     Application.ReferenceStyle = xlA1
    
    Application.DisplayAlerts = False
    
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
