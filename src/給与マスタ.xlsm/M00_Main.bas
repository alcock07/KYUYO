Attribute VB_Name = "M00_Main"
Option Explicit

' グループ賃金リスト マクロ
'2000/06/14 作成 : Shigeo ITOI
'2006/07/19 更新 : takazawa
'2008/04/17 更新 : takazawa
'2011/08/31 更新 : takazawa
'2013/01/29 更新 ：takazawa
'2021/03/04 更新 ：takazawa
'2022/05/13 更新 ：takazawa Git登録

Public Const MYPROVIDERE = "Provider=SQLOLEDB;"
'Public Const MYSERVER = "Data Source=192.168.128.9\SQLEXPRESS;"
Public Const MYSERVER = "Data Source=HB14\SQLEXPRESS;"
Public Const USER = "User ID=sa;"
'Public Const PSWD = "Password=ALCadmin!;"
Public Const PSWD = "Password=admin;"


'Public Const dbK = "\\192.168.128.4\hb\KYUYO\グループ賃金.accdb"
'Public Const dbT = "\\192.168.128.4\hb\ta\給与システム\グループ賃金.accdb"

'=== API 関数宣言 ===
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

'=== ｺﾝﾋﾟｭｰﾀ名取得関数 ===
Public Function CP_NAME() As String

    Const COMPUTERNAMBUFFER_LENGTH = MAX_COMPUTERNAME_LENGTH + 1
    Dim strComputerNameBuffer As String * COMPUTERNAMBUFFER_LENGTH
    Dim lngComputerNameLength As Long
    Dim lngWin32apiResultCode As Long
    
    ' コンピューター名の長さを設定
    lngComputerNameLength = Len(strComputerNameBuffer)
    ' コンピューター名を取得
    lngWin32apiResultCode = GetComputerName(strComputerNameBuffer, _
                                            lngComputerNameLength)
    ' コンピューター名を表示
    CP_NAME = Strings.Left(strComputerNameBuffer, InStr(strComputerNameBuffer, vbNullChar) - 1)
    
End Function

'=== ﾕｰｻﾞｰ名取得関数 ===
Public Function GetUserNameString() As String
    Dim strNAME As String * 256     '文字列バッファ
    Dim lngSize As Long             '文字列の長さ
    Dim lngRet As Long              'API関数の戻り値
    
    On Error GoTo ErrorHandle
    
    lngSize = Len(strNAME) - 1                      '文字列バッファサイズを設定
    lngRet = GetUserName(strNAME, lngSize)          'API関数によりコンピュータ名を取得
    If lngRet = 0 Then
        'エラーが発生した場合
        GetUserNameString = ""
    Else
        'API関数正常終了
        GetUserNameString = Strings.Left(strNAME, lngSize - 1)  '有効文字列のみを返す
        'GetUserNameは第2引数にヌル文字を含めた文字数を格納する
    End If
    
    Exit Function
ErrorHandle:
    'エラーが発生した場合
    GetUserNameString = ""
End Function

'=== EXCELが開いているかチェック ===
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

Sub 終了処理()
    CLR_CELL
    AP_END
End Sub

Sub AP_END()

    Dim myBook As Workbook
    Dim strFN As String
    Dim boolB As Boolean
    
     Application.ReferenceStyle = xlA1
    
    Application.DisplayAlerts = False
    
    strFN = ThisWorkbook.Name 'このブックの名前
    boolB = False
    For Each myBook In Workbooks
        If myBook.Name <> strFN Then boolB = True
    Next
    If boolB Then
        ThisWorkbook.Close False  'ファイルを閉じる
    Else
        Application.Quit  'Excellを終了
        ThisWorkbook.Saved = True
        ThisWorkbook.Close False
    End If
      
    
End Sub
