Attribute VB_Name = "M01_Main"
Option Explicit

' 仕訳データ生成マクロ =========================
' ﾏｸﾛ記録日 : 2001/07/11  ﾕｰｻﾞｰ名 : Shigeo ITOI
' 更新：2006/07/06  t_takazawa
' 更新：2006/11/30  t_takazawa
' 更新：2014/04/09  t_takazawa
' 更新：2022/06/23  t_takazawa
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
Public Const dtW As String = "\\192.168.128.4\hb\kyuyo\PCA給与データ蓄積\"
Public Const dbW As String = "\\192.168.128.4\hb\kyuyo\給与データ.accdb"
Public Const dbM As String = "\\192.168.128.4\hb\kyuyo\グループ賃金.accdb"

Public strSQL    As String
Public strTXT    As String 'ｲﾝﾎﾟｰﾄcsvﾌｧｲﾙ
Public strKBN    As String '支給区分(K or S)
Public DateA     As Date   '支給月
Public strDate   As String '支給月

Sub AP_END()

    Dim myBook As Workbook
    Dim strFN As String
    Dim boolB As Boolean
    
    Application.ReferenceStyle = xlA1
    
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    
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

Public Function CP_NAME() As String

    Const COMPUTERNAMBUFFER_LENGTH = MAX_COMPUTERNAME_LENGTH + 1
    Dim strComputerNameBuffer As String * COMPUTERNAMBUFFER_LENGTH
    Dim lngComputerNameLength As Long
    Dim lngWin32apiResultCode As Long
    
    ' コンピューター名の長さを設定
    lngComputerNameLength = Len(strComputerNameBuffer)
    ' コンピューター名を取得
    lngWin32apiResultCode = GetComputerName(strComputerNameBuffer, lngComputerNameLength)
    ' コンピューター名を表示
    CP_NAME = Left(strComputerNameBuffer, InStr(strComputerNameBuffer, vbNullChar) - 1)

End Function

Function 日付入力() As String
    
    Dim strDAY As String
    Dim DateA  As Date
    Dim lngM   As Long
    
    'データ入力ｼｰﾄへｶﾚﾝﾀﾞｰの日付ｾｯﾄ
    DateA = Now()
    DateA = Strings.Format(DateA, "yyyy/mm") & "/25"
    If Weekday(DateA) = vbSunday Then
        DateA = DateA - 2
    ElseIf Weekday(DateA) = vbSaturday Then
        DateA = DateA - 1
    End If
ReReRe:
    strDAY = InputBox("給与支給日を入力して下さい", "支給日入力", Strings.Format(DateA, "yyyy/mm/dd"))
    If IsError(CDate(strDAY)) Then
        lngM = MsgBox("入力した日付の形式が不正です", vbOKCancel, "日付エラー")
        GoTo ReReRe
    End If
    
    日付入力 = strDAY
    
End Function

Sub シート印刷()
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
End Sub

Sub Go_Siwake()
    Sheets("入力").Select
End Sub
