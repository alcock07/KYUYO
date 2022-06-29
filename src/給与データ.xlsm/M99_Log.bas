Attribute VB_Name = "M99_Log"
Option Explicit

Public Const APP_NAME = "給与データ取込みExcel"

Sub Open_Log()
Dim strLOG As String
Dim boolA  As Boolean
    strLOG = "Start:" & Format(Now(), "yyyy/mm/dd") & " " & Format(Now(), "hh:mm:ss") & " - " & APP_NAME
    boolA = AddText("X:\admin\alcock.Log", strLOG)
End Sub

Sub Close_Log()
Dim strLOG As String
Dim boolA As Boolean
    strLOG = "End  :" & Format(Now(), "yyyy/mm/dd") & " " & Format(Now(), "hh:mm:ss") & " - " & APP_NAME
    boolA = AddText("X:\admin\alcock.Log", strLOG)
End Sub

Public Function AddText(Fname As String, txt As String) As Boolean
'=============================
'ﾃｷｽﾄﾌｧｲﾙ追加
'FName : 出力ﾌｧｲﾙ名
'txt   : 出力ﾃｷｽﾄ
'=============================
    Dim iFNW
    On Error Resume Next
    iFNW = FreeFile
    Open Fname For Append As iFNW
        Print #iFNW, txt
    Close iFNW
End Function

Public Function WriteText(Fname As String, txt As String) As Boolean
'==============================
'ﾃｷｽﾄﾌｧｲﾙ書込み
'FName : 出力ﾌｧｲﾙ名
'txt   : 出力ﾃｷｽﾄ
'==============================
    Dim iFNW
    On Error Resume Next
    iFNW = FreeFile
    Open Fname For Output As iFNW
        Print #iFNW, txt
    Close iFNW
End Function
