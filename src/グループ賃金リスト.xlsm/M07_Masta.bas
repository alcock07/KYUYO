Attribute VB_Name = "M07_Masta"
Option Explicit

Sub Set_KBN()

Dim strKBN As String
Dim Index  As Long

    strKBN = Range("O2")
    For Index = 3 To 8
        If Cells(Index, 15) = strKBN Then
            Range("P2") = Cells(Index, 16)
            Range("Q2") = Cells(Index, 17)
            Range("R2") = Cells(Index, 18)
            Exit For
        End If
    Next Index
    
    Call Get_Masta(strKBN)
    
End Sub

Sub Get_Masta(strKBN As String)

Const SQL1 = "SELECT 事業所区分, 社員コード, 社員名, 社員種類, 等級, 基本給１, 基本給２, 管理職手当, 家族手当, 部門1, 部門2, 部門3, 部門名, 入社年月日 FROM グループ社員マスター WHERE (((事業所区分)='"
Const SQL2 = "')) ORDER BY 社員コード"

Dim cnA As New ADODB.Connection
Dim rsA As New ADODB.Recordset
Dim strSQL As String
Dim lngR   As Long
Dim lngC   As Long
Dim DateA  As Date
Dim DateB  As Date
Dim strYY  As String
Dim lngMM  As Long
Dim strUNM As String
Dim strDB  As String

    Range("A4:J152").ClearContents
    Range("L4:L52").ClearContents
    Range("N4:N52").ClearContents
    
    strUNM = Strings.UCase(GetUserNameString)
    If strUNM = "SCOTT" Or strUNM = "TAKA" Or strUNM = "SIMO" Then
        If strKBN = "5" Or strKBN = "6" Then
            strDB = dbT
        Else
            strDB = dbK
        End If
'    ElseIf strUNM = "SIMO" Then
'        strDB = dbT
    Else
        Call Back_Menu
        GoTo Exit_DB
    End If
    cnA.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & strDB
    cnA.Open
    
    '事業所区分ごと読込み
    strKBN = Range("Q2")
    If strKBN = "" Then GoTo Exit_DB
    strSQL = SQL1 & strKBN & SQL2
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly

    If rsA.EOF = False Then rsA.MoveFirst
    lngR = 4
    Do Until rsA.EOF
        '各項目ｾｯﾄ
        For lngC = 0 To 8
            Cells(lngR, lngC + 1) = rsA.Fields(lngC)
        Next lngC
        '部門区分ｾｯﾄ
        If IsNull(rsA.Fields("部門2")) = False Then Cells(lngR, 10) = rsA.Fields("部門2")
        If IsNull(rsA.Fields("部門3")) = False Then Cells(lngR, 12) = rsA.Fields("部門3")
        '生年月日
        If rsA.Fields("入社年月日") <> "" Then
            DateA = rsA.Fields("入社年月日")
        End If
        '新入社員判定処理
        strYY = Format(Now(), "yyyy")
        lngMM = Format(Now(), "m")
        If lngMM >= 4 And lngMM <= 7 Then
            lngMM = 1
        ElseIf lngMM >= 10 And lngMM <= 12 Then
            lngMM = 5
        Else
            lngMM = 0
        End If
        If lngMM > 0 Then
            DateB = strYY & "/" & Format(lngMM, "00") & "/01"
            If DateA > DateB Then
                Cells(lngR, 14) = "○"
            Else
                Cells(lngR, 14) = ""
            End If
        End If
        rsA.MoveNext
        lngR = lngR + 1
    Loop
    
Exit_DB:
    rsA.Close
    cnA.Close

    Set rsA = Nothing
    Set cnA = Nothing

End Sub

Sub Up_Masta()

Const SQL1 = "SELECT 部門1, 部門2, 部門3, 部門名, 新入社員 FROM グループ社員マスター WHERE (((社員コード)='"
Const SQL2 = "'))"


Dim cnA As New ADODB.Connection
Dim rsA As New ADODB.Recordset
Dim strSQL As String
Dim strCD  As String
Dim strKB1 As String
Dim strKB2 As String
Dim strKB3 As String
Dim strUNM As String
Dim strKBN As String
Dim strDB  As String
Dim lngR   As Long
Dim lngC   As Long
   
    strKBN = Range("O2")
    strUNM = Strings.UCase(GetUserNameString)
    If strUNM = "SCOTT" Or strUNM = "TAKA" Or strUNM = "SIMO" Then
        If strKBN = "5" Or strKBN = "6" Then
            strDB = dbT
        Else
            strDB = dbK
        End If
'    ElseIf strUNM = "SIMO" Then
'        strDB = dbT
    Else
        GoTo Exit_DB
    End If
    cnA.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & strDB
    cnA.Open
    
    lngR = 4
    Do
        strCD = Cells(lngR, 2) '社員ｺｰﾄﾞ
        If strCD = "" Then Exit Do
        strKB1 = Range("P2")
        strKB2 = Cells(lngR, 10)
        strKB3 = Cells(lngR, 12)
        If strCD <> "" Then
            'ﾏｽﾀ呼出
            strSQL = SQL1 & strCD & SQL2
            rsA.Open strSQL, cnA, adOpenStatic, adLockPessimistic
            If rsA.EOF = False Then
                rsA.MoveFirst
                rsA.Fields(0) = strKB1
                rsA.Fields(1) = strKB2
                rsA.Fields(2) = strKB3
                rsA.Fields(3) = Cells(lngR, 13)
                If Cells(lngR, 14) = "○" Then
                    rsA.Fields(4) = "Y"
                Else
                    rsA.Fields(4) = ""
                End If
                rsA.Update
            End If
            rsA.Close
        End If
        lngR = lngR + 1
    Loop
    
    MsgBox "登録しました(^^♪", vbInformation, "マスタ登録"
Exit_DB:

    '接続のクローズ
    cnA.Close

    'オブジェクトの破棄
    Set rsA = Nothing
    Set cnA = Nothing

End Sub

