Attribute VB_Name = "M03_Update"
Option Explicit

Sub 社員更新()

Const SQL1 = "SELECT * FROM グループ社員マスター WHERE (((事業所区分)='"
Const SQL2 = "') AND ((社員コード)='"
Const SQL3 = "'))"

Const SQL4 = "SELECT * FROM 賃金本給表 WHERE (((等級)="
Const SQL5 = ") AND ((号数)="

Dim cnA      As New ADODB.Connection
Dim rsA      As New ADODB.Recordset
Dim rsT      As New ADODB.Recordset
Dim strSQL   As String
Dim strUNM   As String
Dim strKBN   As String
Dim strDB    As String
Dim lngR     As Long   '行ｶｳﾝﾀ
Dim lngC     As Long   '列ｶｳﾝﾀ
Dim strKey1  As String '事業所区分
Dim strKey2  As String '社員コード
Dim strDel   As String '削除ｷｰ
Dim strDAT1  As String '生年月日
Dim strDAT2  As String '入社年月日
Dim DateS    As Date   '生年月日2
Dim DateN    As Date   '入社年月日2
Dim DateA    As Date   '作業用変数
Dim lngTKY   As Long
Dim lngGSU   As Long

    strKBN = Sheets("Menu").Range("AI5")

    '社員分更新
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
    lngR = 7
    Do
        If Cells(lngR, 2) = "" Then Exit Do
        strDel = StrConv(Left(Cells(lngR, 27), 1), vbUpperCase)  '削除ｷｰ
        strKey1 = StrConv(Left(Cells(lngR, 2), 2), vbUpperCase)  '事業所区分
        strKey2 = Format(Cells(lngR, 3), "00000")                '社員ｺｰﾄﾞ
        lngTKY = Cells(lngR, 12)                                 '等級
        lngGSU = Cells(lngR, 14)                                 '号数
        '本給表とチェック
        If Cells(lngR, 11) = "A" Then
            strSQL = SQL4 & lngTKY & SQL5 & lngGSU & "))"
            rsT.Open strSQL, cnA, adOpenStatic, adLockPessimistic
            If Cells(lngR, 17) = rsT.Fields("本給") And Cells(lngR, 18) = rsT.Fields("加給") Then
            Else
                MsgBox "本給あるいは加給が間違っています！" & vbCrLf & "＝＝＝ " & Cells(lngR, 4) & " ＝＝＝", vbCritical
            End If
            rsT.Close
        End If
        
        strSQL = SQL1 & strKey1 & SQL2 & strKey2 & SQL3
        rsA.Open strSQL, cnA, adOpenStatic, adLockPessimistic
        If strDel = "D" Then '削除する
            If rsA.EOF Then
                strDel = ""
            Else
                rsA.Delete
            End If
        Else
            '生年月日
            If Cells(lngR, 7) = "" Then
                strDAT1 = ""
            Else
                DateS = Cells(lngR, 7)
                strDAT1 = Format(DateA, "yyyymmdd")
                
            End If
            '入社年月日
            If Cells(lngR, 10) = "" Then
                strDAT1 = ""
            Else
                DateN = Cells(lngR, 10)
                strDAT2 = Format(DateA, "yyyymmdd")
            End If
            If rsA.EOF Then
                'マスタに無ければ追加
                rsA.AddNew
                rsA.Fields(0) = strKey1
                rsA.Fields(1) = strKey2
                rsA.Fields(20) = 0
            End If
            If rsA.EOF Then
                rsA.AddNew
                rsA.Fields(0) = Cells(lngR, 2)
                rsA.Fields(1) = Cells(lngR, 3)
            Else
                rsA.MoveFirst
            End If
            rsA.Fields("社員名") = Cells(lngR, 4)
            rsA.Fields("性別") = Cells(lngR, 6)
            rsA.Fields("生年月日") = DateS
            rsA.Fields("入社年月日") = DateN
            rsA.Fields("社員種類") = Cells(lngR, 11)
            rsA.Fields("等級") = Cells(lngR, 12)
            rsA.Fields("号数") = Cells(lngR, 14)
            rsA.Fields("管理職区") = Cells(lngR, 16)
            For lngC = 9 To 15
                If Cells(lngR, lngC + 8) = "" Then '本給->手当
                    rsA.Fields(lngC) = 0
                Else
                    rsA.Fields(lngC) = Cells(lngR, lngC + 8)
                End If
            Next lngC
            rsA.Fields("印刷順序") = Cells(lngR, 25)
            rsA.Fields("所属事業所") = Cells(lngR, 26)
            rsA.Fields("役員就任日") = ""
            rsA.Fields("パート時間給") = 0
            rsA.Fields("パート所定時間数") = 0
            If StrConv(rsA![社員種類], vbUpperCase) = "P" Then
                rsA.Fields("パート時間給") = Cells(lngR, 17) / Cells(lngR, 29) '本給 / ﾊﾟｰﾄ所定時間
                rsA.Fields("パート所定時間数") = Cells(lngR, 29) 'ﾊﾟｰﾄ所定時間
            End If
            rsA.Update
        End If
        rsA.Close
        lngR = lngR + 1
        If lngR = 55 Then lngR = 67
        If lngR > 113 Then Exit Do
    Loop
    
    MsgBox "更新しました！(*'ω'*)", vbInformation, "マスタ更新"
    
Exit_DB:
    cnA.Close
    Set rsA = Nothing
    Set cnA = Nothing
    
End Sub
