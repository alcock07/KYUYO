Attribute VB_Name = "M03_Update"
Option Explicit

Sub 社員更新()

'Const SQL1 = "SELECT * FROM グループ社員マスター WHERE (((事業所区分)='"
'Const SQL2 = "') AND ((社員コード)='"
'Const SQL3 = "'))"
'
'Const SQL4 = "SELECT * FROM 賃金本給表 WHERE (((等級)="
'Const SQL5 = ") AND ((号数)="

Dim cnA      As New ADODB.Connection
Dim rsA      As New ADODB.Recordset
Dim Cmd      As New ADODB.Command
Dim rsT      As New ADODB.Recordset
Dim strSQL   As String
Dim strNT    As String
Dim lngR     As Long   '行ｶｳﾝﾀ
Dim lngC     As Long   '列ｶｳﾝﾀ
Dim strKey1  As String '事業所区分
Dim strKey2  As String '社員コード
Dim strDel   As String '削除ｷｰ
Dim lngTKY   As Long   '等級
Dim lngGSU   As Long   '号数

    '社員分更新
    strNT = "Initial Catalog=KYUYO;"
    cnA.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD 'SQLServer
    cnA.Open
    Set Cmd.ActiveConnection = cnA
    lngR = 7
    Do
        If Cells(lngR, 2) = "" Then Exit Do
        strDel = StrConv(Left(Cells(lngR, 27), 1), vbUpperCase)  '削除ｷｰ
        strKey1 = StrConv(Left(Cells(lngR, 2), 2), vbUpperCase)  '事業所区分
        strKey2 = Format(Cells(lngR, 3), "00000")                '社員ｺｰﾄﾞ
        lngTKY = Cells(lngR, 10)                                 '等級
        lngGSU = Cells(lngR, 12)                                 '号数
        '本給表とチェック
        If Cells(lngR, 11) = "A" Then
            strSQL = ""
            strSQL = strSQL & "SELECT * "
            strSQL = strSQL & "       FROM KYUHYO"
            strSQL = strSQL & "            WHERE CLASS = '" & lngTKY & "'"
            strSQL = strSQL & "            AND ISSUE = '" & lngGSU & "'"
            rsT.Open strSQL, cnA, adOpenStatic, adLockPessimistic
            If Cells(lngR, 17) = rsT.Fields("PAY1") And Cells(lngR, 18) = rsT.Fields("PAY2") Then
            Else
                MsgBox "本給あるいは加給が間違っています！" & vbCrLf & "＝＝＝ " & Cells(lngR, 4) & " ＝＝＝", vbCritical
            End If
            If rsT.State = adStateOpen Then rsT.Close
        End If
        
        strSQL = ""
        strSQL = strSQL & "SELECT * "
        strSQL = strSQL & "       FROM KYUMTA"
        strSQL = strSQL & "            WHERE KBN = '" & strKey1 & "'"
        strSQL = strSQL & "            AND SCODE = '" & strKey2 & "'"
        rsA.Open strSQL, cnA, adOpenStatic, adLockPessimistic
        If strDel = "D" Then '削除する
            If rsA.EOF Then
                strDel = ""
            Else
                rsA.Fields(0) = "1"
            End If
        Else
            If rsA.EOF Then
                'マスタに無ければ追加
                rsA.AddNew
                rsA.Fields("DATKB") = "1"
                rsA.Fields("KBN") = strKey1
                rsA.Fields("SCODE") = strKey2
                rsA.Fields("HOUR") = 0
            End If
            If rsA.EOF Then
                rsA.AddNew
                rsA.Fields("KBN") = Cells(lngR, 2)
                rsA.Fields("SCODE") = Cells(lngR, 3)
            Else
                rsA.MoveFirst
            End If
            rsA.Fields("SNAME") = Cells(lngR, 4)
            rsA.Fields("SEX") = Cells(lngR, 5)
            rsA.Fields("DATE1") = CDate(Cells(lngR, 7))
            rsA.Fields("DATE2") = CDate(Cells(lngR, 8))
            rsA.Fields("SKBN") = Cells(lngR, 9)
            rsA.Fields("CLASS") = Cells(lngR, 10)
            rsA.Fields("ISSUE") = Cells(lngR, 12)
            rsA.Fields("MGR") = Cells(lngR, 14)
            For lngC = 10 To 16
                If Cells(lngR, lngC + 5) = "" Then '本給->手当
                    rsA.Fields(lngC) = 0
                Else
                    rsA.Fields(lngC) = Cells(lngR, lngC + 5)
                End If
            Next lngC
            rsA.Fields("PRN") = Cells(lngR, 23)
            rsA.Fields("OFFICE") = Cells(lngR, 24)
            rsA.Fields("JIKYU") = 0
            rsA.Fields("HOUR") = 0
            If StrConv(rsA![SKBN], vbUpperCase) = "P" Then
                rsA.Fields("JIKYU") = Cells(lngR, 15) / Cells(lngR, 27) '本給 / ﾊﾟｰﾄ所定時間
                rsA.Fields("HOUR") = Cells(lngR, 27) 'ﾊﾟｰﾄ所定時間
            End If
            rsA.Update
        End If
        rsA.Close
        lngR = lngR + 1
        If lngR = 54 Then lngR = 66
        If lngR > 112 Then Exit Do
    Loop
    
    MsgBox "更新しました！(*'ω'*)", vbInformation, "マスタ更新"
    
Exit_DB:

    If Not rsT Is Nothing Then
        If rsT.State = adStateOpen Then rsT.Close
        Set rsT = Nothing
    End If
    If Not rsA Is Nothing Then
        If rsA.State = adStateOpen Then rsA.Close
        Set rsA = Nothing
    End If
    If Not cnA Is Nothing Then
        If cnA.State = adStateOpen Then cnA.Close
        Set cnA = Nothing
    End If
    
End Sub
