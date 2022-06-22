Attribute VB_Name = "M07_Masta"
Option Explicit

Sub Set_KBN()

Dim strKBN As String
Dim Index  As Long

    strKBN = Range("O2") '拠点番号
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

Dim cnA As New ADODB.Connection
Dim rsA As New ADODB.Recordset
Dim strSQL As String
Dim strNT  As String
Dim lngR   As Long
Dim lngC   As Long
Dim DateA  As Date
Dim DateB  As Date
Dim strYY  As String
Dim lngMM  As Long

    Range("A4:J152").ClearContents
    Range("L4:L52").ClearContents
    Range("N4:N52").ClearContents

    strNT = "Initial Catalog=KYUYO;"
    cnA.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD 'SQLServer
    cnA.Open

    '事業所区分ごと読込み
    If strKBN = "" Then GoTo Exit_DB
    strKBN = Range("Q2")
    strSQL = ""
    strSQL = strSQL & "SELECT KBN,"
    strSQL = strSQL & "       SCODE,"
    strSQL = strSQL & "       SNAME,"
    strSQL = strSQL & "       SKBN,"
    strSQL = strSQL & "       CLASS,"
    strSQL = strSQL & "       PAY1,"
    strSQL = strSQL & "       PAY2,"
    strSQL = strSQL & "       OPT1,"
    strSQL = strSQL & "       OPT2,"
    strSQL = strSQL & "       BMN1,"
    strSQL = strSQL & "       BMN2,"
    strSQL = strSQL & "       BMN3,"
    strSQL = strSQL & "       BMNNM,"
    strSQL = strSQL & "       DATE2"
    strSQL = strSQL & "    FROM KYUMTA"
    strSQL = strSQL & "         WHERE KBN = '" & strKBN & "'"
    strSQL = strSQL & "         AND DATKB = '1'"
    strSQL = strSQL & "    ORDER BY SCODE"
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly

    If rsA.EOF = False Then rsA.MoveFirst
    lngR = 4
    Do Until rsA.EOF
        '各項目ｾｯﾄ
        For lngC = 0 To 8
            Cells(lngR, lngC + 1) = rsA.Fields(lngC)
        Next lngC
        '部門区分ｾｯﾄ
        If IsNull(rsA.Fields("BMN2")) = False Then Cells(lngR, 10) = rsA.Fields("BMN2")
        If IsNull(rsA.Fields("BMN3")) = False Then Cells(lngR, 12) = rsA.Fields("BMN3")
        '生年月日
        If rsA.Fields("DATE2") <> "" Then
            DateA = rsA.Fields("DATE2")
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

    If Not rsA Is Nothing Then
        If rsA.State = adStateOpen Then rsA.Close
        Set rsA = Nothing
    End If
    If Not cnA Is Nothing Then
        If cnA.State = adStateOpen Then cnA.Close
        Set cnA = Nothing
    End If

End Sub

Sub Up_Masta()

Const SQL1 = "SELECT 部門1, 部門2, 部門3, 部門名, 新入社員 FROM グループ社員マスター WHERE (((社員コード)='"
Const SQL2 = "'))"

Dim cnA As New ADODB.Connection
Dim rsA As New ADODB.Recordset
Dim strSQL As String
Dim strNT  As String
Dim strCD  As String
Dim strKB1 As String
Dim strKB2 As String
Dim strKB3 As String
Dim lngR   As Long

    strNT = "Initial Catalog=KYUYO;"
    cnA.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD 'SQLServer
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
            strSQL = ""
            strSQL = strSQL & "SELECT BMN1,"
            strSQL = strSQL & "       BMN2,"
            strSQL = strSQL & "       BMN3,"
            strSQL = strSQL & "       BMNNM,"
            strSQL = strSQL & "       YKBN"
            strSQL = strSQL & "    FROM KYUMTA"
            strSQL = strSQL & "         WHERE SCODE = '" & strCD & "'"
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

    If Not rsA Is Nothing Then
        If rsA.State = adStateOpen Then rsA.Close
        Set rsA = Nothing
    End If
    If Not cnA Is Nothing Then
        If cnA.State = adStateOpen Then cnA.Close
        Set cnA = Nothing
    End If
    
End Sub
