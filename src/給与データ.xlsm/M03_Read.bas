Attribute VB_Name = "M03_Read"
Option Explicit

'=== PCAﾃﾞｰﾀより抽出 ﾎﾞﾀﾝ===

Sub データ読込()

    '[蓄積データ]から年月毎のデータを抽出
    '部門毎に分けてシートに貼り付ける
    
    Dim cnA            As New ADODB.Connection
    Dim rsS            As ADODB.Recordset
    Dim strYM          As String '年月
    Dim strKS          As String
    Dim lngC           As Long   'ﾙｰﾌﾟｶｳﾝﾀ
    Dim lngR           As Long   '  〃
    Dim lngKIN(17, 21) As Long   '金額格納配列
    Dim K_cell         As Range  '基準ｾﾙ
    
    '初期化
    Erase lngKIN
    Sheets("入力").Select
    Range("C11:S30").ClearContents
    Range("C37:S37").ClearContents
    Range("L2") = 0
    strDate = Range("B4").Value  '支給日取得
    strYM = Strings.Format(strDate, "yyyy/mm")
    
    '[蓄積データ]からﾃﾞｰﾀ取得して配列に格納（支店ごとに集計）
    cnA.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & dbW '給与データ
    cnA.Open
    Set rsS = New ADODB.Recordset
    If Range("C4") = "給料" Then
        strKS = "K"
    Else
        strKS = "S"
    End If
    strSQL = ""
    strSQL = strSQL & "SELECT *"
    strSQL = strSQL & "         FROM 蓄積データ"
    strSQL = strSQL & "                        WHERE 支給年月 = '" & strYM & "'"
    strSQL = strSQL & "                        And 給与区分 = '" & strKS & "'"
    strSQL = strSQL & "         ORDER BY 部門区分"
    rsS.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsS.EOF = False Then rsS.MoveFirst
    Do Until rsS.EOF
        If rsS!差引支給額 <> 0 Then
            '列判定
            lngC = 部門判定(rsS!社員コード)
            If lngC = 0 Then
                Exit Sub
            End If
            'ﾃﾞｰﾀ貼り付け
            lngKIN(lngC, 1) = lngKIN(lngC, 1) + 1  '人数
            lngKIN(lngC, 2) = lngKIN(lngC, 2) + (rsS!差引支給額 - rsS!非課税交通費)  '振込額
            lngKIN(lngC, 4) = lngKIN(lngC, 4) + rsS!非課税交通費  '非課税交通費
            lngKIN(lngC, 5) = lngKIN(lngC, 5) + rsS!健康保険料
            lngKIN(lngC, 6) = lngKIN(lngC, 6) + rsS!介護保険料
            lngKIN(lngC, 7) = lngKIN(lngC, 7) + rsS!厚生年金保険料
            lngKIN(lngC, 8) = lngKIN(lngC, 8) + rsS!確定拠出年金
            lngKIN(lngC, 9) = lngKIN(lngC, 9) + rsS!雇用保険料
            lngKIN(lngC, 10) = lngKIN(lngC, 10) + rsS!源泉所得税
            If Range("C4") = "給料" Then
                lngKIN(lngC, 11) = lngKIN(lngC, 11) + rsS!特徴市民税
                lngKIN(lngC, 12) = lngKIN(lngC, 12) + rsS!貸付金       '貸付金
                lngKIN(lngC, 14) = lngKIN(lngC, 14) + rsS!クック会     'クック会
                lngKIN(lngC, 15) = lngKIN(lngC, 15) + rsS!旅行積立金   '旅行積立金
                lngKIN(lngC, 16) = lngKIN(lngC, 16) + rsS!財形貯蓄     '財形貯蓄
                lngKIN(lngC, 17) = lngKIN(lngC, 17) + rsS!預金預かり金 '家賃等
                lngKIN(lngC, 18) = lngKIN(lngC, 18) + rsS!その他控除分
                lngKIN(lngC, 19) = lngKIN(lngC, 19) + rsS!食事代預かり
                lngKIN(lngC, 20) = lngKIN(lngC, 20) + rsS!その他預かり金
            Else
                lngKIN(lngC, 12) = lngKIN(lngC, 12) + rsS!貸付金       '貸付金
                lngKIN(lngC, 16) = lngKIN(lngC, 16) + rsS!財形貯蓄     '財形貯蓄
            End If
            If rsS!雇用保険料 <> 0 Then lngKIN(lngC, 21) = lngKIN(lngC, 21) + rsS!総支給額 '雇用保険対象支給額
        End If
        rsS.MoveNext
    Loop
    
    If Not rsS Is Nothing Then
        If rsS.State = adStateOpen Then rsS.Close
        Set rsS = Nothing
    End If
    If Not cnA Is Nothing Then
        If cnA.State = adStateOpen Then cnA.Close
        Set cnA = Nothing
    End If
    
    'シートに貼り付けします
    Set K_cell = Sheets("入力").Range("C10")
    For lngC = 1 To 17 Step 1
        If lngKIN(lngC, 1) > 0 Then
            For lngR = 1 To 20 Step 1
                If lngKIN(lngC, lngR) = 0 Then
                    K_cell.Offset(lngR, lngC - 1).Value = ""
                Else
                    K_cell.Offset(lngR, lngC - 1).Value = lngKIN(lngC, lngR)
                End If
            Next lngR
            K_cell.Offset(27, lngC - 1).Value = lngKIN(lngC, 21)
        End If
    Next lngC
    Range("C11").Select
    
End Sub

Function 部門判定(strSCD As String) As Long
    
    Dim cnA     As New ADODB.Connection
    Dim rsS     As New ADODB.Recordset
    Dim Cmd     As New ADODB.Command
    Dim strNT   As String
    Dim lngBMN  As Long
        
    strNT = "Initial Catalog=KYUYO;"
    cnA.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD 'SQLServer
    cnA.Open
    Set Cmd.ActiveConnection = cnA
    
    strSQL = ""
    strSQL = strSQL & "SELECT *"
    strSQL = strSQL & "　　　　FROM KYUMTA"
    strSQL = strSQL & "             WHERE SCODE Like '" & "%" & Right(strSCD, 3) & "'"
    strSQL = strSQL & "             And Left(KBN,1) = 'R'"
    Cmd.CommandText = strSQL
    Set rsS = Cmd.Execute
    If rsS.EOF = False Then
        rsS.MoveFirst
        If rsS.Fields("BMN3") = "" Or IsNull(rsS.Fields("BMN3")) Then
            MsgBox "社員の部門が登録されていません！" & vbCrLf & "登録してから処理をやり直して下さい。(ToT)/~~~" & _
            "社員コード=" & strSCD
            部門判定 = 0
            GoTo Exit_DB
        End If
        Select Case rsS.Fields("KBN")
            Case "RH"
                If rsS.Fields("BMN3") = "41" Then '貿易
                    lngBMN = 3
                Else
                    lngBMN = 1
                End If
            Case "RO"
                If rsS.Fields("BMN3") = "19" Then     '福岡営業
                    lngBMN = 5
                ElseIf rsS.Fields("BMN3") = "20" Then '福岡一般
                        lngBMN = 5
                ElseIf rsS.Fields("BMN3") = "18" Then '大阪加工
                        lngBMN = 7
                ElseIf rsS.Fields("BMN3") = "21" Then '福岡加工
                        lngBMN = 8
                ElseIf rsS.Fields("BMN3") = "22" Then '名古屋営業
                        lngBMN = 6
                ElseIf rsS.Fields("BMN3") = "23" Then '名古屋一般
                        lngBMN = 6
                ElseIf rsS.Fields("BMN3") = "24" Then '名古屋加工
                        lngBMN = 9
                Else
                    lngBMN = 4
                End If
            Case "RT"
                If rsS.Fields("BMN3") = "31" Then '東京加工
                    lngBMN = 11
                ElseIf rsS.Fields("BMN3") = "27" Then '南関東営業
                    lngBMN = 12
                ElseIf rsS.Fields("BMN3") = "32" Then '南関東一般
                    lngBMN = 12
                ElseIf rsS.Fields("BMN3") = "33" Then '南関東加工
                    lngBMN = 13
                ElseIf rsS.Fields("BMN3") = "28" Then '仙台営業
                    lngBMN = 14
                ElseIf rsS.Fields("BMN3") = "34" Then '仙台一般
                    lngBMN = 14
                ElseIf rsS.Fields("BMN3") = "35" Then '仙台加工
                    lngBMN = 15
                ElseIf rsS.Fields("BMN3") = "29" Then '北関東営業
                    lngBMN = 16
                ElseIf rsS.Fields("BMN3") = "36" Then '北関東一般
                    lngBMN = 16
                ElseIf rsS.Fields("BMN3") = "37" Then '北関東加工
                    lngBMN = 17
                Else
                    lngBMN = 10
                End If
            Case "RX", "TA", "KA"
                If rsS.Fields("SCODE") = "00089" Or rsS.Fields("SCODE") = "00490" Or rsS.Fields("SCODE") = "00472" Or rsS.Fields("SCODE") = "00694" Then '本部
                    lngBMN = 2
                ElseIf rsS.Fields("SCODE") = "00497" Then  '大阪支店
                    lngBMN = 4
                ElseIf rsS.Fields("SCODE") = "00526" Then  '東京支店
                    lngBMN = 10
                Else
                    MsgBox "社員コードが不正です！" & vbCrLf & "要調査！！(ToT)/~~~" & _
                    "社員コード=" & strSCD
                End If
        End Select
        部門判定 = lngBMN
    Else
        MsgBox "社員マスターに登録がないデータがあります！" & vbCrLf & "マスタに登録してからやり直して下さい" & _
        "社員コード=" & strSCD
        部門判定 = 0
        GoTo Exit_DB
    End If

Exit_DB:

    If Not rsS Is Nothing Then
        If rsS.State = adStateOpen Then rsS.Close
        Set rsS = Nothing
    End If
    If Not cnA Is Nothing Then
        If cnA.State = adStateOpen Then cnA.Close
        Set cnA = Nothing
    End If
    
End Function
