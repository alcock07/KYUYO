Attribute VB_Name = "M02_Data"
Option Explicit

Private strBMN As String
Private strBKA As String
Private strSCD As String

Sub Proc_Get()
    
    'テキストデータを読み込んでDataシートへ書き込み
    '給与データの[インポートデータ]テーブルへ入れる
    '[インポートデータ]のデータを精査して蓄積データへ移行する
    
    If Range("S11") = 1 Then
        strKBN = "K"
    ElseIf Range("S11") = 2 Then
        strKBN = "S"
    Else
        strKBN = InputBox("支給区分を入力して下さい。" & vbCrLf & "給与=K , 賞与=S", "支給区分ｾｯﾄ", "K")
    End If
    
    'テキストデータを読み込んでDataシートへ書き込み
    READ_TextFile
    
    'DataシートのﾃﾞｰﾀをDBのインポートデータ(Tempﾃｰﾌﾞﾙ)へ書き込み
    If strKBN = "S" Then
        Call 賞与データ移動
    Else
        Call 給与データ移動
    End If
    
    '[インポートデータ]のデータを蓄積データへ移行する
    Call 蓄積処理
    
     MsgBox "データ読み込みが完了しました。"
     
     Sheets("Menu").Select
    
End Sub

Sub READ_TextFile()

    Const cnsTITLE = "テキストファイル読み込み処理"
    Const cnsFILTER = "テキスト形式ファイル (*.txt),*.txt,全てのファイル(*.*),*.*"
    
    Dim xlAPP       As Application ' Applicationオブジェクト
    Dim intFF       As Integer     ' FreeFile値
    Dim strFileName As String      ' OPENするファイル名(フルパス)
    Dim vFileName   As Variant     ' ファイル名受取り用
    Dim X(1 To 54)  As Variant     ' 読み込んだレコード内容
    Dim lngR        As Long        ' 収容するセルの行
    Dim lngCnt      As Long        ' レコード件数カウンタ
    
'    strKBN = Sheets("Wait").Range("S11")
    Sheets("Data").Select
    Range("A1:BB100").ClearContents
    'Applicationオブジェクト取得
    Set xlAPP = Application
    '｢ファイルを開く｣のフォームでファイル名の指定を受ける
    xlAPP.StatusBar = "読み込むファイル名を指定して下さい。"
    ChDrive "K:"
    ChDir dtW
    vFileName = xlAPP.GetOpenFilename(cnsFILTER, 1, cnsTITLE, , False)
    'キャンセルされた場合はFalseが返るので以降の処理は行なわない
    If VarType(vFileName) = vbBoolean Then Exit Sub
    strFileName = vFileName

    'FreeFile値の取得(以降この値で入出力する)
    intFF = FreeFile
    '指定ファイルをOPEN(入力モード)
    Open strFileName For Input As #intFF
    lngR = 0
    'ファイルのEOFまで繰り返す
    Do Until EOF(intFF)
        'レコード件数カウンタの加算
        lngCnt = lngCnt + 1
        xlAPP.StatusBar = "読み込み中です．．．．(" & lngCnt & "レコード目)"
        'レコードを読み込む
        If strKBN = "K" Then
            Input #intFF, X(1), X(2), X(3), X(4), X(5), X(6), X(7), X(8), X(9), X(10), _
                          X(11), X(12), X(13), X(14), X(15), X(16), X(17), X(18), X(19), X(20), _
                          X(21), X(22), X(23), X(24), X(25), X(26), X(27), X(28), X(29), X(30), _
                          X(31), X(32), X(33), X(34), X(35), X(36), X(37), X(38), X(39), X(40), _
                          X(41), X(42), X(43), X(44), X(45), X(46), X(47), X(48), X(49), X(50), _
                          X(51), X(52), X(53), X(54)
        Else
            Input #intFF, X(1), X(2), X(3), X(4), X(5), X(6), X(7), X(8), X(9), X(10), _
                          X(11), X(12), X(13), X(14), X(15), X(16), X(17), X(18), X(19), X(20), _
                          X(21), X(22), X(23), X(24), X(25), X(26), X(27), X(28), X(29), X(30), _
                          X(31), X(32), X(33), X(34), X(35), X(36), X(37), X(38), X(39), X(40), _
                          X(41), X(42), X(43), X(44), X(45)
        End If
        '行を加算しA～E列にレコード内容を表示
        lngR = lngR + 1
        If strKBN = "K" Then
            Range(Cells(lngR, 1), Cells(lngR, 54)).Value = X   ' 配列渡し
        Else
            Range(Cells(lngR, 1), Cells(lngR, 45)).Value = X
        End If
    Loop
    
    xlAPP.StatusBar = False
    
End Sub

Sub 賞与データ移動()

'最初に給与データの中のインポートデータをクリアする
'シートDataにあるデータをインポートデータに入れる

Dim cnW As New ADODB.Connection
Dim rsK As New ADODB.Recordset
Dim lngR   As Long
Dim lngC   As Long

    cnW.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & dbW '給与データ
    cnW.Open
    
    'インポートデータクリア
    strSQL = "DELETE FROM インポートデータ"
    rsK.Open strSQL, cnW, adOpenStatic, adLockPessimistic
    
    'インポートデータオープン
    rsK.Open "インポートデータ", cnW, adOpenStatic, adLockPessimistic
        
    Sheets("Data").Select
    lngR = 3
    Do
        rsK.AddNew
        strBMN = Strings.Format(Cells(lngR, 1), "000")   '部門ｺｰﾄﾞ
        strBKA = Strings.Format(Cells(lngR, 2), "000")   '部課ｺｰﾄﾞ
        strSCD = Strings.Format(Cells(lngR, 3), "00000") '社員ｺｰﾄﾞ
                
        '部門区分の設定(rsK.Fields(0))
        Select Case strBMN
            Case "100"
                rsK.Fields(0) = "000"
            Case "200"
                If strBKA = "010" Then
                    rsK.Fields(0) = "010"
                ElseIf strBKA = "020" Then
                    rsK.Fields(0) = "020"
                ElseIf strBKA = "030" Then
                    rsK.Fields(0) = "030"
                End If
            Case "300"
                If strBKA = "010" Then
                    rsK.Fields(0) = "040"
                ElseIf strBKA = "030" Then
                    rsK.Fields(0) = "050"
                ElseIf strBKA = "040" Then
                    rsK.Fields(0) = "060"
                ElseIf strBKA = "050" Then
                    rsK.Fields(0) = "070"
                End If
            Case "400"
                Select Case strSCD
                    Case "00089"
                        rsK.Fields(0) = "000"
                    Case "00472"
                        rsK.Fields(0) = "000"
                    Case "00490"
                        rsK.Fields(0) = "000"
                    Case "00491"
                        rsK.Fields(0) = "000"
                    Case "00694"
                        rsK.Fields(0) = "000"
                    Case "00497"
                        rsK.Fields(0) = "010"
                    Case "00215"
                        rsK.Fields(0) = "040"
                    Case "00526"
                        rsK.Fields(0) = "040"
                    Case Else
                        MsgBox "識別できない役員のコードがあります！" & vbCrLf & strSCD, vbCritical
                        rsK.Fields(0) = "400"
                End Select
            Case Else
                MsgBox "識別出来ない部門があるようです。" & vbCrLf & "確認して下さい。" & strBMN & " - " & strBKA
        End Select
        
        rsK.Fields(45) = strBKA '部課ｺｰﾄﾞ
        rsK.Fields(1) = strSCD  '社員ｺｰﾄﾞ
        '社員名～支給15
        For lngC = 2 To 25
            rsK.Fields(lngC) = Cells(lngR, lngC + 2)
        Next lngC
        rsK.Fields(32) = Cells(lngR, 28) '貸付金
        rsK.Fields(35) = Cells(lngR, 30) '財形貯蓄
        rsK.Fields(43) = Cells(lngR, 44) '控除合計
        rsK.Fields(44) = Cells(lngR, 45) '差引支給額
        rsK.Update
        
        lngR = lngR + 1
        If Cells(lngR, 1) = "" And lngR > 2 Then Exit Do
    Loop
    
Exit_DB:

    If Not rsK Is Nothing Then
        If rsK.State = adStateOpen Then rsK.Close
        Set rsK = Nothing
    End If
    If Not cnW Is Nothing Then
        If cnW.State = adStateOpen Then cnW.Close
        Set cnW = Nothing
    End If

End Sub

Sub 給与データ移動()

'最初に給与データの中のインポートデータをクリアする
'シートDataにあるデータをインポートデータに入れる

Dim cnW As New ADODB.Connection
Dim rsK As New ADODB.Recordset
Dim lngR   As Long
Dim lngC   As Long
    
    
    
    cnW.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & dbW '給与データ
    cnW.Open
    
    'インポートデータクリア
    strSQL = "DELETE FROM インポートデータ"
    rsK.Open strSQL, cnW, adOpenStatic, adLockPessimistic
    
    'インポートデータオープン
    rsK.Open "インポートデータ", cnW, adOpenStatic, adLockPessimistic
    
    Sheets("Data").Select
    lngR = 3
    Do
        rsK.AddNew
        strBMN = Strings.Format(Cells(lngR, 1), "000")   '部門ｺｰﾄﾞ
        strBKA = Strings.Format(Cells(lngR, 2), "000")   '部課ｺｰﾄﾞ
        strSCD = Strings.Format(Cells(lngR, 3), "00000") '社員ｺｰﾄﾞ
        
        '部門区分の設定(rsK.Fields(0))
        Select Case strBMN
            Case "100"
                rsK.Fields(0) = "000"
            Case "200"
                If strBKA = "010" Then      '大阪支店
                    rsK.Fields(0) = "010"
                ElseIf strBKA = "020" Then  '福岡営業所
                    rsK.Fields(0) = "020"
                ElseIf strBKA = "030" Then  '名古屋営業所
                    rsK.Fields(0) = "030"
                End If
            Case "300"
                If strBKA = "010" Then      '東京支店
                    rsK.Fields(0) = "040"
                ElseIf strBKA = "030" Then  '南関東
                    rsK.Fields(0) = "050"
                ElseIf strBKA = "040" Then  '仙台
                    rsK.Fields(0) = "060"
                ElseIf strBKA = "050" Then  '北関東
                    rsK.Fields(0) = "070"
                End If
            Case "400"
                Select Case strSCD
                    Case "00089"
                        rsK.Fields(0) = "000" '鳥居博一
                    Case "00490"
                        rsK.Fields(0) = "000" '鳥居洋子
                    Case "00694"
                        rsK.Fields(0) = "000" '卯野京子
                    Case "00472"
                        rsK.Fields(0) = "000" '高澤徹
                    Case "00497"
                        rsK.Fields(0) = "010" '森田裕之
                    Case "00526"
                        rsK.Fields(0) = "040" '鳥居新一郎
                    Case Else
                        MsgBox "識別できない役員のコードがあります！" & vbCrLf & strSCD, vbCritical
                        rsK.Fields(0) = "400"
                End Select
            Case Else
                MsgBox "識別出来ない部門があるようです。" & vbCrLf & "確認して下さい。" & strBMN & " - " & strBKA
        End Select
        
        rsK.Fields(45) = strBKA '部課ｺｰﾄﾞ
        rsK.Fields(1) = strSCD  '社員ｺｰﾄﾞ
        '氏名～支給4　--> 社員名～家族手当
        For lngC = 2 To 7
            rsK.Fields(lngC) = Cells(lngR, lngC + 2)
        Next lngC
        '残業計
        rsK.Fields(8) = 0
        For lngC = 18 To 23
            rsK.Fields(8) = rsK.Fields(8) + Cells(lngR, lngC)
        Next lngC
        '遅刻早退
        rsK.Fields(9) = Cells(lngR, 24)
        '赴任手当
        rsK.Fields(10) = Cells(lngR, 12)
        '特別手当
        rsK.Fields(11) = Cells(lngR, 13)
        '欠勤
        rsK.Fields(12) = Cells(lngR, 25)
        '課税合計～差引支給額
        For lngC = 16 To 44
            rsK.Fields(lngC) = Cells(lngR, lngC + 10)
        Next lngC
        rsK.Update
        
        lngR = lngR + 1
        If Cells(lngR, 1) = "" And lngR > 2 Then Exit Do
    Loop

Exit_DB:

    If Not rsK Is Nothing Then
        If rsK.State = adStateOpen Then rsK.Close
        Set rsK = Nothing
    End If
    If Not cnW Is Nothing Then
        If cnW.State = adStateOpen Then cnW.Close
        Set cnW = Nothing
    End If

End Sub

Sub 蓄積処理()

'ｲﾝﾎﾟｰﾄﾃﾞｰﾀを蓄積ﾃﾞｰﾀへ移行する
'（ﾃﾞｰﾀなければ追加、あれば更新）

Dim cnW     As New ADODB.Connection
Dim cnA     As New ADODB.Connection
Dim rsI     As New ADODB.Recordset '蓄積データ
Dim rsK     As New ADODB.Recordset 'インポートデータ
Dim rsM     As New ADODB.Recordset '給与ﾏｽﾀ(KYUMTA)
Dim Cmd     As New ADODB.Command
Dim strNT   As String
Dim lngR    As Long
Dim lngKIN  As Long
    
    DateA = Sheets("Menu").Range("F15")
    If DateA = "0：00：00" Then
        strDate = InputBox("支給月を入力して下さい。", "支給月入力", Strings.Format(Date, "yyyy") & "/" & Strings.Format(Date, "mm"))
    Else
        strDate = Strings.Format(DateA, "yyyy") & "/" & Strings.Format(DateA, "mm")
    End If
    
    cnW.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & dbW '給与データ
    cnW.Open
    strNT = "Initial Catalog=KYUYO;"
    cnA.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD 'SQLServer
    cnA.Open
    Set Cmd.ActiveConnection = cnA
    
    '給与データのインポートデータ読込み
    rsK.Open "インポートデータ", cnW, adOpenStatic, adLockReadOnly
    If rsK.EOF = False Then rsK.MoveFirst
    Do Until rsK.EOF
        If rsK![差引支給額] <> 0 Then
           '給与ﾏｽﾀと給与が一致しているか確認する(支給額から残業遅刻早退欠勤などを差し引きする）
            lngKIN = rsK![課税支給額計] - rsK![残業手当] + rsK![遅刻早退] + rsK![欠勤]
            strBMN = rsK![部門区分]
            strSCD = rsK![社員コード]
            If rsK![部課コード] <> "" And strKBN = "K" Then
                strSQL = ""
                strSQL = strSQL & "SELECT SCODE,"
                strSQL = strSQL & "       PAY1 + PAY2 + OPT1 + OPT2 + OPT3 + OPT4 + OPT5 as PAY,"
                strSQL = strSQL & "       SKBN"
                strSQL = strSQL & "  FROM KYUMTA"
                strSQL = strSQL & "       WHERE SCODE = '" & strSCD & "'"
                Cmd.CommandText = strSQL
                Set rsM = Cmd.Execute
                If rsM.Fields("SKBN") & "" <> "P" Then 'ﾊﾟｰﾄ社員除外
                    If lngKIN <> rsM.Fields("PAY") Then
                        lngR = MsgBox("金額が違います！！" & vbCrLf & "要チェック - " & strSCD & " " & rsK![社員名] & vbCrLf & "続行しますか？", vbYesNo, "給与チェック")
                        If lngR = vbNo Then
                            Exit Sub
                        End If
                    End If
                End If
                rsM.Close
            End If
           '蓄積データを検索
            strSQL = ""
            strSQL = strSQL & "SELECT *"
            strSQL = strSQL & "  FROM 蓄積データ"
            strSQL = strSQL & "       WHERE 支給年月 = '" & strDate & "'"
            strSQL = strSQL & "       AND   給与区分 = '" & strKBN & "'"
            strSQL = strSQL & "       AND   部門区分 = '" & strBMN & "'"
            strSQL = strSQL & "       AND   社員コード = '" & strSCD & "'"
            rsI.Open strSQL, cnW, adOpenStatic, adLockPessimistic
            If rsI.EOF Then
                rsI.AddNew '無ければ追加
                rsI![支給年月] = strDate
                rsI![給与区分] = strKBN
                rsI![部門区分] = strBMN
                rsI![社員コード] = strSCD
            End If
           '以下更新
            rsI![社員名] = rsK![社員名]
            For lngR = 3 To 45 Step 1
                rsI.Fields(lngR + 2) = rsK.Fields(lngR)
            Next lngR
            rsI.Update
            rsI.Close
        End If
        rsK.MoveNext
    Loop

Exit_DB:
    
    If Not rsK Is Nothing Then
        If rsK.State = adStateOpen Then rsK.Close
        Set rsK = Nothing
    End If
    If Not rsI Is Nothing Then
        If rsI.State = adStateOpen Then rsI.Close
        Set rsI = Nothing
    End If
    If Not rsM Is Nothing Then
        If rsM.State = adStateOpen Then rsM.Close
        Set rsM = Nothing
    End If
    If Not cnW Is Nothing Then
        If cnW.State = adStateOpen Then cnW.Close
        Set cnW = Nothing
    End If
    If Not cnA Is Nothing Then
        If cnA.State = adStateOpen Then cnA.Close
        Set cnA = Nothing
    End If

End Sub
