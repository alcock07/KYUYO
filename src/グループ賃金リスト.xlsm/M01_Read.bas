Attribute VB_Name = "M01_Read"
Option Explicit

Sub Select_STN()

'=======================================
'メニュー画面で事業所を選択した時の処理
'=======================================
    
    Sheets("List").Select 'Listシートへ移動
    Range("B3").Select
    
    Call Get_Data 'データ読み込み
    
End Sub

Sub Get_Data()

'=======================================
'賃金データ読み込み
'=======================================

    Dim strSTN As String

    strSTN = Sheets("Menu").Range("AI5") '拠点区分取得(RH,RO,RT,TA,KA)
     
    Call 賃金読込(strSTN)
    
End Sub


Sub 賃金読込(strKBN As String)

Dim cnA    As New ADODB.Connection
Dim rsA    As New ADODB.Recordset
Dim Cmd    As New ADODB.Command
Dim strSQL As String
Dim strUNM As String
Dim strDB  As String
Dim lngR   As Long
Dim lngC   As Long
Dim P_Hant As String
    
    'ユーザ名が給与管理者の場合のみ処理する
    strUNM = Strings.UCase(GetUserNameString)
    If strUNM = "SCOTT" Or strUNM = "TAKA" Or strUNM = "SIMO" Then
        '工場分は別DBｾｯﾄ
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
    Set Cmd.ActiveConnection = cnA
    
    '社員分処理
    Call CLR_CELL          'ﾃﾞｰﾀｼｰﾄｸﾘｱ
        
    strSQL = strSQL & "SELECT *"
    strSQL = strSQL & "     FROM グループ社員マスター"
    strSQL = strSQL & "        WHERE 事業所区分 ='" & strKBN & "'"
    strSQL = strSQL & "     ORDER BY 等級 DESC,"
    strSQL = strSQL & "              号数 DESC,"
    strSQL = strSQL & "              社員種類,"
    strSQL = strSQL & "              社員コード"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    If rsA.EOF = False Then rsA.MoveFirst
    'ｼｰﾄにﾃﾞｰﾀ貼り付け
    lngR = 7
    Do Until rsA.EOF
        If Trim(rsA![管理職区] & "") <> "役員" Then '一般社員
            Cells(lngR, 2) = rsA.Fields("事業所区分")
            Cells(lngR, 3) = rsA.Fields("社員コード")
            Cells(lngR, 4) = rsA.Fields("社員名")
            If rsA.Fields("性別") = "男" Then
                Cells(lngR, 5) = "M"
            Else
                Cells(lngR, 5) = "W"
            End If
            Cells(lngR, 7) = rsA.Fields("生年月日")
            Cells(lngR, 10) = rsA.Fields("入社年月日")
            Cells(lngR, 11) = rsA.Fields("社員種類")
            Cells(lngR, 12) = rsA.Fields("等級")
            Cells(lngR, 14) = rsA.Fields("号数")
            Cells(lngR, 15) = 管理職区探索(rsA.Fields("管理職区") & "")
            Cells(lngR, 17) = rsA.Fields("基本給１") '本給
            Cells(lngR, 18) = rsA.Fields("基本給２") '加給
            Cells(lngR, 19) = rsA.Fields("管理職手当")
            Cells(lngR, 20) = rsA.Fields("家族手当")
            Cells(lngR, 21) = rsA.Fields("大都市勤務手当")
            Cells(lngR, 22) = rsA.Fields("調整手当") '業績手当
            Cells(lngR, 23) = rsA.Fields("特殊作業手当")
            Cells(lngR, 24) = "=SUM(RC[-7]:RC[-1])"
            Cells(lngR, 25) = rsA.Fields("印刷順序")
            Cells(lngR, 26) = rsA.Fields("所属事業所")
            Cells(lngR, 29) = rsA.Fields("パート所定時間数")
        End If
        rsA.MoveNext
        lngR = lngR + 1
        If lngR = 55 Then lngR = 67
        If lngR > 113 Then Exit Do
    Loop
       
    Range("A2").Select

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

Sub CLR_CELL()
    '１枚目クリア
    Range("B7:E53,G7:H53,J7:L53,N7:O53,Q7:W53,Y7:AA53,AC7:AC53").Select
    Selection.ClearContents
    Range("AG7:AR44").Select
    Selection.ClearContents
    '２枚目クリア
    Range("B67:E113,G67:H113,J67:L113,N67:O113,Q67:W113,Y67:AA113,AC67:AC113").Select
    Selection.ClearContents
    Range("A1").Select
End Sub

Function 管理職区探索(strK As String)

    Select Case strK
        Case "役員"
            管理職区探索 = "YY"
        Case "支店長"
            管理職区探索 = "SS"
        Case "部長"
            管理職区探索 = "BB"
        Case "次長"
            管理職区探索 = "JJ"
        Case "課長"
            管理職区探索 = "KK"
        Case "主任"
            管理職区探索 = "KS"
        Case "課長代理"
            管理職区探索 = "HD"
        Case "係長"
            管理職区探索 = "HK"
        Case "班長"
            管理職区探索 = "HH"
        Case Else
            管理職区探索 = ""
    End Select
    
End Function

