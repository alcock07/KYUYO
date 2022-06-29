Attribute VB_Name = "M01_Read"
Option Explicit

Sub Select_STN()

'=======================================
'ƒƒjƒ…[‰æ–Ê‚Å–‹ÆŠ‚ğ‘I‘ğ‚µ‚½‚Ìˆ—
'=======================================
    
    Sheets("List").Select 'ListƒV[ƒg‚ÖˆÚ“®
    Range("B3").Select
    
    Call Get_Data 'ƒf[ƒ^“Ç‚İ‚İ
    
End Sub

Sub Get_Data()

'=======================================
'’À‹àƒf[ƒ^“Ç‚İ‚İ
'=======================================

    Dim strSTN As String

    strSTN = Sheets("Menu").Range("AI5") '‹’“_‹æ•ªæ“¾(RH,RO,RT,TA,KA)
     
    Call ’À‹à“Ç(strSTN)
    
End Sub


Sub ’À‹à“Ç(strKBN As String)

Dim cnA    As New ADODB.Connection
Dim rsA    As New ADODB.Recordset
Dim Cmd    As New ADODB.Command
Dim strSQL As String
Dim strNT As String
Dim lngR   As Long
    
    'Ğˆõ•ªˆ—
    Call CLR_CELL          'ÃŞ°À¼°Ä¸Ø±
    
    strNT = "Initial Catalog=KYUYO;"
    cnA.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD 'SQLServer
    cnA.Open
    Set Cmd.ActiveConnection = cnA
    strSQL = ""
    strSQL = strSQL & "SELECT *"
    strSQL = strSQL & "     FROM KYUMTA"
    strSQL = strSQL & "        WHERE KBN ='" & strKBN & "'"
    strSQL = strSQL & "        AND DATKB ='1'"
    strSQL = strSQL & "     ORDER BY CLASS DESC,"
    strSQL = strSQL & "              ISSUE DESC,"
    strSQL = strSQL & "              SKBN,"
    strSQL = strSQL & "              SCODE"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    If rsA.EOF = False Then rsA.MoveFirst
    '¼°Ä‚ÉÃŞ°À“\‚è•t‚¯
    lngR = 7
    Do Until rsA.EOF
        If Trim(rsA![MGR] & "") <> "–ğˆõ" Then 'ˆê”ÊĞˆõ
            Cells(lngR, 2) = rsA.Fields("KBN")
            Cells(lngR, 3) = rsA.Fields("SCODE")
            Cells(lngR, 4) = rsA.Fields("SNAME")
            Cells(lngR, 5) = Trim(rsA.Fields("SEX"))
            Cells(lngR, 7) = Format(rsA.Fields("DATE1"), "ggge”NmŒd“ú")
            Cells(lngR, 8) = Format(rsA.Fields("DATE2"), "ggge”NmŒd“ú")
            Cells(lngR, 9) = rsA.Fields("SKBN")
            Cells(lngR, 10) = rsA.Fields("CLASS")
            Cells(lngR, 12) = rsA.Fields("ISSUE")
            Cells(lngR, 13) = ŠÇ—E‹æ’Tõ(Trim(rsA.Fields("MGR")) & "")
            Cells(lngR, 15) = rsA.Fields("PAY1") '–{‹‹
            Cells(lngR, 16) = rsA.Fields("PAY2") '‰Á‹‹
            Cells(lngR, 17) = rsA.Fields("OPT1")
            Cells(lngR, 18) = rsA.Fields("OPT2")
            Cells(lngR, 19) = rsA.Fields("OPT3")
            Cells(lngR, 20) = rsA.Fields("OPT4") '‹ÆÑè“–
            Cells(lngR, 21) = rsA.Fields("OPT5")
            Cells(lngR, 22) = "=SUM(RC[-7]:RC[-1])"
            Cells(lngR, 23) = rsA.Fields("PRN")
            Cells(lngR, 24) = rsA.Fields("OFFICE")
            Cells(lngR, 27) = rsA.Fields("HOUR")
        End If
        rsA.MoveNext
        lngR = lngR + 1
        If lngR = 54 Then lngR = 66
        If lngR > 112 Then Exit Do
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
    '‚P–‡–ÚƒNƒŠƒA
    Range("B7:E53,G7:J53,L7:M53,O7:U53,W7:Y53,AA7:AA53").Select
    Selection.ClearContents
    '‚Q–‡–ÚƒNƒŠƒA
    Range("B66:E112,G66:J112,L66:M112,O66:U112,W66:Y112,AA66:AA112").Select
    Selection.ClearContents
    Range("A1").Select
End Sub

Function ŠÇ—E‹æ’Tõ(strK As String)

    Select Case strK
        Case "–ğˆõ"
            ŠÇ—E‹æ’Tõ = "YY"
        Case "x“X’·"
            ŠÇ—E‹æ’Tõ = "SS"
        Case "•”’·"
            ŠÇ—E‹æ’Tõ = "BB"
        Case "Ÿ’·"
            ŠÇ—E‹æ’Tõ = "JJ"
        Case "‰Û’·"
            ŠÇ—E‹æ’Tõ = "KK"
        Case "å”C"
            ŠÇ—E‹æ’Tõ = "KS"
        Case "‰Û’·‘ã—"
            ŠÇ—E‹æ’Tõ = "HD"
        Case "ŒW’·"
            ŠÇ—E‹æ’Tõ = "HK"
        Case "”Ç’·"
            ŠÇ—E‹æ’Tõ = "HH"
        Case Else
            ŠÇ—E‹æ’Tõ = ""
    End Select
    
End Function

