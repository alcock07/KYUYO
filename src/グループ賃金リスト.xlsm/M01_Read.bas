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
Dim strUNM As String
Dim strDB  As String
Dim lngR   As Long
Dim lngC   As Long
Dim P_Hant As String
    
    'ƒ†[ƒU–¼‚ª‹‹—^ŠÇ—Ò‚Ìê‡‚Ì‚İˆ—‚·‚é
    strUNM = Strings.UCase(GetUserNameString)
    If strUNM = "SCOTT" Or strUNM = "TAKA" Or strUNM = "SIMO" Then
        'Hê•ª‚Í•ÊDB¾¯Ä
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
    
    'Ğˆõ•ªˆ—
    Call CLR_CELL          'ÃŞ°À¼°Ä¸Ø±
        
    strSQL = strSQL & "SELECT *"
    strSQL = strSQL & "     FROM ƒOƒ‹[ƒvĞˆõƒ}ƒXƒ^["
    strSQL = strSQL & "        WHERE –‹ÆŠ‹æ•ª ='" & strKBN & "'"
    strSQL = strSQL & "     ORDER BY “™‹‰ DESC,"
    strSQL = strSQL & "              †” DESC,"
    strSQL = strSQL & "              Ğˆõí—Ş,"
    strSQL = strSQL & "              ĞˆõƒR[ƒh"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    If rsA.EOF = False Then rsA.MoveFirst
    '¼°Ä‚ÉÃŞ°À“\‚è•t‚¯
    lngR = 7
    Do Until rsA.EOF
        If Trim(rsA![ŠÇ—E‹æ] & "") <> "–ğˆõ" Then 'ˆê”ÊĞˆõ
            Cells(lngR, 2) = rsA.Fields("–‹ÆŠ‹æ•ª")
            Cells(lngR, 3) = rsA.Fields("ĞˆõƒR[ƒh")
            Cells(lngR, 4) = rsA.Fields("Ğˆõ–¼")
            If rsA.Fields("«•Ê") = "’j" Then
                Cells(lngR, 5) = "M"
            Else
                Cells(lngR, 5) = "W"
            End If
            Cells(lngR, 7) = rsA.Fields("¶”NŒ“ú")
            Cells(lngR, 10) = rsA.Fields("“üĞ”NŒ“ú")
            Cells(lngR, 11) = rsA.Fields("Ğˆõí—Ş")
            Cells(lngR, 12) = rsA.Fields("“™‹‰")
            Cells(lngR, 14) = rsA.Fields("†”")
            Cells(lngR, 15) = ŠÇ—E‹æ’Tõ(rsA.Fields("ŠÇ—E‹æ") & "")
            Cells(lngR, 17) = rsA.Fields("Šî–{‹‹‚P") '–{‹‹
            Cells(lngR, 18) = rsA.Fields("Šî–{‹‹‚Q") '‰Á‹‹
            Cells(lngR, 19) = rsA.Fields("ŠÇ—Eè“–")
            Cells(lngR, 20) = rsA.Fields("‰Æ‘°è“–")
            Cells(lngR, 21) = rsA.Fields("‘å“ss‹Î–±è“–")
            Cells(lngR, 22) = rsA.Fields("’²®è“–") '‹ÆÑè“–
            Cells(lngR, 23) = rsA.Fields("“Áêì‹Æè“–")
            Cells(lngR, 24) = "=SUM(RC[-7]:RC[-1])"
            Cells(lngR, 25) = rsA.Fields("ˆóü‡˜")
            Cells(lngR, 26) = rsA.Fields("Š‘®–‹ÆŠ")
            Cells(lngR, 29) = rsA.Fields("ƒp[ƒgŠ’èŠÔ”")
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
    '‚P–‡–ÚƒNƒŠƒA
    Range("B7:E53,G7:H53,J7:L53,N7:O53,Q7:W53,Y7:AA53,AC7:AC53").Select
    Selection.ClearContents
    Range("AG7:AR44").Select
    Selection.ClearContents
    '‚Q–‡–ÚƒNƒŠƒA
    Range("B67:E113,G67:H113,J67:L113,N67:O113,Q67:W113,Y67:AA113,AC67:AC113").Select
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

