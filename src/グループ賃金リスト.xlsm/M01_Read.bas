Attribute VB_Name = "M01_Read"
Option Explicit

Sub Select_STN()

'=======================================
'j[æÊÅÆðIðµ½Ì
'=======================================
    
    Sheets("List").Select 'ListV[gÖÚ®
    Range("B3").Select
    
    Call Get_Data 'f[^ÇÝÝ
    
End Sub

Sub Get_Data()

'=======================================
'Ààf[^ÇÝÝ
'=======================================

    Dim strSTN As String

    strSTN = Sheets("Menu").Range("AI5") '_æªæ¾(RH,RO,RT,TA,KA)
     
    Call ÀàÇ(strSTN)
    
End Sub


Sub ÀàÇ(strKBN As String)

Dim cnA    As New ADODB.Connection
Dim rsA    As New ADODB.Recordset
Dim Cmd    As New ADODB.Command
Dim strSQL As String
Dim strNT As String
Dim lngR   As Long
    
    'Ðõª
    Call CLR_CELL          'ÃÞ°À¼°Ä¸Ø±
    
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
    '¼°ÄÉÃÞ°À\èt¯
    lngR = 7
    Do Until rsA.EOF
        If Trim(rsA![MGR] & "") <> "ðõ" Then 'êÊÐõ
            Cells(lngR, 2) = rsA.Fields("KBN")
            Cells(lngR, 3) = rsA.Fields("SCODE")
            Cells(lngR, 4) = rsA.Fields("SNAME")
            Cells(lngR, 5) = Trim(rsA.Fields("SEX"))
            Cells(lngR, 7) = Format(rsA.Fields("DATE1"), "gggeNmdú")
            Cells(lngR, 8) = Format(rsA.Fields("DATE2"), "gggeNmdú")
            Cells(lngR, 9) = rsA.Fields("SKBN")
            Cells(lngR, 10) = rsA.Fields("CLASS")
            Cells(lngR, 12) = rsA.Fields("ISSUE")
            Cells(lngR, 13) = ÇEæTõ(Trim(rsA.Fields("MGR")) & "")
            Cells(lngR, 15) = rsA.Fields("PAY1") '{
            Cells(lngR, 16) = rsA.Fields("PAY2") 'Á
            Cells(lngR, 17) = rsA.Fields("OPT1")
            Cells(lngR, 18) = rsA.Fields("OPT2")
            Cells(lngR, 19) = rsA.Fields("OPT3")
            Cells(lngR, 20) = rsA.Fields("OPT4") 'ÆÑè
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
    'PÚNA
    Range("B7:E53,G7:J53,L7:M53,O7:U53,W7:Y53,AA7:AA53").Select
    Selection.ClearContents
    'QÚNA
    Range("B66:E112,G66:J112,L66:M112,O66:U112,W66:Y112,AA66:AA112").Select
    Selection.ClearContents
    Range("A1").Select
End Sub

Function ÇEæTõ(strK As String)

    Select Case strK
        Case "ðõ"
            ÇEæTõ = "YY"
        Case "xX·"
            ÇEæTõ = "SS"
        Case "·"
            ÇEæTõ = "BB"
        Case "·"
            ÇEæTõ = "JJ"
        Case "Û·"
            ÇEæTõ = "KK"
        Case "åC"
            ÇEæTõ = "KS"
        Case "Û·ã"
            ÇEæTõ = "HD"
        Case "W·"
            ÇEæTõ = "HK"
        Case "Ç·"
            ÇEæTõ = "HH"
        Case Else
            ÇEæTõ = ""
    End Select
    
End Function

