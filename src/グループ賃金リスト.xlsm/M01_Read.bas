Attribute VB_Name = "M01_Read"
Option Explicit

Sub Select_STN()
    
    Sheets("List").Select
    Range("B3").Select
    
    Call Get_Data
    
End Sub

Sub Get_Data()

Dim strSTN As String
Dim strSNM As String

    strSTN = Sheets("Menu").Range("AI5")
     
    Call ĐőÇ(strSTN)
    
End Sub


Sub ĐőÇ(strKBN As String)

Const SQL1 = "SELECT * FROM O[vĐő}X^[ WHERE (((ÆæȘ)='"
Const SQL2 = "')) ORDER BY  DESC, ĐőíȚ, ĐőR[h"
Const SQL2T = "')) ORDER BY  DESC,  DESC, ĐőR[h"

Dim cnA    As New ADODB.Connection
Dim rsA    As New ADODB.Recordset
Dim Cmd    As New ADODB.Command

Dim strSQL As String
Dim strUNM As String
Dim strDB  As String
Dim lngR   As Long
Dim lngC   As Long
Dim P_Hant As String
    
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
    Set Cmd.ActiveConnection = cnA
    
    'ĐőȘ
    Call CLR_CELL          'ĂȚ°ÀŒ°ÄžŰ±
        
    strSQL = SQL1 & strKBN & SQL2
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    If rsA.EOF = False Then rsA.MoveFirst
    lngR = 7
    Do Until rsA.EOF
        If Trim(rsA![ÇEæ] & "") <> "đő" Then 'êÊĐő
            Cells(lngR, 2) = rsA.Fields("ÆæȘ")
            Cells(lngR, 3) = rsA.Fields("ĐőR[h")
            Cells(lngR, 4) = rsA.Fields("ĐőŒ")
            If rsA.Fields("«Ê") = "j" Then
                Cells(lngR, 5) = "M"
            Else
                Cells(lngR, 5) = "W"
            End If
            Cells(lngR, 7) = rsA.Fields("¶Nú")
            Cells(lngR, 10) = rsA.Fields("üĐNú")
            Cells(lngR, 11) = rsA.Fields("ĐőíȚ")
            Cells(lngR, 12) = rsA.Fields("")
            Cells(lngR, 14) = rsA.Fields("")
            Cells(lngR, 15) = ÇEæTő(rsA.Fields("ÇEæ") & "")
            Cells(lngR, 17) = rsA.Fields("î{P") '{
            Cells(lngR, 18) = rsA.Fields("î{Q") 'Á
            Cells(lngR, 19) = rsA.Fields("ÇEè")
            Cells(lngR, 20) = rsA.Fields("Æ°è")
            Cells(lngR, 21) = rsA.Fields("ćssÎ±è")
            Cells(lngR, 22) = rsA.Fields("Čźè") 'ÆŃè
            Cells(lngR, 23) = rsA.Fields("ÁêìÆè")
            Cells(lngR, 24) = "=SUM(RC[-7]:RC[-1])"
            Cells(lngR, 25) = rsA.Fields("óü")
            Cells(lngR, 26) = rsA.Fields("źÆ")
            Cells(lngR, 29) = rsA.Fields("p[gèÔ")
        End If
        rsA.MoveNext
        lngR = lngR + 1
        If lngR = 55 Then lngR = 67
        If lngR > 113 Then Exit Do
    Loop
       
    Range("A2").Select

    'Ú±ÌN[Y
    rsA.Close
    cnA.Close

Exit_DB:

    'IuWFNgÌjü
    Set rsA = Nothing
    Set cnA = Nothing
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
    Range("B7:E53,G7:H53,J7:L53,N7:O53,Q7:W53,Y7:AA53,AC7:AC53").Select
    Selection.ClearContents
    Range("AG7:AR44").Select
    Selection.ClearContents
    'QÚNA
    Range("B67:E113,G67:H113,J67:L113,N67:O113,Q67:W113,Y67:AA113,AC67:AC113").Select
    Selection.ClearContents
    Range("A1").Select
End Sub

Function ÇEæTő(strK As String)

    Select Case strK
        Case "đő"
            ÇEæTő = "YY"
        Case "xX·"
            ÇEæTő = "SS"
        Case "·"
            ÇEæTő = "BB"
        Case "·"
            ÇEæTő = "JJ"
        Case "Û·"
            ÇEæTő = "KK"
        Case "ćC"
            ÇEæTő = "KS"
        Case "Û·ă"
            ÇEæTő = "HD"
        Case "W·"
            ÇEæTő = "HK"
        Case "Ç·"
            ÇEæTő = "HH"
        Case Else
            ÇEæTő = ""
    End Select
    
End Function

