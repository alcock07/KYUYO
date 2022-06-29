Attribute VB_Name = "M06_Function"
Option Explicit

Function ŠÇ—E‹æ”»’è(ShainKU As String, KanriKU As String, Kihon As Long, P_gkn As Single)
Attribute ŠÇ—E‹æ”»’è.VB_ProcData.VB_Invoke_Func = " \n14"

Dim Fjb As Long

    If ShainKU = "P" Then
        'Êß°ÄĞˆõ
        If P_gkn = 0 Or Kihon = 0 Then
            ŠÇ—E‹æ”»’è = ""
        Else
            Fjb = Round(Kihon / P_gkn, 0) '‹‹
            ŠÇ—E‹æ”»’è = "@\" & Format(Fjb, "#,##0") & " "
        End If
    Else
        ŠÇ—E‹æ”»’è = ""
        If UCase(KanriKU) = "YY" Then ŠÇ—E‹æ”»’è = "–ğˆõ"
        If UCase(KanriKU) = "SS" Then ŠÇ—E‹æ”»’è = "x“X’·"
        If UCase(KanriKU) = "BB" Then ŠÇ—E‹æ”»’è = "•”’·"
        If UCase(KanriKU) = "JJ" Then ŠÇ—E‹æ”»’è = "Ÿ’·"
        If UCase(KanriKU) = "KK" Then ŠÇ—E‹æ”»’è = "‰Û’·"
        If UCase(KanriKU) = "KS" Then ŠÇ—E‹æ”»’è = "å”C"
        If UCase(KanriKU) = "HD" Then ŠÇ—E‹æ”»’è = "‰Û’·‘ã—"
        If UCase(KanriKU) = "HK" Then ŠÇ—E‹æ”»’è = "ŒW’·"
        If UCase(KanriKU) = "HH" Then ŠÇ—E‹æ”»’è = "”Ç’·"
    End If

End Function

