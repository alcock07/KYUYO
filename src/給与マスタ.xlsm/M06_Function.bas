Attribute VB_Name = "M06_Function"
Option Explicit

Function ΗEζ»θ(ShainKU As String, KanriKU As String, Kihon As Long, P_gkn As Single)
Attribute ΗEζ»θ.VB_ProcData.VB_Invoke_Func = " \n14"

Dim Fjb As Long

    If ShainKU = "P" Then
        'Κί°ΔΠυ
        If P_gkn = 0 Or Kihon = 0 Then
            ΗEζ»θ = ""
        Else
            Fjb = Round(Kihon / P_gkn, 0) '
            ΗEζ»θ = "@\" & Format(Fjb, "#,##0") & " "
        End If
    Else
        ΗEζ»θ = ""
        If UCase(KanriKU) = "YY" Then ΗEζ»θ = "πυ"
        If UCase(KanriKU) = "SS" Then ΗEζ»θ = "xX·"
        If UCase(KanriKU) = "BB" Then ΗEζ»θ = "·"
        If UCase(KanriKU) = "JJ" Then ΗEζ»θ = "·"
        If UCase(KanriKU) = "KK" Then ΗEζ»θ = "Ϋ·"
        If UCase(KanriKU) = "KS" Then ΗEζ»θ = "εC"
        If UCase(KanriKU) = "HD" Then ΗEζ»θ = "Ϋ·γ"
        If UCase(KanriKU) = "HK" Then ΗEζ»θ = "W·"
        If UCase(KanriKU) = "HH" Then ΗEζ»θ = "Η·"
    End If

End Function

