Attribute VB_Name = "M04_BS"
Option Explicit

Const Fname = "Z:\会計システム\仕訳\賞与仕訳.txt"
Const Kname = "Z:\会計システム\仕訳\給与仕訳.txt"

Private lngKIN(17, 6) As Long '金額格納配列(事業所,項目)
Private lngKYU       As Long '報酬給料振替額
Private strKMC       As String  '科目ｺｰﾄﾞ
Private strKMN       As String  '科目名
Private strHJC       As String  '補助ｺｰﾄﾞ
Private strHJN       As String  '補助名
Private strKMC2      As String  '科目ｺｰﾄﾞ
Private strKMN2      As String  '科目名
Private strHJC2      As String  '補助ｺｰﾄﾞ
Private strHJN2      As String  '補助名
Private strBMC       As String  '部門ｺｰﾄﾞ
Private strBMC2      As String  '部門ｺｰﾄﾞ
Private strTXC       As String  '税区分
Private lngR         As Long    '仕訳行ｶｳﾝﾀ
Private lngCR        As Long    '行ｶｳﾝﾀ
Private lngCC        As Long    '拠点ｶｳﾝﾀ
Private lngNO        As Long    '伝票№
Private lngGNO       As Long    '行№
Private boolR        As Boolean

'=== 仕訳印刷&更新 ﾎﾞﾀﾝ===
Sub SEL_KS()
    boolR = False
    If Sheets("入力").Range("C4") = "給料" Then
        Call 仕訳印刷("K")
    Else
        If Sheets("入力").Range("C4") = "臨時賞与" Then
            boolR = True
            Call 仕訳印刷("S")
        Else
            Call 仕訳印刷("S")
        End If
    End If
    Range("B4").Select
End Sub

Sub 仕訳印刷(strKBN As String)

Dim lngCock As Long

With Sheets("入力")
.Select
DateA = Range("B4")
Erase lngKIN

'金額を配列にｾｯﾄ =====
'人数・振込額・交通費・貸付金・雇用保険料会社負担分
For lngCC = 0 To 16
    lngKIN(lngCC, 0) = lngKIN(lngCC, 0) + Cells(11, lngCC + 3) '人数
    lngKIN(lngCC, 1) = lngKIN(lngCC, 1) + Cells(12, lngCC + 3) '振込額
    lngKIN(lngCC, 2) = lngKIN(lngCC, 2) + Cells(14, lngCC + 3) '交通費
    lngKIN(lngCC, 4) = lngKIN(lngCC, 4) + Cells(22, lngCC + 3) '貸付金
    lngKIN(lngCC, 5) = lngKIN(lngCC, 5) + Cells(33, lngCC + 3) '雇用保険料会社負担分
    lngKIN(lngCC, 6) = lngKIN(lngCC, 6) + Cells(24, lngCC + 3) 'クック会
Next lngCC
'控除計
For lngCC = 0 To 16
    For lngCR = 0 To 15
        lngKIN(lngCC, 3) = lngKIN(lngCC, 3) + Cells(lngCR + 15, lngCC + 3)
    Next lngCR
Next lngCC

'貸付金の金額をﾁｪｯｸ
If strKBN = "K" Then
    If (lngKIN(0, 4) + lngKIN(1, 4) + lngKIN(2, 4)) <> Range("AB19") Then
        MsgBox "本部の貸付金が一致しません。貸付金明細を保守してからやり直して下さい。", vbCritical
        Exit Sub
    ElseIf (lngKIN(3, 4) + lngKIN(4, 4) + lngKIN(5, 4) + lngKIN(6, 4) + lngKIN(7, 4) + lngKIN(8, 4)) <> Range("AE19") Then
        MsgBox "大阪の貸付金が一致しません。貸付金明細を保守してからやり直して下さい。", vbCritical
        Exit Sub
    ElseIf (lngKIN(9, 4) + lngKIN(10, 4) + lngKIN(11, 4) + lngKIN(12, 4) + lngKIN(13, 4) + lngKIN(14, 4)) <> Range("AH19") Then
        MsgBox "東京の貸付金が一致しません。貸付金明細を保守してからやり直して下さい。", vbCritical
        Exit Sub
    End If
    If Dir(Kname) <> "" Then Kill Kname
Else
    If (lngKIN(0, 4) + lngKIN(1, 4) + lngKIN(2, 4)) <> Range("AL19") Then
        MsgBox "本部の貸付金が一致しません。貸付金明細を保守して下さい。", vbCritical
        Exit Sub
    ElseIf (lngKIN(3, 4) + lngKIN(4, 4) + lngKIN(5, 4) + lngKIN(6, 4) + lngKIN(7, 4) + lngKIN(8, 4)) <> Range("AO19") Then
        MsgBox "大阪の貸付金が一致しません。貸付金明細を保守して下さい。", vbCritical
        Exit Sub
    ElseIf (lngKIN(9, 4) + lngKIN(10, 4) + lngKIN(11, 4) + lngKIN(12, 4) + lngKIN(13, 4) + lngKIN(14, 4)) <> Range("AR19") Then
        MsgBox "東京の貸付金が一致しません。貸付金明細を保守して下さい。", vbCritical
        Exit Sub
    End If
    If Dir(Fname) <> "" Then Kill Fname
End If

Sheets("仕訳").Select
Call CLS_仕訳

If strKBN = "K" Then
    Range("B1") = "給与仕訳"
    strKMC = .Range("U11")
    strKMN = .Range("W11")
    strHJC = ""
    strHJN = ""
    strBMC = .Range("C6")
    strTXC = "00"
Else
    Range("B1") = "賞与仕訳"
    If boolR Then
        strKMC = "713"
        strKMN = "賞与"
        strHJC = ""
        strHJN = ""
        strBMC = .Range("C6")
    Else
        strKMC = "323"
        strKMN = "未払賞与"
        strHJC = "601"
        strHJN = "本部"
        strBMC = ""
    End If
    strTXC = "00"
End If
lngR = 5  '開始行
lngNO = 1 '伝票№
lngGNO = 1

'=== 一旦総額を本部で計上する ===
'振込総額（振込額+交通費）
Cells(lngR, 1) = lngNO
Cells(lngR, 2) = strKMC
Cells(lngR, 3) = strKMN
If strHJC <> "" Then
    Cells(lngR + 1, 2) = strHJC
    Cells(lngR + 1, 3) = strHJN
End If
Cells(lngR, 4) = strTXC
Cells(lngR, 5) = Strings.Format(DateA, "ggge年m月分") & .Range("C4")
Cells(lngR + 1, 5) = "振込総額  " & .Range("T11") & "名分"
Cells(lngR, 6) = .Range("T12") + .Range("T14")
Cells(lngR, 7) = .Range("U12") '科目ｺｰﾄﾞ
Cells(lngR, 8) = .Range("W12") '科目名
Cells(lngR, 10) = "00" '税区分
Cells(lngR + 1, 7) = .Range("V12") '補助ｺｰﾄﾞ
Cells(lngR + 1, 8) = .Range("X12") '補助名
Cells(lngR + 1, 4) = strBMC
Cells(lngR + 1, 10) = ""
lngNO = lngNO + 1
lngR = lngR + 2

'社会保険料
For lngCR = 15 To 30
    If .Cells(lngCR, 20) <> 0 And lngCR <> 22 Then
        Cells(lngR, 1) = lngNO
        Cells(lngR, 5) = Strings.Format(DateA, "ggge年m月分") & .Range("C4")
        Cells(lngR + 1, 5) = .Cells(lngCR, 2) & " 預り"
        lngNO = lngNO + 1
        Cells(lngR, 6) = .Cells(lngCR, 20)
        Cells(lngR, 2) = strKMC
        Cells(lngR, 3) = strKMN
        Cells(lngR, 4) = strTXC
        Cells(lngR, 7) = .Cells(lngCR, 21)
        Cells(lngR, 8) = .Cells(lngCR, 23)
        Cells(lngR, 10) = "00"
        If strHJC <> "" Then
            Cells(lngR + 1, 2) = strHJC
            Cells(lngR + 1, 3) = strHJN
        End If
        Cells(lngR + 1, 4) = strBMC
        Cells(lngR + 1, 7) = .Cells(lngCR, 22)
        Cells(lngR + 1, 8) = .Cells(lngCR, 24)
        Cells(lngR + 1, 10) = ""
        lngR = lngR + 2
    End If
Next lngCR

Call Data_Export(DateA)
Call シート印刷
Call CLS_仕訳
lngR = 5
lngNO = 1

'=== ここから振替 ==========
'本部内-部門振替
If strKBN = "S" Then
    '臨時賞与の場合は直接賞与勘定で上げる
    If boolR Then
        strKMC = "713"
        strKMN = "賞与"
    Else
        strKMC = "713"
        strKMN = "賞与月割額"
    End If
End If

For lngCC = 0 To 2
     If strKBN = "K" Then
        If lngCC > 0 Then Call 給料振替("101")
        Call 交通費振替
    Else
        If lngCC > 0 Then Call 給料振替("101")
    End If
    Call 雇用保険振替
Next lngCC

'貸付金個別計上
If strKBN = "K" And .Range("AB19") <> 0 Then  '本部
    For lngCR = 13 To 18
        If .Cells(lngCR, 19) = "" Then Exit For
        Cells(lngR, 1) = lngNO
        Cells(lngR, 2) = strKMC
        Cells(lngR, 3) = strKMN
        Cells(lngR, 4) = "00"
        Cells(lngR, 5) = Strings.Format(DateA, "ggge年m月分") & .Range("C4")
        Cells(lngR + 1, 5) = .Cells(lngCR, 27) & "　貸付金計上"
        Cells(lngR, 6) = .Cells(lngCR, 28)
        Cells(lngR, 7) = .Range("U22")
        Cells(lngR, 8) = .Range("W22")
        Cells(lngR, 10) = "00"
        Cells(lngR + 1, 7) = .Cells(lngCR, 26) '貸付金補助ｺｰﾄﾞ
        Cells(lngR + 1, 8) = .Cells(lngCR, 27) '貸付金補助名
        Cells(lngR + 1, 4) = strBMC
        Cells(lngR + 1, 10) = ""
        lngNO = lngNO + 1
        lngR = lngR + 2
    Next lngCR
ElseIf strKBN = "S" And .Range("AL19") <> 0 Then
    For lngCR = 13 To 18
        If .Cells(lngCR, 29) = "" Then Exit For
        Cells(lngR, 1) = lngNO
        Cells(lngR, 2) = "323"
        Cells(lngR, 3) = "未払賞与"
        Cells(lngR, 4) = "00"
        Cells(lngR + 1, 2) = "601"
        Cells(lngR + 1, 3) = "本部"
        Cells(lngR, 5) = Strings.Format(DateA, "ggge年m月分") & .Range("C4")
        Cells(lngR + 1, 5) = .Cells(lngCR, 37) & "　貸付金計上"
        Cells(lngR, 6) = .Cells(lngCR, 38)
        Cells(lngR, 7) = .Range("U22")
        Cells(lngR, 8) = .Range("W22")
        Cells(lngR, 7) = .Cells(lngCR, 36)
        Cells(lngR, 8) = .Cells(lngCR, 37)
        Cells(lngR, 10) = "00"
        Cells(lngR + 1, 10) = ""
        lngNO = lngNO + 1
        lngR = lngR + 2
    Next lngCR
End If
Call Data_Export(DateA)
Call シート印刷
Call CLS_仕訳
lngR = 5
lngNO = 1

'大阪支店振替
Range("B1") = "大阪振替"
If strKBN = "S" Then Call 賞与振替("大阪")
For lngCC = 3 To 8
    strBMC = Cells(6, lngCC + 3)
    If strKBN = "K" Then
        Call 給料振替("101")
        Call 交通費振替
    Else
        If lngCC > 3 Then Call 営業所振替("201")
    End If
    Call 雇用保険振替
Next lngCC
'貸付金個別計上
If strKBN = "K" And .Range("AE19") <> 0 Then
    For lngCR = 13 To 18
        If .Cells(lngCR, 29) = "" Then Exit For
        Cells(lngR, 1) = lngNO
        Cells(lngR, 2) = strKMC
        Cells(lngR, 3) = strKMN
        Cells(lngR, 4) = "00"
        Cells(lngR + 1, 4) = "101"
        Cells(lngR, 5) = Strings.Format(DateA, "ggge年m月分") & .Range("C4")
        Cells(lngR + 1, 5) = .Cells(lngCR, 30) & "　貸付金振替"
        Cells(lngR, 6) = .Cells(lngCR, 31)
        Cells(lngR, 7) = .Range("U22")
        Cells(lngR, 8) = .Range("W22")
        Cells(lngR + 1, 7) = .Cells(lngCR, 29)
        Cells(lngR + 1, 8) = .Cells(lngCR, 30)
        Cells(lngR, 10) = "00"
        Cells(lngR + 1, 10) = ""
        lngNO = lngNO + 1
        lngR = lngR + 2
        Next lngCR
ElseIf strKBN = "S" And .Range("AR19") <> 0 Then
    For lngCR = 13 To 18
        If .Cells(lngCR, 39) = "" Then Exit For
        Cells(lngR, 1) = lngNO
        If boolR Then
            Cells(lngR, 2) = "713"
            Cells(lngR, 3) = "賞与"
            Cells(lngR + 1, 4) = "201"
        Else
            Cells(lngR, 2) = "323"
            Cells(lngR, 3) = "未払賞与"
            Cells(lngR + 1, 2) = "611"
            Cells(lngR + 1, 3) = "大阪"
        End If
        Cells(lngR, 5) = Strings.Format(DateA, "ggge年m月分") & .Range("C4")
        Cells(lngR + 1, 5) = .Cells(lngCR, 38) & "　貸付金計上"
        Cells(lngR, 6) = .Cells(lngCR, 41)
        Cells(lngR, 7) = .Range("U22")
        Cells(lngR, 8) = .Range("W22")
        Cells(lngR + 1, 7) = .Cells(lngCR, 39)
        Cells(lngR + 1, 8) = .Cells(lngCR, 40)
        Cells(lngR, 10) = "00"
        Cells(lngR + 1, 10) = ""
        lngNO = lngNO + 1
        lngR = lngR + 2
    Next lngCR
End If
Call Data_Export(DateA)
Call シート印刷
Call CLS_仕訳
lngR = 5
lngNO = 1

'東京支店振替
Range("B1") = "東京振替"
If strKBN = "S" Then Call 賞与振替("東京")
For lngCC = 9 To 17
    If strKBN = "K" Then
        Call 給料振替("101")
        Call 交通費振替
        lngCock = lngCock + lngKIN(lngCC, 6)
    Else
         If lngCC > 9 Then Call 営業所振替("301")
    End If
    Call 雇用保険振替
Next lngCC

'クック会振替
If strKBN = "K" And lngCock <> 0 Then
    Cells(lngR, 1) = lngNO
    Cells(lngR, 2) = "326"
    Cells(lngR, 3) = "預り金"
    Cells(lngR + 1, 2) = "707"
    Cells(lngR + 1, 3) = "クック会"
    Cells(lngR, 4) = "00"
    Cells(lngR, 5) = Strings.Format(DateA, "ggge年m月分") & .Range("C4")
    Cells(lngR + 1, 5) = "東京分クック会費振替"
    Cells(lngR, 6) = lngCock
    Cells(lngR, 7) = "326"
    Cells(lngR, 8) = "預り金"
    Cells(lngR, 10) = "00"
    Cells(lngR + 1, 7) = "717"
    Cells(lngR + 1, 8) = "クック会-東京"
    lngNO = lngNO + 1
    lngR = lngR + 2
End If

'貸付金個別計上
If strKBN = "K" And .Range("AH19") <> 0 Then
    For lngCC = 13 To 18
        If .Cells(lngCC, 32) = "" Then Exit For
        Cells(lngR, 1) = lngNO
        Cells(lngR, 2) = strKMC
        Cells(lngR, 3) = strKMN
        Cells(lngR, 4) = "00"
        Cells(lngR + 1, 4) = "101"
        Cells(lngR, 5) = Strings.Format(DateA, "ggge年m月分") & .Range("C4")
        Cells(lngR + 1, 5) = .Cells(lngCC, 33) & "　貸付金振替"
        Cells(lngR, 6) = .Cells(lngCC, 34)
        Cells(lngR, 7) = .Range("U22")
        Cells(lngR, 8) = .Range("W22")
        Cells(lngR + 1, 7) = .Cells(lngCC, 32)
        Cells(lngR + 1, 8) = .Cells(lngCC, 33)
        Cells(lngR, 10) = "00"
        Cells(lngR + 1, 10) = ""
        lngNO = lngNO + 1
        lngR = lngR + 2
    Next lngCC
ElseIf strKBN = "S" And .Range("AR19") <> 0 Then
    For lngCR = 13 To 18
        If .Cells(lngCR, 44) = "" Then Exit For
        Cells(lngR, 1) = lngNO
        If boolR Then
            Cells(lngR, 2) = "713"
            Cells(lngR, 3) = "賞与"
            Cells(lngR + 1, 4) = "301"
        Else
            Cells(lngR, 2) = "323"
            Cells(lngR, 3) = "未払賞与"
            Cells(lngR + 1, 2) = "631"
            Cells(lngR + 1, 3) = "東京"
        End If
        Cells(lngR, 4) = "00"
        Cells(lngR, 5) = Strings.Format(DateA, "ggge年m月分") & .Range("C4")
        Cells(lngR + 1, 5) = .Cells(lngCR, 43) & "　貸付金計上"
        Cells(lngR, 6) = .Cells(lngCR, 44)
        Cells(lngR, 7) = .Range("U22")
        Cells(lngR, 8) = .Range("W22")
        Cells(lngR + 1, 7) = .Cells(lngCR, 42)
        Cells(lngR + 1, 8) = .Cells(lngCR, 43)
        Cells(lngR, 10) = "00"
        Cells(lngR + 1, 10) = ""
        lngNO = lngNO + 1
        lngR = lngR + 2
    Next lngCR
End If
Call Data_Export(DateA)
Call シート印刷

Sheets("振替").Select
Call シート印刷

.Select
End With

End Sub

Sub 給料振替(strRMC As String)

'報酬給料算出（振込額+控除計）
lngKYU = 0
lngKYU = lngKYU + lngKIN(lngCC, 1) + lngKIN(lngCC, 3)

If lngKYU <> 0 Then
    Cells(lngR, 1) = lngNO
    Cells(lngR, 2) = strKMC
    Cells(lngR, 3) = strKMN
    Cells(lngR, 4) = "00"
    Cells(lngR, 5) = Strings.Format(DateA, "ggge年m月分") & Sheets("入力").Range("C4")
    Cells(lngR + 1, 5) = Sheets("入力").Cells(5, lngCC + 3) & "分計上"
    Cells(lngR, 6) = lngKYU
    Cells(lngR, 7) = strKMC
    Cells(lngR, 8) = strKMN
    Cells(lngR, 10) = "00"
    Cells(lngR + 1, 4) = Sheets("入力").Cells(6, lngCC + 3)
    Cells(lngR + 1, 10) = strRMC
    lngNO = lngNO + 1
    lngR = lngR + 2
    If lngR > 43 Then
        Call Data_Export(DateA)
        Call シート印刷
        Call CLS_仕訳
        lngR = 5
    End If
End If
    
End Sub

Sub 営業所振替(strRMC As String)

'報酬給料算出（振込額+控除計）
lngKYU = 0
lngKYU = lngKYU + lngKIN(lngCC, 1) + lngKIN(lngCC, 3)

strKMC = "323"
strKMN = "未払賞与"
If strRMC = "201" Then
    strHJC = "611"
    strHJN = "大阪"
Else
    strHJC = "631"
    strHJN = "東京"
End If

If lngKYU <> 0 Then
    Cells(lngR, 1) = lngNO
    Cells(lngR, 2) = strKMC
    Cells(lngR, 3) = strKMN
    Cells(lngR, 4) = "00"
    Cells(lngR, 5) = Strings.Format(DateA, "ggge年m月分") & Sheets("入力").Range("C4")
    Cells(lngR + 1, 5) = Sheets("入力").Cells(5, lngCC + 3) & "分計上"
    Cells(lngR, 6) = lngKYU
    Cells(lngR, 7) = strKMC
    Cells(lngR, 8) = strKMN
    Cells(lngR, 10) = "00"
    Cells(lngR + 1, 2) = strHJC
    Cells(lngR + 1, 3) = strHJN
    Cells(lngR + 1, 7) = strHJC
    Cells(lngR + 1, 8) = strHJN
    lngNO = lngNO + 1
    lngR = lngR + 2
    If lngR > 43 Then
        Call Data_Export(DateA)
        Call シート印刷
        Call CLS_仕訳
        lngR = 5
    End If
End If
    
End Sub

Sub 賞与振替(strSTN As String)
'賞与算出（振込額+控除計）

Dim strKCD  As String
Dim strKNM  As String
Dim strKCD2 As String
Dim strKNM2 As String

lngKYU = 0
If strSTN = "大阪" Then
    For lngCC = 3 To 8
        lngKYU = lngKYU + lngKIN(lngCC, 1) + lngKIN(lngCC, 3)
    Next lngCC
    If boolR Then '臨時賞与処理
        strKCD = "713"
        strKNM = "賞与"
        strHJC = ""
        strHJN = ""
        strBMC = "201"
        strKCD2 = "713"
        strKNM2 = "賞与"
        strHJC2 = ""
        strHJN2 = ""
        strBMC2 = "101"
    Else
        strKCD = "323"
        strKNM = "未払賞与"
        strHJC = "611"
        strHJN = "大阪"
        strBMC = ""
        strKCD2 = "323"
        strKNM2 = "未払賞与"
        strHJC2 = "601"
        strHJN2 = "本部"
        strBMC2 = ""
    End If
Else
    For lngCC = 9 To 14
        lngKYU = lngKYU + lngKIN(lngCC, 1) + lngKIN(lngCC, 3)
    Next lngCC
    If boolR Then
        strKCD = "713"
        strKNM = "賞与"
         strHJC = ""
        strHJN = ""
        strBMC = "301"
        strKCD2 = "713"
        strKNM2 = "賞与"
        strHJC2 = ""
        strHJN2 = ""
        strBMC2 = "101"
    Else
        strKCD = "323"
        strKNM = "未払賞与"
        strHJC = "631"
        strHJN = "東京"
        strBMC = ""
        strKCD2 = "323"
        strKNM2 = "未払賞与"
        strHJC2 = "601"
        strHJN2 = "本部"
        strBMC2 = ""
    End If
End If
    
If lngKYU <> 0 Then
    Cells(lngR, 1) = lngNO
    Cells(lngR, 2) = strKCD
    Cells(lngR, 3) = strKNM
    Cells(lngR + 1, 2) = strHJC
    Cells(lngR + 1, 3) = strHJN
    Cells(lngR, 4) = "00"
    Cells(lngR + 1, 4) = strBMC
    Cells(lngR, 5) = Strings.Format(DateA, "ggge年m月分") & Sheets("入力").Range("C4")
    Cells(lngR + 1, 5) = strSTN & "分計上"
    Cells(lngR, 6) = lngKYU
    Cells(lngR, 7) = strKCD2
    Cells(lngR, 8) = strKNM2
    Cells(lngR + 1, 7) = strHJC2
    Cells(lngR + 1, 8) = strHJN2
    Cells(lngR, 10) = "00"
    Cells(lngR + 1, 10) = strBMC2
    lngNO = lngNO + 1
    lngR = lngR + 2
    If lngR > 43 Then
        Call Data_Export(DateA)
        Call シート印刷
        Call CLS_仕訳
        lngR = 5
    End If
End If
    
End Sub

Sub 交通費振替()

If lngKIN(lngCC, 2) <> 0 Then
    Cells(lngR, 1) = lngNO
    Cells(lngR, 2) = Sheets("入力").Range("U14")
    Cells(lngR, 3) = Sheets("入力").Range("W14")
    Cells(lngR, 4) = "Q5"
    Cells(lngR, 5) = Strings.Format(DateA, "ggge年m月分") & Sheets("入力").Range("C4")
    Cells(lngR + 1, 5) = "交通費振替 " & " " & Sheets("入力").Cells(5, lngCC + 3) & "分計上"
    Cells(lngR, 6) = lngKIN(lngCC, 2)
    Cells(lngR + 1, 6) = Round((lngKIN(lngCC, 2) / 110) * 10, 0)
    Cells(lngR, 7) = Sheets("入力").Range("U11")
    Cells(lngR, 8) = Sheets("入力").Range("W11")
    Cells(lngR, 10) = "00"
    Cells(lngR + 1, 4) = Sheets("入力").Cells(6, lngCC + 3)
    Cells(lngR + 1, 10) = "101"
    lngNO = lngNO + 1
    lngR = lngR + 2
    If lngR > 43 Then
        Call Data_Export(DateA)
        Call シート印刷
        Call CLS_仕訳
        lngR = 5
    End If
End If

End Sub

Sub 雇用保険振替()

If lngKIN(lngCC, 5) <> 0 Then
    Cells(lngR, 1) = lngNO
    Cells(lngR, 2) = Sheets("入力").Range("U33")
    Cells(lngR, 3) = Sheets("入力").Range("W33")
    Cells(lngR, 4) = "P0"
    Cells(lngR, 5) = Strings.Format(DateA, "ggge年m月分") & Sheets("入力").Range("C4")
    Cells(lngR + 1, 5) = Sheets("入力").Cells(33, 2) & " " & Sheets("入力").Cells(5, lngCC + 3) & "分計上"
    
    Cells(lngR, 6) = lngKIN(lngCC, 5)
    Cells(lngR, 7) = Sheets("入力").Range("U19")
    Cells(lngR, 8) = Sheets("入力").Range("W19")
    Cells(lngR + 1, 7) = Sheets("入力").Range("V19")
    Cells(lngR + 1, 8) = Sheets("入力").Range("X19")
    Cells(lngR, 10) = "00"
    Cells(lngR + 1, 8) = Sheets("入力").Range("X19")
    Cells(lngR + 1, 4) = Sheets("入力").Cells(6, lngCC + 3)
    Cells(lngR + 1, 10) = ""
    lngNO = lngNO + 1
    lngR = lngR + 2
    If lngR > 43 Then
        Call Data_Export(DateA)
        Call シート印刷
        Call CLS_仕訳
        lngR = 5
    End If
End If

End Sub

Sub CLS_仕訳()
    Range("A5:J46") = ""
End Sub

Public Sub Data_Export(DateA As Date)

Dim boolT    As Boolean  '戻り値
Dim strNO    As String   '伝票番号
Dim strTEXT  As String   'ﾃﾞｰﾀﾃｷｽﾄ
Dim lngC     As Long     'ｶｳﾝﾀ
Dim lngTaxL  As Long     '借方消費税
Dim lngTaxR  As Long     '貸方消費税
Dim strBMNL  As String   '借方部門
Dim strBMNR  As String   '貸方部門
Dim strKMKL  As String   '借方科目
Dim strKMKR  As String   '貸方科目
Dim strHCDL  As String   '借方取引先ｺｰﾄﾞ
Dim strHCDR  As String   '貸方取引先ｺｰﾄﾞ
Dim strTXBL  As String   '借方税区分
Dim strTXBR  As String   '貸方税区分
Dim strTKYO  As String   '摘要
Dim strKINL  As String   '借方金額
Dim strKINR  As String   '貸方金額
Dim strTaxL  As String   '借方金額
Dim strTaR   As String   '貸方金額

    '伝票番号&仕訳ﾌｧｲﾙ名作成
    strNO = "4" & Strings.Format(DateA, "mmdd")
    strDate = Format(DateA, "yyyymmdd")
    
    '汎用仕訳データｴｸｽﾎﾟｰﾄ処理
    For lngC = 5 To 45 Step 2
        If Cells(lngC, 3) = "" Or Cells(lngC, 6) = 0 Then
        Else
            strBMNL = Cells(lngC + 1, 4)
            strBMNR = Cells(lngC + 1, 10)
            strKMKL = Cells(lngC, 2)
            strKMKR = Cells(lngC, 7)
            strHCDL = Cells(lngC + 1, 2)
            strHCDR = Cells(lngC + 1, 7)
            strTXBL = Cells(lngC, 4)
            strTXBR = Cells(lngC, 10)
            strTKYO = Cells(lngC, 5) & " " & Cells(lngC + 1, 5)
            strKINL = Cells(lngC, 6)
            strKINR = Cells(lngC, 6)
            If Right(strTXBL, 1) = "0" Then
                lngTaxL = 0
            Else
                lngTaxL = Cells(lngC + 1, 6)
            End If
            If Right(strTXBR, 1) = "0" Then
                lngTaxR = 0
            Else
                lngTaxR = Cells(lngC + 1, 6)
            End If
                        
            strTEXT = strDate & "," & strNO & ",21,0,1," & _
            strBMNL & ",," & strKMKL & ",," & strHCDL & ",," & strTXBL & ",," & strKINL & "," & lngTaxL & ",1," & _
            strBMNR & ",," & strKMKR & ",," & strHCDR & ",," & strTXBR & ",," & strKINR & "," & lngTaxR & "," & _
            strTKYO & ",,,1,,,,,,,,,,,,,,,,,,,,,,,,,"
            
            If Sheets("入力").Range("C4") = "給料" Then
                boolT = AddText(Kname, strTEXT)
            Else
                boolT = AddText(Fname, strTEXT)
            End If
            lngGNO = lngGNO + 1
        End If
    Next lngC
    
End Sub

