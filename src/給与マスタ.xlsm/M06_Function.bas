Attribute VB_Name = "M06_Function"
Option Explicit

Function �Ǘ��E�攻��(ShainKU As String, KanriKU As String, Kihon As Long, P_gkn As Single)
Attribute �Ǘ��E�攻��.VB_ProcData.VB_Invoke_Func = " \n14"

Dim Fjb As Long

    If ShainKU = "P" Then
        '�߰ĎЈ�
        If P_gkn = 0 Or Kihon = 0 Then
            �Ǘ��E�攻�� = ""
        Else
            Fjb = Round(Kihon / P_gkn, 0) '����
            �Ǘ��E�攻�� = "@\" & Format(Fjb, "#,##0") & " "
        End If
    Else
        �Ǘ��E�攻�� = ""
        If UCase(KanriKU) = "YY" Then �Ǘ��E�攻�� = "����"
        If UCase(KanriKU) = "SS" Then �Ǘ��E�攻�� = "�x�X��"
        If UCase(KanriKU) = "BB" Then �Ǘ��E�攻�� = "����"
        If UCase(KanriKU) = "JJ" Then �Ǘ��E�攻�� = "����"
        If UCase(KanriKU) = "KK" Then �Ǘ��E�攻�� = "�ے�"
        If UCase(KanriKU) = "KS" Then �Ǘ��E�攻�� = "��C"
        If UCase(KanriKU) = "HD" Then �Ǘ��E�攻�� = "�ے��㗝"
        If UCase(KanriKU) = "HK" Then �Ǘ��E�攻�� = "�W��"
        If UCase(KanriKU) = "HH" Then �Ǘ��E�攻�� = "�ǒ�"
    End If

End Function

