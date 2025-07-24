Attribute VB_Name = "mdl_Tratamentos"
 '' M�dulo de Tratamentos
 
Function Wval(WxVariante As Variant) As Double
    '' Fun��o para tratar campos num�ricos

   Dim WxRetorno As String
   Dim WnPosicao As Integer
    'Verifica se � vazio ou nulo
    If (WxVariante = Empty Or WxVariante = Null) Then WxVariante = Format(0, Formato01)
    'Verifica se existe V�RGULA
    If WxVariante <> Empty Then
        WxVariante = Replace(WxVariante, ".", "")
        WnPosicao = InStr(WxVariante, ",")
        If WnPosicao = 0 Then
            WxRetorno = WxVariante
        Else
            WxRetorno = Left(WxVariante, WnPosicao - 1) & "." & Right(WxVariante, Len(WxVariante) - WnPosicao)
        End If
        Wval = Val(WxRetorno)
    End If
   Exit Function
   
ErroFuncao:
   MsgBox "Ocorreu um erro durante a transforma��o do campo " & vbCrLf & Err & vbCrLf & Error, 16, "Fun��o Wval"
   
End Function
