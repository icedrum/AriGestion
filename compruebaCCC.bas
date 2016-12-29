Attribute VB_Name = "CompruebaCCC"
'-- Esta librería contiene un conjunto de funciones de utilidad general
Public Function Comprueba_CC(CC As String) As Boolean
    Dim Ent As String ' Entidad
    Dim Suc As String ' Oficina
    Dim DC As String ' Digitos de control
    Dim I, i2, i3, i4 As Integer
    Dim NumCC As String ' Número de cuenta propiamente dicho
    '-- Esta función comprueba la corrección de un número de cuenta pasado en CC
    If Len(CC) <> 20 Then Exit Function '-- Las cuentas deben contener 20 dígitos en total
    
    
    '-- Calculamos el primer dígito de control
    I = Val(Mid(CC, 1, 1)) * 4
    I = I + Val(Mid(CC, 2, 1)) * 8
    I = I + Val(Mid(CC, 3, 1)) * 5
    I = I + Val(Mid(CC, 4, 1)) * 10
    I = I + Val(Mid(CC, 5, 1)) * 9
    I = I + Val(Mid(CC, 6, 1)) * 7
    I = I + Val(Mid(CC, 7, 1)) * 3
    I = I + Val(Mid(CC, 8, 1)) * 6
    i2 = Int(I / 11)
    i3 = I - (i2 * 11)
    i4 = 11 - i3
    Select Case i4
        Case 11
            i4 = 0
        Case 10
            i4 = 1
    End Select
    If i4 <> Val(Mid(CC, 9, 1)) Then Exit Function '-- El primer dígito de control no coincide
    '-- Calculamos el segundo dígito de control
    I = Val(Mid(CC, 11, 1)) * 1
    I = I + Val(Mid(CC, 12, 1)) * 2
    I = I + Val(Mid(CC, 13, 1)) * 4
    I = I + Val(Mid(CC, 14, 1)) * 8
    I = I + Val(Mid(CC, 15, 1)) * 5
    I = I + Val(Mid(CC, 16, 1)) * 10
    I = I + Val(Mid(CC, 17, 1)) * 9
    I = I + Val(Mid(CC, 18, 1)) * 7
    I = I + Val(Mid(CC, 19, 1)) * 3
    I = I + Val(Mid(CC, 20, 1)) * 6
    i2 = Int(I / 11)
    i3 = I - (i2 * 11)
    i4 = 11 - i3
    Select Case i4
        Case 11
            i4 = 0
        Case 10
            i4 = 1
    End Select
    If i4 <> Val(Mid(CC, 10, 1)) Then Exit Function '-- El segundo dígito de control no coincide
    '-- Si llega aquí ambos figitos de control son correctos
    Comprueba_CC = True
End Function


'---- Añade Laura: 04/10/05
Public Function Comprueba_CuentaBan(CC As String) As Boolean
    'Validar que la cuenta bancaria es correcta
    If Trim(CC) <> "" Then
        If Not Comprueba_CC(CC) Then
            MsgBox "La cuenta bancaria no es correcta", vbInformation
        End If
    End If
End Function
'------------------------------


'[Monica]02/09/2013: me la traigo de tesoreria de david
Public Function CodigoDeControl(ByVal strBanOfiCuenta As String) As String

Dim conPesos
Dim lngPrimerCodigo As Long, lngSegundoCodigo As Long
Dim I As Long, J As Long
conPesos = "06030709100508040201"
J = 1
lngPrimerCodigo = 0
lngSegundoCodigo = 0

' Banco(4) + Oficina(4) nos dará el primer dígito de control
For I = 8 To 1 Step -1
  lngPrimerCodigo = lngPrimerCodigo + (Mid$(strBanOfiCuenta, I, 1) * Mid$(conPesos, J, 2))
  J = J + 2
Next I

J = 1 ' reiniciar el contador de pesos

' Número de cuenta nos dará el segundo digito de control
For I = 18 To 9 Step -1
  lngSegundoCodigo = lngSegundoCodigo + (Mid$(strBanOfiCuenta, I, 1) * Mid$(conPesos, J, 2))
  J = J + 2
Next I


' ajustar el primer dígito de control
lngPrimerCodigo = 11 - (lngPrimerCodigo Mod 11)
If lngPrimerCodigo = 11 Then
    lngPrimerCodigo = 0
ElseIf lngPrimerCodigo = 10 Then
    lngPrimerCodigo = 1
End If


' ajustar el segundo dígito de control
lngSegundoCodigo = 11 - (lngSegundoCodigo Mod 11)
If lngSegundoCodigo = 11 Then
    lngSegundoCodigo = 0
ElseIf lngSegundoCodigo = 10 Then
    lngSegundoCodigo = 1
End If

' convertirlos en cadenas y concatenarlos
CodigoDeControl = Format(lngPrimerCodigo) & Format(lngSegundoCodigo)

End Function
