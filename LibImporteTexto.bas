Attribute VB_Name = "LibImporteTexto"
Option Explicit



Public Function EscribeImporteLetra(impo As Currency) As String
Dim N As Long
Dim Cad1 As String
Dim Cad2 As String
Dim Men As String

    Men = ""
    N = Int(impo)
    Cad1 = DevuelveImporteLetra(N)
    If N <> 0 Then Men = Men & Cad1 & " euros"
    
    impo = impo - N
    N = impo * 100
    Cad2 = DevuelveImporteLetra(N)
    
    If N <> 0 Then
        If Cad1 <> "" Then Men = Men & " con "   'Por si no hay k escribir el con
        Men = Men & Cad2
        Men = Men & " céntimos de euro"
    End If
    EscribeImporteLetra = Men
End Function



Public Function DevuelveImporteLetra(Importe As Long) As String
Dim Entera As Long
Dim Aux As Long
Dim aux2 As Integer
Dim cadena1 As String
Dim cadena2 As String
Dim cadena3 As String


Entera = Importe 'tenemos la parte entera
'estamos en millones
cadena1 = ""
Aux = Entera \ 1000000
If Aux > 0 Then
    If Aux = 1 Then
        cadena1 = " Un millón"
        Else
            cadena1 = GrupoCien(CInt(Aux), 2) & " millones"
    End If
End If
Entera = Entera - (Aux * 1000000)
'estamos en miles
Aux = Entera \ 1000
cadena2 = ""
If Aux > 0 Then
    If Aux = 1 Then
        cadena2 = cadena2 & " mil "
        Else
            cadena2 = GrupoCien(CInt(Aux), 1) & " mil"
    End If
End If
'estamos en cientos
Entera = Entera - (Aux * 1000)
Aux = Entera
cadena3 = ""
If Aux > 0 Then cadena3 = GrupoCien(CInt(Aux), 0)

'estamos aqui
DevuelveImporteLetra = cadena1 & cadena2 & cadena3
End Function


Private Function GrupoCien(Importe As Integer, SonMillones As Byte) As String
'**************+
'*
'*  Son millones: 2 Millones  /0 Unidades /1 Miles


Dim vCien  As Integer
Dim cadena1 As String
Dim cadena2 As String
Dim cadena3 As String
Dim Aux As String
Dim vDec As Integer
Dim vUni As Integer
Dim nexo As String 'Palabra que unira
Dim aux2 As Integer

    
cadena1 = ""

'Primera comprobacion
If Importe = 100 Then
    GrupoCien = " cien "
    Exit Function
End If

vCien = Importe \ 100
If vCien > 0 Then
    Select Case vCien
        Case 1
            cadena1 = " cien"
        Case 2
            cadena1 = " doscient"
        Case 3
            cadena1 = " trescient"
        Case 4
            cadena1 = " cuatrocient"
        Case 5
            cadena1 = " quinient"
        Case 6
            cadena1 = " seiscient"
        Case 7
            cadena1 = " setecient"
        Case 8
            cadena1 = " ochocient"
        Case 9
            cadena1 = " novecient"
    End Select
    cadena1 = cadena1 & "os"
    
End If
nexo = " "
aux2 = Importe - (vCien * 100)

'Si Vdec >0 entonces cien pasa a ser ciento
If aux2 > 0 And vCien = 1 Then cadena1 = " ciento "
vDec = aux2 \ 10
If vDec > 0 Then
    Select Case vDec
        Case 1
               GrupoCien = cadena1 & NumeroEspecial(aux2)
               Exit Function
        Case 2
            If aux2 = 20 Then
                cadena2 = " veinte"
                Else
                cadena2 = " veinti"
            End If
            nexo = ""
        Case 3
            cadena2 = " treinta"
        Case 4
            cadena2 = " cuarenta"
        Case 5
            cadena2 = " cincuenta"
        Case 6
            cadena2 = " sesenta"
        Case 7
            cadena2 = " setenta"
        Case 8
            cadena2 = " ochenta"
        Case 9
            cadena2 = " noventa"
    End Select
End If

'Unidades
'aux = Mid(Importe, Len(Importe), 1)
vUni = Importe - (vDec * 10) - (vCien * 100)
If vUni > 0 And vDec > 2 Then nexo = " y "
cadena3 = ""
Select Case vUni
        Case 1
            If SonMillones = 2 Then
                cadena3 = "un"
                Else
                    If SonMillones = 1 Then
                        cadena3 = "una"
                        Else
                            cadena3 = "un"
                    End If
            End If
        Case 2
            cadena3 = "dos"
        Case 3
            cadena3 = "tres"
        Case 4
            cadena3 = "cuatro"
        Case 5
            cadena3 = "cinco"
        Case 6
            cadena3 = "seis"
        Case 7
            cadena3 = "siete"
        Case 8
            cadena3 = "ocho"
        Case 9
            cadena3 = "nueve"
End Select
GrupoCien = cadena1 & cadena2 & nexo & cadena3
End Function


Private Function NumeroEspecial(v As Integer) As String
Select Case v
Case 10
    NumeroEspecial = " diez"
Case 11
    NumeroEspecial = " once"
Case 12
    NumeroEspecial = " doce"
Case 13
    NumeroEspecial = " trece"
Case 14
    NumeroEspecial = " catorce"
Case 15
    NumeroEspecial = " quince"
Case 16
    NumeroEspecial = " dieciseis"
Case 17
    NumeroEspecial = " diecisiete"
Case 18
    NumeroEspecial = " dieciocho"
Case 19
    NumeroEspecial = " diecinueve"
End Select
End Function


