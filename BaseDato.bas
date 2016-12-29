Attribute VB_Name = "BaseDato"
Option Explicit

Private Sql As String

Dim ImpD As Currency
Dim ImpH As Currency
Dim RT As ADODB.Recordset


Dim d As String
Dim H As String
'Para los balances
Dim M1 As Integer   ' años y kmeses para el balance
Dim M2 As Integer
Dim M3 As Integer
Dim A1 As Integer
Dim A2 As Integer
Dim A3 As Integer
Dim vCta As String
Dim vDig As Byte
Dim ImAcD As Currency  'importes
Dim ImAcH As Currency
Dim ImPerD As Currency  'importes
Dim ImPerH As Currency
Dim ImCierrD As Currency  'importes
Dim ImCierrH As Currency
Dim Contabilidad As Integer
Dim Aux As String
Dim vFecha1 As Date
Dim vFecha2 As Date
Dim VFecha3 As Date
Dim Codigo As String
Dim EjerciciosCerrados As Boolean
Dim NumAsiento As Integer
Dim Nulo1 As Boolean
Dim Nulo2 As Boolean

Dim FIniPeriodo As Date
Dim FFinPeriodo As Date

Dim VarConsolidado(2) As String

Dim EsBalancePerdidas_y_ganancias As Boolean

'Para la precarga de datos del balance de sumas y saldos
Dim RsBalPerGan As ADODB.Recordset


'--------------------------------------------------------------------
'--------------------------------------------------------------------
Private Function ImporteASQL(ByRef Importe As Currency) As String
ImporteASQL = ","
If Importe = 0 Then
    ImporteASQL = ImporteASQL & "NULL"
Else
    ImporteASQL = ImporteASQL & TransformaComasPuntos(CStr(Importe))
End If
End Function



'--------------------------------------------------------------------
'--------------------------------------------------------------------
' El dos sera para k pinte el 0. Ya en el informe lo trataremos.
' Con esta opcion se simplifica bastante la opcion de totales
Private Function ImporteASQL2(ByRef Importe As Currency) As String
    ImporteASQL2 = "," & TransformaComasPuntos(CStr(Importe))
End Function



'--------------------------------------------------------------------
'--------------------------------------------------------------------



Public Sub CommitConexion()
    On Error Resume Next
    Conn.Execute "Commit"
    If Err.Number <> 0 Then Err.Clear
End Sub





Public Function SeparaCampoBusqueda(Tipo As String, Campo As String, CADENA As String, ByRef DevSQL As String) As Byte
Dim Cad As String
Dim Aux As String
Dim Ch As String
Dim Fin As Boolean
Dim I, J As String

On Error GoTo ErrSepara
SeparaCampoBusqueda = 1
DevSQL = ""
Cad = ""
Select Case Tipo
Case "N"
    '----------------  NUMERICO  ---------------------
    I = CararacteresCorrectos(CADENA, "N")
    If I > 0 Then Exit Function  'Ha habido un error y salimos
    'Comprobamos si hay intervalo ':'
    I = InStr(1, CADENA, ":")
    If I > 0 Then
        'Intervalo numerico
        Cad = Mid(CADENA, 1, I - 1)
        Aux = Mid(CADENA, I + 1)
        If Not IsNumeric(Cad) Or Not IsNumeric(Aux) Then Exit Function  'No son numeros
        'Intervalo correcto
        'Construimos la cadena
        DevSQL = Campo & " >= " & Cad & " AND " & Campo & " <= " & Aux
        '----
        'ELSE
        Else
            'Prueba
            'Comprobamos que no es el mayor
            If CADENA = ">>" Or CADENA = "<<" Then
                DevSQL = "1=1"
             Else
                    Fin = False
                    I = 1
                    Cad = ""
                    Aux = "NO ES NUMERO"
                    While Not Fin
                        Ch = Mid(CADENA, I, 1)
                        If Ch = ">" Or Ch = "<" Or Ch = "=" Then
                            Cad = Cad & Ch
                            Else
                                Aux = Mid(CADENA, I)
                                Fin = True
                        End If
                        I = I + 1
                        If I > Len(CADENA) Then Fin = True
                    Wend
                    'En aux debemos tener el numero
                    If Not IsNumeric(Aux) Then Exit Function
                    'Si que es numero. Entonces, si Cad="" entronces le ponemos =
                    If Cad = "" Then Cad = " = "
                    DevSQL = Campo & " " & Cad & " " & Aux
            End If
        End If
Case "F"
     '---------------- FECHAS ------------------
    I = CararacteresCorrectos(CADENA, "F")
    If I = 1 Then Exit Function
    'Comprobamos si hay intervalo ':'
    I = InStr(1, CADENA, ":")
    If I > 0 Then
        'Intervalo de fechas
        Cad = Mid(CADENA, 1, I - 1)
        Aux = Mid(CADENA, I + 1)
        If Not EsFechaOKString(Cad) Or Not EsFechaOKString(Aux) Then Exit Function  'Fechas incorrectas
        'Intervalo correcto
        'Construimos la cadena
        Cad = Format(Cad, FormatoFecha)
        Aux = Format(Aux, FormatoFecha)
        'En my sql es la ' no el #
        'DevSQL = Campo & " >=#" & Cad & "# AND " & Campo & " <= #" & AUX & "#"
        DevSQL = Campo & " >='" & Cad & "' AND " & Campo & " <= '" & Aux & "'"
        '----
        'ELSE
        Else
            'Comprobamos que no es el mayor
            If CADENA = ">>" Or CADENA = "<<" Then
                  DevSQL = "1=1"
            Else
                Fin = False
                I = 1
                Cad = ""
                Aux = "NO ES FECHA"
                While Not Fin
                    Ch = Mid(CADENA, I, 1)
                    If Ch = ">" Or Ch = "<" Or Ch = "=" Then
                        Cad = Cad & Ch
                        Else
                            Aux = Mid(CADENA, I)
                            Fin = True
                    End If
                    I = I + 1
                    If I > Len(CADENA) Then Fin = True
                Wend
                'En aux debemos tener el numero
                If Not EsFechaOKString(Aux) Then Exit Function
                'Si que es numero. Entonces, si Cad="" entronces le ponemos =
                Aux = "'" & Format(Aux, FormatoFecha) & "'"
                If Cad = "" Then Cad = " = "
                DevSQL = Campo & " " & Cad & " " & Aux
            End If
        End If
    
    
    
    
Case "T"
    '---------------- TEXTO ------------------
    I = CararacteresCorrectos(CADENA, "T")
    If I = 1 Then Exit Function
    
    'Comprobamos que no es el mayor
     If CADENA = ">>" Or CADENA = "<<" Then
        DevSQL = "1=1"
        Exit Function
    End If
    
    
    I = InStr(1, CADENA, ":")
    If I > 0 Then
        'Intervalo numerico

        Cad = Mid(CADENA, 1, I - 1)
        Aux = Mid(CADENA, I + 1)
        
        'Intervalo correcto
        'Construimos la cadena
        Cad = DevNombreSQL(Cad)
        Aux = DevNombreSQL(Aux)
        'En my sql es la ' no el #
        'DevSQL = Campo & " >=#" & Cad & "# AND " & Campo & " <= #" & AUX & "#"
        DevSQL = Campo & " >='" & Cad & "' AND " & Campo & " <= '" & Aux & "'"
    
    
    Else
    
        'Cambiamos el * por % puesto que en ADO es el caraacter para like
        I = 1
        Aux = CADENA
        
        '++
        If Len(Aux) <> 0 Then
            If InStr(1, Aux, "*") = 0 Then
                Aux = "*" & Aux & "*"
            End If
        End If
        '++
        
        
        While I <> 0
            I = InStr(1, Aux, "*")
            If I > 0 Then Aux = Mid(Aux, 1, I - 1) & "%" & Mid(Aux, I + 1)
        Wend
        'Cambiamos el ? por la _ pue es su omonimo
        I = 1
        While I <> 0
            I = InStr(1, Aux, "?")
            If I > 0 Then Aux = Mid(Aux, 1, I - 1) & "_" & Mid(Aux, I + 1)
        Wend
        Cad = Mid(CADENA, 1, 2)
        If Cad = "<>" Then
            Aux = Mid(CADENA, 3)
            DevSQL = Campo & " LIKE '!" & Aux & "'"
            Else
            DevSQL = Campo & " LIKE '" & Aux & "'"
        End If
    End If


    
Case "B"
    'Como vienen de check box o del option box
    'los escribimos nosotros luego siempre sera correcta la
    'sintaxis
    'Los booleanos. Valores buenos son
    'Verdadero , Falso, True, False, = , <>
    'Igual o distinto
    I = InStr(1, CADENA, "<>")
    If I = 0 Then
        'IGUAL A valor
        Cad = " = "
        Else
            'Distinto a valor
        Cad = " <> "
    End If
    'Verdadero o falso
    I = InStr(1, CADENA, "V")
    If I > 0 Then
            Aux = "True"
            Else
            Aux = "False"
    End If
    'Ponemos la cadena
    DevSQL = Campo & " " & Cad & " " & Aux
    
Case Else
    'No hacemos nada
        Exit Function
End Select
SeparaCampoBusqueda = 0
ErrSepara:
    If Err.Number <> 0 Then MuestraError Err.Number
End Function


Private Function CararacteresCorrectos(vCad As String, Tipo As String) As Byte
Dim I As Integer
Dim Ch As String
Dim Error As Boolean

CararacteresCorrectos = 1
Error = False
Select Case Tipo
Case "N"
    'Numero. Aceptamos numeros, >,< = :
    For I = 1 To Len(vCad)
        Ch = Mid(vCad, I, 1)
        Select Case Ch
            Case "0" To "9"
            Case "<", ">", ":", "=", ".", " ", "-"
            Case Else
                Error = True
                Exit For
        End Select
    Next I
Case "T"
    'Texto aceptamos numeros, letras y el interrogante y el asterisco
    For I = 1 To Len(vCad)
        Ch = Mid(vCad, I, 1)
        Select Case Ch
            Case "a" To "z"
            Case "è", "é", "í" 'Añade Laura: 16/03/06
            Case "A" To "Z"
            Case "0" To "9"
            'QUITAR#### o no.
            'Modificacion hecha 26-OCT-2006.  Es para que meta la coma como caracter en la busqueda
            Case "*", "%", "?", "_", "\", "/", ":", ".", " ", "-", "," ' estos son para un caracter sol no esta demostrado , "%", "&"
            'Esta es opcional
            Case "#", "@", "$"
            Case "<", ">"
            Case "Ñ", "ñ"
            Case Else
                Error = True
                Exit For
                
        End Select
    Next I
Case "F"
    'Numeros , "/" ,":"
    For I = 1 To Len(vCad)
        Ch = Mid(vCad, I, 1)
        Select Case Ch
            Case "0" To "9"
            Case "<", ">", ":", "/", "="
            Case Else
                Error = True
                Exit For
        End Select
    Next I
Case "B"
    'Numeros , "/" ,":"
    For I = 1 To Len(vCad)
        Ch = Mid(vCad, I, 1)
        Select Case Ch
            Case "0" To "9"
            Case "<", ">", ":", "/", "=", " "
            Case Else
                Error = True
                Exit For
        End Select
    Next I
End Select
'Si no ha habido error cambiamos el retorno
If Not Error Then CararacteresCorrectos = 0
End Function






'le pasamos el SQL y vemos si tiene algun dato
Private Function TieneDatosSQL(ByRef Rs As ADODB.Recordset, vSql As String) As Boolean
    TieneDatosSQL = False
    Rs.Open vSql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then TieneDatosSQL = True
    Rs.Close

End Function

'
'   I N F O R M E S        C R I S T A L
Public Function DevNombreInformeCrystal(QueInforme As Integer) As String

    DevNombreInformeCrystal = DevuelveDesdeBD("informe", "scryst", "codigo", CStr(QueInforme), "N")
    If DevNombreInformeCrystal = "" Then
        MsgBox "Opcion NO encontrada: " & QueInforme, vbExclamation
        DevNombreInformeCrystal = "ERROR"
    End If

End Function





