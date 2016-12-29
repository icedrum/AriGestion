Attribute VB_Name = "LibIVA"
Option Explicit

Dim Linea As String  'Solo se genera una linea

Dim Rs As ADODB.Recordset
Dim IdentificacionPresentador As String
Dim I As Integer
Dim NF As Integer


    
    
    
    
    If Not Generaidentificacion(False) Then GoTo Salida
    'PEriodo devengo
    Linea = Linea & RecuperaValor(vPeriodo, 3)  'AÑO
    I = CInt(RecuperaValor(vPeriodo, 1)) 'El periodo
    
    

    If vParam.periodos = 0 Then
        'Trimestral
        Linea = Linea & I & "T"
    Else
        Linea = Linea & Format(I, "00")
    End If
    
    '2008. Importes. Se queda en el punto: Resultado. Falta; a deducir, resultado de la declracion, a compensar
    Linea = Linea & CadenaImportes
    
    
    If Not GeneraPIE_303(Importe, EsACompensar) Then GoTo Salida
    
    'Final

    'Fecha
    Linea = Linea & Format(Day(vFecha), "00")
    Linea = Linea & DatosTexto(Format(vFecha, "mmmm"), 10)
    Linea = Linea & Format(vFecha, "yyyy")
    Linea = Linea & "</T30301>"
    Linea = Linea & Chr(13)
    Linea = Linea & Chr(10)
    
    If Not ImprimeFichero Then GoTo Salida
    
    GenerarFicheroIVA_303 = True
Salida:
    If Err.Number <> 0 Then MuestraError Err.Number
    
End Function


Public Function GenerarFicheroIVA_303_2014(ByRef CadenaImportes As String, Importe As Currency, vFecha As Date, vPeriodo As String, EsACompensar As Byte, CadRegistroAdicional03 As String) As Boolean
Dim Aux As String
Dim k As Integer
On Error GoTo Salida '

    GenerarFicheroIVA_303_2014 = False
    Linea = ""
    Linea = Linea & "<T3030"
    Linea = Linea & RecuperaValor(vPeriodo, 3)  'AÑO
    I = CInt(RecuperaValor(vPeriodo, 1)) 'El periodo
    If vParam.periodos = 0 Then
        'Trimestral
        Linea = Linea & I & "T"
    Else
        Linea = Linea & Format(I, "00")
    End If
    Linea = Linea & "0000><AUX>"
    
    'Blancos
    Linea = Linea & Space(70)  'reservado admon
    Aux = "    "
    Aux = Mid(App.Minor & "." & App.Revision & "   ", 1, 4)  'Version programa
    Linea = Linea & Aux
    
    
    Linea = Linea & "    "  'Nº campo: 10
    
    'Nif empresa desarrollo
    Aux = "B96470190"
    Linea = Linea & Aux
    
    'Campo 12
    Linea = Linea & Space(213)
    
    'Linea = Linea & "</AUX><VECTOR>"
    Linea = Linea & "</AUX>"
    
    
    'Abril 2015. Suprime vector paginas
    'VECTOR PAGINAS
    'Una pagina 1,  una pagina 3   NOSOTROS no consolidamos modelos. Es nuestro antiguo 303
    'Aux = "00100010030001FIN"
   '
   ' Linea = Linea & DatosTexto(Aux, 300)
   ' Linea = Linea & "</VECTOR>"

    
    '***  No se imprimen las lineas, se imprimira una unica linea al final
    
    '************************************
    'AHORA YA VAN LAS DECLARACIONES
    'RESGISTRO >T30301
    
   
    
    Linea = Linea & "<T30301>"
    
    Select Case EsACompensar
    Case 0
        'Devolver
        Linea = Linea & "D"
    Case 1
        
        'Compensar
        Linea = Linea & "C"
        
    Case 2
        'Ing en cuenta corriente
        Linea = Linea & "G"
        
    Case 3
        'ingrso
        Linea = Linea & "I"
    End Select
    '?????
    ' FALTA N=SIN ACTIVIDAD RESULTADO 0
    '       V=CUENTA CORRIENTE TRIBUTARIA DEVOLUCION
    '       U=DOMICILIACION DEL INGRESO EN CCC
    
    If Not Generaidentificacion(False) Then GoTo Salida
   
    Linea = Linea & "2"  'Inscrito en el registro de devol mensual  1.Si   2:NO
    Linea = Linea & "3"  'tributa exclusiva  1.Si   2:NO (Regimen gnral + simplifa)    3: Sol Reg Gnral
    Linea = Linea & "2"  'autoliquidacion conjunta  1.Si   2:NO
    Linea = Linea & "2"  'declarado concurso acreedores 1.Si   2:NO
    Linea = Linea & Space(8)  'Fecha que entra el ERE
    Linea = Linea & " "
    Linea = Linea & "2"  'Criterio de caja 1.Si   2:NO
    Linea = Linea & "222"  'resto opciones critero caja
    
    'PEriodo devengo
    'Campo 19
    Linea = Linea & RecuperaValor(vPeriodo, 3)  'AÑO
    I = CInt(RecuperaValor(vPeriodo, 1)) 'El periodo
    If vParam.periodos = 0 Then
        'Trimestral
        Linea = Linea & I & "T"
    Else
        Linea = Linea & Format(I, "00")
    End If
    
    'Cadena importes ivas deducible y devengado
    Linea = Linea & CadenaImportes
    
    
    'Final IVA
    Linea = Linea & Space(582)  'reservado para la AEAT
    Linea = Linea & Space(13)  'reservado para el sello de la AEAT
    
    Linea = Linea & "</T30301>" & Chr(13) & Chr(10)
    
    
    
    
       
    '***************************************************
    'Registro adicional 303_03    el que lleva los totales
    
   
    
    Linea = Linea & "<T30303>"
    
    
    Linea = Linea & CadRegistroAdicional03
    
    
    'Campo 16 . Liquidacion adicional y campos 17 18 19
    For k = 1 To 4
        Linea = Linea & String(17, "0")
    Next k
    
    
    'Campo 20. Declaracion complementaria y numero justificante anterior
    Linea = Linea & " " & Space(13)
    'Sin actividad
    Linea = Linea & " "
    
    
    'Domiciliacion devolucion . IBAN
    Linea = Linea & Space(34)
    
    
    
    'Informacion aadicional unicamente a cumplimentar en el utlimo trimestre
    ' 6 parejas de "0" +  "    "  '4pos
    For k = 1 To 6
        Linea = Linea & "0    "
    Next k
    
    'Información adicional - Exclusivamente a cumplimentar en el último periodo exonerados de la Declaración-re
    Linea = Linea & " "
    
    'Campos del 37 al 47
    'Decimales
    For k = 1 To 11
        Linea = Linea & String(17, "0")
    Next k
    
    
    '2015. Ya no lleva esto
    '    If Not GeneraPIE_303(Importe, EsACompensar) Then GoTo Salida
    
    'Campo 23. Reservado AEAT
    Linea = Linea & Space(573)
    
    
    Linea = Linea & "</T30303>" & Chr(13) & Chr(10)
       
    
    
    
    
    
    
    
    
    
    'Final GENERAL
    Linea = Linea & "</T3030"
    Linea = Linea & RecuperaValor(vPeriodo, 3)  'AÑO
    If vParam.periodos = 0 Then
        'Trimestral
        Linea = Linea & I & "T"
    Else
        Linea = Linea & Format(I, "00")
    End If
    Linea = Linea & "0000>" & Chr(13) & Chr(10)
    
    
    If Not ImprimeFichero Then GoTo Salida
    
    GenerarFicheroIVA_303_2014 = True
Salida:
    If Err.Number <> 0 Then MuestraError Err.Number
    
End Function




Private Function ImprimeFichero() As Boolean
Dim NF As Integer
    NF = FreeFile
    On Error GoTo EImp
    Open App.Path & "\Hacienda\mod300\" & "miIVA.txt" For Output As #NF
    Print #NF, Linea;
    Close #NF
    ImprimeFichero = True
    
    Exit Function
EImp:
    MuestraError Err.Number, "Imprime fichero"
End Function


Private Function DatosTexto(Datos As String, Espacios As Integer) As String
Dim Aux As String
Dim I As Integer
    Datos = UCase(Datos)
    Aux = Trim(Datos)
    I = Espacios - Len(Aux)
    If I = 0 Then
        DatosTexto = Aux
    Else
        If I > 0 Then
            DatosTexto = Aux & Space(I)
        Else
            DatosTexto = Mid(Datos, 1, Espacios)
        End If
    End If
End Function


'Quitamos el datos texto 2
Private Function DatosTexto3(Datos As Currency, Espacios As Integer) As String
Dim Aux As String

    Aux = "000000000000000000" & Datos
    DatosTexto3 = Right(Aux, Espacios)
     
End Function


Private Function DatosNumeroDec(Dato As Currency, Espacios As Integer) As String
Dim Aux As String
Dim Num As String
Dim I As Integer
Dim Fmat As String


    'Seran en total "Espacios" posiciones, teniendo en cuenta k el signo menos sera una N
    Fmat = String(Espacios - 3, "0") & "." & "00"  'FORMATO
    If Dato < 0 Then
        DatosNumeroDec = "N"
    Else
        DatosNumeroDec = "0"
    End If
    
    Aux = CStr(Abs(Dato))
    I = Len(Aux)
    If InStr(1, Aux, ",") > 1 Then I = I - 1
    If I = Espacios Then
        'Caso especial. La longitud son 13
        If Dato < 0 Then
            MsgBox "Numero excede de las posciones para el programa de hacienda", vbExclamation
            End
            
            
        Else
            
            Num = Format(Abs(Dato), Fmat)
            DatosNumeroDec = Mid(Num, 1, 11) & Right(Num, 2)
        End If
    
    
    Else
        If I < Espacios Then
            Num = Format(Abs(Dato), Fmat)
            I = InStr(1, Num, ",")
            
            DatosNumeroDec = DatosNumeroDec & Mid(Num, 1, I - 1) & Right(Num, 2)
        Else
            MsgBox "El número excede de las posiciones solicitdas en programa de Hacienda. La aplicación finalizará.", vbExclamation
            End
        End If
    End If
End Function

'Es la misma funcion solo que si no es negativo NO pinta Cero, sino un espacio en blanco
Private Function DatosNumeroDec340(Dato As Currency, Espacios As Integer) As String
Dim Aux As String
Dim Num As String
Dim I As Integer
Dim Fmat As String


    'Seran en total "Espacios" posiciones, teniendo en cuenta k el signo menos sera una N
    Fmat = String(Espacios - 3, "0") & "." & "00"  'FORMATO
    If Dato < 0 Then
        DatosNumeroDec340 = "N"
    Else
        DatosNumeroDec340 = " "
    End If
    
    Aux = CStr(Abs(Dato))
    I = Len(Aux)
    If InStr(1, Aux, ",") > 1 Then I = I - 1
    If I = Espacios Then
        'Caso especial. La longitud son 13
        If Dato < 0 Then
            MsgBox "Numero excede de las posciones para el programa de hacienda", vbExclamation
            End
            
            
        Else
            
            Num = Format(Abs(Dato), Fmat)
            DatosNumeroDec340 = Mid(Num, 1, 11) & Right(Num, 2)
        End If
    
    
    Else
        If I < Espacios Then
            Num = Format(Abs(Dato), Fmat)
            I = InStr(1, Num, ",")
            
            DatosNumeroDec340 = DatosNumeroDec340 & Mid(Num, 1, I - 1) & Right(Num, 2)
        Else
            MsgBox "El número excede de las posiciones solicitdas en programa de Hacienda. La aplicación finalizará.", vbExclamation
            End
        End If
    End If
End Function





'///////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////



'-------------------------------------------------
'Genera los primeros datos del archivo, esto es
Private Function GeneraEncabezados() As Boolean
    GeneraEncabezados = True
    Linea = Linea & "300" & "01  "
    
End Function



'-------------------------------------------------
'Identificacion declarante
Private Function Generaidentificacion(Modelo300 As Boolean) As Boolean
    Generaidentificacion = False
    Set Rs = New ADODB.Recordset
    Rs.Open "empresa2", Conn, adOpenDynamic, adLockPessimistic, adCmdTable
    If Rs.EOF Then
        MsgBox "Datos de empresa mal configurados", vbExclamation
    Else
        If Modelo300 Then
            'Administracion
            Linea = Linea & DatosTexto(DBLet(Rs!administracion), 5)
            Linea = Linea & DatosTexto(DBLet(Rs!nifempre), 9)
            Linea = Linea & DatosTexto(DBLet(Rs!siglaempre), 4)
            Linea = Linea & DatosTexto(vEmpresa.nomempre, 30)
            
            Linea = Linea & DatosTexto("", 15)   'es el nombre, pero va vacio puesto k aqui son empresas
            
        Else
            Linea = Linea & DatosTexto(DBLet(Rs!nifempre), 9)
            Linea = Linea & DatosTexto(vEmpresa.nomempre, 30)
            Linea = Linea & DatosTexto(DBLet(Rs!apoderado), 15)
            Linea = Linea & "2"  'Inscrito en el registro de devol mensual  1.Si   2:NO
        End If
        Generaidentificacion = True
    End If
    Rs.Close
    Set Rs = Nothing
End Function



'-------------------------------------------------
'Identificacion declarante
Private Function GeneraPIE(ByRef Importe As Currency, Compensa As Boolean) As Boolean

    GeneraPIE = False
    Set Rs = New ADODB.Recordset
    Rs.Open "empresa2", Conn, adOpenDynamic, adLockPessimistic, adCmdTable
    'Este el el campo: resultado de la declaracion
    Linea = Linea & DatosNumeroDec(Importe, 13)
    If Rs.EOF Then
        MsgBox "Datos de empresa mal configurados", vbExclamation
    Else
        If Importe < 0 Then
            'Modificado el 25 de enero 2005
            'SOLO DEVOLVEMOS
            
            If Compensa Then
                Linea = Linea & DatosNumeroDec(Abs(Importe), 13)  'a compensar
                Linea = Linea & "0000000000000"
            Else
                Linea = Linea & "0000000000000"
                Linea = Linea & DatosNumeroDec(Abs(Importe), 13)  'a devolver
            End If
            'Si es a compensar no pongo cuenta banco
            If Compensa Then
                Linea = Linea & Space(20)
            Else
                Linea = Linea & DatosTexto(DBLet(Rs!banco1), 4)
                Linea = Linea & DatosTexto(DBLet(Rs!oficina1), 4)
                Linea = Linea & DatosTexto(DBLet(Rs!dc1), 2)
                If Not IsNull(Rs!cuenta1) Then
                    Linea = Linea & DatosTexto(Format(Rs!cuenta1, "0000000000"), 10)
                Else
                    Linea = Linea & Space(10)
                End If
            End If
                    
            'Forma pago
            Linea = Linea & "0"    'En 2002 estaba al reves "X "
            
            'El resto va en blanco
            Linea = Linea & "0000000000000" & Space(20)
        Else
            ' Son los campos 56 del informe
            Linea = Linea & DatosNumeroDec(0, 13)  'a compensar
            Linea = Linea & DatosNumeroDec(0, 13)  'devolucion
            Linea = Linea & DatosTexto("", 20)  'cuenta devolucion
            
            If Val(DBLet(Rs!banco2, "T") = 0) Then
                'No tiene cuenta bancaria en configuracion empresa
                Linea = Linea & "1"
                'cuenta del banco del infrome renta
                Linea = Linea & DatosNumeroDec(Importe, 13)
                'LINEA = LINEA & String(20, "0")
                Linea = Linea & Space(20)
            Else
                'OK. Tiene puesta la cuenta bancaria
                Linea = Linea & "2"
                'cuenta del banco del infrome renta
                Linea = Linea & DatosNumeroDec(Importe, 13)
                Linea = Linea & CuentaBancaria(DBLet(Rs!banco2), 4)
                Linea = Linea & CuentaBancaria(DBLet(Rs!oficina2), 4)
                Linea = Linea & CuentaBancaria(DBLet(Rs!dc2), 2)   '
                Linea = Linea & CuentaBancaria(DBLet(Rs!cuenta2), 10)   '
            End If
            
        End If
        
        'Complementaria, num declaracion complementaria
        Linea = Linea & DatosTexto("", 29)
        
        'Persona contacto
        Linea = Linea & DatosTexto(UCase(DBLet(Rs!contacto)), 100)
        Linea = Linea & DatosTexto(DBLet(Rs!tfnocontacto), 9)
        
        
        'Observaciones
        Linea = Linea & Space(350)
        'Localidad
        Linea = Linea & DatosTexto(DBLet(Rs!pobempre), 16)
        
        
        GeneraPIE = True
    End If
    Rs.Close
    Set Rs = Nothing
End Function








Private Function GeneraPIE_303(ByRef Importe As Currency, EsCompensar As Byte) As Boolean
Dim PintaBanco As Boolean
    GeneraPIE_303 = False
    Set Rs = New ADODB.Recordset
    Rs.Open "empresa2", Conn, adOpenDynamic, adLockPessimistic, adCmdTable
    
    Linea = Linea & DatosNumeroDec(0, 17)  'a deducir Casilla 47
    Linea = Linea & DatosNumeroDec(Importe, 17)  'resultado casilla 48
    
    
    If Rs.EOF Then
        MsgBox "Datos de empresa mal configurados", vbExclamation
    Else
        If Importe < 0 Then

            'Nos tiene que devolver hacienda.
            'O bien compensamos, o que nos ingrese(cuando les venga bien claro)
            'A dedcuir
            
            
            
            'Si es a compensar no pongo cuenta banco
            If EsCompensar = 1 Then
                Linea = Linea & DatosNumeroDec(Importe, 17)  'compensar
                Linea = Linea & "0"  'Sin activadidad
                Linea = Linea & String(17, "0") & Space(20)  'importe devolucion y cuenta
            Else
                'DEVOLUCION
                Linea = Linea & DatosNumeroDec(0, 17)  'compensar
                Linea = Linea & "0"  'Sin activadidad
                Linea = Linea & DatosNumeroDec(Importe, 17)  'devolucion
                Linea = Linea & DatosTexto(DBLet(Rs!banco1), 4)
                Linea = Linea & DatosTexto(DBLet(Rs!oficina1), 4)
                Linea = Linea & DatosTexto(DBLet(Rs!dc1), 2)
                If Not IsNull(Rs!cuenta1) Then
                    Linea = Linea & DatosTexto(Format(Rs!cuenta1, "0000000000"), 10)
                Else
                    Linea = Linea & Space(10)
                End If
            End If
                    
            'Forma pago
            Linea = Linea & "0"    'No consta
            
            'Campos del ingreso
            Linea = Linea & String(17, "0") & Space(20)
        Else
            '
            '
            'TENGO QUE PAGAR A LA HACIENDA PUBLICA
            'Dos opciones. Si tengo la cuenta y me han dicho que es adeudo en cuenta y si no sera ingreso
            'El trozo de devolucion y todo eso va a blanco
            Linea = Linea & DatosNumeroDec(0, 17)  'compensar
            Linea = Linea & "0"  'Sin activadidad
            Linea = Linea & String(17, "0") & Space(20)  'importe devolucion y cuenta
            
            'Ahora
            PintaBanco = False
            If EsCompensar = 3 Then
                If Val(DBLet(Rs!banco2, "N")) > 0 Then PintaBanco = True
            End If
            
            
            If Not PintaBanco Then
                'No tiene cuenta bancaria en configuracion empresa
                Linea = Linea & "1"
                'cuenta del banco del infrome renta
                Linea = Linea & DatosNumeroDec(Importe, 17)
                'LINEA = LINEA & String(20, "0")
                Linea = Linea & Space(20)
            Else
                'OK. Tiene puesta la cuenta bancaria
                Linea = Linea & "2"
                'cuenta del banco del infrome renta
                Linea = Linea & DatosNumeroDec(Importe, 17)
                Linea = Linea & CuentaBancaria(DBLet(Rs!banco2), 4)
                Linea = Linea & CuentaBancaria(DBLet(Rs!oficina2), 4)
                Linea = Linea & CuentaBancaria(DBLet(Rs!dc2), 2)   '
                Linea = Linea & CuentaBancaria(DBLet(Rs!cuenta2), 10)   '
            End If
            
        End If
        
        'Complementaria, num declaracion complementaria
        Linea = Linea & "0" & DatosTexto("", 13)
        
        'Persona contacto
        'LINEA = LINEA & DatosTexto(UCase(DBLet(RS!contacto)), 100)
        'LINEA = LINEA & DatosTexto(DBLet(RS!tfnocontacto), 9)
        
        
        'Observaciones
        Linea = Linea & Space(400)
        'Localidad
        Linea = Linea & DatosTexto(DBLet(Rs!pobempre), 16)
        
        
        GeneraPIE_303 = True
    End If
    Rs.Close
    Set Rs = Nothing
End Function

















Private Function CuentaBancaria(Datos As String, L As Integer) As String
If Datos = "" Then
   CuentaBancaria = DatosTexto("", L)
   Else
        CuentaBancaria = Format(Datos, Mid("0000000000", 1, L))
End If
End Function



'///////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////



Public Sub ImpresionModeloOficial()
On Error GoTo EImpresionModeloOficial


    If Dir(App.Path & "\Hacienda\mod300\MIPF32.exe") = "" Then
        MsgBox "No existe la aplicación de hacienda: MIPF32", vbExclamation
        Exit Sub
    End If
    
    'Borramos el archivo Errores
    If Dir(App.Path & "\Hacienda\mod300\" & "IvaErr.txt") <> "" Then Kill App.Path & "\Hacienda\mod300\" & "IvaErr.txt"
    
    'Lanzamos la aplicacion
    Linea = App.Path & "\Hacienda\mod300\MIPF32.exe  /E:""" & App.Path & "\Hacienda\mod300\" & "miIVA.txt"""
    Linea = Linea & " /R:""" & App.Path & "\Hacienda\mod300\" & "IvaErr.txt"""

    Screen.MousePointer = vbHourglass
    Shell Linea
    espera 2
    Do
        espera 0.1
        Linea = Dir(App.Path & "\Hacienda\mod300\FLAG.TXT")
    Loop Until Linea = ""

    Linea = Dir(App.Path & "\IvaErr.txt")
    Linea = ""
    espera 2
    
    If Dir(App.Path & "\Hacienda\mod300\" & "IvaErr.txt") <> "" Then espera 2
    
    'Comprobamos k no se ha generado el fichero de error
    If Dir(App.Path & "\Hacienda\mod300\" & "IvaErr.txt") <> "" Then
        'Si k hay error
        Linea = "Se ha producido un error. Lear el archivo: " & vbCrLf & vbCrLf
        Linea = Linea & App.Path & "\Hacienda\mod300\" & "IvaErr.txt" & vbCrLf & vbCrLf
        Linea = Linea & " para obtener más datos."
        MsgBox Linea, vbExclamation
    End If
EImpresionModeloOficial:
    If Err.Number <> 0 Then MuestraError Err.Number
    Screen.MousePointer = vbDefault
End Sub






'---------------------------------------------------------------------------
'---------------------------------------------------------------------------
'               Modelo 347
'---------------------------------------------------------------------------
'---------------------------------------------------------------------------

Public Function Modelo347(Impresion As Boolean, Anyo As Integer) As Boolean

    Modelo347 = False
    If Not PrimerosPasos(Anyo) Then Exit Function 'Ha fallado generando el fichero
        
    If Impresion Then
        ImpresionModeloOficial347
    End If
    Modelo347 = True
End Function

'dEL 347
Private Function PrimerosPasos(Anyo As Integer) As Boolean
Dim Importe As Currency
Dim Cad As String
Dim YaMostrado As Boolean
Dim RI As ADODB.Recordset
Dim importe2 As Currency

    On Error GoTo EGen347
    PrimerosPasos = False
    
    'Conteo previo
    Set Rs = New ADODB.Recordset
    Linea = "SELECT COUNT(*) FROM tmp347tot WHERE codusu=" & vUsu.Codigo
    Rs.Open Linea, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    If Not Rs.EOF Then I = DBLet(Rs.Fields(0), "N")
    Rs.Close
    If I = 0 Then
        MsgBox "Ninguna entrada para el 347 generada.", vbExclamation
        Exit Function
    End If
        
    'Suma
    Linea = "SELECT SUM(importe) FROM tmp347tot WHERE codusu=" & vUsu.Codigo
    Rs.Open Linea, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Importe = 0
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then Importe = Rs.Fields(0)
    End If
    Rs.Close
        
    Linea = "Select * from empresa2"
    Rs.Open Linea, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NF = 0
    If Rs.EOF Then
        Rs.Close
        MsgBox "Mal configurada la aplicacion. No existen datos fiscales de la empresa", vbExclamation
        Exit Function
    End If

    NF = FreeFile
    'Open App.Path & "\Hacienda\mod347\mod347.txt" For Output As #NF
    Open App.Path & "\docum.pdf" For Output As #NF
    
    IdentificacionPresentador = ""
    Linea = "347"
    Linea = Linea & Anyo
    Linea = Linea & DatosTexto(DBLet(Rs!nifempre), 9)
    IdentificacionPresentador = Linea
    
  
    'Febrero 2012
    '--------------------------------------
    '--------------------------------------
    'Registro tipo 1.- DECLARANTE
    '--------------------------------------
    '--------------------------------------
     '--------------------------------------
    '--------------------------------------
    Linea = "1" 'Obligado
    Linea = Linea & IdentificacionPresentador
    
    Linea = Linea & DatosTexto(vEmpresa.nomempre, 40)
    Linea = Linea & "T"  'tipo presentcion C- cinta   D- diskette T.- TElematica
    Linea = Linea & DatosTexto(DBLet(Rs!telefono), 9)
    Linea = Linea & DatosTexto(DBLet(Rs!contacto), 40)
    Linea = Linea & "3470000000000"   'Numero justificante la declaracion. Empieza por 347
    Linea = Linea & "  "
    Linea = Linea & "0000000000000"   'Numero justificante la declaracion anterior
    Linea = Linea & Format(I, "000000000")
    Linea = Linea & DatosNumeroDec340((Importe), 16)
    Linea = Linea & "000000000000000000000000"   'Inmuebles, total y suma
    Linea = Linea & DatosTexto(" ", 67)
    Cad = Space(500 - Len(Linea))
    Linea = Linea & Cad
    Print #NF, Linea
    Rs.Close
    
    
    '--------------------------------------
    '--------------------------------------
    'Registro tipo 2.- DECLARADO
    '--------------------------------------
    '--------------------------------------
    '--------------------------------------
    '--------------------------------------
    Set RI = New ADODB.Recordset
    Linea = "SELECT * FROM tmp347tot WHERE codusu=" & vUsu.Codigo
    Linea = Linea & " ORDER BY cliprov,nif"
    Rs.Open Linea, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    YaMostrado = False 'varable que nos indicara si han llegado datos incorrectos al 347 en el campo cliprov
    While Not Rs.EOF
        Linea = "2" 'Obligado
        Linea = Linea & IdentificacionPresentador
        
        Linea = Linea & DatosTexto(Rs!NIF, 9)
        Linea = Linea & DatosTexto(" ", 9)
        Linea = Linea & DatosTexto(Rs!razosoci, 40)
        Linea = Linea & "D"
        Linea = Linea & Mid(Rs!codposta, 1, 2)
        Linea = Linea & "  "   'PAIS
        Linea = Linea & " "   'BLANCO psocion 81
        
        
        'SEGUN AEAT. Una linea quedaria asi mas o menos
        '2347201024348588YF46160875         ARIADNA                                 D46   B000000000600000
        
        
        If Rs!cliprov = 48 Then
            Cad = "B"  'ventas
            
        Else
            If Rs!cliprov = 49 Then
                Cad = "A"  'compras
                
            Else
                'Agencias
                'Si no es una agencia AQUI no deberia entrar
                If Not vParam.AgenciaViajes Then
                    'Muestro el mensaje una unica vez
                    If Not YaMostrado Then
                        MsgBox "Letra incorrecta en el campo 'cliprov' ( compra /venta)", vbExclamation
                        YaMostrado = True
                    End If
                End If
                Cad = Chr(Rs!cliprov)
            End If
        End If
        
        
        
        
        
        Linea = Linea & Cad
        
        
        'LINEA = LINEA & DatosTexto3((RS!Importe * 100), 16)
        Linea = Linea & DatosNumeroDec340(Rs!Importe, 16)
        
        
        'Operacion segur
        Linea = Linea & " "
        Linea = Linea & " "
        
        

        
        'Nuevo Febrero 2012
        'Los IVAs trimiestrales
        Cad = "SELECT * FROM tmp347trimestral WHERE codusu=" & vUsu.Codigo
        Cad = Cad & " AND cliprov =" & Rs!cliprov & " AND nif = '" & Rs!NIF & "'"
        importe2 = 0
        RI.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RI.EOF Then
        
        
                    'Febrero 2009
            ' Importe en metalico e importe por trasmisiones sujetas a IVA
            importe2 = DBLet(RI!metalico, "N")
            Cad = "0"
            If importe2 > 0 Then Cad = Anyo
                
            Linea = Linea & DatosNumeroDec(importe2, 15)   'metalico
            Linea = Linea & " " & DatosTexto3(0, 15)   'trasmisiones
            Linea = Linea & DatosTexto3(CCur(Cad), 4)    'ejercicio  NUEVO Febrero 2011
        
        
        
            importe2 = RI!trim1 + RI!trim2 + RI!trim3 + RI!trim4
            If importe2 <> Rs!Importe Then
                Cad = "Importe total y por trimestres distinto" & vbCrLf
                Cad = Cad & "Total " & Importe & vbCrLf
                Cad = Cad & "T1 " & RI!trim1 & "     " & "T2 " & RI!trim2 & "     "
                Cad = Cad & "T3 " & RI!trim3 & "     " & "T4 " & RI!trim4 & "     "
                Err.Raise 513, Cad
            End If
              
            'OK pintamos los trimestrales
            Cad = DatosNumeroDec340(RI!trim1, 16)
            Cad = Cad & " " & DatosTexto3(0, 15)   'trim 1 inmueble
            Cad = Cad & DatosNumeroDec340(RI!trim2, 16)
            Cad = Cad & " " & DatosTexto3(0, 15)   'trim 2 inmueble
            Cad = Cad & DatosNumeroDec340(RI!trim3, 16)
            Cad = Cad & " " & DatosTexto3(0, 15)   'trim 3 inmueble
            Cad = Cad & DatosNumeroDec340(RI!trim4, 16)
            Cad = Cad & " " & DatosTexto3(0, 15)   'trim 4 inmueble
        Else
            Cad = "No se encuentran valores trimestrales para: " & Rs!razosoci
            Err.Raise 513, Cad
        End If
        RI.Close
        Linea = Linea & Cad
        
        
        'Hasta final de lineas
       
        Cad = Space(300)
        Linea = Mid(Linea & Cad, 1, 500)
        Print #NF, Linea
        
        
        'Sig
        Rs.MoveNext
    Wend
    Rs.Close
    
    Close #NF
    PrimerosPasos = True
EGen347:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Generando datos 347"
        If NF <> 0 Then Close #NF
    End If
    Set Rs = Nothing
    Set RI = Nothing
End Function



Public Sub ImpresionModeloOficial347()
On Error GoTo EImpresionModeloOficial


    If Dir(App.Path & "\Hacienda\mod347\347pdf.exe") = "" Then
        MsgBox "No existe la aplicación de hacienda: 347pdf.exe", vbExclamation
        Exit Sub
    End If
    
    'Borramos el archivo Errores
    If Dir(App.Path & "\Hacienda\mod347\" & "IvaErr.txt") <> "" Then Kill App.Path & "\Hacienda\mod347\" & "IvaErr.txt"
    
    'Lanzamos la aplicacion
    Linea = App.Path & "\Hacienda\mod347\347pdf.exe  /E:""" & App.Path & "\Hacienda\mod347\" & "mod347.txt"""
    Linea = Linea & " /R:""" & App.Path & "\Hacienda\mod347\" & "IvaErr.txt"""

    Screen.MousePointer = vbHourglass
    Shell Linea
    espera 1
    Do
        IdentificacionPresentador = Dir(App.Path, vbDirectory)
        espera 0.1
        Linea = Dir(App.Path & "\Hacienda\mod347\FLAG.TXT")
    Loop Until Linea = ""

    espera 2
    
    If Dir(App.Path & "\Hacienda\mod347\" & "IvaErr.txt") <> "" Then
        Linea = Dir(App.Path & "\IvaErr.txt")
        espera 3
    End If
    
    'Comprobamos k no se ha generado el fichero de error
    If Dir(App.Path & "\Hacienda\mod347\" & "IvaErr.txt") <> "" Then
        'Si k hay error
        Linea = "Se ha producido un error. Lea el archivo: " & vbCrLf & vbCrLf
        Linea = Linea & App.Path & "\Hacienda\mod347\" & "IvaErr.txt" & vbCrLf & vbCrLf
        Linea = Linea & " para obtener más información."
        MsgBox Linea, vbExclamation
    End If
EImpresionModeloOficial:
    If Err.Number <> 0 Then MuestraError Err.Number
    Screen.MousePointer = vbDefault
End Sub

Public Function ComprobarNifs347() As Boolean
    On Error GoTo EComprobarNifs347
    ComprobarNifs347 = False
    
    Linea = "select nif,cliprov,razosoci,dirdatos, importe from tmp347 where codusu =  " & vUsu.Codigo
    Linea = Linea & " group by nif"
    Set Rs = New ADODB.Recordset
    Rs.Open Linea, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Linea = ""
    While Not Rs.EOF
        
            'ESPAÑA. Comprobamos NIF"
            If Not Comprobar_NIF(Rs!NIF) Then
                Linea = Linea & Chr(Rs!cliprov) & "      " & Trim(Rs!NIF) & "  " & Rs!razosoci & vbCrLf
            End If
        
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    If Linea <> "" Then
        Linea = "Los siguientes NIFs parecen incorrectos. ¿Continuar de igual modo ?" & vbCrLf & Linea
        If MsgBox(Linea, vbQuestion + vbYesNo) = vbNo Then Exit Function
    End If
    ComprobarNifs347 = True
    Exit Function
EComprobarNifs347:
    MuestraError Err.Number, Err.Description
    Linea = ""
End Function










'---------------------------------------------------------------------------
'---------------------------------------------------------------------------
'               fin 347
'---------------------------------------------------------------------------
'---------------------------------------------------------------------------

'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'---------------------------------------------------------------------------
'---------------------------------------------------------------------------
'               Modelo 349
'---------------------------------------------------------------------------
'---------------------------------------------------------------------------

Public Function MODELO349(Impresion As Boolean, Periodo As String, AnyoPres As Integer) As Boolean

    MODELO349 = False
    If Not PrimerosPasos349(1, Periodo, AnyoPres) Then Exit Function  'Ha fallado generando el fichero
        
        
        
    'Abril 2009
    '------------------------
    'Ya no hay modulo de impresion. Siempre genrar fichero
     MODELO349 = True
     Exit Function
    If Impresion Then
        If ImpresionModeloOficial349 Then MODELO349 = True
    End If
End Function



' 0.- Intentara generar en papel.
'       Si no puede pasara a generar en diskette
' 1.- Diskette
' 2.- Telematica
Private Function PrimerosPasos349(presentacion As Byte, vPeriodo As String, AnyoPres As Integer) As Boolean
Dim Importe As Currency
Dim Contador As Integer
Dim Cad As String


    On Error GoTo EGen347
    PrimerosPasos349 = False
    
    'Conteo previo
    Set Rs = New ADODB.Recordset
    Linea = "SELECT COUNT(*) FROM tmp347tot WHERE codusu=" & vUsu.Codigo
    Rs.Open Linea, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Contador = 0
    If Not Rs.EOF Then Contador = DBLet(Rs.Fields(0), "N")
    Rs.Close
    If Contador = 0 Then
        Linea = "Ninguna entrada para el 349 generada." & vbCrLf & "¿Continuar?"
        If MsgBox(Linea, vbQuestion + vbYesNo) = vbNo Then Exit Function
    End If
    
    
    'No podemos presentar 25 a papel
    If Contador > 25 Then
        If presentacion = 0 Then
            MsgBox "Mas de 25 registros. No puede presentarse en papel", vbExclamation
            Exit Function
        End If
    End If
        
    'Suma
    Linea = "SELECT SUM(importe) FROM tmp347tot WHERE codusu=" & vUsu.Codigo
    Rs.Open Linea, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Importe = 0
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then Importe = Rs.Fields(0)
    End If
    Rs.Close
        
    Linea = "Select * from empresa2"
    Rs.Open Linea, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NF = 0
    If Rs.EOF Then
        Rs.Close
        MsgBox "Mal configurada la aplicacion. No existen datos fiscales de la empresa", vbExclamation
        Exit Function
    End If

    NF = FreeFile
    Open App.Path & "\Hacienda\mod349\mod349.txt" For Output As #NF
    
    IdentificacionPresentador = ""
    Linea = "349"
    'MODIFICADO 11 Abril 2006
    'LINEA = LINEA & Year(vParam.fechaini)
    Linea = Linea & Format(AnyoPres, "0000") 'Year(Fecha)
    Linea = Linea & DatosTexto(DBLet(Rs!nifempre), 9)
    IdentificacionPresentador = Linea
    
    
    Linea = "0"
    Linea = Linea & "349"
    Linea = Linea & DatosTexto("", 246)
    'Print #NF, LINEA
    
    '--------------------Modificacion en Enero 2012
    
    '--------------------------------------
    '--------------------------------------
    'Registro tipo 1.- DECLARANTE
    '--------------------------------------
    '--------------------------------------
     '--------------------------------------
    '--------------------------------------
    Linea = "1" 'Obligado
    Linea = Linea & IdentificacionPresentador
    
    Linea = Linea & DatosTexto(vEmpresa.nomempre, 40)
    
    'enero 2012-- > Siempre una T
    'Select Case presentacion
    'Case 0
    '    Cad = "P"
    'Case 1
    '    Cad = "D"
    'Case Else
    '    Cad = "T"
    'End Select
    Cad = "T"
       
    Linea = Linea & Cad
    Linea = Linea & DatosTexto(DBLet(Rs!telefono), 9)
    Linea = Linea & DatosTexto(DBLet(Rs!contacto), 40)
    Linea = Linea & "3490000000000"   'Numero justificante la declaracion. Empieza por 343. ENERO> 349
    Linea = Linea & "  "
    Linea = Linea & "0000000000000"   'Numero justificante la declaracion anterior s. Empieza por 348
    'Periodo
    Linea = Linea & vPeriodo
    
    'Total registros 2
    Linea = Linea & Format(Contador, "000000000")
    Linea = Linea & DatosTexto3((Importe * 100), 15)
    
    'Total rectificaciones e importe de las mismas
    Linea = Linea & "000000000"
    Linea = Linea & "000000000000000"
    Linea = Linea & DatosTexto(" ", 65)
    'enero 2012
    'blancos hasta 500
    Linea = Linea & Space(500)
    Linea = Mid(Linea, 1, 500)
    Print #NF, Linea
    Rs.Close
    
    
    '--------------------------------------
    '--------------------------------------
    'Registro tipo 2.- OPERACION
    '--------------------------------------
    '--------------------------------------
    '--------------------------------------
    '--------------------------------------
    Linea = "SELECT * FROM tmp347tot WHERE codusu=" & vUsu.Codigo
    Rs.Open Linea, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        Linea = "2" 'Obligado
        Linea = Linea & IdentificacionPresentador
        Linea = Linea & DatosTexto(" ", 58)
        'NIF INTACOM
        'MOdificacion de marzo 2009
        'El NIF lleva las letras del pais... que cojera del pais, NO del NIF
        'LINEA = LINEA & DatosTexto(RS!NIF, 17)
        
        Cad = DBLet(Rs!desPobla) & "  "   'Llevara el pais
        Cad = Trim(Mid(Cad, 1, 2)) & Rs!NIF
        
        Linea = Linea & DatosTexto(Cad, 17)
        Linea = Linea & DatosTexto(Rs!razosoci, 40)
        
        
        'JULIO 2012
        'Estaba al reves
        'If RS!cliprov = 1 Then
        If Rs!cliprov = 0 Then
            Cad = "E"  'ventas o entregas
        Else
            Cad = "A"  'compras o adquisiciones
        End If
        Linea = Linea & Cad
        
        Linea = Linea & DatosTexto3((Rs!Importe * 100), 13)
        
        
        'Hasta final de lineas
       ' LINEA = LINEA & DatosTexto(" ", 104)
        Linea = Linea & Space(500)
        Linea = Mid(Linea, 1, 500)
        Print #NF, Linea
        
        
        'Sig
        Rs.MoveNext
    Wend
    Rs.Close
    
    Close #NF
    PrimerosPasos349 = True
EGen347:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Generando datos 349"
        If NF <> 0 Then Close #NF
    End If
    Set Rs = Nothing
End Function




Public Function ImpresionModeloOficial349() As Boolean
Dim NombreExe As String

On Error GoTo EImpresionModeloOficial


    ImpresionModeloOficial349 = False
    
    'Busco el nombre de la aplicacion
    If Dir(App.Path & "\Hacienda\mod349", vbDirectory) = "" Then
        MsgBox "No existe carpeta: " & App.Path & "\Hacienda\mod349", vbExclamation
        Exit Function
    End If
    
    NombreExe = Dir(App.Path & "\Hacienda\mod349\*.exe", vbArchive)
    If NombreExe = "" Then
        MsgBox "No se ha encontrado ningun exe en la carpeta: " & App.Path & "\Hacienda\mod349", vbExclamation
        Exit Function
    End If
    
    NombreExe = App.Path & "\Hacienda\mod349\" & NombreExe
    
    
    
    If Dir(NombreExe) = "" Then
        MsgBox "No existe la aplicación de hacienda: " & NombreExe, vbExclamation
        Exit Function
    End If
    
    'Borramos el archivo Errores
    If Dir(App.Path & "\Hacienda\mod349\" & "349Err.txt") <> "" Then Kill App.Path & "\Hacienda\mod349\" & "349Err.txt"
    
    'Lanzamos la aplicacion
    'LINEA = App.path & "\Hacienda\mod349\vs349.exe  /E:""" & App.path & "\Hacienda\mod349\" & "mod349.txt"""
    Linea = NombreExe & "  /E:""" & App.Path & "\Hacienda\mod349\" & "mod349.txt"""
    Linea = Linea & " /R:""" & App.Path & "\Hacienda\mod349\" & "349Err.txt"""

    Screen.MousePointer = vbHourglass
    Shell Linea
    espera 1
    Do
        IdentificacionPresentador = Dir(App.Path, vbDirectory)
        espera 0.1
        Linea = Dir(App.Path & "\Hacienda\mod347\FLAG.TXT")
    Loop Until Linea = ""

    espera 2
    
    If Dir(App.Path & "\Hacienda\mod349\" & "349Err.txt") <> "" Then
        Linea = Dir(App.Path & "\349Err.txt")
        espera 3
    End If
    
    'Comprobamos k no se ha generado el fichero de error
    If Dir(App.Path & "\Hacienda\mod349\" & "349Err.txt") <> "" Then
        'Si k hay error
        Linea = "Se ha producido un error. Lea el archivo: " & vbCrLf & vbCrLf
        Linea = Linea & App.Path & "\Hacienda\mod349\" & "349Err.txt" & vbCrLf & vbCrLf
        Linea = Linea & " para obtener más información."
        MsgBox Linea, vbExclamation
    Else
        ImpresionModeloOficial349 = True
    End If
EImpresionModeloOficial:
    If Err.Number <> 0 Then MuestraError Err.Number
    Screen.MousePointer = vbDefault
End Function





'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
'
'       MODELO   340
'       MODELO   340
'       MODELO   340
'       MODELO   340
'
'
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
' FraPro_MismaFecha:  Facturas de proveedores. Grabamos la misma fecha para factura y recepcion (RECEPCION)
'
'   Los tickets:   VectorDatosEmpresa340. LlevaraSerieOCta|k serie o cta1|cta2|
Public Function Modelo340(ByRef ListadoEmpresas As ListView, Anyo As Integer, Periodo As Byte, Opciones As String, lbl As Label, QuitarREA As Boolean, aFichero As Boolean, ByRef VectorDatosEmpresa340(), UtlPeriodoLiquidacion As Boolean) As Boolean
Dim Sql As String
Dim I As Integer
Dim NIFError As Integer  'Un fichero de texto
Dim NIFErrores As Boolean
Dim Minimo As Long
Dim SerieTickets As String
Dim varTicketsEn340LetraSerie As Boolean
Dim vAux As Integer
Dim ImportePagosMetalico As Currency


    On Error GoTo EModelo340
    Modelo340 = False
    
    

    
    
    'Borramos temporal
    Conn.Execute "DELETE FROM tmp340 WHERE codusu = " & vUsu.Codigo
    
    NIFError = FreeFile
    NIFErrores = False 'No tiene errores
    
    Open App.Path & "\tmpNiferr.txt" For Output As #NIFError
    Set Rs = New ADODB.Recordset
    
    
    
    
    NumRegElim = 0
    'Para todas las empresas que tenga el list
    For I = 1 To ListadoEmpresas.ListItems.Count
        If ListadoEmpresas.ListItems(I).Checked Then
    
            lbl.Caption = "Ariconta " & ListadoEmpresas.ListItems(I).Text & " - Preparando"  '.ItemData(I - 1) & " - Preparando"
            lbl.Refresh
            
            'Dic. 2012
            'Metalico
            'Veremos si esta empresa tiene el parametro de declaracion de efectivo.  DEBERIAN tenerlo todas
            ImportePagosMetalico = 0
            Sql = "ariconta" & ListadoEmpresas.ListItems(I).Text & ".parametros" '& ListadoEmpresas.ItemData(I - 1) & ".parametros"
            Sql = DevuelveDesdeBD("ImporteMaxEfec340", Sql, "1", "1")
            If Sql <> "" Then ImportePagosMetalico = CCur(Sql)
        
                
                
'            SerieTickets = VectorDatosEmpresa340(I - 1)
'            Sql = RecuperaValor(SerieTickets, 1)
'            varTicketsEn340LetraSerie = (Sql = "1")
'            'Si es serie lleva directamente la serie, sin ppipes
'            'si es cta llevara cta1|cta|
'            If varTicketsEn340LetraSerie Then
'                Sql = RecuperaValor(SerieTickets, 2) 'Sin pipes ni nada
'            Else
'                Sql = RecuperaValor(SerieTickets, 2) & "|" & RecuperaValor(SerieTickets, 3) & "|"
'            End If
'            SerieTickets = Sql

            SerieTickets = "J"
            Sql = ""
            
            Minimo = NumRegElim
            
    
            'Facturas emitidas normales
            lbl.Caption = "Ariconta " & ListadoEmpresas.ListItems(I).Text & "  Emitidas" '.ItemData(I - 1) & "  Emitidas"
            lbl.Refresh
            Sql = CadenaSelect340(True, False, Periodo, Anyo, False)
            CargaFacturasEmitidas ListadoEmpresas.ListItems(I).Text, Sql, SerieTickets, varTicketsEn340LetraSerie
            
            
            'Facturas recibidas NORMALES
            lbl.Caption = "Ariconta " & ListadoEmpresas.ListItems(I).Text & "  Recibidas"
            lbl.Refresh
            Sql = CadenaSelect340(False, False, Periodo, Anyo, False)
            CargaFacturasRecibidas2 ListadoEmpresas.ListItems(I).Text, Sql, QuitarREA
            
            
            
            'Facturas recibidas con bien de inversion
            
            'If UtlPeriodoLiquidacion Then
            CargaFacturasRecibidasBienInversion ListadoEmpresas.ListItems(I).Text, Anyo
            
            
            
            
            
            'DICIEMBRE 2012.
            ' METALICO
            '   En los conceptos hay una marca para decir si el pago es en metalico o no
            '   Si hay pagos en metalico habra que buscar aquellos
            If ImportePagosMetalico > 0 Then
                'Vemos que conceptos tienen la marca de "efectivo"
                Sql = "ariconta" & ListadoEmpresas.ListItems(I).Text & ".conceptos"
                Sql = "Select * from " & Sql & " WHERE EsEfectivo340 = 1"
                Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                Sql = ""
                While Not Rs.EOF
                    Sql = Sql & ", " & Rs!codconce
                    Rs.MoveNext
                Wend
                Rs.Close
                
                If Sql <> "" Then
                    Sql = Mid(Sql, 2)
                    
                    'En esta funcion guardaremos los datos de los pagos en efectivo
                    CargarPagosEnEfectivo CByte(ListadoEmpresas.ListItems(I).Text), ImportePagosMetalico, Sql, Periodo, Anyo, lbl
                End If
            End If
            
            'Comprobamos NIFS
            '-------------------------------------
            ' Si linea="" entonces NO tiene errores
            lbl.Caption = "Ariconta " & ListadoEmpresas.ListItems(I).Text & "  comprueba NIF"
            lbl.Refresh
            vAux = ComprobarNifs340(Minimo, NumRegElim)
            If vAux <> 0 Then
                Sql = "ariconta" & ListadoEmpresas.ListItems(I).Text & ".empresa"
                Sql = DevuelveDesdeBD("nomempre", Sql, "codempre", ListadoEmpresas.ListItems(I).Text)
                If Sql = "" Then Sql = "Ariconta: " & ListadoEmpresas.ListItems(I).Text
                Sql = "Empresa(" & ListadoEmpresas.ListItems(I).Text & ").    " & Sql & vbCrLf & String(40, "=") & vbCrLf
                Sql = Sql & Linea & vbCrLf & vbCrLf & vbCrLf
                
                If Minimo = 0 Then
                    'Es la primera vez que imprime lo de los nifs
                    Linea = String(50, "*") & vbCrLf & vbCrLf
                    Linea = Linea & "NIF con erroes" & vbCrLf & Linea
                    Sql = Linea & Sql
                End If
                
                If vAux < 0 Then Sql = Sql & vbCrLf & vbCrLf & "Error comprobando NIFs"
                Print #NIFError, Sql
                NIFErrores = True
            End If
        End If
    Next I
    
    Close #NIFError
    NIFError = -1
    
    If aFichero Then
        If NIFErrores Then
            'Tiene Errores
            LanzaNotepad App.Path & "\tmpNiferr.txt", "NIFs con errores"
            espera 0.5
        End If
    End If
    
    'Si llega aqui, veremos si HAY alguna
    If NumRegElim > 0 Then
    
        'Agrupamos el efectivo
        'Por team de consolidacion, ahora en el tmp340 estaran los pagos en efectivo
        'Harab que hacer un select group by where clavelbro='D' para poder extraerlos
        'Luego borraremos la 'D'  y los datos extraidos los grabaremos con la 'Z'
        'para que salgan al final
        If ImportePagosMetalico > 0 Then AgrupaPagosEfectivo lbl, ImportePagosMetalico, aFichero
    
        
        Modelo340 = True
'        Sql = "INSERT INTO usuarios.z340 select * from tmp340 WHERE codusu = " & vUsu.Codigo
'        EjecutaSQL Sql
    Else
        MsgBox "No se han generado datos", vbExclamation
    End If
        
        
    'Cerrar
    Set Rs = Nothing
    Exit Function
EModelo340:
    MuestraError Err.Number, Err.Description
    Set Rs = Nothing
    If NIFError > 0 Then Close #NIFError
End Function


Private Sub CargaFacturasEmitidas(NumeroEmpresa As Byte, CadWhere As String, SerieTicket As String, vTicketsEn340LetraSerie As Boolean)
Dim PAIS As String
Dim CadenaInsert As String
Dim Identificacion As Byte
Dim TotalLin As Currency
Dim EsTicket As Boolean
' PARA cu
Dim SerieAnt As String
Dim EsPorCtaAjena As Boolean
Dim ErroresCtaAjena As String  'Cuando en lugar de codfaccl va a co
Dim SqlNew As String

    '0: ESPAÑA
    '1: De momento van juntos intracom y extranjero. Ya veremos com separamos

    Linea = "select factcli.*,cuentas.razosoci, factcli_totales.numlinea, factcli_totales.baseimpo, factcli_totales.codigiva, factcli_totales.porciva, factcli_totales.porcrec, factcli_totales.impoiva, factcli_totales.imporec"
    Linea = Linea & " from ariconta" & NumeroEmpresa & ".factcli as factcli,"
    Linea = Linea & "ariconta" & NumeroEmpresa & ".cuentas cuentas, "
    Linea = Linea & "ariconta" & NumeroEmpresa & ".factcli_totales factcli_totales "
    Linea = Linea & " where factcli.codmacta=cuentas.codmacta AND " & CadWhere
    Linea = Linea & " and factcli.numserie = factcli_totales.numserie "
    Linea = Linea & " and factcli.numfactu = factcli_totales.numfactu "
    Linea = Linea & " and factcli.anofactu = factcli_totales.anofactu "
    'Voy a ordenar por numserie para no leer tantas veces de contadores
    Linea = Linea & " ORDER BY factcli.numserie"
    CadenaInsert = ""
    ErroresCtaAjena = ""
    Rs.Open Linea, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
            'Para cada factura si tiene varias bases, el trzo sera igual menos el importe final(de la cuto % y totaliva)
    
    
            EsTicket = (DBLet(Rs!codconce340, "T") = "J")
'            If vTicketsEn340LetraSerie Then
'                EsTicket = (Rs!NumSerie = SerieTicket)
'            Else
'                If InStr(1, SerieTicket, Rs!codmacta) > 0 Then EsTicket = True
'            End If
            '`nifdeclarado`,`nifrepresante`,`razosoci`,
            Linea = DBLet(Rs!razosoci, "T")
            If Linea = "" Then Linea = Rs!Nommacta 'por si acaso no tiene razon social
            
            Linea = Devnombresql340(Linea)
            
                        
            'Si es ticket aqui no va nada
            If EsTicket Then
                Linea = "'',NULL,'" & DevNombreSQL(Linea) & "',"
            Else
                Linea = "'" & DBLet(Rs!nifdatos, "T") & "',NULL,'" & DevNombreSQL(Linea) & "',"
            End If
                
            '`codpais`,`idenpais`,`nifresidencia`
            PAIS = UCase(DBLet(Rs!codPAIS, "T"))
            If PAIS = "" Then PAIS = "ES"
            
            If PAIS = "ES" Then
                Linea = Linea & "'ES','1',NULL"
                Identificacion = 0
            Else
                'ESTRANJERO o INTRACOM
                Linea = Linea & "'" & Mid(PAIS, 1, 2) & "','2','"
                Linea = Linea & DBLet(Rs!nifdatos, "T") & "'"
                Identificacion = 1
            End If
            
            '`clavelibro`,`claveoperacion`,
'            NF = 1
'            If Not IsNull(Rs!tp2faccl) Then NF = NF + 1
'            If Not IsNull(Rs!tp3faccl) Then NF = NF + 1
            SqlNew = "select count(*) from factcli_totales where numserie = " & DBSet(Rs!NumSerie, "T")
            SqlNew = SqlNew & " and numfactu = " & DBSet(Rs!NumFactu, "N")
            SqlNew = SqlNew & " and anofactu = " & DBSet(Rs!anofactu, "N")
            
            NF = TotalRegistros(SqlNew)
            
            
            If EsTicket Then
                PAIS = "J"
            Else
                If NF = 1 Then
                    'SOLO HAY UNA base
                    PAIS = " "
                Else
                    PAIS = "C" 'mas de una base
                End If
                
            End If
            Linea = Linea & ",'E','" & PAIS & "',"
            
            
            '`fechaexp`,`fechaop`,
            PAIS = "'" & Format(Rs!FecFactu, FormatoFecha) & "'"
            Linea = Linea & PAIS & "," & PAIS   'fechaexp y fechaop
           
           
            If Rs!NumSerie <> SerieAnt Then
                 SerieAnt = Rs!NumSerie
                 PAIS = DevuelveDesdeBD("facliajena", "ariconta" & NumeroEmpresa & ".contadores", "tiporegi", SerieAnt, "T")
                 EsPorCtaAjena = PAIS = "1"
            End If
                
            If EsPorCtaAjena Then
                'COJE LO QUE HAYA EN confaccl
                PAIS = DBLet(Rs!observa, "T") ' antes confaccl
                If PAIS = "" Then
                    PAIS = Rs!NumSerie & Format(Rs!NumFactu, "00000000")
                    ErroresCtaAjena = ErroresCtaAjena & "   - " & PAIS & vbCrLf
                End If
                    
            Else
                'LO NORMAL, es decir codfaccl
                PAIS = Rs!NumSerie & Format(Rs!NumFactu, "00000000")
                
            End If
            Linea = Linea & ",'" & PAIS & "',"
           
            'rectifica,dom_intracom,pob_intracom,cp_intracom,"
            If Rs!totfaccl < 0 Then
                 'Rectifica  FALTA###
                 Linea = Linea & "'" & DevNombreSQL(DBLet(Rs!confaccl, "T")) & "'"
            Else
                 'NULL
                 Linea = Linea & "NULL"
            End If
            
           'If Identificacion = 1 Then
           Linea = Linea & ",NULL,NULL,NULL,"
           
           'Base UNO. SIEMPRE EXISTE
           TotalLin = Rs!Impoiva + Rs!Baseimpo + DBLet(Rs!ImpoRec, "N")
           PAIS = "NULL," & NF & "," & TransformaComasPuntos(CStr(Rs!porciva)) & "," & TransformaComasPuntos(CStr(Rs!Baseimpo)) _
                & "," & TransformaComasPuntos(CStr(Rs!Impoiva)) & "," & TransformaComasPuntos(CStr(TotalLin))
           PAIS = PAIS & "," & TransformaComasPuntos(CStr(DBLet(Rs!porcrec, "N"))) & "," & TransformaComasPuntos(CStr(DBLet(Rs!ImpoRec, "N"))) & ")"
            
           'Insertar
           CadenaInsert = CadenaInsert & ",(" & vUsu.Codigo & "," & NumRegElim & "," & Linea & PAIS
           NumRegElim = NumRegElim + 1
           
'           'Base DOS si no es null
'           If Not IsNull(Rs!ba2faccl) Then
'                TotalLin = Rs!ti2faccl + Rs!ba2faccl + DBLet(Rs!tr2faccl, "N")
'                PAIS = "NULL," & NF & "," & TransformaComasPuntos(CStr(Rs!pi2faccl)) & "," & TransformaComasPuntos(CStr(Rs!ba2faccl)) _
'                & "," & TransformaComasPuntos(CStr(Rs!ti2faccl)) & "," & TransformaComasPuntos(CStr(TotalLin))
'                PAIS = PAIS & "," & TransformaComasPuntos(CStr(DBLet(Rs!pr2faccl, "N"))) & "," & TransformaComasPuntos(CStr(DBLet(Rs!tr2faccl, "N"))) & ")"
'
'                'Insertar
'                CadenaInsert = CadenaInsert & ",(" & vUsu.Codigo & "," & NumRegElim & "," & Linea & PAIS
'                NumRegElim = NumRegElim + 1
'           End If
'
'
'           'Base TRES si no es null
'           If Not IsNull(Rs!ba3faccl) Then
'                TotalLin = Rs!ti3faccl + Rs!ba3faccl + DBLet(Rs!tr3faccl, "N")
'                PAIS = "NULL," & NF & "," & TransformaComasPuntos(CStr(Rs!pi3faccl)) & "," & TransformaComasPuntos(CStr(Rs!ba3faccl)) _
'                & "," & TransformaComasPuntos(CStr(Rs!ti3faccl)) & "," & TransformaComasPuntos(CStr(TotalLin))
'                PAIS = PAIS & "," & TransformaComasPuntos(CStr(DBLet(Rs!pr3faccl, "N"))) & "," & TransformaComasPuntos(CStr(DBLet(Rs!tr3faccl, "N"))) & ")"
'
'                'Insertar
'                CadenaInsert = CadenaInsert & ",(" & vUsu.Codigo & "," & NumRegElim & "," & Linea & PAIS
'                NumRegElim = NumRegElim + 1
'           End If
           
           
           'HACEMOS EL INSERT
           
           If Len(CadenaInsert) > 100000 Then
                PAIS = DevuelveInsertTmp340(0)
                PAIS = PAIS & Mid(CadenaInsert, 2)
                Conn.Execute PAIS
                CadenaInsert = ""
           End If
           
           
           Rs.MoveNext
        Wend
        Rs.Close
        If ErroresCtaAjena <> "" Then
            ErroresCtaAjena = "Errores en facturas por cuenta ajena en empresa(" & NumeroEmpresa & "): " & vbCrLf & ErroresCtaAjena & vbCrLf & "Cargara en Nª factura asociado"
            MsgBox ErroresCtaAjena, vbExclamation
        End If
        If CadenaInsert <> "" Then
            PAIS = DevuelveInsertTmp340(0)
            PAIS = PAIS & Mid(CadenaInsert, 2)
            Conn.Execute PAIS
        End If
End Sub


'aui aaui aqui
'BIEN DE INVERSION EN EL UTL PERIODO UNICAMENTE. NUNCA las cargare desde aqui
Private Sub CargaFacturasRecibidas2(NumeroEmpresa As Byte, CadenaWhere As String, QuitarREA As Boolean)
'FraPro_MismaFecha
Dim PAIS As String
Dim CadenaInsert As String
Dim IvaBienesInversion As String
Dim IvaREA As String
Dim IVA_BI As Boolean
Dim Identificacion As Byte
Dim TotalLin As Currency
Dim B As Boolean
Dim SqlNew As String


    '0: ESPAÑA
    '1: De momento van juntos intracom y extranjero. Ya veremos com separamos
    
    
    'Cargare los IVAS que seand e bienes de inversion
    Linea = "Select * from ariconta" & NumeroEmpresa & ".tiposiva where tipodiva = 2"
    Rs.Open Linea, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    IvaBienesInversion = "|"
    While Not Rs.EOF
        IvaBienesInversion = IvaBienesInversion & Rs!codigiva & "|"
        Rs.MoveNext
    Wend
    Rs.Close
    
    
    'Enero 2012.
    'El REA
    
    IvaREA = "|"
    
    Linea = "Select * from ariconta" & NumeroEmpresa & ".tiposiva where tipodiva = 3"  'REA
    Rs.Open Linea, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    While Not Rs.EOF
        IvaREA = IvaREA & Rs!codigiva & "|"
        Rs.MoveNext
    Wend
    Rs.Close

    
    Linea = "select factpro.*,razosoci, factpro_totales.numlinea, factpro_totales.baseimpo, factpro_totales.codigiva, factpro_totales.porciva, factpro_totales.porcrec, factpro_totales.impoiva, factpro_totales.imporec "
    Linea = Linea & " from ariconta" & NumeroEmpresa & ".factpro factpro,"
    Linea = Linea & "ariconta" & NumeroEmpresa & ".cuentas cuentas, "
    Linea = Linea & "ariconta" & NumeroEmpresa & ".factpro_totales factpro_totales "
    Linea = Linea & " where factpro.codmacta=cuentas.codmacta AND " & CadenaWhere
    Linea = Linea & " and factpro.numserie = factpro_totales.numserie "
    Linea = Linea & " and factpro.numregis = factpro_totales.numregis "
    Linea = Linea & " and factpro.anofactu = factpro_totales.anofactu "
    'FALTA el periodod

    CadenaInsert = ""
    Rs.Open Linea, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
            'Para cada factura si tiene varias bases, el trzo sera igual menos el importe final(de la cuto % y totaliva)
                    
    
            'Campos en la tmp
            '`tmp340` (`codusu`,`codigo`,`nifdeclarado`,`nifrepresante`,`razosoci`,
            '`codpais`,`idenpais`,`nifresidencia`,
            '`clavelibro`,`claveoperacion`,
            '`fechaexp`,`fechaop`,`idfactura`,`totalfac`,
            'numreg`,`tipo`,`base`,`totiva`,
            '`,`numiva`,`ivare`,`cuotare`
                        
            '`nifdeclarado`,`nifrepresante`,`razosoci`,
            Linea = DBLet(Rs!razosoci, "T")
            If Linea = "" Then Linea = Rs!Nommacta 'por si acaso no tiene razon social
            
            Linea = Devnombresql340(Linea)
            
            
            Linea = "'" & DBLet(Rs!nifdatos, "T") & "',NULL,'" & DevNombreSQL(Linea) & "',"
            
                
            '`codpais`,`idenpais`,`nifresidencia`
            PAIS = UCase(DBLet(Rs!codPAIS, "T"))
            If PAIS = "" Then PAIS = "ESPAÑA"
            
            If PAIS = "ESPAÑA" Then
                Linea = Linea & "'ES','1',NULL"
                Identificacion = 0
            Else
                'ESTRANJERO o INTRACOM
                Linea = Linea & "'" & Mid(PAIS, 1, 2) & "',2,'" & DBLet(Rs!nifdatos, "T") & "'"
                Identificacion = 1
            End If
            Linea = Linea & ",'"
            '`clavelibro`,`claveoperacion`,
'            NF = 1
'            If Not IsNull(Rs!tp2facpr) Then NF = NF + 1
'            If Not IsNull(Rs!tp3facpr) Then NF = NF + 1

            SqlNew = "select count(*) from factpro_totales where numserie = " & DBSet(Rs!NumSerie, "T")
            SqlNew = SqlNew & " and numregis = " & DBSet(Rs!NumRegis, "N")
            SqlNew = SqlNew & " and anofactu = " & DBSet(Rs!anofactu, "N")
            
            NF = TotalRegistros(SqlNew)

            If NF = 1 Then
                'SOLO HAY UNA base
                PAIS = " "
            Else
                PAIS = "C" 'mas de una base
            End If
            'IMPORTANTE.
            'Hemos cargado en IvaBienesInversion el codigo de IVA que este marcado como Bien de inversion.
            'Si el IVA de la factura esta dentro de este tipo de via, marcaremos esta factura como BI
            If InStr(1, IvaBienesInversion, "|" & Rs!codigiva & "|") > 0 Then
                IVA_BI = True
                Linea = Linea & "I"
            Else
                IVA_BI = False
'                If InStr(1, IvaREA, "|" & CStr(Rs!codigiva) & "|") > 0 Then
                If DBLet(Rs!codopera) = 5 Then
                    Linea = Linea & "X"
                Else
                    Linea = Linea & "R"
                End If
            End If
            
            Linea = Linea & "','" & PAIS & "',"
            Linea = Linea & "'" & Format(Rs!fecharec, FormatoFecha) & "','" & Format(Rs!fecharec, FormatoFecha) & "'"
           
            '`idfactura`
            Linea = Linea & ",'" & Rs!NumFactu & "'"
           
           
           'Numero registro
           Linea = Linea & ",'" & Format(Rs!NumRegis, "00000000") & "',"
           
           
           'rectifica,dom_intracom,pob_intracom,cp_intracom,
           'rectifica,dom_intracom,pob_intracom,cp_intracom,"
           If Rs!totfacpr < 0 Then
                'Rectifica
                Linea = Linea & "'" & DevNombreSQL(DBLet(Rs!observa, "T")) & "'"
           Else
                'NULL
                Linea = Linea & "NULL"
           End If
           
           'If Identificacion = 1 Then
           Linea = Linea & ",NULL,NULL,NULL,"
           
           
           
           'Van los importes
           'numreg,`numiva`,`tipo`,`base`,`totiva`,
           '`ivare`,`cuotare`
           
           'Si es bien de INVERSION NO la meto
           If IVA_BI Then
                'NO hacemos nada
                
                
           Else
                    'Base UNO. SIEMPRE EXISTE
                    TotalLin = Rs!Impoiva + Rs!Baseimpo + DBLet(Rs!ImpoRec, "N")
                    PAIS = NF & "," & TransformaComasPuntos(CStr(Rs!porciva)) & "," & TransformaComasPuntos(CStr(Rs!Baseimpo)) _
                         & "," & TransformaComasPuntos(CStr(Rs!Impoiva)) & "," & TransformaComasPuntos(CStr(TotalLin))
                    PAIS = PAIS & "," & TransformaComasPuntos(CStr(DBLet(Rs!porcrec, "N"))) & "," & TransformaComasPuntos(CStr(DBLet(Rs!ImpoRec, "N"))) & ")"
                     
                    'Insertar
                    B = True
                    If QuitarREA Then
                        'Si hay que quitar REA, y este iva es REA no insertamos
                        'If InStr(1, IvaREA, "|" & CStr(Rs!codigiva) & "|") > 0 Then B = False
                        If DBLet(Rs!codopera) = 5 Then B = False
                    End If
                    
                    If B Then
                         CadenaInsert = CadenaInsert & ",(" & vUsu.Codigo & "," & NumRegElim & "," & Linea & PAIS
                         NumRegElim = NumRegElim + 1
                     End If
                     
                     
'                    'Base DOS si no es null
'                    If Not IsNull(Rs!ba2facpr) Then
'                         TotalLin = Rs!ti2facpr + Rs!ba2facpr + DBLet(Rs!tr2facpr, "N")
'                         PAIS = NF & "," & TransformaComasPuntos(CStr(Rs!pi2facpr)) & "," & TransformaComasPuntos(CStr(Rs!ba2facpr)) _
'                         & "," & TransformaComasPuntos(CStr(Rs!ti2facpr)) & "," & TransformaComasPuntos(CStr(TotalLin))
'                         PAIS = PAIS & "," & TransformaComasPuntos(CStr(DBLet(Rs!pr2facpr, "N"))) & "," & TransformaComasPuntos(CStr(DBLet(Rs!tr2facpr, "N"))) & ")"
'
'                         'Insertar
'                         CadenaInsert = CadenaInsert & ",(" & vUsu.Codigo & "," & NumRegElim & "," & Linea & PAIS
'                         NumRegElim = NumRegElim + 1
'                    End If
'
'
'                    'Base TRES si no es null
'                    If Not IsNull(Rs!ba3facpr) Then
'                         TotalLin = Rs!ti3facpr + Rs!ba3facpr + DBLet(Rs!tr3facpr, "N")
'                         PAIS = NF & "," & TransformaComasPuntos(CStr(Rs!pi3facpr)) & "," & TransformaComasPuntos(CStr(Rs!ba3facpr)) _
'                         & "," & TransformaComasPuntos(CStr(Rs!ti3facpr)) & "," & TransformaComasPuntos(CStr(TotalLin))
'                         PAIS = PAIS & "," & TransformaComasPuntos(CStr(DBLet(Rs!pr3facpr, "N"))) & "," & TransformaComasPuntos(CStr(DBLet(Rs!tr3facpr, "N"))) & ")"
'
'                         'Insertar
'                         CadenaInsert = CadenaInsert & ",(" & vUsu.Codigo & "," & NumRegElim & "," & Linea & PAIS
'                         NumRegElim = NumRegElim + 1
'                    End If
                    
            End If 'de IVA_BI
            
                        
            
            
            'HACEMOS EL INSERT
            If Len(CadenaInsert) > 100000 Then
                 PAIS = DevuelveInsertTmp340(0)
                 PAIS = PAIS & Mid(CadenaInsert, 2)
                 Conn.Execute PAIS
                 CadenaInsert = ""
            End If
            
           
           Rs.MoveNext
        Wend
        Rs.Close
           
        If CadenaInsert <> "" Then
            PAIS = DevuelveInsertTmp340(0)
            PAIS = PAIS & Mid(CadenaInsert, 2)
            Conn.Execute PAIS
        End If
End Sub



'Siempre llmaran a la funcion.
'Lo primero que hace es borrar las de bien de inversion para cargarlas todas
'
Private Sub CargaFacturasRecibidasBienInversion(NumeroEmpresa As Byte, Anyo As Integer)
'FraPro_MismaFecha
Dim PAIS As String
Dim CadenaInsert As String
Dim IvaBienesInversion2 As String   'Para añadir al where
Dim Identificacion As Byte
Dim TotalLin As Currency
Dim SqlNew As String
    
    'Cargare los IVAS que seand e bienes de inversion
    Linea = "Select * from ariconta" & NumeroEmpresa & ".tiposiva where tipodiva = 2"
    Rs.Open Linea, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    IvaBienesInversion2 = ""
    While Not Rs.EOF
        IvaBienesInversion2 = IvaBienesInversion2 & ", " & Rs!codigiva
        Rs.MoveNext
    Wend
    Rs.Close
    
    
    'No tiene
    If IvaBienesInversion2 = "" Then Exit Sub
    
    IvaBienesInversion2 = Mid(IvaBienesInversion2, 2)
    IvaBienesInversion2 = " AND factpro_totales.codigiva IN (" & IvaBienesInversion2 & ")"
    Linea = CadenaSelect340(False, False, 1, Anyo, True)
    IvaBienesInversion2 = Linea & IvaBienesInversion2
    
    
    Linea = "select factpro.*,cuentas.razosoci, factpro_totales.numlinea, factpro_totales.baseimpo, factpro_totales.codigiva, factpro_totales.porciva, factpro_totales.porcrec, factpro_totales.impoiva, factpro_totales.imporec"
    Linea = Linea & " from ariconta" & NumeroEmpresa & ".factpro factpro,"
    Linea = Linea & " ariconta" & NumeroEmpresa & ".cuentas cuentas,  "
    Linea = Linea & " ariconta" & NumeroEmpresa & ".factpro_totales factpro_totales  "
    Linea = Linea & " where factpro.codmacta=cuentas.codmacta AND " & IvaBienesInversion2
    Linea = Linea & " and factpro.numserie = factpro_totales.numserie and factpro.numregis = factpro_totales.numregis "
    Linea = Linea & " and factpro.anofactu = factpro_totales.anofactu "
    Linea = Linea & " ORDER BY fecharec"
    'FALTA el periodod
    
    CadenaInsert = ""
    Rs.Open Linea, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
            'Para cada factura si tiene varias bases, el trzo sera igual menos el importe final(de la cuto % y totaliva)
                    
    
            'Campos en la tmp
            '`tmp340` (`codusu`,`codigo`,`nifdeclarado`,`nifrepresante`,`razosoci`,
            '`codpais`,`idenpais`,`nifresidencia`,
            '`clavelibro`,`claveoperacion`,
            '`fechaexp`,`fechaop`,`idfactura`,`totalfac`,
            'numreg`,`tipo`,`base`,`totiva`,
            '`,`numiva`,`ivare`,`cuotare`
                        
            '`nifdeclarado`,`nifrepresante`,`razosoci`,
            Linea = DBLet(Rs!razosoci, "T")
            If Linea = "" Then Linea = Rs!Nommacta 'por si acaso no tiene razon social
            
            Linea = Devnombresql340(Linea)
            
            
            Linea = "'" & DBLet(Rs!nifdatos, "T") & "',NULL,'" & DevNombreSQL(Linea) & "',"
            
                
            '`codpais`,`idenpais`,`nifresidencia`
            PAIS = UCase(DBLet(Rs!codPAIS, "T"))
            If PAIS = "" Then PAIS = "ESPAÑA"
            
            If PAIS = "ESPAÑA" Then
                Linea = Linea & "'ES','1',NULL"
                Identificacion = 0
            Else
                'ESTRANJERO o INTRACOM
                Linea = Linea & "'" & Mid(PAIS, 1, 2) & "',2,'" & DBLet(Rs!nifdatos, "T") & "'"
                Identificacion = 1
            End If
            Linea = Linea & ",'"
            '`clavelibro`,`claveoperacion`,
'            NF = 1
'            If Not IsNull(Rs!tp2facpr) Then NF = NF + 1
'            If Not IsNull(Rs!tp3facpr) Then NF = NF + 1
            SqlNew = "select count(*) from factpro_totales where numserie = " & DBSet(Rs!NumSerie, "T")
            SqlNew = SqlNew & " and numregis = " & DBSet(Rs!NumRegis, "N")
            SqlNew = SqlNew & " and anofactu = " & DBSet(Rs!anofactu, "N")
            
            NF = TotalRegistros(SqlNew)


            If NF = 1 Then
                'SOLO HAY UNA base
                PAIS = " "
            Else
                PAIS = "C" 'mas de una base
            End If
            'IMPORTANTE.
            'Solo estamos cargando Bienes de inversion
            
            
            Linea = Linea & "I"
        
            
            Linea = Linea & "','" & PAIS & "',"
            '`fechaexp`,`fechaop`,
            'Graba en los dos campos la misma fecha m es decir, la de RECEPCION
            Linea = Linea & "'" & Format(Rs!fecharec, FormatoFecha) & "','" & Format(Rs!fecharec, FormatoFecha) & "'"
            '`idfactura`
            Linea = Linea & ",'" & Rs!NumFactu & "'"
           
           
           'Numero registro
           Linea = Linea & ",'" & Format(Rs!NumRegis, "00000000") & "',"
           
           
           'rectifica,dom_intracom,pob_intracom,cp_intracom,
           'rectifica,dom_intracom,pob_intracom,cp_intracom,"
           If Rs!totfacpr < 0 Then
                'Rectifica
                Linea = Linea & "'" & DevNombreSQL(DBLet(Rs!observa, "T")) & "'"
           Else
                'NULL
                Linea = Linea & "NULL"
           End If
           
           'If Identificacion = 1 Then
           Linea = Linea & ",NULL,NULL,NULL,"
           
           
           
           'Van los importes
           'numreg,`numiva`,`tipo`,`base`,`totiva`,
           '`ivare`,`cuotare`
           
           'Base UNO. SIEMPRE EXISTE
           TotalLin = Rs!Impoiva + Rs!Baseimpo + DBLet(Rs!ImpoRec, "N")
           PAIS = NF & "," & TransformaComasPuntos(CStr(Rs!porciva)) & "," & TransformaComasPuntos(CStr(Rs!Baseimpo)) _
                & "," & TransformaComasPuntos(CStr(Rs!Impoiva)) & "," & TransformaComasPuntos(CStr(TotalLin))
           PAIS = PAIS & "," & TransformaComasPuntos(CStr(DBLet(Rs!porcrec, "N"))) & "," & TransformaComasPuntos(CStr(DBLet(Rs!ImpoRec, "N"))) & ")"
            
           
            CadenaInsert = CadenaInsert & ",(" & vUsu.Codigo & "," & NumRegElim & "," & Linea & PAIS
            NumRegElim = NumRegElim + 1

           
'           'Base DOS si no es null
'           If Not IsNull(Rs!ba2facpr) Then
'                TotalLin = Rs!ti2facpr + Rs!ba2facpr + DBLet(Rs!tr2facpr, "N")
'                PAIS = NF & "," & TransformaComasPuntos(CStr(Rs!pi2facpr)) & "," & TransformaComasPuntos(CStr(Rs!ba2facpr)) _
'                & "," & TransformaComasPuntos(CStr(Rs!ti2facpr)) & "," & TransformaComasPuntos(CStr(TotalLin))
'                PAIS = PAIS & "," & TransformaComasPuntos(CStr(DBLet(Rs!pr2facpr, "N"))) & "," & TransformaComasPuntos(CStr(DBLet(Rs!tr2facpr, "N"))) & ")"
'
'                'Insertar
'                CadenaInsert = CadenaInsert & ",(" & vUsu.Codigo & "," & NumRegElim & "," & Linea & PAIS
'                NumRegElim = NumRegElim + 1
'           End If
'
'
'           'Base TRES si no es null
'           If Not IsNull(Rs!ba3facpr) Then
'                TotalLin = Rs!ti3facpr + Rs!ba3facpr + DBLet(Rs!tr3facpr, "N")
'                PAIS = NF & "," & TransformaComasPuntos(CStr(Rs!pi3facpr)) & "," & TransformaComasPuntos(CStr(Rs!ba3facpr)) _
'                & "," & TransformaComasPuntos(CStr(Rs!ti3facpr)) & "," & TransformaComasPuntos(CStr(TotalLin))
'                PAIS = PAIS & "," & TransformaComasPuntos(CStr(DBLet(Rs!pr3facpr, "N"))) & "," & TransformaComasPuntos(CStr(DBLet(Rs!tr3facpr, "N"))) & ")"
'
'                'Insertar
'                CadenaInsert = CadenaInsert & ",(" & vUsu.Codigo & "," & NumRegElim & "," & Linea & PAIS
'                NumRegElim = NumRegElim + 1
'           End If
'
           
           
                       
           
           
           'HACEMOS EL INSERT
           If Len(CadenaInsert) > 100000 Then
                PAIS = DevuelveInsertTmp340(0)
                PAIS = PAIS & Mid(CadenaInsert, 2)
                Conn.Execute PAIS
                CadenaInsert = ""
           End If
           
           
           Rs.MoveNext
        Wend
        Rs.Close
           
        If CadenaInsert <> "" Then
            PAIS = DevuelveInsertTmp340(0)
            PAIS = PAIS & Mid(CadenaInsert, 2)
            Conn.Execute PAIS
        End If
End Sub








Private Function CadenaSelect340(Clientes As Boolean, Intracomunitarias As Boolean, Periodo As Byte, Ano As Integer, paraBienDeInversion As Boolean) As String
Dim FI As Date
Dim FF As Date
Dim I As Integer
    If vParam.periodos = 0 Then
      'Liquidacion TRIMESTRAL
      I = (Periodo - 1) * 3 + 1
      FI = CDate("01/" & I & "/" & Ano)
      I = I + 2 'Fin trimestre
      FF = CDate(DiasMes(CByte(I), Ano) & "/" & I & "/" & Ano)
    Else
      I = Periodo
      FI = CDate("01/" & I & "/" & Ano)
      FF = CDate(DiasMes(CByte(I), Ano) & "/" & I & "/" & Ano)
    End If
    
    If Clientes Then
        'FRACLI normal
        'CadenaSelect340 = " intracom = " & Abs(Intracomunitarias)
        CadenaSelect340 = " 1 = 1"
        CadenaSelect340 = CadenaSelect340 & " AND fecliqcl >= '" & Format(FI, FormatoFecha) & "' AND fecliqcl <= '" & Format(FF, FormatoFecha) & "'"
    Else
        'PROVEEDIRES
        'CadenaSelect340 = " extranje = " & Abs(Intracomunitarias)
        CadenaSelect340 = " 1 = 1"
        If paraBienDeInversion Then
            'Bien de inversion.
            'Todo el año
            FI = "01/01/" & Ano
            FF = "31/12/" & Ano
        End If
        
        '
        CadenaSelect340 = CadenaSelect340 & " AND fecliqpr >= '" & Format(FI, FormatoFecha) & "' AND fecliqpr <= '" & Format(FF, FormatoFecha) & "'"

    End If
End Function


Private Function DevuelveInsertTmp340(Tipo As Byte) As String

    If Tipo = 0 Then
        DevuelveInsertTmp340 = "INSERT INTO tmp340 (`codusu`,`codigo`,`nifdeclarado`,`nifrepresante`,`razosoci`,`codpais`,`idenpais`,`nifresidencia`,"
        DevuelveInsertTmp340 = DevuelveInsertTmp340 & "`clavelibro`,`claveoperacion`,`fechaexp`,`fechaop`,`idfactura`,"
        DevuelveInsertTmp340 = DevuelveInsertTmp340 & "rectifica,dom_intracom,pob_intracom,cp_intracom,"
        DevuelveInsertTmp340 = DevuelveInsertTmp340 & "numreg,numiva,tipo,`base`,`totiva`,`totalfac`,`ivare`,`cuotare`) VALUES "
    End If
End Function


Private Function Devnombresql340(CADENA As String) As String
    'QUitaremos algunos caracteres NO validos para el modelo 340
    'Ejmplo: ª º
    
    Dim I As Integer
    Dim Aux As String
    Dim C As String
    
    Aux = ""
    For I = 1 To Len(CADENA)
        C = Mid(CADENA, I, 1)
        
        If C = "º" Or C = "ª" Then
            C = "."
        Else
            If C = "Ñ" Then
                C = Chr(209)
            ElseIf C = "Ç" Or C = "ç" Then
                C = Chr(199)
            End If
        End If
        Aux = Aux & C
    Next I
    
    Devnombresql340 = Aux

End Function

Public Function GeneraFichero340(PresentaInternet As Boolean, anoPeriodo As String, UtlPeriodoLiquidacion As Boolean) As Boolean
    
    On Error GoTo EGeneraFichero340
    GeneraFichero340 = False
    
    
    Set Rs = New ADODB.Recordset
    
    Linea = App.Path & "\tmp340.dat"
    If Dir(Linea, vbArchive) <> "" Then Kill Linea
    
    NF = FreeFile
    Open Linea For Output As NF
    
    
 '   GrabaRegistros340Pandata
  
    
    'Grabaremos el fichero de cabecera
    IdentificacionPresentador = "340" & Mid(anoPeriodo, 1, 4)
    GrabaCabecera340 PresentaInternet, anoPeriodo, UtlPeriodoLiquidacion 'Le añadire el nif a la cabcera
    
    GrabaRegistros340
    
    Close #NF
    
    
    
    
    
    GeneraFichero340 = True
    Set Rs = Nothing
    Exit Function
EGeneraFichero340:
    Linea = Err.Description
    Linea = "Generando fichero" & vbCrLf & Linea
    MuestraError Err.Number, Linea
    Set Rs = Nothing
    On Error Resume Next
    Close #NF
    Err.Clear
    On Error GoTo 0
End Function



'anoPeriodo:  Llevara  yyyymmpp
'                      año
'                          mes,  cuando sea por trimestres: marzo er trimestr  jun 2º trimes ...
'                             pp period 1..12  o 1T 2T ....
Private Sub GrabaCabecera340(vPresentaInternet As Boolean, anoPeriodo As String, UtlimoPeriodoPresentacion As Boolean)
Dim vAux As String

    Linea = "Select * from empresa2"
    Rs.Open Linea, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Linea = DatosTexto(Rs!nifempre, 9)
    IdentificacionPresentador = IdentificacionPresentador & Linea   'Este sera el identificador para todos los tipos de regstro
    
    
    
    'Empzamos a generar
    Linea = "1" & IdentificacionPresentador
    Linea = Linea & DatosTexto(DBLet(vEmpresa.nomempre, "T"), 40)
    If vPresentaInternet Then
        Linea = Linea & "T"  'telematica
    Else
        Linea = Linea & "C"  'CD
    End If
    Linea = Linea & DatosTexto(DBLet(Rs!telefono, "T"), 9)
    Linea = Linea & DatosTexto(DBLet(Rs!contacto, "T"), 40)
    
    
    'Cierro el RS
    Rs.Close
    
    'nn1 declar  340EEEEPPSSSS
    Linea = Linea & "340" & Mid(anoPeriodo, 1, 6) & "0001"
    
    'Campos que no relleanmos
    'dec compen,sustitutiva,nº declar anteriro
    Linea = Linea & " " & " " & String(13, "0")
    
    'Periodo
    Linea = Linea & Mid(anoPeriodo, 7, 2)
    

    
    
    'Numero total de registros
    Rs.Open "Select count(*) from tmp340 where codusu = " & vUsu.Codigo, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    If Not Rs.EOF Then NumRegElim = DBLet(Rs.Fields(0), "N")
    Rs.Close
    Linea = Linea & Format(NumRegElim, "000000000")
    
    
    
    
    
    'La suma del campo totalfac
    Rs.Open "select sum(base) , sum(totiva),sum(totalfac) from tmp340 where codusu = " & vUsu.Codigo, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Linea = Linea & DatosNumeroDec340(Rs.Fields(0), 18) & DatosNumeroDec340(Rs.Fields(1), 18)
    'La suma de facturas puede ser problema ya que para una misma factura pueden haber varias lineas
    Linea = Linea & DatosNumeroDec340(Rs.Fields(2), 18)
    
    
    
    'Cierro el RS
    Rs.Close
    
    I = 500 - Len(Linea)
    Linea = Linea & Space(I)
    Print #NF, Linea
    
End Sub



Private Sub GrabaRegistros340()
    
    Linea = "Select * from tmp340 where codusu = " & vUsu.Codigo
    Linea = Linea & " ORDER BY clavelibro,numreg,idfactura"
    Rs.Open Linea, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs.EOF
        'Casi todos los encabezados s
        Linea = "2" & IdentificacionPresentador
         
        'Nif declarado  y repres legal
        'si es intracom no pinto este FALTA###
        If Val(DBLet(Rs!idenpais, "N")) > 1 Then
            'ES EXTRANJERO
            Linea = Linea & String(9, " ")
        Else
            Linea = Linea & DatosTexto(Rs!nifdeclarado, 9)
        End If
        'Representatne
        Linea = Linea & DatosTexto(DBLet(Rs!nifrepresante), 9)
        'Razosoci
        Linea = Linea & DatosTexto(Rs!razosoci, 40)
        'Pais,idenpais,niresidencia                 'Dira que es el documento  del id en pais de referencia (nif, passporte...)
        Linea = Linea & DatosTexto(Rs!codPAIS, 2)
        Linea = Linea & DBLet(Rs!idenpais, "N")
        If Val(DBLet(Rs!idenpais, "N")) > 1 Then
            Linea = Linea & DatosTexto(Rs!codPAIS, 2) & DatosTexto(DBLet(Rs!nifresidencia), 18)
        Else
            Linea = Linea & String(20, " ")
        End If
        'Clave libro, claveoperacion
        If Rs!clavelibro = "X" Then
            Linea = Linea & "R"  'las REA son recibidas
            Linea = Linea & "X" 'rea
        Else
            If Rs!clavelibro = "Z" Then
                'Son efectivos. Parece ser que hay que pintar una E
                Linea = Linea & "E"
                Linea = Linea & " "
                
            Else
                Linea = Linea & Rs!clavelibro
                Linea = Linea & DatosTexto(DBLet(Rs!claveoperacion), 1)
            End If
        End If
        
        
        
       
        'Tipo impositivo, base imponible  base totiva totalfac
        'El % iva siempre es positivo y no tiene espacio para el sigono
        If Rs!clavelibro = "X" Then
             'Fecha expedicion, fecha operacion
            Linea = Linea & Format(Rs!fechaexp, "yyyymmdd") & Format(Rs!fechaop, "yyyymmdd")
            'fEBRERO 2012
            'El IVA REA hay que ponerlo asin..  'TIpo impositov=00000   BI total fra   total fra=totalfra
            'LINEA = LINEA & "00000" & DatosNumeroDec340(RS!totalfac, 14) & DatosNumeroDec340(0, 14) & DatosNumeroDec340(RS!totalfac, 14)
            'Febreo 2012, mas tarde. Opcion Llutxent
            Linea = Linea & "00000" & DatosNumeroDec340(Rs!Base, 14) & DatosNumeroDec340(Rs!totiva, 14) & DatosNumeroDec340(Rs!TotalFac, 14)
            
        Else
            If Rs!clavelibro = "Z" Then
                
                 'Fecha expedicion, fecha operacion
                Linea = Linea & Format(Rs!fechaexp, "yyyymmdd") & "00000000"
                Linea = Linea & "00000" & DatosNumeroDec340(0, 14) & DatosNumeroDec340(0, 14) & DatosNumeroDec340(Rs!TotalFac, 14)
            Else
                 'Fecha expedicion, fecha operacion
                Linea = Linea & Format(Rs!fechaexp, "yyyymmdd") & Format(Rs!fechaop, "yyyymmdd")
                Linea = Linea & DatosNumeroDec(Rs!Tipo, 5) & DatosNumeroDec340(Rs!Base, 14) & DatosNumeroDec340(Rs!totiva, 14) & DatosNumeroDec340(Rs!TotalFac, 14)
            End If
        End If
         'Base imponible a cuoste
        Linea = Linea & " " & String(13, "0")
        
        '--------------------------------------
        If Rs!clavelibro = "Z" Then
            'Para los efectivos NO pinto nada
            Linea = Linea & Space(58)
        Else
            'Para el resto
            '-------------------
           
            'Identificacion de la factura  y numero registro
            Linea = Linea & DatosTexto(Rs!idfactura, 40) & DatosTexto(DBLet(Rs!NumReg, "N"), 18)
        End If
        
        
        'Aqui llegamos con un len(linea) de 235
        
        'Hasta aqui trozo comun
        'AHora segun sea el registro
        Select Case Rs!clavelibro
        Case "E"
                'EMITIDAS o EXPEDIDAS
                '-------------------------------------
                Linea = Linea & String(7, "0") & "1"   'nº facturas
                'Si es ticket pongo uno, sea cual sea el numero de ivas
                If CStr(Rs!claveoperacion) = "J" Then
                    Linea = Linea & "01" 'Desglose de ivas
                Else
                    Linea = Linea & Format(Rs!numiva, "00") 'Desglose de ivas
                End If
                Linea = Linea & String(80, " ")   'intervalo  id facturas
                Linea = Linea & String(40, " ")   'Factura rectificada
                
                'Recargo de equivalencia
                Linea = Linea & DatosNumeroDec(Rs!ivare, 5) & DatosNumeroDec340(Rs!cuotare, 14)
                
                Linea = Linea & "0" 'situ inmueble
                Linea = Linea & Space(25) 'catastral
                Linea = Linea & DatosNumeroDec(0, 15)
                Linea = Linea & "0000"
                Linea = Linea & DatosNumeroDec(0, 15)
                
                
        Case "R", "X"
                'RECIBIDAS
                '----------------------------------------------
          
               Linea = Linea & String(17, "0") & "1"   'nº facturas
               Linea = Linea & Format(Rs!numiva, "00")
               Linea = Linea & String(80, " ")   'intervalo  id facturas
               Linea = Linea & DatosNumeroDec340(Rs!cuotare, 14)
        Case "I"
                'Bien de INVERSION. En la poscion
                'MsgBox "Bienes de inversion UNICAMENTE en ultimo periodo liquidacion", vbExclamation
                
              
                        
               Linea = Linea & "001"  'prorrata definitiva (me lo Dijo Picassent)
               Linea = Linea & DatosNumeroDec340(0, 14) 'regulariz anual deduccion
               Linea = Linea & String(40, " ")   'id entrega
               Linea = Linea & DatosNumeroDec340(0, 14) 'regulariz de la deduc efectuada
               Linea = Linea & Format(Rs!fechaop, "yyyymmdd")
               Linea = Linea & DatosTexto(Rs!idfactura, 17)
        Case "Z"
                'Efectivos que superan una cantidad y haya algun pago en el periodo
                
                Linea = Linea & String(7, "0") & "0"   'nº facturas
                'Si es ticket pongo uno, sea cual sea el numero de ivas
                Linea = Linea & "00" 'Desglose de ivas
                Linea = Linea & String(80, " ")   'intervalo  id facturas
                Linea = Linea & String(40, " ")   'Factura rectificada
                
                'Recargo de equivalencia
                Linea = Linea & DatosNumeroDec(0, 5) & DatosNumeroDec340(0, 14)
                
                Linea = Linea & "0" 'situ inmueble
                Linea = Linea & Space(25) 'catastral
                ''importe en metalico percibido
                'Se declara lo anterior y el periodo. El periodo SEGURo que tiene movimientos
                Linea = Linea & DatosNumeroDec(DBLet(Rs!Base, "N") + Rs!totiva, 15)
                'Ejercicio del pago en la cadena IdentificacionPresentador la posicion 4,5,6,7 son el anño
                Linea = Linea & Mid(IdentificacionPresentador, 4, 4)
                Linea = Linea & DatosNumeroDec(0, 15)
        End Select
        
        
        'Trozo final, comun tambien
        I = 500 - Len(Linea)
        Linea = Linea & Space(I)
        

        
        'Los bienes de inversion seran
        'If RS!clavelibro <> "I" Then Print #NF, LINEA
        Print #NF, Linea
        
        'Sig
        Rs.MoveNext
        
    Wend
    Rs.Close
    
        
End Sub



'Como el preoceso es multiempresa(consolidado)
'enviaremos el minimo y maximo para no hacer la comrpoabacion de los NIFS cada vez
'Ej.  Si la empresa A va de 1 a 10   y la B del 11 al 16,
'cuando compreube la B no tengo que empezar en el 1 si no en el 11
Private Function ComprobarNifs340(Minimo As Long, maximo As Long) As Integer
    On Error GoTo EComprobarNifs340
    ComprobarNifs340 = -1
    
    Linea = "select nifdeclarado,codpais,clavelibro,idfactura,fechaexp from tmp340 where codusu = " & vUsu.Codigo
    Linea = Linea & " AND  codigo>=" & Minimo & " and codigo <= " & maximo
    Linea = Linea & " group by nifdeclarado"
    Rs.Open Linea, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Linea = ""
    While Not Rs.EOF
        If Rs!codPAIS = "ES" Then
            'ESPAÑA. Comprobamos NIF"
            If Not Comprobar_NIF(Rs!nifdeclarado) Then
                Linea = Linea & DevFacturasTmp340DeEseNIF(Rs!nifdeclarado)
                ComprobarNifs340 = ComprobarNifs340 + 1
            End If
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    
    'Empezaba en menos1
    ComprobarNifs340 = ComprobarNifs340 + 1
    Exit Function
EComprobarNifs340:
    MuestraError Err.Number, Err.Description
    Linea = ""
End Function



Private Function DevFacturasTmp340DeEseNIF(vNif As String) As String
Dim C As String
Dim RT As ADODB.Recordset
    C = "select clavelibro,idfactura,fechaexp from tmp340 where nifdeclarado='" & vNif & "'"
    C = C & " AND codusu = " & vUsu.Codigo
    Set RT = New ADODB.Recordset
    RT.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    C = "- " & Rs!nifdeclarado & "  " & vbCrLf
    While Not RT.EOF
        C = C & Space(13) & RT!clavelibro & "   " & Format(RT!fechaexp, "dd/mm/yyyy") & "  " & RT!idfactura & vbCrLf
        RT.MoveNext
    Wend
    RT.Close
    DevFacturasTmp340DeEseNIF = C
    
End Function



'Hayque pasarlo a bus.bas
Public Sub LanzaNotepad(ByVal NombreFichero As String, Descripcion As String)
    On Error Resume Next
    'Solo sirve para los archivos
    Shell "notepad " & NombreFichero, vbNormalFocus
    If Err.Number <> 0 Then
        NombreFichero = "Imposible mostrar archivo para: " & Descripcion & vbCrLf & "Fichero: " & NombreFichero
        MsgBox NombreFichero, vbExclamation
        Err.Clear
    End If
End Sub



'Concetos.>  ByVal pq SQl la reutilziare
Private Sub CargarPagosEnEfectivo(NConta As Byte, Importemax As Currency, ByVal Conceptos As String, Periodo As Byte, Anyo As Integer, ByRef Lb As Label)
Dim Fin As Boolean
Dim CtaLeida As Boolean
Dim Cta As String
Dim Sql As String
Dim I As Integer
Dim ImporteAnt As Currency
Dim ImportePer As Currency
Dim RCIF As ADODB.Recordset
Dim CadInsert As String
Dim F2 As Date

    Lb.Caption = "Leyendo datos en efectivo"
    Lb.Refresh

    'Sacaremos aquellos pagos
    'Efectuados en concepto efectivo. Los de principio de año en un sum y los del peridodo en otro
    'En el select habra un registro para cada cuenta
    'el periodo anterior sera en el Field(1) un 0
    
    If vParam.periodos = 0 Then
        'trimestral
        I = (Periodo - 1) * 3 + 1
    Else
        'mensual
        I = Periodo
    End If
    Sql = "select codmacta,if (month(fechaent)<" & I & ",0,1) Periodo"
    'HABER -DEBE
    Sql = Sql & ",sum(if(timporteh is null,0,timporteh))-sum(if(timported is null,0,timported)) importe"
    Sql = Sql & ",max(fechaent) ultimafecha" 'ultima fecha cobro efectuado
    Sql = Sql & " from ariconta" & NConta & ".hlinapu where "
    'Desde fecha incio ejercicio hasta el ultimo dia del periodo
    Sql = Sql & " fechaent>= '" & Anyo & "-01-01' AND fechaent <='"
    If vParam.periodos = 0 Then
      'Liquidacion TRIMESTRAL
      I = (Periodo * 3)
      F2 = CDate(DiasMes(CByte(I), Anyo) & "/" & I & "/" & Anyo)
    Else
      I = Periodo
      F2 = CDate(DiasMes(CByte(I), Anyo) & "/" & I & "/" & Anyo)
    End If
    Sql = Sql & Format(F2, FormatoFecha) & "'"
    Sql = Sql & " AND codconce IN (" & Conceptos & ")"
    

    
    Sql = Sql & " group by 1,2 order by 1,2"

    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Fin = Rs.EOF
   
   
    'Leemos todas las cuentas que tengan CIF
    Set RCIF = New ADODB.Recordset
    Sql = "Select codmacta,razosoci,nifdatos,pais,nommacta from ariconta" & NConta & ".cuentas where apudirec='S' and nifdatos<>''"
    Sql = Sql & " and model347=1" 'par evitar sacar los de varios
    RCIF.Open Sql, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    CadInsert = ""
    CtaLeida = False
    While Not Fin
            Lb.Caption = Rs!codmacta
            Lb.Refresh

            
    
            ImporteAnt = 0
            ImportePer = 0
            Sql = Rs!codmacta
            I = Val(Rs!Periodo)
            F2 = Rs!ultimafecha
            If I = 0 Then
                'Es el anterior
                'pruebo a ver si el siguiente registro el del periodo actual
                ImporteAnt = Rs!Importe
                CtaLeida = True
                Rs.MoveNext
            Else
                'No tienen periodo anterior
                
                'No hacemos  nada y entrara en el siguiente IF
                ImportePer = Rs!Importe
                CtaLeida = False
                Rs.MoveNext
            End If
           
                
            If CtaLeida Then
                If Not Rs.EOF Then
                    If Rs!codmacta = Sql Then
                        'OK es la misma cuenta
                        ImportePer = Rs!Importe
                        
                        Rs.MoveNext
                    Else
                        'NO, es otra cuenta
                        'NO hago nada
                        
                    End If
                End If
            End If
    
            'Como ppuede consolidar lo ponemos insertamos si o si
            If ImporteAnt + ImportePer <> 0 Then
                'Ziiiiiiiii, lo supera
                'OK este deberia entrar
                
                Sql = "codmacta ='" & Sql & "'"
                RCIF.Find Sql, , adSearchForward, 1
                If Not RCIF.EOF Then
                    'Existe la cta y tiene nif
                    CadenaInsertDeclaracionesEfectivo RCIF, ImporteAnt, ImportePer, F2
                    CadInsert = CadInsert & ", (" & Linea & ")"
                    
                    If Len(CadInsert) > 60000 Then
                        Lb.Caption = "Insert efectivo"
                        Lb.Refresh

                        DoEvents
                        'INSERTAMOS
                        InsertaMultipleTMP340 CadInsert
                        
                        'A cero otra vez
                        CadInsert = ""
                    End If
                    
                End If
            End If
            Fin = Rs.EOF
    Wend
    Rs.Close
    InsertaMultipleTMP340 CadInsert

    Set RCIF = Nothing
End Sub

'REULTILIZO LINEA
Private Sub InsertaMultipleTMP340(Cad As String)

    If Cad = "" Then Exit Sub
    
    Cad = Mid(Cad, 2) 'quitamos la primera coma
    Linea = DevuelveInsertTmp340(0)
    Linea = Linea & Cad
    
    
    Conn.Execute Linea

End Sub



Private Sub CadenaInsertDeclaracionesEfectivo(ByRef R As ADODB.Recordset, ImporAnt As Currency, ImporPer As Currency, FechaUltimo As Date)
Dim Aux As String
    
    
            
            '`nifdeclarado`,`nifrepresante`,`razosoci`,
            Linea = DBLet(R!razosoci, "T")
            If Linea = "" Then Linea = R!Nommacta 'por si acaso no tiene razon social
            
            Linea = Devnombresql340(Linea)
            
                        
           
            Linea = "'" & R!nifdatos & "',NULL,'" & DevNombreSQL(Linea) & "',"
           
                
            '`codpais`,`idenpais`,`nifresidencia`
            Aux = UCase(DBLet(R!PAIS, "T"))
            If Aux = "" Then Aux = "ESPAÑA"
            
            If Aux = "ESPAÑA" Then
                Linea = Linea & "'ES','1',NULL"
            Else
                'ESTRANJERO o INTRACOM
                Linea = Linea & "'" & Mid(Aux, 1, 2) & "','2','"
                Linea = Linea & DBLet(R!nifdatos, "T") & "'"
            End If
            
            '`clavelibro`,`claveoperacion`,   !!!ATENCION!!!! POnemos como clave libro una D  <<<<---- UNA D
            ' Antes de declara habrar que consolidar por NIF para todas las empresas declaradas,
            'quitanod los que no llegen
            Linea = Linea & ",'D',''"
            
            
            '`fechaexp`,`fechaop`,   '28 Enero 2013 Paso la ultima fecha con variable
            'Aux = ",'" & Format(RS!ultimafecha, FormatoFecha) & "'"
            Aux = ",'" & Format(FechaUltimo, FormatoFecha) & "'"
            Linea = Linea & Aux & Aux
           
               
            Linea = Linea & ",'',NULL"
            
            
            'If Identificacion = 1 Then
            Linea = Linea & ",NULL,NULL,NULL,"
           
            'Base UNO. SIEMPRE EXISTE
            Linea = Linea & "NULL,NULL,NULL," & TransformaComasPuntos(CStr(ImporAnt)) & "," & TransformaComasPuntos(CStr(ImporPer))
            
            Linea = vUsu.Codigo & "," & NumRegElim & "," & Linea & ",0,0,0"
            'Insertar
            NumRegElim = NumRegElim + 1
           
   
End Sub


Private Sub AgrupaPagosEfectivo(ByRef L As Label, MaximoImporteMetalico As Currency, EsAlFichero As Boolean)
Dim CADENA As String

    L.Caption = "Agrupando efectivo x nif"
    L.Refresh
    
    'Obtengo el ultimo registro
    Linea = "Select max(codigo) from tmp340 WHERE codusu = " & vUsu.Codigo
    Rs.Open Linea, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    'NO pue ser EOF
    NumRegElim = DBLet(Rs.Fields(0), "N") + 1
    Rs.Close
    Linea = "Select nifdeclarado,razosoci,codpais,idenpais,nifresidencia,max(fechaexp) lafecha,sum(base) anterior,sum(totiva) Periodo"
    Linea = Linea & " from tmp340 WHERE codusu = " & vUsu.Codigo
    Linea = Linea & " AND clavelibro = 'D' group by nifdeclarado"
    
    Rs.Open Linea, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CADENA = ""
    While Not Rs.EOF
        If DBLet(Rs!Anterior, "N") + DBLet(Rs!Periodo, "N") >= MaximoImporteMetalico Then
            L.Caption = Rs!nifdeclarado
            L.Refresh
            
            
            'Metemos el registro con clave de operacione Z (definida por mi) para los efectivos
            Linea = ", (" & vUsu.Codigo & "," & NumRegElim
            Linea = Linea & ",'" & Rs!nifdeclarado & "',NULL,'" & DevNombreSQL(Rs!razosoci) & "',"
           
                
            '`codpais`,`idenpais`,`nifresidencia`
            Linea = Linea & "'" & Rs!codPAIS & "','" & Rs!idenpais & "','" & DBLet(Rs!nifresidencia, "T") & "'"
            
            
            '`clavelibro`,`claveoperacion`,   !!!ATENCION!!!! POnemos como clave libro una Z  <<<<---- UNA z
            Linea = Linea & ",'Z',''"
            
            
            '`fechaexp`,`fechaop`,
            Linea = Linea & ",'" & Format(Rs!lafecha, FormatoFecha) & "','" & Format(Rs!lafecha, FormatoFecha) & "'"
            Linea = Linea & ",'',NULL"
            'If Identificacion = 1 Then
            Linea = Linea & ",NULL,NULL,NULL,"
           
            'Base UNO. SIEMPRE EXISTE
            Linea = Linea & "NULL,NULL,NULL," & TransformaComasPuntos(DBLet(Rs!Anterior, "N")) & "," & TransformaComasPuntos(DBLet(Rs!Periodo, "N"))
            
            Linea = Linea & ",0,0,0)"
            
            CADENA = CADENA & Linea
            
            'Insertar
            NumRegElim = NumRegElim + 1
            
            If Len(CADENA) > 50000 Then
                InsertaMultipleTMP340 CADENA
                CADENA = ""
            End If
        
        End If
        
        Rs.MoveNext
    Wend
    Rs.Close
    
    InsertaMultipleTMP340 CADENA
    
    
    
    L.Caption = "Eliminar datos tmp nif efectivo"
    L.Refresh
    'borramos los clavelibro='D'
    Linea = " DELETE from tmp340 WHERE codusu = " & vUsu.Codigo
    Linea = Linea & " AND clavelibro = 'D' "
    Conn.Execute Linea
    
    
    'Si es al fichero eliminaremos los pagos que no tengan nada en este periodo
    If EsAlFichero Then
        Linea = " DELETE from tmp340 WHERE codusu = " & vUsu.Codigo
        Linea = Linea & " AND clavelibro = 'Z' AND totiva=0 "  'totiva es im`porte periodo
        Conn.Execute Linea
    End If
        
End Sub
