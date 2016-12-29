Attribute VB_Name = "libFactura"
Option Explicit

'***********************************************************************************************
'***********************************************************************************************
'***********************************************************************************************
'***********************************************************************************************
'***********************************************************************************************
'
'
'                   Esto  habria que crear un clase Factura
'
'
'
'
'
'***********************************************************************************************
'***********************************************************************************************
'***********************************************************************************************
'***********************************************************************************************
'***********************************************************************************************
'***********************************************************************************************
'***********************************************************************************************







'======================================================================
'GRABAR EN TESORERIA
'======================================================================

Public Function InsertarEnTesoreria(Es1Cuota As Boolean, ByRef rsFactura As ADODB.Recordset, CuentaPrev As String, vTextosCSB As String, MenError As String) As Boolean
'Guarda datos de Tesoreria en tablas: ariges.svenci y en conta.scobros
Dim B As Boolean
Dim RS As ADODB.Recordset
Dim rsVenci As ADODB.Recordset
Dim Sql As String, Codmacta_ As String, textcsb33 As String
Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAuxConta As String 'para insertar en conta.scobro
Dim CadValues3 As String
Dim FecVenci As Date, FecVenci1 As Date
Dim ImpVenci As Currency 'importe para insertar en la svenci
Dim ImpVenci2 As Currency 'importe para insertar en conta.scobro
Dim Knumerovenci As Byte
Dim TotalFactura3 As Currency   'Por si acaso lleva aportacion al terminal
Dim ImporteDeLaFactura As Currency  'por si lleva pago por adelantado

'1 Julio 2009. Los graba en scobro
Dim CadenaDatosFiscales As String
Dim J As Integer
Dim NumeroDeVencimientos As Byte
Dim NuevaNorma19 As Boolean

Dim FormapagoAportacion  As Integer   'De momento NO la lee de parametros
Dim AuxIBAN As String

Dim TextoAuxiliar As String
Dim TipForPago As Byte

'por si acaso necesito
Dim ImpCheque As Currency
Dim Aportacion As Currency
Dim Agente As Integer

    On Error GoTo EInsertarTesoreria

    
    ImpCheque = 0   'por si las necesitamos en otro momento
    Aportacion = 0
    Agente = 1
    
    
    
    
    Set rsVenci = New ADODB.Recordset
    AuxIBAN = DBLet(rsFactura!IBAN, "T")
    
    
    
 
    vTextosCSB = DBSet(vTextosCSB, "T", "S")
     
    B = False
    InsertarEnTesoreria = False
    CadValues3 = ""
    CadValues = ""
    CadValues2 = ""
    TipForPago = DevuelveDesdeBDNew(1, "ariconta" & vParam.Numconta & ".formapago", "tipforpa", "codforpa", rsFactura!codforpa, "N")
    
    'campo para insertar en conta.scobro de Tesoreria. pAra las de telefonia ya lo ha creado arriba
    textcsb33 = "'FACTURA: " & rsFactura!numserie & "-" & Format(rsFactura!NumFactu, "0000000") & " de Fecha " & Format(rsFactura!FecFactu, "dd mmm yyyy") & "'"


    'Datos fiscales en scobro     Julio 2009
    'nomclien,domclien,pobclien, cpclien,proclien
    CadenaDatosFiscales = DBSet(rsFactura!NomClien, "T") & "," & DBSet(rsFactura!DomClien, "T") & "," & DBSet(rsFactura!PobClien, "T")
    CadenaDatosFiscales = CadenaDatosFiscales & "," & DBSet(rsFactura!codposta, "T") & "," & DBSet(rsFactura!ProClien, "T")
    
    J = vUsu.Codigo Mod 1000  'usuario real
    CadenaDatosFiscales = DBSet(rsFactura!NIFClien, "T") & "," & J & "," & DBSet(rsFactura!codpais, "T") & "," & CadenaDatosFiscales
    
    ImporteDeLaFactura = 0
    If Not IsNull(rsFactura!totfaccl) Then ImporteDeLaFactura = rsFactura!totfaccl
    Knumerovenci = 1
    
    'Obtener el Nº de Vencimientos de la forma de pago
    

    Sql = "SELECT numerove, primerve, restoven FROM ariconta" & vParam.Numconta & ".formapago WHERE codforpa=" & rsFactura!codforpa
    rsVenci.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not rsVenci.EOF Then
    
        If rsFactura!totfaccl < 0 Then
            'Fras NEGATIVAS solo hay un vencimiento
            NumeroDeVencimientos = 1
        Else
            NumeroDeVencimientos = CByte(rsVenci!numerove)
        End If
    
    
        
        If rsVenci!numerove > 0 And CCur(ImporteDeLaFactura) <> 0 Then
        
            'Comporbamos si el importe es <>0
            'Obtener los dias de pago del cliente,de momento no esta, y
            'la codmacta viene de la matricula. Ya veremos como
            Sql = " SELECT  0 diapago1, 0  diapago2,0 diapago3,0 mesnogir,0 diavtoat, matricula "
            Sql = Sql & " FROM clientes "
            Sql = Sql & " WHERE codclien=" & rsFactura!codclien
            Set RS = New ADODB.Recordset
            RS.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            Codmacta_ = DevuelveCuentaContableCliente(Es1Cuota, RS!Matricula)
           
            
            If Not RS.EOF Then
                cadValuesAux = "('" & rsFactura!numserie & "', " & rsFactura!NumFactu & ", '" & Format(rsFactura!FecFactu, FormatoFecha) & "', "
                CadValuesAuxConta = "('" & rsFactura!numserie & "', " & rsFactura!NumFactu & ", '" & Format(rsFactura!FecFactu, FormatoFecha) & "', "
                '                    Añadire a la cadena fija esta los valores de textcsb41,txcs
                CadValuesAuxConta = CadValuesAuxConta & vTextosCSB & ","
                
               
                'FECHA VTO
                FecVenci = CDate(rsFactura!FecFactu)
                FecVenci = DateAdd("d", DBLet(rsVenci!primerve, "N"), FecVenci)
                '===
                'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                If TipForPago <> 0 Then
                    FecVenci = ComprobarFechaVenci(FecVenci, DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))
                Else
                    FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                End If
                'Comprobar si cliente tiene mes a no girar
                FecVenci1 = FecVenci
                If CInt(DBLet(RS!mesnogir, "N")) <> 0 Then
                    FecVenci1 = ComprobarMesNoGira(FecVenci1, DBLet(RS!mesnogir, "N"), DBLet(RS!DiaVtoAt, "N"), DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))
                End If
                
                'Comprobar si cliente tiene dia de vencimiento atrasado
                CadValues = cadValuesAux & Knumerovenci & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                CadValues2 = CadValuesAuxConta & Knumerovenci & ", "
                CadValues2 = CadValues2 & Codmacta_ & ", " & rsFactura!codforpa & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                
                'IMPORTE del Vencimiento
                TotalFactura3 = ImporteDeLaFactura - 0 'Aportacion
                If NumeroDeVencimientos = 1 Then
                    ImpVenci = TotalFactura3
                    ImpVenci2 = TotalFactura3 - ImpCheque
                Else
                    
                    ImpVenci = Round2(TotalFactura3 / NumeroDeVencimientos, 2)
                    ImpVenci2 = Round2((TotalFactura3 - ImpCheque) / NumeroDeVencimientos, 2)
                    'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                    If ImpVenci * NumeroDeVencimientos <> TotalFactura3 Then
                        ImpVenci = Round(ImpVenci + (TotalFactura3 - ImpVenci * NumeroDeVencimientos), 2)
                    End If
                    'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                    If (ImpVenci2 * NumeroDeVencimientos) + ImpCheque <> TotalFactura3 Then
                        ImpVenci2 = Round(ImpVenci2 + (TotalFactura3 - ImpCheque - (ImpVenci2 * NumeroDeVencimientos)), 2)
                    End If
                End If
                
                CadValues = CadValues & DBSet(ImpVenci, "N") & ")"
                CadValues2 = CadValues2 & DBSet(ImpVenci2, "N") & ", '" & CuentaPrev & "', "
                
                
                CadValues2 = CadValues2 & DBSet(AuxIBAN, "T", "S") & ", "
                CadValues2 = CadValues2 & textcsb33 & ", " & DBSet(Agente, "N")
                
                'departamento y transfer
                'CadValues2 = CadValues2 & "," & DBSet(Me.DirDpto, "N", "S") & ",NULL"
                CadValues2 = CadValues2 & ",NULL,NULL"
                
                
                ' Datos fiscales en scobro nomclien , domclien, pobclien, cpclien, proclien
                 CadValues2 = CadValues2 & "," & CadenaDatosFiscales & ")"
                
                'Resto Vencimientos
                '--------------------------------------------------------------------
                For J = 2 To NumeroDeVencimientos
                   'FECHA Resto Vencimientos
                    '=== Laura 23/01/2007
                    'FecVenci = FecVenci + DBSet(rsVenci!restoven, "N")
                    FecVenci = DateAdd("d", DBLet(rsVenci!restoven, "N"), FecVenci)
                    '===
                    'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                    If TipForPago <> 0 Then
                        FecVenci = ComprobarFechaVenci(FecVenci, DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))
                    Else
                        FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                    End If
                    'Comprobar si cliente tiene mes a no girar
                    FecVenci1 = FecVenci
                    If DBLet(RS!mesnogir, "N") <> "0" Then
                        FecVenci1 = ComprobarMesNoGira(FecVenci1, DBLet(RS!mesnogir, "N"), DBLet(RS!DiaVtoAt, "N"), DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))
                    End If
                    Knumerovenci = Knumerovenci + 1
                    CadValues = CadValues & ", " & cadValuesAux & Knumerovenci & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                    CadValues2 = CadValues2 & ", " & CadValuesAuxConta & Knumerovenci & ", " & Codmacta_ & ", " & rsFactura!codforpa & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                    
                    'IMPORTE Resto de Vendimientos
                    ImpVenci = Round2(TotalFactura3 / rsVenci!numerove, 2)
                    ImpVenci2 = Round2((TotalFactura3 - ImpCheque) / rsVenci!numerove, 2)
                    CadValues = CadValues & DBSet(ImpVenci, "N") & ")"
                    CadValues2 = CadValues2 & DBSet(ImpVenci2, "N") & ", " & DBSet(CuentaPrev, "T") & ", "
                    CadValues2 = CadValues2 & DBSet(AuxIBAN, "T", "S") & ", " & textcsb33 & ", " & DBSet(Agente, "N") & ", "
                    
                    'CadValues2 = CadValues2 & DBSet(Me.DirDpto, "N", "S") & ",NULL"
                    CadValues2 = CadValues2 & "NULL,NULL"
                    
                    ' Datos fiscales en scobro nomclien , domclien, pobclien, cpclien, proclien
                    CadValues2 = CadValues2 & "," & CadenaDatosFiscales & ")"
                    
                Next J
                
                '--- Cheque regalo: laura 1/12/2006   y/o    Aportacion terminal
                'si hay cheque regalo insertar una linea más para la forma de pago correspondiente y el importe del cheque
                If ImpCheque > 0 Then
                
                    'Knumerovenci = J
                    'CadValues2 = CadValues2 & ", " & CadValuesAuxConta & Knumerovenci & "," & Codmacta & ", " & vParamAplic.ForPagoChequeRegalo & ", "
                   
                    ''FECHA VTO

                    'TextoAuxiliar = "primerve"
                    'TipForPago = DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", vParamAplic.ForPagoChequeRegalo, "N", TextoAuxiliar)
                    'FecVenci = CDate(FecFactu)
                    'FecVenci = FecVenci + CInt(TextoAuxiliar)
                    ''comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                    'If TipForPago <> 0 Then
                    '            FecVenci = ComprobarFechaVenci(FecVenci, DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))
                    '           MsgBox "FALTA cheque regalo con forma de pago no en EFECTIVO", vbInformation
                    'Else
                    '    FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                    'End If
                    
                    'CadValues2 = CadValues2 & DBSet(FecVenci, "F") & ", "
                    'CadValues2 = CadValues2 & DBSet(ImpCheque, "N") & ", '" & CuentaPrev & "', "
                    'If Not vParamAplic.ContabilidadNueva Then
                    '    CadValues2 = CadValues2 & DBSet(Banco, "N") & ", " & DBSet(Sucursal, "N") & ", " & DBSet(DigControl, "T") & ", " & DBSet(CuentaBan, "T") & ", "
                    'End If
                    'CadValues2 = CadValues2 & DBSet(AuxIBAN, "T", "S") & ", " & textcsb33 & ", " & DBSet(Agente, "N")
                    ''departamento
                    'CadValues2 = CadValues2 & "," & DBSet(Me.DirDpto, "N", "S") & ",NULL)"
                End If
                
                
                'Aportacion al terminal
                'Si tiene cuenta aportacion entonces añadiremos un cobro
                If Aportacion > 0 Then
                    'If vParamAplic.ctaAportacion = "" Then
                    '    MsgBox "Error cta aportacion NULL", vbExclamation
                    '    Exit Function
                    'End If
                        
                    ''Montamos EL SQL para el cobro de la aportacion al termina
                    ''ForPago
                    'FecVenci = CDate(FecFactu)
                    'Knumerovenci = Knumerovenci + 1
                    'CadValues2 = CadValues2 & ", " & CadValuesAuxConta & Knumerovenci & ",'" & vParamAplic.ctaAportacion & "', "
                   
                    'FormapagoAportacion = -1
                    
                    ''Vemos primero Efectivos con texto efectivo o contador
                    'TextoAuxiliar = "(nomforpa like '%efec%' or nomforpa like '%conta%') and tipforpa"
                    'TextoAuxiliar = DevuelveDesdeBDNew(conAri, "sforpa", "codforpa", TextoAuxiliar, "0", "N", TextoAuxiliar)
                    
                    ''CadValuesAuxConta = "(nomforpa like '%efec%' or nomforpa like '%conta%') and tipforpa"
                    ''CadValuesAuxConta = DevuelveDesdeBDNew(conAri, "sforpa", "codforpa", CadValuesAuxConta, "0", "N", CadValuesAuxConta)
                    
                    
                    'If TextoAuxiliar <> "" Then
                    '    'OK ya tenemos la forma de pago
                    '    TipForPago = 0
                    '    FormapagoAportacion = TextoAuxiliar
                    'Else
                    '    'Provamos de otro modo
                    '    TextoAuxiliar = "primerve"
                    '    'TipForPago = DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", ForPago, "N", CadValuesAuxConta)
                    '    TipForPago = DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", ForPago, "N", TextoAuxiliar)
                    '    FormapagoAportacion = ForPago
                    '    FecVenci = FecVenci + CInt(TextoAuxiliar)
                    'End If
                    
                  
                    'CadValues2 = CadValues2 & FormapagoAportacion & ", "
                    'CadValues2 = CadValues2 & DBSet(FecVenci, "F") & ", "
                    'CadValues2 = CadValues2 & DBSet(Aportacion, "N") & ", '" & CuentaPrev & "', "
                    'If Not vParamAplic.ContabilidadNueva Then
                    '    CadValues2 = CadValues2 & DBSet(Banco, "N") & ", " & DBSet(Sucursal, "N") & ", " & DBSet(DigControl, "T") & ", " & DBSet(CuentaBan, "T") & ", "
                    'End If
                    'CadValues2 = CadValues2 & DBSet(AuxIBAN, "T", "S") & ", " & textcsb33 & ", " & DBSet(Agente, "N")
                    ''departamento
                    'CadValues2 = CadValues2 & "," & DBSet(Me.DirDpto, "N", "S") & ",NULL"
                    
                    'CadValues2 = CadValues2 & "," & CadenaDatosFiscales & ")"
                End If
                
            End If
            RS.Close
            
            'Si habia un primer pago como aportacion entonces lo metemos aqui, con el numero venci=1
            'If ImporteAdelantado <> 0 Then
            '        FecVenci1 = Now
            '        Knumerovenci = 1
            '        CadValues = CadValues & ", " & cadValuesAux & Knumerovenci & ", '" & Format(FecVenci1, FormatoFecha) & "', "
            '        CadValues2 = CadValues2 & ", " & CadValuesAuxConta & Knumerovenci & ", " & Codmacta & ", " & FormaPagoAdelantado & ", '" & Format(FecVenci1, FormatoFecha) & "', "
            '
            '        'IMPORTE Resto de Vendimientos
            '        ImpVenci = ImporteAdelantado
            '        ImpVenci2 = ImporteAdelantado
            '        CadValues = CadValues & DBSet(ImpVenci, "N") & ")"
            '        CadValues2 = CadValues2 & DBSet(ImpVenci2, "N") & ", " & DBSet(CuentaPrev, "T") & ", "
            '        If Not vParamAplic.ContabilidadNueva Then
            '            CadValues2 = CadValues2 & DBSet(Banco, "N") & ", " & DBSet(Sucursal, "N") & ", " & DBSet(DigControl, "T") & ", " & DBSet(CuentaBan, "T") & ", "
            '        End If
            '        CadValues2 = CadValues2 & DBSet(AuxIBAN, "T", "S") & ", " & textcsb33 & ", " & DBSet(Agente, "N") & ", "
            '        CadValues2 = CadValues2 & DBSet(Me.DirDpto, "N", "S") & ",NULL"
            '        '1 Julio 2009
            '        ' Datos fiscales en scobro nomclien , domclien, pobclien, cpclien, proclien
            '        CadValues2 = CadValues2 & "," & CadenaDatosFiscales & ")"
            '
            'End If
            
            
            
        Else
            'totalfac =0 and numerovtos >=1
            B = True
        End If
        
        Set RS = Nothing
    End If
    rsVenci.Close
    Set rsVenci = Nothing
    
    If CadValues <> "" Then
        Sql = "INSERT INTO factcli_vtos (numserie,numfactu,fecfactu,numlinea,fecefect,impefect)"
        Sql = Sql & " VALUES " & CadValues
        Conn.Execute Sql
    End If
    
    
    'Grabar tabla scobro de la CONTABILIDAD
    '-------------------------------------------------
    If CadValues2 <> "" Then

        If CuentaPrev <> "" Then
            
        
            Sql = "INSERT INTO ariconta" & vParam.Numconta & ".cobros (numserie, numfactu, fecfactu,text41csb, "
            Sql = Sql & "numorden , Codmacta, codforpa, FecVenci, ImpVenci, ctabanc1, "
            Sql = Sql & "iban,text33csb,agente,departamento,transfer "
            Sql = Sql & ", nifclien,codusu,codpais"  'Junio 16
            Sql = Sql & ",nomclien,domclien,pobclien, cpclien,proclien)"   '=Datos fiscales. para conta nueva meto el NIF mNIFClien
            Sql = Sql & " VALUES " & CadValues2
            Conn.Execute Sql

        End If
    End If
    
    B = True
  '  If UtilizaFormaPagoAlternativa Then
  '      SQL = "UPDATE scafac set codforpa = " & ForPago
  '      SQL = SQL & " WHERE codtipom='" & Me.codtipom & "' AND numfactu = " & Me.NumFactu & " and fecfactu=" & DBSet(Me.FecFactu, "F")
  '      ejecutar SQL, False
  '  End If

    
EInsertarTesoreria:
    If Err.Number <> 0 Then
        B = False
        MenError = "Insertar en Tesoreria: " & vbCrLf & Err.Description
    End If
    InsertarEnTesoreria = B
End Function







Private Function ComprobarFechaVenci(FechaVenci As Date, Dia1 As Byte, Dia2 As Byte, Dia3 As Byte) As Date
Dim newFecha As Date
Dim B As Boolean

'=== Modificada Laura: 23/01/2007
    On Error GoTo ErrObtFec
    B = False
    
    '--- comprobar que tiene dias de pago para obtener nueva fecha
    If Not (Dia1 > 0 Or Dia2 > 0 Or Dia3 > 0) Then
        'si no tiene dias de pago la fecha es OK y fin
        ComprobarFechaVenci = FechaVenci
        Exit Function
    End If
        
    
    '--- Obtener nueva fecha del vencimiento
    newFecha = FechaVenci
    
    Do
        'si dia de la fecha vencimiento es uno de los 3 dias de pagos fecha es OK
        If Day(newFecha) = Dia1 Or Day(newFecha) = Dia2 Or Day(newFecha) = Dia3 Then
'            newFecha = CStr(newFecha)
            B = True
        Else
            'mientras esta en el mismo mes vamos aumentando dias hasta encontrar un dia de pago
            newFecha = DateAdd("d", 1, CDate(newFecha))
        End If
    Loop Until B = True Or Year(newFecha) = Year(FechaVenci) + 3
    
    ComprobarFechaVenci = newFecha
    Exit Function
    
ErrObtFec:
    MuestraError Err.Number, "Obtener Fecha vencimiento según dias de pago.", Err.Description
End Function


Public Function ComprobarMesNoGira(FecVenci As Date, MesNG As Byte, DiaVtoAt As Byte, Dia1 As Byte, Dia2 As Byte, Dia3 As Byte) As Date
Dim F As String
Dim diaPago As Byte

    If Month(FecVenci) = MesNG Then
        '### LAURA 14/08/2008
'        If DiaVtoAt > 0 Then
'            F = DiaVtoAt & "/"
'        Else
'            F = Day(FecVenci) & "/"
'        End If
        
'        If Month(FecVenci) + 1 < 13 Then
'            F = F & Month(FecVenci) + 1 & "/" & Year(FecVenci)
'        Else
'            F = F & "01/" & Year(FecVenci) + 1
'        End If

        If DiaVtoAt > 0 Then
            'si tiene dia de vto atrasado a ese dia del mes siguiente
            'al mes a no girar
            F = DiaVtoAt & "/"
            F = F & Month(FecVenci) & "/" & Year(FecVenci)
            F = DateAdd("m", 1, F)
        Else
            'si no tiene dia de vto atrasado el primer dia de pago
            'del mes siguiente si tiene o sino el siguiente mes del
            'vencimiento obtenido
            If Dia1 > 0 Or Dia2 > 0 Or Dia3 > 0 Then
                'tiene dias de pago: el menor dia del mes siguiente
                diaPago = Dia1
                If (diaPago = 0) Or ((Dia2 < diaPago) And Dia2 <> 0) Then diaPago = Dia2
                If (diaPago = 0) Or ((Dia3 < diaPago) And Dia3 <> 0) Then diaPago = Dia3
                
                F = diaPago & "/"
                F = F & Month(FecVenci) & "/" & Year(FecVenci)
            Else
                'no tiene dias de pago: al mes siguiente
                F = Day(FecVenci) & "/"
                F = F & Month(FecVenci) & "/" & Year(FecVenci)
            End If
            
            F = DateAdd("m", 1, F)
        End If
        '###
        
        FecVenci = Format(F, "dd/mm/yyyy")
    End If
    
    ComprobarMesNoGira = FecVenci
End Function



Public Function DevuelveCuentaContableCliente(EsCuota As Boolean, Matricula As String) As String
Dim C As String
Dim N As Integer
     
    N = vEmpresa.DigitosUltimoNivel
    C = Matricula
    N = N - Len(C)
    If EsCuota Then
        N = N - Len(vParam.Raizcuotas)
        C = vParam.Raizcuotas & String(N, "0") & C
    Else
        N = N - Len(vParam.Raiztasas)
        C = vParam.Raiztasas & String(N, "0") & C
    End If
    DevuelveCuentaContableCliente = C
End Function




