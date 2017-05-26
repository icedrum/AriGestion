Attribute VB_Name = "libContabilizar"
Option Explicit


' Nueva contabilizacion
' Los IVAS de la cabceera se los paso a las lineas
' ya que agrupara por cuenta, y codigiva
Private vTipoIva(2) As Currency
Private vPorcIva(2) As Currency
Private vPorcRec(2) As Currency
Private vBaseIva(2) As Currency
Private vImpIva(2) As Currency
Private vImpRec(2) As Currency


'******************************************************************************************
'******************************************************************************************
'******************************************************************************************
'
'
'   Creara un apunte a partir de un collection
'   col: 'codmacta | docum | codconce | ampliaci | imported|importeH |ctacontrpar| numseri| numfaccl
'   Despues de contrapartida llevara tambien numserie numfaccl  para poder actualizar el campo   'opcional
'
'
'******************************************************************************************
'******************************************************************************************
'******************************************************************************************
Public Function CrearApunteDesdeColeccion(Fecha As Date, Observaciones As String, ByRef ColApuntes As Collection) As Boolean
Dim Mc As ContadoresConta
Dim Actual As Boolean
Dim CADENA As String
Dim Aux As String
Dim k As Integer

    On Error GoTo eCrearApunteDesdeColeccion
    CrearApunteDesdeColeccion = False

    Set Mc = New ContadoresConta
    Actual = (Fecha < DateAdd("yyyy", 1, vEmpresa.FechaInicioEjercicio))
    CADENA = ""
    If Mc.ConseguirContador("0", Actual, True) = 1 Then Err.Raise 513, , "Error consiguiendo contador"
    
    'hcabapu(numdiari,fechaent,numasien,obsdiari,feccreacion,usucreacion,desdeaplicacion)
    CADENA = "INSERT INTO ariconta" & vParam.Numconta & ".hcabapu(numdiari,fechaent,numasien,obsdiari,feccreacion,usucreacion,desdeaplicacion) "
    CADENA = CADENA & " VALUES (1," & DBSet(Fecha, "F") & "," & Mc.Contador & "," & DBSet(Observaciones, "T")
    CADENA = CADENA & "," & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'arigestion')"
    Conn.Execute CADENA
    
    'Las lineas
    'codmacta | docum | codconce | ampliaci | imported|importeH |ctacontrpar|
    'hlinapu(numdiari,fechaent,numasien,linliapu,codmacta,numdocum,codconce,ampconce,timporteD,timporteH,ctacontr,numserie,numfaccl,fecfactu,tipforpa,numorden)
    CADENA = ""
    For k = 1 To ColApuntes.Count
        
        
        CADENA = CADENA & ", (1," & DBSet(Fecha, "F") & "," & Mc.Contador & "," & k & ","
        CADENA = CADENA & DBSet(RecuperaValor(ColApuntes.Item(k), 1), "T") & "," 'codmacta
        CADENA = CADENA & DBSet(RecuperaValor(ColApuntes.Item(k), 2), "T") & "," 'numdocum
        CADENA = CADENA & RecuperaValor(ColApuntes.Item(k), 3) & "," 'codconce
        CADENA = CADENA & DBSet(RecuperaValor(ColApuntes.Item(k), 4), "T") & "," 'ampconce
        Aux = RecuperaValor(ColApuntes.Item(k), 5) 'ImporteD
        If Aux = "" Then
            Aux = RecuperaValor(ColApuntes.Item(k), 6) 'ImporteD
            If Aux = "" Then Aux = "0"
            Aux = ImporteFormateado(Aux)
            CADENA = CADENA & "NULL," & DBSet(Aux, "N")
        Else
            Aux = ImporteFormateado(Aux)
            CADENA = CADENA & DBSet(Aux, "N") & ",NULL"
        End If
        CADENA = CADENA & "," & DBSet(RecuperaValor(ColApuntes.Item(k), 7), "T") & "," 'contrapartida
        
        Aux = RecuperaValor(ColApuntes.Item(k), 8)
        If Aux = "" Then
            CADENA = CADENA & "null,null,null,null,null"
        Else
            'numSerie , numfaccl, Fecfactu
            CADENA = CADENA & DBSet(Aux, "T") & "," & Val(RecuperaValor(ColApuntes.Item(k), 9)) & ","
            Aux = RecuperaValor(ColApuntes.Item(k), 10)
            CADENA = CADENA & DBSet(Aux, "F", "S") & ",0,1"   'tipforpa,numorden
        End If
        
        CADENA = CADENA & ")"
    Next
    Aux = "INSERT INTO ariconta" & vParam.Numconta & ".hlinapu(numdiari,fechaent,numasien,linliapu,codmacta,numdocum,codconce,ampconce,timporteD,timporteH,ctacontr,numserie,numfaccl,fecfactu,tipforpa,numorden) VALUES "
    CADENA = Mid(CADENA, 2) 'quitamos la primera coma
    Aux = Aux & CADENA
    Conn.Execute Aux
    
    
    CrearApunteDesdeColeccion = True
    
    
eCrearApunteDesdeColeccion:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set Mc = Nothing
End Function















'******************************************************************************************
'******************************************************************************************
'******************************************************************************************
'
'
'   Pasar facturas clientes a contabilidad
'
'
'******************************************************************************************
'******************************************************************************************
'******************************************************************************************



Public Function PasarFactura(CadWhere As String, ByRef LB As Label) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
' ariges.scafac --> conta.cabfact
' ariges.slifac --> conta.linfact
'Actualizar la tabla ariges.scafac.inconta=1 para indicar que ya esta contabilizada
Dim B As Boolean
Dim cadMen As String
Dim Sql As String
Dim ErrorContab As String
Dim RN As ADODB.Recordset
Dim AmpliacionFacurasCli  As String
Dim AmpliacionParametros As Byte
Dim Abononeg As Boolean

    Set RN = New ADODB.Recordset
    
     
    LB.Caption = "Leyendo parametros"
    LB.Refresh
    
    'Parametros
    Sql = "SELECT concefcl,conceacl,nctafact,Abononeg FROM ariconta" & vParam.Numconta & ".parametros "
    RN.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    AmpliacionParametros = RN!nctafact
    Abononeg = DBLet(RN!Abononeg, "N") = 1
    ErrorContab = "tipoconce"
    Sql = DevuelveDesdeBD("nomconce", "ariconta" & vParam.Numconta & ".conceptos", "codconce", CStr(RN!concefcl), "N", ErrorContab)
    If Sql = "" And ErrorContab = "tipoconce" Then Exit Function
    AmpliacionFacurasCli = RN!concefcl & "|" & Sql & "|"
        
    ErrorContab = "tipoconce"
    Sql = DevuelveDesdeBD("nomconce", "ariconta" & vParam.Numconta & ".conceptos", "codconce", CStr(RN!conceacl), "N", ErrorContab)
    If Sql = "" And ErrorContab = "tipoconce" Then Exit Function
    AmpliacionFacurasCli = AmpliacionFacurasCli & RN!conceacl & "|" & Sql & "|"
    
    RN.Close
    
    Sql = ""
    ErrorContab = ""
    
    LB.Caption = "Leyendo registros"
    LB.Refresh
    
    Sql = "Select numserie,numfactu,fecfactu from factcli  WHERE  "
    If CadWhere <> "" Then Sql = Sql & CadWhere & " AND "
    Sql = Sql & " intconta =0"
    Sql = Sql & " ORDER BY 1,2 FOR UPDATE "
    RN.Open Sql, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    
    If Not RN.EOF Then
        While Not RN.EOF
            k = k + 1
            RN.MoveNext
        Wend
        RN.MoveFirst
        
    End If
    
    J = 0
    While Not RN.EOF
        J = J + 1
        LB.Caption = "(" & J & "/" & k & ")    Factura: " & RN!numSerie & " " & RN!NumFactu
        LB.Refresh
        
        
        Conn.BeginTrans
        
        'Insertar en la conta Cabecera Factura
        Sql = "numserie = " & DBSet(RN!numSerie, "T") & " AND numfactu =" & RN!NumFactu & " AND fecfactu =" & DBSet(RN!Fecfactu, "F")
        B = InsertarFacturaConta(Sql)
        If B Then
           
            If B Then
                ErrorContab = IntegraLaFacturaCliente(RN!NumFactu, Year(RN!Fecfactu), RN!numSerie, AmpliacionFacurasCli, AmpliacionParametros, Abononeg)
                'ErrorContab = vContaFra.IntegraLaFacturaCliente(vContaFra.NumeroFactura, vContaFra.Anofac, vContaFra.Serie)
                'vContaFra.AnyadeElError ErrorContab
                If ErrorContab <> "" Then
                    B = False
                    cadMen = cadMen & " " & ErrorContab & vbCrLf
                End If
            End If
        Else
            cadMen = cadMen & ". " & Sql & vbCrLf
        End If


        If B Then
            Sql = "UPDATE factcli set intconta=1 WHERE "
            Sql = Sql & "numserie = " & DBSet(RN!numSerie, "T") & " AND numfactu =" & RN!NumFactu & " AND fecfactu =" & DBSet(RN!Fecfactu, "F")
            If Not Ejecuta(Sql) Then B = False
        End If
    
        If B Then
            
            
            'B = ActualizarCabFact("scafac", cadWhere, cadMen)
            'cadMen = "Actualizando Factura: " & cadMen
            Conn.CommitTrans
            If (J Mod 20) = 0 Then
                DoEvents
                Screen.MousePointer = vbHourglass
            End If
            
        Else
            Conn.RollbackTrans
        End If
        
        'Siguiente
        RN.MoveNext
    Wend
    RN.Close
        
    Set RN = Nothing
        
    If cadMen <> "" Then MsgBox "Se han producido errores. " & vbCrLf & cadMen, vbExclamation
End Function



Private Function InsertarFacturaConta(CadWhere As String) As Boolean
'Insertando en tabla conta.cabfact
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim Nulo2 As String
Dim Nulo3 As String
Dim FraRectifica As String
Dim I As Integer
Dim Aux As String
Dim BaseImp As Currency
Dim TotalFac As Currency


'Nueva contabilidad
Dim ImporAux As Currency
Dim TipoOpera As Byte   'Nueva contabilidad. Tipo operacion (usuarios.wtipopera)
Dim CadenaInsertFaclin2 As String
Dim Sql2 As String



    On Error GoTo EInsertar
    
    Set Rs = New ADODB.Recordset
    
    FraRectifica = ""
    If InStr(1, CadWhere, "'FRT'") > 0 Then
        
        '¡Voy a intentar sacar le numero de factura a la que rectifica. Sera de laobservacion
        FraRectifica = "S"
    End If
    Sql = " SELECT  numserie,numfactu,fecfactu, clientes.codclien, 0 cliabono,year(fecfactu) as anofaccl,"
    Sql = Sql & " 1 codagent,clientes.nomclien,clientes.domclien,clientes.codposta,clientes.pobclien,"
    Sql = Sql & " clientes.proclien,clientes.nifclien, codpais,factcli.codforpa,"
    Sql = Sql & " TotBases , totbasesret, totrecargo, totfaccl, retfaccl, trefaccl, cuereten, tiporeten,totivas ,RectSer,RectNum,RectFecha"
    Sql = Sql & " FROM factcli INNER JOIN clientes ON factcli.codclien=clientes.codclien "
    Sql = Sql & " WHERE " & CadWhere
    
    
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    CadenaInsertFaclin2 = ""
    If Not Rs.EOF Then
        
        TotalFac = Rs!totfaccl

        
        Sql = DevuelveCuentaContableCliente(Rs!numSerie = "CUO", Rs!CodClien)
        
        Sql = "'" & Rs!numSerie & "'," & Rs!NumFactu & "," & DBSet(Rs!Fecfactu, "F") & "," & DBSet(Sql, "T") & "," & Year(Rs!Fecfactu) & ","
        
        
        
        
        'MAYO 2009
        'Si es una factura rectificativa, y hemos encontrado
        ' a k factura rectifica entonces meto esto, sino sigue como antes
        If FraRectifica = "" Then
            
            
                'Observacion factura
                'vParamAplic.ObsFactura
                Select Case 2
                Case 0
                    'Vacio
                    Sql = Sql & ValorNulo
                Case 1
                    'Nº Factura
                    Sql = Sql & "'" & DevNombreSQL("N/Fra " & Rs!NumFactu) & "'"
                Case 2
                    'Fecha integracion
                    Sql = Sql & "'" & Format(Now, FormatoFecha) & "'"
                End Select
            
           
        Else
            
            If DBLet(Rs!RectSer, "T") <> "" Then
                FraRectifica = Rs!RectSer & " " & DBLet(Rs!RectNum, "N") & " " & DBLet(Rs!RectFecha, "T")
            Else
                FraRectifica = "No encontrado"
            End If
            Sql = Sql & "'" & FraRectifica & "'"
        End If
        
        
        'TotBases , totbasesret, totrecargo, totfaccl, retfaccl, trefaccl, cuereten, tiporeten
        'totbases ,totbasesret,
        ImporAux = Rs!TotBases
        Sql = Sql & "," & DBSet(ImporAux, "N") & "," & ValorNulo & ","
        'totivas
        ImporAux = DBLet(Rs!TotIvas, "N")
        Sql = Sql & DBSet(ImporAux, "N") & ","
        ',totrecargo,totfaccl,retfaccl,trefaccl,cuereten,tiporeten,
        ImporAux = DBLet(Rs!totrecargo, "N")
        Sql = Sql & DBSet(ImporAux, "N") & "," & DBSet(Rs!totfaccl, "N")
        Sql = Sql & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0,"
        
            
        'fecliqcl,nommacta,dirdatos,codpobla,despobla,desprovi,nifdatos,codpais,dpto,codagente,codforpa,escorrecta,
        Sql = Sql & DBSet(Rs!Fecfactu, "F") & "," & DBSet(Rs!NomClien, "T") & "," & DBSet(Rs!domclien, "T", "S") & ","
        Sql = Sql & DBSet(Rs!codposta, "T", "S") & "," & DBSet(Rs!pobclien, "T", "S") & "," & DBSet(Rs!proclien, "T", "S") & ","
        Sql = Sql & DBSet(Rs!NIFClien, "F", "S") & "," & DBSet(Rs!codpais, "T", "S") & ",NULL,"
        Sql = Sql & DBSet(Rs!codagent, "N", "S") & "," & Rs!Codforpa & ",1,"
        
        
        
        'codopera,codconce340,codintra
        '*****
        ' Tipo de operacion
        '  GENERAL // INTRACOMUNITARIA // EXPORT. - IMPORT. //   INTERIOR EXENTA   // INV. SUJETO PASIVO   // R.E.A.
        'Si es una factura con IVA 0%
        If Rs!TotIvas = 0 Then
            'IVA ES CERO
            Aux = DBLet(Rs!codpais, "T")
            If Aux = "" Then Aux = "ES"
            
            
            If Aux = "ES" Then
                'NACIONAL. Facturas exenta de iva
                TipoOpera = 3
                
            Else
                'FALTA
                Stop
                'Aux = DevuelveDesdeBD(conConta, "intracom", "paises", "codpais", Aux, "T")
                Aux = "1"
                If Aux = "1" Then
                    'intracomunitaria
                    TipoOpera = 1
                Else
                    'Exstranjero
                    TipoOpera = 2
                End If
            End If
        Else
            'Factura NORMAL
            TipoOpera = 0
        End If
        
        'Concepto 340
        '---------------------
        ' 0 Habitual                B  Ticuet agrupado         C  Varios tipos impositivos
        ' D Rectificativa           I Sujeto pasivo             J Tikets
        ' P adquisiciones de bienes y servicios
        Select Case Rs!numSerie
        Case "FTG"
            Aux = "B"
        Case "FTI"
            Aux = "J"
        Case "FRT"
            Aux = "D"
        Case Else
            'HABITUAL
            'If Not IsNull(RS!porciva2) Then
            If False Then
                Aux = "C" 'varios tipos de iVA
            Else
                Aux = "0"
            End If
        End Select
        
        'codopera,codconce340,codintra
        Sql = Sql & TipoOpera & "," & DBSet(Aux, "T") & ","
        Aux = ValorNulo
        If TipoOpera = 1 Then Aux = "'E'" 'Entregas intracomunitarias extenas de IVA
        Sql = Sql & Aux
        
        Cad = "(" & Sql & ")"
            
    End If 'rs.eof
    Rs.Close
    
    
    
    Sql = "INSERT INTO ariconta" & vParam.Numconta & ".factcli (numserie,numfactu,fecfactu,codmacta,anofactu,observa,totbases ,totbasesret,totivas,totrecargo,"
    Sql = Sql & "totfaccl,retfaccl,trefaccl,cuereten,tiporeten,fecliqcl,nommacta,dirdatos,codpobla"
    Sql = Sql & ",despobla,desprovi,nifdatos,codpais,dpto,codagente,codforpa,escorrecta,codopera,codconce340,codintra) "
    Sql = Sql & " VALUES " & Cad
    Conn.Execute Sql

    
    Sql = "SELECT * FROM factcli_totales "
    Sql = Sql & " WHERE " & CadWhere
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    I = 0
    While Not Rs.EOF
        I = I + 1
        'numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)
        'IVA 1, siempre existe
        
        Sql2 = "'" & Rs!numSerie & "'," & Rs!NumFactu & "," & DBSet(Rs!Fecfactu, "F") & "," & Year(Rs!Fecfactu) & ","
        Sql2 = Sql2 & I & "," & DBSet(Rs!Baseimpo, "N") & "," & Rs!codigiva & "," & DBSet(Rs!porciva, "N") & ","
        Sql2 = Sql2 & DBSet(Rs!porcrec, "N") & "," & DBSet(Rs!Impoiva, "N") & "," & DBSet(Rs!ImpoRec, "N")
        CadenaInsertFaclin2 = CadenaInsertFaclin2 & ", (" & Sql2 & ")"
        Rs.MoveNext
    Wend
    Rs.Close
    
    If CadenaInsertFaclin2 <> "" Then
        CadenaInsertFaclin2 = Mid(CadenaInsertFaclin2, 2)
        Sql = "INSERT INTO ariconta" & vParam.Numconta & ".factcli_totales(numserie,numfactu,fecfactu,anofactu,numlinea,baseimpo,codigiva,"
        Sql = Sql & "porciva,porcrec,impoiva,imporec) VALUES " & CadenaInsertFaclin2
        Conn.Execute Sql
    End If
    
    
    'Las lineas
    Sql = " SELECT  numserie , NumFactu, Fecfactu, numlinea, factcli_lineas.codconce, Importe,"
    Sql = Sql & " factcli_lineas.codigiva, porciva, porcrec, Impoiva, ImpoRec, aplicret,codmacta"
    Sql = Sql & " FROM factcli_lineas INNER JOIN conceptos ON factcli_lineas.codconce=conceptos.codconce "
    Sql = Sql & " WHERE " & CadWhere
    
    
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    I = 0
    CadenaInsertFaclin2 = ""
    While Not Rs.EOF
        'factcli_lineas
        'numserie , NumFactu, Fecfactu, anofactu, numlinea, codmacta, Baseimpo, codigiva,
        'porciva, porcrec, Impoiva, ImpoRec, aplicret, codccost
        I = I + 1
        Sql = "'" & Rs!numSerie & "'," & Rs!NumFactu & "," & DBSet(Rs!Fecfactu, "F") & "," & Year(Rs!Fecfactu) & "," & I & ","
        Sql = Sql & DBSet(Rs!codmacta, "T") & "," & DBSet(Rs!Importe, "N") & "," & Rs!codigiva & "," & DBSet(Rs!porciva, "N") & ","
        Sql = Sql & DBSet(Rs!porcrec, "N") & "," & DBSet(Rs!Impoiva, "N") & "," & DBSet(Rs!ImpoRec, "N")
        Sql = Sql & "," & Val(Rs!aplicret) & ",NULL"  'codccost
        CadenaInsertFaclin2 = CadenaInsertFaclin2 & ", (" & Sql & ")"
        Rs.MoveNext
    Wend
    Rs.Close
    If CadenaInsertFaclin2 <> "" Then
        CadenaInsertFaclin2 = Mid(CadenaInsertFaclin2, 2)
        
        Sql = "INSERT INTO ariconta" & vParam.Numconta & ".factcli_lineas(numserie , NumFactu, Fecfactu, anofactu, numlinea, codmacta, Baseimpo, codigiva"
        Sql = Sql & ",porciva, porcrec, Impoiva, ImpoRec, aplicret, codccost) VALUES " & CadenaInsertFaclin2
        Conn.Execute Sql
    End If
    
    
    
    
    
    
    
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarFacturaConta = False
        Msg = Err.Description
        Err.Clear
        'cadErr = Err.Description
    Else
        InsertarFacturaConta = True
    End If
    Set Rs = Nothing
End Function


'
'AmpliacionFacurasCli:  concepto normal|nomconce|conceto abono|nomconce|
'AmpliacionParametros:
Public Function IntegraLaFacturaCliente(numFac As Long, Anofac As Integer, numSerie As String, AmpliacionFacurasCli2 As String, AmpliacionParametros As Byte, Abononeg As Boolean) As String
Dim Cad As String
Dim cad2 As String
Dim Cad3 As String
Dim Amplia2 As String
Dim DocConcAmp As String
Dim RF As Recordset
Dim ImporteNegativo As Boolean
Dim Importe0 As Boolean
Dim PrimeraContrapartida As String
Dim A_Donde As String
Dim numlinea As Integer
Dim Sql As String
Dim Importe As Currency
Dim DatosRetencion As String
Dim Cliente As String
Dim FechaAsi As Date
Dim DiarioFacturas As Integer
Dim NumAsiento As Long

    On Error GoTo EIntegraLaFactura
    'Sabemos que
    'numfac     --> CODIGO FACTURA
    'anofac      --> AÑO FACTURA
    'NUmSerie       --> SERIE DE LA FACTURA

    
    
    'Obtenemos los datos de la factura
    A_Donde = "Apunte: " & numSerie & Format(numFac, "000000") & " / " & Anofac
    Set RF = New ADODB.Recordset
    
    
    Sql = "SELECT numserie, numfactu codfaccl, fecfactu fecfaccl,factcli.codmacta"
    Sql = Sql & ", totfaccl ,factcli.observa confaccl  ,cuereten  , trefaccl"
    Sql = Sql & " ,numasien,cuentas.nommacta,factcli.nommacta as nomclien FROM "
    Sql = Sql & "ariconta" & vParam.Numconta & ".factcli left join  ariconta" & vParam.Numconta & ".cuentas"
    Sql = Sql & " ON factcli.codmacta = cuentas.codmacta"
    Sql = Sql & " WHERE numserie='" & numSerie
    Sql = Sql & "' AND numfactu= " & numFac
    Sql = Sql & " AND anofactu=" & Anofac
    
    
    RF.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If RF.EOF Then
        IntegraLaFacturaCliente = "No se encuentra la factura: " & vbCrLf & A_Donde
        RF.Close
        Exit Function
    End If
    
    
    If Not IsNull(RF!numasien) Then
        IntegraLaFacturaCliente = "Factura contabilizada: " & vbCrLf & A_Donde & "  Num: " & RF!numasien
        RF.Close
        Exit Function
    End If
    
    
    'Creamos la cuentas
    If IsNull(RF!nommacta) Then
        A_Donde = "Cuenta cliente"
        
        Sql = "INSERT INTO ariconta" & vParam.Numconta & ".cuentas(codmacta,nommacta,apudirec,razosoci) VALUES "
        Sql = Sql & "(" & DBSet(RF!codmacta, "T") & "," & DBSet(RF!NomClien, "T") & ",'S'," & DBSet(RF!NomClien, "T") & ")"
        Conn.Execute Sql
    End If
    
    'COnseguir contador y eso
    FechaAsi = RF!fecfaccl
    
    DiarioFacturas = 1 'Val(RecuperaValor(Diarios, 1)) 'Clientes 1,proveedores 2
    'Consigo contador
    '****************************************+
   
    Do
        NumAsiento = ConseguirContador(FechaAsi <= vEmpresa.FechaFinEjercicio, A_Donde)
        If NumAsiento = 0 Then
            'Ha habido algun tipo de error.
            
           
                IntegraLaFacturaCliente = A_Donde
            
                RF.Close
                Exit Function
       End If
    Loop Until NumAsiento > 0
    
    'Cabecera del hco de apuntes
    A_Donde = "Inserta cabecera hco apuntes"
    Sql = "INSERT INTO ariconta" & vParam.Numconta & ".hcabapu (numdiari, fechaent, numasien, obsdiari"
    Sql = Sql & ",feccreacion,usucreacion,desdeaplicacion) VALUES ("
    Sql = Sql & DiarioFacturas & ",'" & Format(FechaAsi, FormatoFecha) & "'," & NumAsiento
    Sql = Sql & "," & DBSet(RF!confaccl, "T", "S")
    Sql = Sql & "," & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARIGESTION'"
    Conn.Execute Sql & ")"
    
    'Lineas fijas, es decir la linea de cliente, importes y tal y tal
    'Para el sql
    
    Cad = "INSERT INTO ariconta" & vParam.Numconta & ".hlinapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, "
    Cad = Cad & "codconce,ampconce, timporteD, timporteH,codccost, ctacontr, idcontab, punteada)"
    
    
    
    
    Cad = Cad & " VALUES (" & DiarioFacturas & ",'" & Format(FechaAsi, FormatoFecha) & "'," & NumAsiento & ","
    numlinea = 1 'Contador de lineas
    
    
    A_Donde = "Linea cliente"
    
    'Guardo unos datos
    DatosRetencion = ""
    If Not IsNull(RF!cuereten) Then DatosRetencion = RF!cuereten & "|" & RF!trefaccl & "|"
    
    '-------------------------------------------------------------------
    'LINEA Cliente
    Sql = numlinea & ",'" & RF!codmacta & "',"
    

    ' en AmpliacionFacurasCli_ van:
    '  concepto normal|nomconce|conceto abono|nomconce|
    If RF!totfaccl >= 0 Then
        DocConcAmp = RecuperaValor(AmpliacionFacurasCli2, 1)
    Else
        DocConcAmp = RecuperaValor(AmpliacionFacurasCli2, 3)
    End If
    DocConcAmp = "'" & Format(numFac, "00000") & "'," & DocConcAmp & ",'"
    
    
    'Ampliacion segun parametros
    Select Case AmpliacionParametros
    Case 1
        If RF!totfaccl < 0 Then
            cad2 = RecuperaValor(AmpliacionFacurasCli2, 4)
        Else
            cad2 = RecuperaValor(AmpliacionFacurasCli2, 2)
        End If
        '28/02/2007.
        'Añado numerie
        cad2 = cad2 & " " & numSerie & Format(numFac, "00000")
    Case 2
        cad2 = DevNombreSQL(DBLet(RF!nommacta))
    Case Else
        cad2 = DBLet(RF!confaccl)
    End Select
    
    '   Modificacion para k aparezca en la ampliacio el CC en la ampliacion de codmacta
    '
    Amplia2 = cad2
'    If CCenFacturas Then
'        A_Donde = "CC en Facturas."
'        Cad3 = DevuelveCentroCosteFactura(True, PrimeraContrapartida, numFac, Anofac, numSerie)
'        If Cad3 <> "" Then
'            If Len(Amplia2) > 21 Then Amplia2 = Mid(Amplia2, 1, 21)
'            'Opcion1
'            'Amplia2 = Amplia2 & " .CC:" & Cad3
'            'Opcion2
'            Amplia2 = Amplia2 & " [" & Cad3 & "]"
'        End If
'    End If
    A_Donde = "Linea cliente. Ampliacion2"
    
    
    Sql = Sql & DocConcAmp & Amplia2 & "'"
    DocConcAmp = DocConcAmp & cad2 & "'"   'DocConcAmp Sirve para el IVA
    
    'Esta variable sirve para las demas
    ImporteNegativo = (RF!totfaccl < 0)
    
    'Importes, atencion importes negativos
    '  antes --> Cad2 = CadenaImporte(ImporteNegativo, True, RF!totfaccl)
    cad2 = CadenaImporte(True, RF!totfaccl, Importe0, Abononeg)
    Sql = Sql & "," & cad2 & ",NULL,"
    
    'Contrpartida. 28 Marzo 2006
    If PrimeraContrapartida <> "" Then
        Sql = Sql & "'" & PrimeraContrapartida & "'"
    Else
        Sql = Sql & "NULL"
    End If
    Sql = Sql & ",'FRACLI',0)"
    
    
    Conn.Execute Cad & Sql
    numlinea = numlinea + 1 'Es el contador de lineaapunteshco
    Cliente = RF!codmacta
    
      
    RF.Close
      
    'Abrimos el RF con las lineas del IVA
    Sql = "SELECT numserie, numfactu codfaccl, fecfactu fecfaccl"
    Sql = Sql & ", baseimpo,codigiva,porciva,porcrec,impoiva,imporec  FROM ariconta" & vParam.Numconta & ".factcli_totales "  'totales intregados . Dben ser los mismos que en arigestion
    Sql = Sql & " WHERE numserie='" & numSerie
    Sql = Sql & "' AND numfactu= " & numFac
    Sql = Sql & " AND anofactu=" & Anofac
    RF.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RF.EOF
        
        
        A_Donde = "IVA " & numlinea - 1
        Cad3 = "cuentarr"
        cad2 = DevuelveDesdeBD("cuentare", "ariconta" & vParam.Numconta & ".tiposiva", "codigiva", RF!codigiva, "N", Cad3)
        If cad2 <> "" Then
            Sql = numlinea & ",'" & cad2 & "'," & DocConcAmp
            cad2 = CadenaImporte(False, RF!Impoiva, Importe0, Abononeg)
            Sql = Sql & "," & cad2 & ","
            Sql = Sql & "NULL,'" & Cliente & "','FRACLI',0)"
            'If Not Importe0 Then
            If True Then
                Conn.Execute Cad & Sql
                numlinea = numlinea + 1
            End If
            
            'La de recargo  1-----------------
            If Not IsNull(RF!ImpoRec) Then
                     Sql = numlinea & "," & Cad3 & "," & DocConcAmp
                    'Importes, atencion importes negativos
                    cad2 = CadenaImporte(False, RF!ImpoRec, Importe0, Abononeg)
                    Sql = Sql & "," & cad2 & ","
                    Sql = Sql & "NULL,'" & Cliente & "','FRACLI',0)"
                    If Not Importe0 Then
                        Conn.Execute Cad & Sql
                        numlinea = numlinea + 1
                    End If
            End If
        
        End If
    
    
    
        numlinea = numlinea + 1 'Es el contador de lineaapunteshco
        RF.MoveNext
    Wend
    RF.Close
    
    '-------------------------------------
    ' RETENCION
    A_Donde = "Retencion"
    
    'hay una cadena con los datos, pq el rf deberaimos haberlo cerrado ya
    
    If DatosRetencion <> "" Then
        
        Sql = numlinea & ",'" & RF!cuereten & "'," & DocConcAmp
        'Importes, atencion importes negativos
        cad2 = CadenaImporte(True, RF!trefaccl, Importe0, Abononeg)
        Sql = Sql & "," & cad2 & ","
        Sql = Sql & "NULL,NULL,'FRACLI',0)"
       
        Conn.Execute Cad & Sql
        numlinea = numlinea + 1 'Es el contador de lineaapunteshco
    End If
    
    
    
    '------------------------------------------------------------
    'Las lineas de la factura.
    
    A_Donde = "Leyendo lineas factura"
    Sql = "SELECT codmacta codtbase, baseimpo impbascl,codccost  FROM ariconta" & vParam.Numconta & ".factcli_lineas "
    Sql = Sql & " WHERE numserie='" & numSerie
    Sql = Sql & "' AND numfactu= " & numFac
    Sql = Sql & " AND anofactu=" & Anofac
    RF.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    'Para cada linea insertamos
    cad2 = ""
    A_Donde = "Procesando lineas"
    While Not RF.EOF
        'Importes, atencion importes negativos
        If cad2 = "" Then PrimeraContrapartida = RF!codtbase
        Sql = numlinea & ",'" & RF!codtbase & "'," & DocConcAmp
        cad2 = CadenaImporte(False, RF!impbascl, Importe0, Abononeg)
        Sql = Sql & "," & cad2 & ","
        If IsNull(RF!CodCCost) Then
            cad2 = "NULL"
        Else
            cad2 = "'" & RF!CodCCost & "'"
        End If
        
        Sql = Sql & cad2 & ",'" & Cliente & "','FRACLI',0)"
    
        Conn.Execute Cad & Sql
        numlinea = numlinea + 1 'Es el contador de lineaapunteshco
        

        RF.MoveNext
        If Not RF.EOF Then PrimeraContrapartida = ""
    Wend
    RF.Close
    
    
    
    
    'AHora viene lo bueno.  MARZO 2006
    'Si el valor fuera true YA lo habria insertado en la cabcera
    'If CCenFacturas Then
    '    If PrimeraContrapartida <> "" Then
    '        SQL = "UPDATE hlinapu SET ctacontr ='" & PrimeraContrapartida & "'"
    '        SQL = SQL & " WHERE numdiari = " & DiarioFacturas & " AND fechaent ='" & Format(FechaAsi, FormatoFecha) & "' and numasien = " & NumAsiento
    '        SQL = SQL & " AND linliapu =1 " 'LA PRIMERA LINEA SIEMPRE ES LA DE LA CUENTA
    '        miConexion.Execute SQL  '
    '    End If
    'End If
        
    
    
    
    'Actualimos en factura, el nº de asiento
    Sql = "UPDATE ariconta" & vParam.Numconta & ".factcli SET numdiari = " & DiarioFacturas & ", fechaent = '" & Format(FechaAsi, FormatoFecha) & "', numasien =" & NumAsiento
    Sql = Sql & " WHERE numserie='" & numSerie
    Sql = Sql & "' AND numfactu= " & numFac
    Sql = Sql & " AND anofactu=" & Anofac
    Conn.Execute Sql
    
    
    Set RF = Nothing
    
    Exit Function
EIntegraLaFactura:
    IntegraLaFacturaCliente = A_Donde & vbCrLf & Err.Description & vbCrLf & Sql
    Set RF = Nothing
End Function



Private Function ConseguirContador(EjercicioActual As Boolean, ByRef PosibleError As String) As Long
Dim RT As ADODB.Recordset
Dim C1 As Long
Dim F1 As Date
Dim Sql As String

    On Error GoTo Err1
    'Abrimos bloqueando
    ConseguirContador = 0 'ERROR
    Sql = "Select * from ariconta" & vParam.Numconta & ".contadores WHERE TipoRegi = '0' "
    C1 = 0
    Set RT = New ADODB.Recordset
    RT.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RT.EOF Then
        If EjercicioActual Then
            C1 = RT!Contado1
        Else
            C1 = RT!Contado2
        End If
    Else
        PosibleError = "No se ecuentra Contador( 0)"
    End If
    RT.Close
    If C1 = 0 Then Exit Function
    C1 = C1 + 1
    
    
    

    

        Sql = "Select numasien from ariconta" & vParam.Numconta & ".hcabapu where numasien = " & C1
        
        'Las fechas
        If EjercicioActual Then
            F1 = DateAdd("yyyy", -1, vEmpresa.FechaFinEjercicio)
            Sql = Sql & " AND fechaent > " & DBSet(F1, "F") 'Mayor estricto
            Sql = Sql & " AND fechaent <= " & DBSet(vEmpresa.FechaFinEjercicio, "F")
        Else
            F1 = DateAdd("yyyy", 1, vEmpresa.FechaFinEjercicio)
            Sql = Sql & " AND fechaent > " & DBSet(DateAdd("yyyy", 1, vEmpresa.FechaFinEjercicio), "F")
            Sql = Sql & " AND fechaent <= " & DBSet(DateAdd("yyyy", 1, F1), "F")
        End If

        
        RT.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If RT.EOF Then
            'OK. Todo bien
        Else
            
            'Error. Saldremos con el error
            PosibleError = "Ya existe el asiento: " & C1 & "    No se ha podido contabilizar desde el registro de facturas"
            Sql = ""
        End If
        RT.Close
    
    
    If Sql = "" Then
        'Ha habido un error . Ya exise asiento
        C1 = 0
        Exit Function
    End If
    
    'Actualizamos el contador
    Sql = "UPDATE ariconta" & vParam.Numconta & ".contadores set "
    If EjercicioActual Then
        Sql = Sql & " contado1="
    Else
        Sql = Sql & " contado2="
    End If
    Sql = Sql & C1
    Sql = Sql & " WHERE TipoRegi = '0'"
    Conn.Execute Sql
    
    ConseguirContador = C1
    
    Exit Function
Err1:
    PosibleError = "Error: " & Err.Number & " : " & Err.Description
    Err.Clear
End Function

Private Function CadenaImporte(VaAlDebe As Boolean, ByRef Importe As Currency, ElImporteEsCero As Boolean, Abononeg As Boolean) As String
Dim CadImporte As String

'Si va al debe, pero el importe es negativo entonces va al haber a no ser que la contabilidad admita importes negativos
    If Importe < 0 Then
        If Not Abononeg Then
            VaAlDebe = Not VaAlDebe
            Importe = Abs(Importe)
        End If
    End If
    ElImporteEsCero = (Importe = 0)
    CadImporte = TransformaComasPuntos(CStr(Importe))
    If VaAlDebe Then
        CadenaImporte = CadImporte & ",NULL"
    Else
        CadenaImporte = "NULL," & CadImporte
    End If
End Function


