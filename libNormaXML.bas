Attribute VB_Name = "libNormaXML"
Option Explicit



'//////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////
'
'
'
' SEPA en XML
'
'
'
'//////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////


Dim NFic As Integer   'Para no tener que pasarlo a todas las funciones

Private Function XML(CADENA As String) As String
Dim I As Integer
Dim Aux As String
Dim Le As String
Dim C As Integer
    'Carácter no permitido en XML  Representación ASCII
    '& (ampersand)          &amp;
    '< (menor que)          &lt;
    ' > (mayor que)         &gt;
    '“ (dobles comillas)    &quot;
    '' (apóstrofe)          &apos;
    
    'La ISO recomienda trabajar con los carcateres:
    'a b c d e f g h i j k l m n o p q r s t u v w x y z
    'A B C D E F G H I J K L M N O P Q R S T U V W X Y Z
    '0 1 2 3 4 5 6 7 8 9
    '/ - ? : ( ) . , ' +
    'Espacio
    Aux = ""
    For I = 1 To Len(CADENA)
        Le = Mid(CADENA, I, 1)
        C = Asc(Le)
        
        
        Select Case C
        Case 40 To 57
            'Caracteres permitidos y numeros
            
        Case 65 To 90
            'Letras mayusculas
            
        Case 97 To 122
            'Letras minusculas
            
        Case 32
            'espacio en balanco
            
        Case Else
            Le = " "
        End Select
        Aux = Aux & Le
    Next
    XML = Aux
End Function


Public Function GeneraFicheroNorma34SEPA_XML(CIF As String, Fecha As Date, CuentaPropia2 As String, NumeroTransferencia As Long, Pagos As Boolean, ConceptoTr As String, Anyo As String) As Boolean
Dim Regs As Integer
Dim Importe As Currency
Dim Im As Currency
Dim cad As String
Dim Aux As String
Dim SufijoOEM As String
Dim NFic As Integer
Dim EsPersonaJuridica2 As Boolean

    On Error GoTo EGen3
    GeneraFicheroNorma34SEPA_XML = False
    
'    NFic = -1
    
    
    'Cargamos la cuenta
    cad = "Select * from bancos where codmacta='" & CuentaPropia2 & "'"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        cad = ""
    Else
        If IsNull(miRsAux!IBAN) Then
            cad = ""
        Else
            SufijoOEM = "000" ''Sufijo3414
            cad = miRsAux!IBAN
            If DBLet(miRsAux!Sufijo3414, "T") <> "" Then SufijoOEM = Right("000" & miRsAux!Sufijo3414, 3)
            CuentaPropia2 = cad
        End If
        
        
    End If
    miRsAux.Close
  
    If cad = "" Then
        MsgBox "Error leyendo datos para: " & CuentaPropia2, vbExclamation
        Exit Function
    End If
    
    NFic = FreeFile
    CerrarFichero NFic
    Open App.Path & "\norma34.txt" For Output As #NFic
    
    
    
    
    Print #NFic, "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
    Print #NFic, "<Document xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""urn:iso:std:iso:20022:tech:xsd:pain.001.001.03"">"
    Print #NFic, "<CstmrCdtTrfInitn>"
    Print #NFic, "   <GrpHdr>"
    cad = "TRAN" & IIf(Pagos, "PAG", "ABO") & Format(NumeroTransferencia, "000000") & "F" & Format(Now, "yyyymmddThhnnss")
    Print #NFic, "      <MsgId>" & cad & "</MsgId>"
    Print #NFic, "      <CreDtTm>" & Format(Now, "yyyy-mm-ddThh:nn:ss") & "</CreDtTm>"
    
    
    If Pagos Then
        Aux = "ImpEfect - coalesce(imppagad ,0)"
        cad = "pagos"
        cad = "Select count(*),sum(" & Aux & ") FROM " & cad & " WHERE nrodocum = " & NumeroTransferencia & " and  anyodocum = " & DBSet(Anyo, "N")
    Else
        Aux = "abs(impvenci + coalesce(Gastos, 0) - coalesce(impcobro, 0))"
        cad = "cobros"
        cad = "Select count(*),sum(" & Aux & ") FROM " & cad & " WHERE transfer = " & NumeroTransferencia & " and  anyorem = " & DBSet(Anyo, "N")
    End If
    Aux = "0|0|"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(1)) Then Aux = miRsAux.Fields(0) & "|" & Format(miRsAux.Fields(1), "#.00") & "|"
    End If
    miRsAux.Close
    
    Print #NFic, "      <NbOfTxs>" & RecuperaValor(Aux, 1) & "</NbOfTxs>"
    Print #NFic, "      <CtrlSum>" & TransformaComasPuntos(RecuperaValor(Aux, 2)) & "</CtrlSum>"
    Print #NFic, "      <InitgPty>"
    Print #NFic, "         <Nm>" & XML(vEmpresa.nomempre) & "</Nm>"
    Print #NFic, "         <Id>"
    cad = Mid(CIF, 1, 1)
    
    EsPersonaJuridica2 = Not IsNumeric(cad)

    
    
    
    cad = "PrvtId"
    If EsPersonaJuridica2 Then cad = "OrgId"
    
    Print #NFic, "           <" & cad & ">"
    Print #NFic, "               <Othr>"
    Print #NFic, "                  <Id>" & CIF & SufijoOEM & "</Id>"
    Print #NFic, "               </Othr>"
    Print #NFic, "           </" & cad & ">"
    
    Print #NFic, "         </Id>"
    Print #NFic, "      </InitgPty>"
    Print #NFic, "   </GrpHdr>"

    Print #NFic, "   <PmtInf>"
    
    Print #NFic, "      <PmtInfId>" & Format(Now, "yyyymmddhhnnss") & CIF & "</PmtInfId>"
    Print #NFic, "      <PmtMtd>TRF</PmtMtd>"
    Print #NFic, "      <ReqdExctnDt>" & Format(Fecha, "yyyy-mm-dd") & "</ReqdExctnDt>"
    Print #NFic, "      <Dbtr>"
    
     'Nombre
    miRsAux.Open "Select siglasvia ,direccion ,numero ,codpobla,pobempre,provempre,provincia from empresa2"
    cad = cad & FrmtStr(vEmpresa.nomempre, 70)
    If miRsAux.EOF Then Err.Raise 513, , "Error obteniendo datos empresa(empresa2)"
    
    Print #NFic, "         <Nm>" & XML(vEmpresa.nomempre) & "</Nm>"
    Print #NFic, "         <PstlAdr>"
    Print #NFic, "            <Ctry>ES</Ctry>"

    cad = DBLet(miRsAux!siglasvia, "T") & " " & miRsAux!Direccion & " " & DBLet(miRsAux!numero, "T") & " "
    cad = cad & Trim(DBLet(miRsAux!CodPobla, "T") & " " & miRsAux!pobempre) & " "
    cad = cad & DBLet(miRsAux!provincia, "T")
    miRsAux.Close
    Print #NFic, "            <AdrLine>" & XML(Trim(cad)) & "</AdrLine>"
    
    Print #NFic, "         </PstlAdr>"
    Print #NFic, "         <Id>"
    
    Aux = "PrvtId"
    If EsPersonaJuridica2 Then Aux = "OrgId"
   
    
    Print #NFic, "            <" & Aux & ">"
    
    Print #NFic, "               <Othr>"
    Print #NFic, "                  <Id>" & CIF & SufijoOEM & "</Id>"
    Print #NFic, "               </Othr>"
    Print #NFic, "            </" & Aux & ">"
    Print #NFic, "         </Id>"
    Print #NFic, "    </Dbtr>"
    
    
    Print #NFic, "    <DbtrAcct>"
    Print #NFic, "       <Id>"
    Print #NFic, "          <IBAN>" & Trim(CuentaPropia2) & "</IBAN>"
    Print #NFic, "       </Id>"
    Print #NFic, "       <Ccy>EUR</Ccy>"
    Print #NFic, "    </DbtrAcct>"
    Print #NFic, "    <DbtrAgt>"
    Print #NFic, "       <FinInstnId>"
    
    cad = Mid(CuentaPropia2, 5, 4)
    cad = DevuelveDesdeBD("bic", "bics", "entidad", cad)
    Print #NFic, "          <BIC>" & Trim(cad) & "</BIC>"
    Print #NFic, "       </FinInstnId>"
    Print #NFic, "    </DbtrAgt>"
    
    
    
    'Para ello abrimos la tabla tmpNorma34
    If Pagos Then
'        cad = "Select pagos.*,nommacta,dirdatos,codposta,dirdatos,desprovi,pais,cuentas.despobla,bic,nifdatos from pagos "
        cad = "Select mid(pagos.iban,5,4) as entidad,mid(pagos.iban,9,4) as oficina,mid(pagos.iban,15,10) cuentaba,mid(pagos.iban,13,2) as CC,pagos.iban, "
        cad = cad & "nomprove nommacta,domprove dirdatos,cpprove codposta,pobprove despobla,impefect,pagos.codmacta,codpais,0 Gastos,imppagad,proprove desprovi"
        cad = cad & " ,NUmSerie,numfactu,fecfactu,numorden,text1csb,text2csb,bic,nifprove nifdatos from pagos"
        
        cad = cad & " left join bics on mid(pagos.iban,5,4)=bics.entidad "
        cad = cad & " WHERE nrodocum =" & NumeroTransferencia & " and anyodocum = " & DBSet(Anyo, "N")
    Else
        'ABONOS
         '
        cad = "Select mid(cobros.iban,5,4) as entidad,mid(cobros.iban,9,4) as oficina,mid(cobros.iban,15,10) cuentaba,mid(cobros.iban,13,2) as CC,cobros.iban"
        cad = cad & ",nomclien nommacta,domclien dirdatos,cpclien codposta,pobclien despobla,impvenci,cobros.codmacta,codpais,Gastos,impcobro,proclien desprovi"
        cad = cad & " ,NUmSerie,numfactu,fecfactu,numorden,text33csb,text41csb,bic,nifclien nifdatos from cobros"
        cad = cad & " LEFT JOIN bics on mid(cobros.iban,5,4)=bics.entidad "
        cad = cad & " WHERE transfer =" & NumeroTransferencia & " and anyorem = " & DBSet(Anyo, "N")
    End If
    miRsAux.Open cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    Regs = 0
    While Not miRsAux.EOF
        Print #NFic, "   <CdtTrfTxInf>"
        Print #NFic, "      <PmtId>"
        
         
        If Pagos Then
            'numfactu fecfactu numorden
            Aux = FrmtStr(miRsAux!NumFactu, 10)
            Aux = Aux & Format(miRsAux!FecFactu, "yyyymmdd") & Format(miRsAux!numorden, "000")
        
        Else
            'fecfaccl
            Aux = FrmtStr(miRsAux!NUmSerie, 3) & Format(miRsAux!NumFactu, "00000000")
            Aux = Aux & Format(miRsAux!FecFactu, "yyyymmdd") & Format(miRsAux!numorden, "000")
        End If
        
        Print #NFic, "         <EndToEndId>" & Aux & "</EndToEndId>"
        Print #NFic, "      </PmtId>"
        Print #NFic, "      <PmtTpInf>"
        If Pagos Then
            Im = DBLet(miRsAux!imppagad, "N")
            Im = miRsAux!ImpEfect - Im
            Aux = miRsAux!codmacta

        Else
            Im = Abs(miRsAux!ImpVenci + DBLet(miRsAux!Gastos, "N")) - DBLet(miRsAux!impcobro, "N")
            Aux = miRsAux!codmacta
        End If
        
        'Persona fisica o juridica
        cad = Mid(miRsAux!nifdatos, 1, 1)
        EsPersonaJuridica2 = Not IsNumeric(cad)
        'Como da problemas Cajamar, siempre ponemos Perosna juridica. Veremos
        EsPersonaJuridica2 = True
        
        
        Importe = Importe + Im
        Regs = Regs + 1
        
        Print #NFic, "          <SvcLvl><Cd>SEPA</Cd></SvcLvl>"
        'Print #NFic, "          <LclInstrm><Cd>SDCL</Cd></LclInstrm>"
        If ConceptoTr = "1" Then
            Aux = "SALA"
        ElseIf ConceptoTr = "0" Then
            Aux = "PENS"
        Else
            Aux = "TRAD"
        End If
        Print #NFic, "          <CtgyPurp><Cd>" & Aux & "</Cd></CtgyPurp>"
        Print #NFic, "       </PmtTpInf>"
        Print #NFic, "       <Amt>"
        cad = Format(Im, "#.00")
        Print #NFic, "          <InstdAmt Ccy=""EUR"">" & TransformaComasPuntos(cad) & "</InstdAmt>"
        Print #NFic, "       </Amt>"
        Print #NFic, "       <CdtrAgt>"
        Print #NFic, "          <FinInstnId>"
        cad = DBLet(miRsAux!BIC, "T")
        If cad = "" Then Err.Raise 513, , "No existe BIC: " & miRsAux!Nommacta & vbCrLf & "Entidad: " & miRsAux!Entidad
        Print #NFic, "             <BIC>" & DBLet(miRsAux!BIC, "T") & "</BIC>"
        Print #NFic, "          </FinInstnId>"
        Print #NFic, "       </CdtrAgt>"
        Print #NFic, "       <Cdtr>"
        Print #NFic, "          <Nm>" & XML(miRsAux!Nommacta) & "</Nm>"
        
        
        'Como cajamar da problemas, lo quitamos para todos
        'Print #NFic, "          <PstlAdr>"
        '
        'Cad = "ES"
        'If Not IsNull(miRsAux!PAIS) Then Cad = Mid(miRsAux!PAIS, 1, 2)
        'Print #NFic, "              <Ctry>" & Cad & "</Ctry>"
        '
        'If Not IsNull(miRsAux!dirdatos) Then Print #NFic, "              <AdrLine>" & XML(miRsAux!dirdatos) & "</AdrLine>"
        'Cad = XML(Trim(DBLet(miRsAux!codposta, "T") & " " & DBLet(miRsAux!despobla, "T")))
        'If Cad <> "" Then Print #NFic, "              <AdrLine>" & Cad & "</AdrLine>"
        'If Not IsNull(miRsAux!desprovi) Then Print #NFic, "              <AdrLine>" & XML(miRsAux!desprovi) & "</AdrLine>"
        'Print #NFic, "           </PstlAdr>"
        
        
        
        Print #NFic, "           <Id>"
        Aux = "PrvtId"
        If EsPersonaJuridica2 Then Aux = "OrgId"
      
        Print #NFic, "               <" & Aux & ">"
        Print #NFic, "                  <Othr>"
        Print #NFic, "                     <Id>" & miRsAux!nifdatos & "</Id>"
        'Da problemas.... con Cajamar
        'Print #NFic, "                     <Issr>NIF</Issr>"
        Print #NFic, "                  </Othr>"
        Print #NFic, "               </" & Aux & ">"
        Print #NFic, "           </Id>"
        Print #NFic, "        </Cdtr>"
        Print #NFic, "        <CdtrAcct>"
        Print #NFic, "           <Id>"
        Print #NFic, "              <IBAN>" & IBAN_Destino & "</IBAN>"
        Print #NFic, "           </Id>"
        Print #NFic, "        </CdtrAcct>"
        Print #NFic, "      <Purp>"
        
        If ConceptoTr = "1" Then
            Aux = "SALA"
        ElseIf ConceptoTr = "0" Then
            Aux = "PENS"
        Else
            Aux = "TRAD"
        End If
        
        Print #NFic, "         <Cd>" & Aux & "</Cd>"
        Print #NFic, "      </Purp>"
        Print #NFic, "      <RmtInf>"
        'Print #NFic, "      <Ustrd>ESTE ES EL CONCEPTO, POR TANTO NO SE SI SERA EL TEXTO QUE IRA DONDE TIENE QUE IR, O EN OTRO LADAO... A SABER. LO QUE ESTA CLARO ES QUE VA.</Ustrd>
        
        If Pagos Then
            ''`text1csb` `text2csb`
            Aux = DBLet(miRsAux!text1csb, "T") & " " & DBLet(miRsAux!text2csb, "T")
        Else
            '`text33csb` `text41csb`
            Aux = DBLet(miRsAux!text33csb, "T") & " " & DBLet(miRsAux!text41csb, "T")
        End If
        If Trim(Aux) = "" Then Aux = miRsAux!Nommacta
        Print #NFic, "         <Ustrd>" & XML(Trim(Aux)) & "</Ustrd>"
        Print #NFic, "      </RmtInf>"
        Print #NFic, "   </CdtTrfTxInf>"
 
       
    
            
        miRsAux.MoveNext
    Wend
    Print #NFic, "   </PmtInf>"
    Print #NFic, "</CstmrCdtTrfInitn></Document>"
    
    
    miRsAux.Close
    Set miRsAux = Nothing
    
    Close #NFic
    
    NFic = -1
    
    If Regs > 0 Then GeneraFicheroNorma34SEPA_XML = True
    Exit Function
    
EGen3:
    MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
'    If NFic > 0 Then Close (NFic)
    CerrarFichero NFic
    
End Function


Public Function GrabarDisketteNorma19_SEPA_XML(NomFichero As String, Remesa_ As String, FecPre As String, TipoReferenciaCliente As Byte, Sufijo As String, FechaCobro As String, SEPA_EmpresasGraboNIF As Boolean, Norma19_15 As Boolean, DatosBanco As String, NifEmpresa As String, esAnticipoCredito As Boolean) As Boolean
    Dim ValorEnOpcionales As Boolean
    '-- Genera_Remesa: Esta función genera la remesa indicada, en el fichero correspondiente
    
    
    Dim Sql As String
    Dim ImpEfe As Currency

    '
    Dim IdDeudor As String
    Dim Cuenta As String
    Dim Fecha2 As Date
    Dim FinFecha As Boolean


    Dim EsPersonaJuridica As Boolean
    
    Dim J As Integer
    'Dim IdNorma As String  '1914 o 1915
    
    On Error GoTo Err_Remesa19sepa
    
    

    '-- Abrir el fichero a enviar
    NFic = FreeFile()
    Open NomFichero For Output As #NFic
    
    Sql = "select  numserie,numfactu,fecfactu,numorden,cobros.codmacta,codrem,anyorem,Tiporem,"
    
    'SEPTIEMBRE 2015
    'Todos van a la fecha de vencimiento
'    If vParam.Norma19xFechaVto Then
'        SQL = SQL & " fecvenci"
'    Else
'        SQL = SQL & "'" & Format(FecCobro, FormatoFecha) & "'"
'    End If
    'OCTUBRE. Si no indica fecha cobro, ira cada una con su vencimiento, si no la fecha de cobro
    
    If FechaCobro = "" Then
        Sql = Sql & " fecvenci"
    Else
        Sql = Sql & "'" & Format(FechaCobro, FormatoFecha) & "'"
    End If

    
    Sql = Sql & " as fecvenci,impvenci,ctabanc1,"
    Sql = Sql & "text33csb,text41csb,cobros.gastos,cobros.iban,"
    Sql = Sql & "cobros.nomclien,cobros.nifclien,cobros.domclien,cobros.cpclien,cobros.pobclien,cobros.proclien,cobros.codpais,bics.bic,cobros.referencia,cuentas.SEPA_Refere,cuentas.SEPA_FecFirma  from cobros"
    Sql = Sql & "  left join bics on mid(cobros.iban,5,4)=bics.entidad inner join cuentas on "
    Sql = Sql & " cobros.codmacta = cuentas.codmacta WHERE "
    Sql = Sql & " codrem = " & RecuperaValor(Remesa_, 1)
    Sql = Sql & " AND anyorem=" & RecuperaValor(Remesa_, 2)
    
    
    'sepa
    Sql = Sql & " order by  fecvenci,nifdatos,cobros.codmacta"
    
    
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Sql = ""
    If Not miRsAux.EOF Then
        
            'Encabezado
            Print #NFic, "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
            Print #NFic, "<Document xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""urn:iso:std:iso:20022:tech:xsd:pain.008.001.02"">"
            Print #NFic, "<CstmrDrctDbtInitn>"
                        
            Print #NFic, "<GrpHdr>"
            
            If esAnticipoCredito Then
                Sql = "FSDD"
            Else
                Sql = "PRE"
            End If
            
            Sql = Sql & Format(Now, "yyyymmddhhnnss")
            
            'Los milisegundos
            Sql = Sql & Format((Timer - Int(Timer)) * 10000, "0000") & "0"
            'Idententificacion propia
            '   tiporem,codrem,anyorem
            Sql = Sql & "RE" & miRsAux!Tiporem & Format(miRsAux!CodRem, "000000") & Format(miRsAux!AnyoRem, "0000")
                    
            
            Print #NFic, "<MsgId>" & Sql & "</MsgId>"
            
            Sql = Format(Now, "yyyy-mm-dd") & "T" & Format(Now, "hh:mm:ss")   '<CreDtTm>2015-09-10T16:26:56</CreDtTm>
            Print #NFic, "   <CreDtTm>" & Sql & "</CreDtTm>"
            
            'Control sumatorio y numero de registro
            
            Sql = " codrem = " & RecuperaValor(Remesa_, 1) & " AND anyorem=" & RecuperaValor(Remesa_, 2) & " AND 1"
            Sql = DevuelveDesdeBD("concat(count(*),'|',sum(coalesce(gastos,0)+impvenci),'|')", "cobros", Sql, "1")
            Print #NFic, "   <NbOfTxs>" & RecuperaValor(Sql, 1) & "</NbOfTxs>"
            Sql = RecuperaValor(Sql, 2)
            Print #NFic, "   <CtrlSum>" & Sql & "</CtrlSum>"
            
            
            'Empezamos datos
            Print #NFic, "   <InitgPty>"
            Print #NFic, "     <Nm>" & XML(vEmpresa.nomempre) & "</Nm>"
            Print #NFic, "     <Id>"
                        
            'Tipo identificador deudor.  Persona fisica (2) o juridica (1)
            Sql = Mid(NifEmpresa, 1, 1)
            EsPersonaJuridica = Not IsNumeric(Sql)
            If EsPersonaJuridica Then
                Print #NFic, "        <OrgId>"
            Else
                Print #NFic, "        <PrvtId>"
            End If
            
            Sql = Trim(NifEmpresa) + "ES00"   'Identificacion acreedor
            Sql = CadenaTextoMod97(Sql)
            'Si no es dos digitos es un mensaje de error
            If Len(Sql) <> 2 Then Err.Raise 513, , Sql
            Sql = "ES" & Sql & Sufijo & NifEmpresa
            Print #NFic, "           <Othr>"
            Print #NFic, "              <Id>" & Sql & "</Id>"   'Ejemplo: ES3100024348588Y
            Print #NFic, "           </Othr>"
            
            If EsPersonaJuridica Then
                Print #NFic, "        </OrgId>"
            Else
                Print #NFic, "        </PrvtId>"
            End If
            
            
            Print #NFic, "      </Id>"
            Print #NFic, "   </InitgPty>"
            Print #NFic, "</GrpHdr>"
        
        
            
            
            Fecha2 = "01/01/1900"
            FinFecha = False
            While Not miRsAux.EOF
            
                'Informacion del PAGO.
                ' Se imprime una vez cada FECHA
                If Fecha2 <> miRsAux!FecVenci Then
                        
                        If Fecha2 > CDate("01/02/1900") Then Print #NFic, "</PmtInf>"
                        Fecha2 = miRsAux!FecVenci
                        
                        
                        'Previo envio vtos
                       Print #NFic, "<PmtInf>"

                        'SQL = "RE" & miRsAux!Tiporem & Format(miRsAux!CodRem, "000000") & Format(miRsAux!AnyoRem, "0000") & " " & Format(Fecha2, "dd/mm/yyyy")
                        Sql = "RE" & Format(miRsAux!CodRem, "00000") & Format(miRsAux!AnyoRem, "0000") & " " & Format(FecPre, "dd/mm/yy") & NifEmpresa
                        
                        Print #NFic, "   <PmtInfId>" & Sql & "</PmtInfId>"
                        Print #NFic, "   <PmtMtd>DD</PmtMtd>"             'DirectDebit
                        Print #NFic, "   <BtchBookg>false</BtchBookg>"    'True: un apunte por cada recib   False: Por el total
                        Print #NFic, "   <PmtTpInf>"
                        Print #NFic, "      <SvcLvl>"
                        Print #NFic, "          <Cd>SEPA</Cd>"
                        Print #NFic, "      </SvcLvl>"
                        Print #NFic, "      <LclInstrm>"
                        Print #NFic, "         <Cd>COR1</Cd>"   'CORE o COR1
                        Print #NFic, "      </LclInstrm>"
                        Print #NFic, "      <SeqTp>RCUR</SeqTp>"
                        Print #NFic, "      <CtgyPurp>"
                        Print #NFic, "         <Cd>TRAD</Cd>"
                        Print #NFic, "      </CtgyPurp>"
                        Print #NFic, "   </PmtTpInf>"
                        'Print #NFic, "   <ReqdColltnDt>" & Format(FecCobro, "yyyy-mm-dd") & "</ReqdColltnDt>"
                        Print #NFic, "   <ReqdColltnDt>" & Format(Fecha2, "yyyy-mm-dd") & "</ReqdColltnDt>"
                        Print #NFic, "   <Cdtr>"
                        Print #NFic, "      <Nm>" & XML(vEmpresa.nomempre) & "</Nm>"
                        Print #NFic, "      <PstlAdr>"
                        Print #NFic, "          <Ctry>ES</Ctry>"
                        
                        Dim RsDirec As ADODB.Recordset
                        Dim SqlDirec As String
                        Dim Direccion As String
                        
                        Direccion = ""
                        
                        SqlDirec = "select direccion, numero, escalera, piso, puerta from empresa2"
                        Set RsDirec = New ADODB.Recordset
                        RsDirec.Open SqlDirec, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        If Not RsDirec.EOF Then
                            Direccion = DBLet(RsDirec!Direccion) & " " & DBLet(RsDirec!numero) & " " & DBLet(RsDirec!escalera) & " " & DBLet(RsDirec!piso) & " " & DBLet(RsDirec!puerta)
                        End If
                        Set RsDirec = Nothing
                        
                        Sql = Direccion
                        If Sql <> "" Then Print #NFic, "          <AdrLine>" & XML(Sql) & "</AdrLine>"
                        Print #NFic, "      </PstlAdr>"
                        Print #NFic, "   </Cdtr>"
                        Print #NFic, "   <CdtrAcct>"
                        Print #NFic, "      <Id>"
                        'IBAN

                        Print #NFic, "         <IBAN>" & DatosBanco & "</IBAN>"
                        Print #NFic, "      </Id>"
                        Print #NFic, "   </CdtrAcct>"
                        Print #NFic, "   <CdtrAgt>"
                        Print #NFic, "      <FinInstnId>"
                        Sql = Mid(DatosBanco, 5, 4)
                        Sql = DevuelveDesdeBD("bic", "bics", "entidad", Sql)
                        Print #NFic, "         <BIC>" & Trim(Sql) & "</BIC>"
                        Print #NFic, "      </FinInstnId>"
                        Print #NFic, "   </CdtrAgt>"
                        
                        Print #NFic, "   <CdtrSchmeId>"
                        Print #NFic, "       <Id>"
                        Print #NFic, "          <PrvtId>"
                        Print #NFic, "             <Othr>"
                        
                        Sql = Trim(NifEmpresa) + "ES00"   'Identificacion acreedor
                        Sql = CadenaTextoMod97(Sql)
                        'Si no es dos digitos es un mensaje de error
                        If Len(Sql) <> 2 Then Err.Raise 513, , Sql
                        Sql = "ES" & Sql & Sufijo & NifEmpresa
                        Print #NFic, "                 <Id>" & Sql & "</Id>"
                        Print #NFic, "                 <SchmeNm><Prtry>SEPA</Prtry></SchmeNm>"
                        Print #NFic, "             </Othr>"
                        Print #NFic, "          </PrvtId>"
                        Print #NFic, "       </Id>"
                        Print #NFic, "   </CdtrSchmeId>"
                End If
                
            
            
            
            
                'Tipo identificador deudor.  Persona fisica (2) o juridica (1)
                Sql = Mid(miRsAux!nifclien, 1, 1)
                EsPersonaJuridica = Not IsNumeric(Sql)
                
                
                
            
            
                Print #NFic, "   <DrctDbtTxInf>"
                Print #NFic, "      <PmtId>"
                
                'Referencia del adeudo
                Sql = FrmtStr(miRsAux!codmacta, 10) & FrmtStr(miRsAux!NUmSerie, 3) & Format(miRsAux!NumFactu, "00000000")
                Sql = Sql & Format(miRsAux!FecFactu, "yyyymmdd") & Format(miRsAux!numorden, "000")
                Sql = FrmtStr(Sql, 35)
                Print #NFic, "          <EndToEndId>" & Sql & "</EndToEndId>"
                Print #NFic, "      </PmtId>"
                
                
                ImpEfe = DBLet(miRsAux!Gastos, "N")
                ImpEfe = miRsAux!ImpVenci + ImpEfe
                Sql = TransformaComasPuntos(Format(ImpEfe, "####0.00"))
                Print #NFic, "      <InstdAmt Ccy=""EUR"">" & Sql & "</InstdAmt>"
                Print #NFic, "      <DrctDbtTx>"
                Print #NFic, "         <MndtRltdInf>"
                
                'Si la cuenta tiene ORDEN de mandato, coge este
                Sql = DBLet(miRsAux!SEPA_Refere, "T")
                If Sql = "" Then
                    Select Case TipoReferenciaCliente
                    Case 0
                        'ALZIRA. La referencia final de 12 es el ctan bancaria del cli + su CC
                        Sql = Mid(DBLet(miRsAux!IBAN), 13, 2) ' Dígitos de control
                        Sql = Sql & Mid(DBLet(miRsAux!IBAN), 15, 10) ' Código de cuenta
                    Case 1
                        'NIF
                        Sql = DBLet(miRsAux!nifclien, "T")
                        
                    Case 2
                        'Referencia en el VTO. No es Nula
                        Sql = DBLet(miRsAux!Referencia, "T")
                        
                    End Select
                End If
                Print #NFic, "            <MndtId>" & Sql & "</MndtId>"   'Orden de mandato
                
                'Si tiene fecha firma de mandato
                Sql = "2009-10-31"
                If Not IsNull(miRsAux!SEPA_FecFirma) Then Sql = Format(miRsAux!SEPA_FecFirma, "yyyy-mm-dd")
                Print #NFic, "            <DtOfSgntr>" & Sql & "</DtOfSgntr>"
                
                Print #NFic, "         </MndtRltdInf>"
                Print #NFic, "      </DrctDbtTx>"
                Print #NFic, "      <DbtrAgt>"
                Print #NFic, "         <FinInstnId>"
                Sql = FrmtStr(DBLet(miRsAux!BIC, "T"), 11)
                Print #NFic, "            <BIC>" & Sql & "</BIC>"
                Print #NFic, "         </FinInstnId>"
                Print #NFic, "      </DbtrAgt>"
                Print #NFic, "      <Dbtr>"
                
                Print #NFic, "         <Nm>" & XML(miRsAux!nomclien) & "</Nm>"
                Print #NFic, "         <PstlAdr>"
                
                Sql = "ES"
                If Not IsNull(miRsAux!codPAIS) Then Sql = Mid(miRsAux!codPAIS, 1, 2)
                Print #NFic, "            <Ctry>" & Sql & "</Ctry>"
                
                
                If Not IsNull(miRsAux!domclien) Then Print #NFic, "              <AdrLine>" & XML(miRsAux!domclien) & "</AdrLine>"
                
                Sql = ""
                'SQL = XML(Trim(DBLet(miRsAux!codposta, "T") & " " & DBLet(miRsAux!despobla, "T")))
                'If SQL <> "" Then Print #NFic, "              <AdrLine>" & SQL & "</AdrLine>"If Not IsNull(miRsAux!desprovi) Then Print #NFic, "              <AdrLine>" & XML(miRsAux!desprovi) & "</AdrLine>"
                If DBLet(miRsAux!pobclien, "T") = DBLet(miRsAux!proclien, "N") Then
                    Sql = Trim(DBLet(miRsAux!cpclien, "T") & "   " & DBLet(miRsAux!pobclien, "T"))
                
                Else
                    Sql = Trim(DBLet(miRsAux!pobclien, "T") & "   " & DBLet(miRsAux!cpclien, "T"))
                    If Not IsNull(miRsAux!proclien) Then Sql = Sql & "     " & miRsAux!proclien
                End If
                If Sql <> "" Then Print #NFic, "              <AdrLine>" & XML(Mid(Sql, 1, 70)) & "</AdrLine>"
                
                
                
                Print #NFic, "         </PstlAdr>"
                Print #NFic, "         <Id>"
                Print #NFic, "            <PrvtId>"
                Print #NFic, "               <Othr>"
                
                
                'Opcion nueva: 3   Quiere el campo referencia de cobros
'??             SQL = DBLet(miRsAux!SEPA_Refere, "T")
'??             If SQL = "" Then
                   Select Case TipoReferenciaCliente
                   Case 0
                       'ALZIRA. La referencia final de 12 es el ctan bancaria del cli + su CC
                       Sql = Mid(DBLet(miRsAux!IBAN), 13, 2) ' Dígitos de control
                       Sql = Sql & Mid(DBLet(miRsAux!IBAN), 15, 10) ' Código de cuenta
                   Case 1
                       'NIF
                       Sql = DBLet(miRsAux!nifclien, "T")
                
                   Case 2
                       'Referencia en el VTO. No es Nula
                       Sql = DBLet(miRsAux!Referencia, "T")
                   
                   End Select
'??             End If
                
                Print #NFic, "                   <Id>" & Sql & "</Id>"
                If TipoReferenciaCliente = 1 Then Print #NFic, "                   <Issr>NIF</Issr>"
                Print #NFic, "               </Othr>"
                Print #NFic, "            </PrvtId>"
                Print #NFic, "         </Id>"
                Print #NFic, "      </Dbtr>"
                Print #NFic, "      <DbtrAcct>"
                Print #NFic, "         <Id>"
                
                Sql = IBAN_Destino   'Hay que poner TRUE aunque sea cobro
                Print #NFic, "            <IBAN>" & Sql & "</IBAN>"
                Print #NFic, "         </Id>"
                Print #NFic, "      </DbtrAcct>"
                Print #NFic, "      <Purp>"
                Print #NFic, "         <Cd>TRAD</Cd>"
                Print #NFic, "      </Purp>"
                Print #NFic, "      <RmtInf>"
                
                Sql = Trim(DBLet(miRsAux!text33csb, "T") & " " & FrmtStr(DBLet(miRsAux!text41csb, "T"), 60))
                If Sql = "" Then Sql = miRsAux!nomclien
                Print #NFic, "         <Ustrd>" & XML(Sql) & "</Ustrd>"
                Print #NFic, "      </RmtInf>"
                Print #NFic, "   </DrctDbtTxInf>"
        
            
            
            'Siguiente
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
        
              
              
        Print #NFic, "</PmtInf>"
        Print #NFic, "</CstmrDrctDbtInitn></Document>"
        
        
        GrabarDisketteNorma19_SEPA_XML = True
            
    End If  'De EOF
    Close #NFic
        
    
    Exit Function
Err_Remesa19sepa:

    MsgBox "Err: " & Err.Number & vbCrLf & _
        Err.Description, vbCritical, "Grabación del diskette de Remesa SEPA"
        
    CerrarFichero NFic
End Function



Private Sub CerrarFichero(nFile As Integer)

    On Error Resume Next
    
    Close #nFile
    
    Err.Clear

End Sub






Private Function IBAN_Destino() As String
    IBAN_Destino = FrmtStr(DBLet(miRsAux!IBAN, "T"), 4) ' ES00
    IBAN_Destino = IBAN_Destino & Mid(DBLet(miRsAux!IBAN, "T"), 5, 4) ' Código de entidad receptora
    IBAN_Destino = IBAN_Destino & Mid(DBLet(miRsAux!IBAN, "T"), 9, 4) ' Código de oficina receptora
    IBAN_Destino = IBAN_Destino & Mid(DBLet(miRsAux!IBAN, "T"), 13, 2) ' Dígitos de control
    IBAN_Destino = IBAN_Destino & Mid(DBLet(miRsAux!IBAN, "T"), 15, 10) ' Código de cuenta
End Function












'Devolucion SEPA
'
Public Sub ProcesaFicheroDevolucionSEPA_XML(Fichero As String, ByRef Remesa As String)
Dim aux2 As String  'Para buscar los vencimientos
Dim FinLecturaLineas As Boolean

Dim ErroresVto As String

Dim posicion As Long
Dim L2 As Long
Dim Sql As String
Dim ContenidoFichero As String
Dim NF As Integer
Dim CadenaComprobacionDevueltos As String  'cuantos|importe|


    On Error GoTo eProcesaCabeceraFicheroDevolucionSEPA_XML
    Remesa = ""
    
    
    
   

    NF = FreeFile
    Open Fichero For Input As #NF
    ContenidoFichero = ""
    While Not EOF(NF)
        Line Input #NF, aux2
        ContenidoFichero = ContenidoFichero & aux2
    Wend
    Close #NF
    
    Set miRsAux = New ADODB.Recordset
    
    'Vamos a obtener el ID de la remesa  enviada
    ' Buscaremos la linea
    'Idententificacion propia  Ejemplo: <MsgId>PRE2015093012481641020RE10000802015</MsgId>  de donde RE mesa, 1 tipo 000080 Nº   ano;2015
    posicion = PosicionEnFichero(1, ContenidoFichero, "<CstmrPmtStsRpt>")
    
    posicion = PosicionEnFichero(posicion, ContenidoFichero, "<OrgnlMsgId>")
    L2 = PosicionEnFichero(posicion, ContenidoFichero, "</OrgnlMsgId>")
    
    aux2 = Mid(ContenidoFichero, posicion, L2 - posicion)
    aux2 = Mid(aux2, InStr(10, aux2, "RE") + 3) 'QUTIAMO EL RE y ye tipo RE1(ejemp)
    
    'Los 4 ultimos son año
    Remesa = Mid(aux2, 1, 6) & "|" & Mid(aux2, 7, 4) & "|"
    
    
    'Voy a buscar el numero total de vencimientos que devuelven y el importe total(comproabacion ultima
    posicion = PosicionEnFichero(posicion, ContenidoFichero, "<OrgnlPmtInfAndSts>")
    posicion = PosicionEnFichero(posicion, ContenidoFichero, "<OrgnlNbOfTxs>")
    '<OrgnlNbOfTxs>1</OrgnlNbOfTxs>
    L2 = PosicionEnFichero(posicion, ContenidoFichero, "</OrgnlNbOfTxs>")
    CadenaComprobacionDevueltos = Mid(ContenidoFichero, posicion, L2 - posicion)
    
    '<OrgnlCtrlSum>5180.98</OrgnlCtrlSum>
    posicion = PosicionEnFichero(posicion, ContenidoFichero, "<OrgnlCtrlSum>")
    L2 = PosicionEnFichero(posicion, ContenidoFichero, "</OrgnlCtrlSum>")
    CadenaComprobacionDevueltos = CadenaComprobacionDevueltos & Mid(ContenidoFichero, posicion, L2 - posicion)
            
    
    
    'Primera comprobacion. Existe la remesa obtenida
    
    
    'Vamos con los vtos  4300106840T  0001180220150925001

    Do
        posicion = InStr(posicion, ContenidoFichero, "<TxInfAndSts>")
        If posicion > 0 Then
            
            'Si existe un grupo de registros TxInfAndSts, los de abjo deben existir SI o SI
            posicion = PosicionEnFichero(posicion, ContenidoFichero, "<OrgnlEndToEndId>")
            L2 = PosicionEnFichero(posicion, ContenidoFichero, "</OrgnlEndToEndId>")
            aux2 = Mid(ContenidoFichero, posicion, L2 - posicion)
            
            'Id del recibo devuleto. Ejemplo
            '4300106840T  0001180220150925001
            ' asi es como se monta el el generador de remesa
            '           FrmtStr(miRsAux!codmacta, 10) & FrmtStr(miRsAux!NUmSerie, 3) & Format(miRsAux!codfaccl, "00000000")
            '           Format(miRsAux!fecfaccl, "yyyymmdd") & Format(miRsAux!numorden, "000")
            
            Sql = "Select codrem,anyorem,siturem from cobros where fecfactu='" & Mid(aux2, 22, 4) & "-" & Mid(aux2, 26, 2) & "-" & Mid(aux2, 28, 2)
            Sql = Sql & "' AND numserie = '" & Trim(Mid(aux2, 11, 3)) & "' AND numfactu = " & Val(Mid(aux2, 14, 8)) & " AND numorden=" & Mid(aux2, 30, 3)

            miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Sql = Mid(Sql, InStr(1, UCase(Sql), " WHERE ") + 7)
            Sql = Replace(Sql, "fecfactu", "F.Fac:")
            Sql = Replace(Sql, "numserie", "Serie:")
            Sql = Replace(Sql, "numfactu", "NºFac:")
            Sql = Replace(Sql, "numorden", "Ord:")
            Sql = Replace(Sql, "AND", ""): Sql = Replace(Sql, "=", "")
            Sql = "Vto no encontrado: " & Mid(Sql, InStr(1, UCase(Sql), " WHERE ") + 7)
            If Not miRsAux.EOF Then
                If IsNull(miRsAux!CodRem) Then
                    Sql = "Vencimiento sin Remesa: " & aux2
                Else
        
                    Sql = ""
                End If
            End If
            miRsAux.Close
            
            If Sql <> "" Then ErroresVto = ErroresVto & vbCrLf & Sql
            
            
            posicion = InStr(posicion, ContenidoFichero, "</TxInfAndSts>") + 11 'Para que pase al siguiente registro, si es que existe
            
        
        Else
           posicion = Len(ContenidoFichero) + 1
        End If  'posicion>0 de OrgnlTxRef
        
        
    Loop Until posicion > Len(ContenidoFichero)
    

    If ErroresVto <> "" Then
        MsgBox ErroresVto, vbExclamation
        Remesa = ""
    Else
        


    
        'En aux2 tendre codrem|anñorem|
        aux2 = RecuperaValor(Remesa, 1) & " AND anyo = " & RecuperaValor(Remesa, 2)
        aux2 = "Select situacion from remesas where codigo = " & aux2
        miRsAux.Open aux2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If miRsAux.EOF Then
            aux2 = "-No se encuentra remesa"
            
        Else
            'Si que esta.
            'Situacion
            If CStr(miRsAux!Situacion) <> "Q" And CStr(miRsAux!Situacion) <> "Y" Then
                aux2 = "- Situacion incorrecta : " & miRsAux!Situacion
            Else
                aux2 = "" 'TODO OK
            End If
        End If

        If aux2 <> "" Then
            aux2 = aux2 & " ->" & Mid(miRsAux.Source, InStr(1, UCase(miRsAux.Source), " WHERE ") + 7)
            aux2 = Replace(aux2, " AND ", " ")
            aux2 = Replace(aux2, "anyo", "año")
            ErroresVto = ErroresVto & vbCrLf & aux2
        End If
        miRsAux.Close

    
    


        If ErroresVto <> "" Then
            aux2 = "Error remesas " & vbCrLf & String(30, "=") & ErroresVto
            MsgBox aux2, vbExclamation

            'Pongo REMESA=""
            Remesa = "" 'para que no continue el preoceso de devolucion
        End If

    End If
    Set miRsAux = Nothing
    Exit Sub
eProcesaCabeceraFicheroDevolucionSEPA_XML:
    Remesa = ""
    MuestraError Err.Number, "Procesando fichero devolucion SEPA XML" & Err.Description
    Set miRsAux = Nothing
End Sub

'Si no se encuentra lo que busco saltara un error
Private Function PosicionEnFichero(ByVal Inicio As Long, ContenidoDelFichero As String, QueBusco As String) As Long
    PosicionEnFichero = InStr(Inicio, ContenidoDelFichero, QueBusco)
    If PosicionEnFichero = 0 Then
        Err.Raise 513, "No se encuentra cadena: " & QueBusco
    Else
        If InStr(1, QueBusco, "</") Then
            'PosicionEnFichero = PosicionEnFichero - Len(QueBusco)
        Else
            PosicionEnFichero = PosicionEnFichero + Len(QueBusco)
        End If
    End If
        
End Function


'XML
Public Sub ProcesaLineasFicheroDevolucionXML(Fichero As String, ByRef Listado As Collection)
Dim NF As Integer
Dim ContenidoFichero As String
Dim posicion As Long
Dim L2 As Long
Dim aux2 As String

    NF = FreeFile
    Open Fichero For Input As #NF
    ContenidoFichero = ""
    While Not EOF(NF)
        Line Input #NF, aux2
        ContenidoFichero = ContenidoFichero & aux2
    Wend
    Close #NF
    
   
    posicion = 1
    Do
        posicion = InStr(posicion, ContenidoFichero, "<TxInfAndSts>")
        If posicion > 0 Then
            
            'Si existe un grupo de registros TxInfAndSts, los de abjo deben existir SI o SI
            posicion = PosicionEnFichero(posicion, ContenidoFichero, "<OrgnlEndToEndId>")
            L2 = PosicionEnFichero(posicion, ContenidoFichero, "</OrgnlEndToEndId>")
            aux2 = Mid(ContenidoFichero, posicion, L2 - posicion)
            
            'Id del recibo devuleto. Ejemplo
            '4300106840T  0001180220150925001
            ' asi es como se monta el el generador de remesa
            '           FrmtStr(miRsAux!codmacta, 10) & FrmtStr(miRsAux!NUmSerie, 3) & Format(miRsAux!codfaccl, "00000000")
            '           Format(miRsAux!fecfaccl, "yyyymmdd") & Format(miRsAux!numorden, "000")
            
            'Vamos a guardar en el col la linea en formato antiguo SEPA y asi no toco el programa
            'M  0330047820131201001   430000061
            aux2 = Mid(aux2, 11, 23) & "   " & Mid(aux2, 1, 10)
            Listado.Add aux2
            posicion = InStr(posicion, ContenidoFichero, "</TxInfAndSts>") + 11 'Para que pase al siguiente registro, si es que existe
            
        
        Else
           posicion = Len(ContenidoFichero) + 1
        End If  'posicion>0 de OrgnlTxRef
        
    Loop Until posicion > Len(ContenidoFichero)
    
End Sub


Public Sub LeerLineaDevolucionSEPA_XML(Fichero As String, ByRef Remesa As String, ByRef lwCobros As ListView)
Dim aux2 As String  'Para buscar los vencimientos
Dim AUX3 As String
Dim FinLecturaLineas As Boolean

Dim ErroresVto As String

Dim posicion As Long
Dim L2 As Long
Dim Sql As String
Dim ContenidoFichero As String
Dim NF As Integer
Dim CadenaComprobacionDevueltos As String  'cuantos|importe|

Dim VtoEncontrado As Boolean
Dim DatosXMLVto As String
Dim Itm As ListItem
Dim Rs As ADODB.Recordset

Dim RegistroErroneo As Boolean


    On Error GoTo eLeerLineaDevolucionSEPA_XML
    Remesa = ""
    
   

    NF = FreeFile
    Open Fichero For Input As #NF
    ContenidoFichero = ""
    While Not EOF(NF)
        Line Input #NF, aux2
        ContenidoFichero = ContenidoFichero & aux2
    Wend
    Close #NF
    
    
    
    
    
    'Comprobacion 1
    'El NIF del fichero enviado es el de la empresa
    'Lo busco acotandolo por etiquetas XML
    posicion = PosicionEnFichero(1, ContenidoFichero, "<OrgnlPmtInfAndSts>")
    L2 = PosicionEnFichero(posicion, ContenidoFichero, "</OrgnlPmtInfAndSts>")
    If posicion > 0 And L2 > 0 Then
        '
        aux2 = Mid(ContenidoFichero, posicion, L2 - posicion)
        posicion = PosicionEnFichero(1, aux2, "<StsRsnInf>")
        L2 = PosicionEnFichero(posicion, aux2, "</StsRsnInf>")
        
        If posicion > 0 And L2 > 0 Then
            aux2 = Mid(aux2, posicion, L2 - posicion)
            posicion = PosicionEnFichero(1, aux2, "<Id>ES")   'de momento todos los clientes seran de españa
            L2 = PosicionEnFichero(posicion, aux2, "</Id>")
    
            aux2 = Mid(aux2, posicion, L2 - posicion)
            If Len(aux2) > 5 Then
                Sql = DevuelveDesdeBD("nifempre", "empresa2", "1", "1")
                'Es CCSSSNNNNNN
                '   contro
                '     SUFIJO
                '        NIF
                aux2 = Mid(aux2, 6)
                If aux2 <> Sql Then
'                    Stop
                    Err.Raise 513, , "NIF empresa del fichero no coincide con el de la empresa en Ariconta"
                End If
            End If
        End If
    End If
    
    Set miRsAux = New ADODB.Recordset
    
    'Vamos a obtener el ID de la remesa  enviada
    ' Buscaremos la linea
    'Idententificacion propia  Ejemplo: <MsgId>PRE2015093012481641020RE10000802015</MsgId>  de donde RE mesa, 1 tipo 000080 Nº   ano;2015
    posicion = PosicionEnFichero(1, ContenidoFichero, "<CstmrPmtStsRpt>")
    
    posicion = PosicionEnFichero(posicion, ContenidoFichero, "<OrgnlMsgId>")
    L2 = PosicionEnFichero(posicion, ContenidoFichero, "</OrgnlMsgId>")
    
    aux2 = Mid(ContenidoFichero, posicion, L2 - posicion)
    aux2 = Mid(aux2, InStr(10, aux2, "RE") + 3) 'QUTIAMO EL RE y ye tipo RE1(ejemp)
    
    'Los 4 ultimos son año
    Remesa = Mid(aux2, 1, 6) & "|" & Mid(aux2, 7, 4) & "|"
    
    
    'Voy a buscar el numero total de vencimientos que devuelven y el importe total(comproabacion ultima
    posicion = PosicionEnFichero(posicion, ContenidoFichero, "<OrgnlPmtInfAndSts>")
    posicion = PosicionEnFichero(posicion, ContenidoFichero, "<OrgnlNbOfTxs>")
    '<OrgnlNbOfTxs>1</OrgnlNbOfTxs>
    L2 = PosicionEnFichero(posicion, ContenidoFichero, "</OrgnlNbOfTxs>")
    CadenaComprobacionDevueltos = Mid(ContenidoFichero, posicion, L2 - posicion) & "|"
    
    '<OrgnlCtrlSum>5180.98</OrgnlCtrlSum>
    posicion = PosicionEnFichero(posicion, ContenidoFichero, "<OrgnlCtrlSum>")
    L2 = PosicionEnFichero(posicion, ContenidoFichero, "</OrgnlCtrlSum>")
    CadenaComprobacionDevueltos = CadenaComprobacionDevueltos & Mid(ContenidoFichero, posicion, L2 - posicion) & "|"
            
    'Primera comprobacion. Existe la remesa obtenida
    
    
    'Vamos con los vtos  4300106840T  0001180220150925001
    
    Dim jj As Long
    jj = 1
    Set Rs = New ADODB.Recordset
    
    Do
        posicion = InStr(posicion, ContenidoFichero, "<TxInfAndSts>")
        If posicion > 0 Then
            L2 = PosicionEnFichero(posicion, ContenidoFichero, "</TxInfAndSts>")
            DatosXMLVto = Mid(ContenidoFichero, posicion, L2 - posicion)
            
            ContenidoFichero = Mid(ContenidoFichero, L2 + 14)
            
            
            'Si existe un grupo de registros TxInfAndSts, los de abjo deben existir SI o SI
            posicion = PosicionEnFichero(1, DatosXMLVto, "<OrgnlEndToEndId>")
            L2 = PosicionEnFichero(posicion, DatosXMLVto, "</OrgnlEndToEndId>")
            aux2 = Mid(DatosXMLVto, posicion, L2 - posicion)
            
            
            Set Itm = lwCobros.ListItems.Add(, "C" & jj)
            Itm.Text = Trim(Mid(aux2, 11, 3))  'miRsAux!NUmSerie
            
            Itm.SubItems(1) = Mid(aux2, 14, 8) ' numfactu
            Itm.SubItems(2) = Mid(aux2, 30, 3) ' miRsAux!numorden
            Itm.SubItems(3) = Mid(aux2, 1, 10) 'miRsAux!codmacta
            Itm.Tag = Format(Mid(aux2, 22, 4) & "-" & Mid(aux2, 26, 2) & "-" & Mid(aux2, 28, 2), "dd/mm/yyyy")
            
            Itm.SubItems(8) = RecuperaValor(Remesa, 1) ' remesa
            Itm.SubItems(9) = RecuperaValor(Remesa, 2) ' año de remesa
            Itm.SubItems(10) = DevuelveValor("select codmacta from remesas where codigo = " & RecuperaValor(Remesa, 1) & " and anyo = " & RecuperaValor(Remesa, 2))
            
            Sql = "select * from cobros where "
            Sql = Sql & " numserie = " & DBSet(Trim(Mid(aux2, 11, 3)), "T") & " and numfactu = " & DBSet(Val(Mid(aux2, 14, 8)), "N")
            Sql = Sql & " and fecfactu = '" & Mid(aux2, 22, 4) & "-" & Mid(aux2, 26, 2) & "-" & Mid(aux2, 28, 2) & "'"
            
            Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            VtoEncontrado = False
            If Not Rs.EOF Then
                Itm.SubItems(4) = DBLet(Rs!nomclien, "T")    'miRsAux!nomclien
                If Rs!Devuelto = 1 Then
                    Itm.Bold = True
                    Itm.ForeColor = vbRed
                End If
                VtoEncontrado = True
            Else
                Itm.SubItems(4) = " "    'miRsAux!nomclien    'AVISAR A MONICA--> Si no pones espacio en blanco cuando lo selecciona sale raro
            End If
            
            posicion = PosicionEnFichero(1, DatosXMLVto, "<InstdAmt Ccy=""EUR"">")
            L2 = PosicionEnFichero(posicion, DatosXMLVto, "</InstdAmt>")
            AUX3 = Mid(DatosXMLVto, posicion, L2 - posicion)
            
            If posicion > 0 Then
            
            
                AUX3 = TransformaPuntosComas(AUX3)
                Itm.SubItems(5) = Format(AUX3, FormatoImporte)
                If VtoEncontrado Then
                    'El importe deberia coincidir. Si no lo marcariamos como error
'                    Stop
'                    Stop
                    
                    Dim ImporteRemesado As Currency
                    '[[[[[[[[[[[[[[PREGUNTAR a DAVID
                    'antes cobros_realizados
                    Sql = "select impcobro cobros where "
                    Sql = Sql & " numserie = " & DBSet(Trim(Mid(aux2, 11, 3)), "T") & " and numfactu = " & DBSet(Val(Mid(aux2, 14, 8)), "N")
                    Sql = Sql & " and fecfactu = '" & Mid(aux2, 22, 4) & "-" & Mid(aux2, 26, 2) & "-" & Mid(aux2, 28, 2) & "' "
                    
                    ImporteRemesado = DevuelveValor(Sql)
                    
                    If ImporteRemesado <> AUX3 Then
                    
                        MsgBox "La factura " & DBSet(Trim(Mid(aux2, 11, 3)), "T") & "-" & DBSet(Val(Mid(aux2, 14, 8)), "N") & " de fecha " & Mid(aux2, 28, 2) & "/" & Mid(aux2, 26, 2) & "/" & Mid(aux2, 22, 4) & " es de " & aux2 & " euros", vbExclamation
                    
                    Else
                        
                    End If
                End If
            Else
                Itm.SubItems(5) = " "
            End If
           
           
            'Motivo devolucion   EJEMPLO
            '<Rsn>
            '   <Cd>AM04</Cd>
            '</Rsn>
            posicion = PosicionEnFichero(1, DatosXMLVto, "<Rsn>")
            L2 = PosicionEnFichero(posicion, DatosXMLVto, "</Rsn>")
            aux2 = Mid(DatosXMLVto, posicion, L2 - posicion)
            
            posicion = PosicionEnFichero(1, DatosXMLVto, "<Cd>")
            L2 = PosicionEnFichero(posicion, DatosXMLVto, "</Cd>")
            If posicion > 0 And L2 > 0 Then
                aux2 = Mid(DatosXMLVto, posicion, L2 - posicion)
                
                aux2 = DevuelveDesdeBD("concat(codigo,' - ', descripcion)", "usuarios.wdevolucion", "codigo", aux2, "T")
                
                If aux2 = "" Then aux2 = " "
           
            Else
                'MOTIVO no encontrado
                'Ver por que
                'Ver que poner
                aux2 = " "
                
                
            End If
            Itm.SubItems(11) = aux2
           
           
            If Not VtoEncontrado Then
                Itm.ForeColor = vbRed
'                Itm.Ghosted = True
                For posicion = 1 To Itm.ListSubItems.Count
                    Debug.Print lwCobros.ColumnHeaders(posicion).Text & ":" & Itm.ListSubItems(posicion).Text
                    Itm.ListSubItems(posicion).ForeColor = vbRed
                Next
                
            Else
                Itm.Checked = True
            End If
            
            'posicion = InStr(posicion, ContenidoFichero, "</TxInfAndSts>") + 11 'Para que pase al siguiente registro, si es que existe
            posicion = 1
            jj = jj + 1 'numero de item
            Rs.Close
        Else
           posicion = Len(ContenidoFichero) + 1
        End If  'posicion>0 de OrgnlTxRef
        
        
    Loop Until posicion > Len(ContenidoFichero)
    
    
    Exit Sub
eLeerLineaDevolucionSEPA_XML:
    Remesa = ""
    MuestraError Err.Number, "Procesando fichero devolucion SEPA XML" & vbCrLf & Err.Description
    Set miRsAux = Nothing
    Set Rs = New ADODB.Recordset
           
End Sub




