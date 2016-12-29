Attribute VB_Name = "modTesoreria"
Option Explicit

Public Function CargarCobrosTemporal(Forpa As String, FecFactu As String, TotalFac As Currency) As Boolean
Dim SQL As String
Dim CadValues As String
Dim rsVenci As ADODB.Recordset
Dim FecVenci As String
Dim ImpVenci As Currency

    On Error GoTo eCargarCobros

    CargarCobrosTemporal = False

    SQL = "SELECT numerove, primerve, restoven FROM formapago WHERE codforpa=" & DBSet(Forpa, "N")
    
    Set rsVenci = New ADODB.Recordset
    rsVenci.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CadValues = ""
    
    If Not rsVenci.EOF Then
        If rsVenci.Fields(0).Value > 0 Then
            '-------- Primer Vencimiento
            I = 1
            'FECHA VTO
            FecVenci = CDate(FecFactu)
            FecVenci = DateAdd("d", DBLet(rsVenci!primerve, "N"), FecVenci)
            '===
            
            'IMPORTE del Vencimiento
            If rsVenci!numerove = 1 Then
                ImpVenci = TotalFac
            Else
                ImpVenci = Round(TotalFac / rsVenci.Fields(0).Value, 2)
                'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                If ImpVenci * rsVenci!numerove <> TotalFac Then
                    ImpVenci = Round(ImpVenci + (TotalFac - ImpVenci * rsVenci.Fields(0).Value), 2)
                End If
            End If
            CadValues = "(" & vUsu.Codigo & "," & DBSet(I, "N") & "," & DBSet(FecVenci, "F") & "," & DBSet(ImpVenci, "N") & "),"
            
            'Resto Vencimientos
            '--------------------------------------------------------------------
            For I = 2 To rsVenci!numerove
                FecVenci = DateAdd("d", DBLet(rsVenci!restoven, "N"), FecVenci)
                    
                'IMPORTE Resto de Vendimientos
                ImpVenci = Round(TotalFac / rsVenci.Fields(0).Value, 2)
                
                CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(I, "N") & "," & DBSet(FecVenci, "F") & "," & DBSet(ImpVenci, "N") & "),"
            Next I
        End If
    End If
    
    Set rsVenci = Nothing
    
    If CadValues <> "" Then
        SQL = "INSERT INTO tmpcobros (codusu, numorden, fecvenci, impvenci)"
        SQL = SQL & " VALUES " & Mid(CadValues, 1, Len(CadValues) - 1)
        Conn.Execute SQL
    End If
    
    CargarCobrosTemporal = True
    Exit Function

eCargarCobros:
    MuestraError Err.Number, "Cargar Cobros en Temporal", Err.Description
End Function


'Cargara sobre un collection los cobros.
'Cada linea el SQL
'       insert into cobros(numserie,numfactu,fecfactu,codmacta,codforpa,ctabanc1,iban,text33csb,numorden,fecvenci,impvenci)
'    Para ello enviaremos TODO el sql menos y numorden fecvenci e impvenci
Public Function CargarCobrosSobreCollectionConSQLInsert(ByRef ColCobros As Collection, Forpa As String, FecFactu As String, TotalFac As Currency, PartFijaSQL As String) As Boolean
Dim SQL As String
Dim rsVenci As ADODB.Recordset
Dim FecVenci As String
Dim ImpVenci As Currency

    On Error GoTo eCargarCobros

    CargarCobrosSobreCollectionConSQLInsert = False

    Set ColCobros = New Collection
    
    SQL = "SELECT numerove, primerve, restoven FROM formapago WHERE codforpa=" & DBSet(Forpa, "N")
    
    Set rsVenci = New ADODB.Recordset
    rsVenci.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    

    
    If Not rsVenci.EOF Then
        If rsVenci.Fields(0).Value > 0 Then
            '-------- Primer Vencimiento
            
            'FECHA VTO
            FecVenci = CDate(FecFactu)
            FecVenci = DateAdd("d", DBLet(rsVenci!primerve, "N"), FecVenci)
            '===
            
            'IMPORTE del Vencimiento
            If rsVenci!numerove = 1 Then
                ImpVenci = TotalFac
            Else
                ImpVenci = Round(TotalFac / rsVenci.Fields(0).Value, 2)
                'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                If ImpVenci * rsVenci!numerove <> TotalFac Then
                    ImpVenci = Round(ImpVenci + (TotalFac - ImpVenci * rsVenci.Fields(0).Value), 2)
                End If
            End If
            'CadValues = "(" & vUsu.Codigo & "," & DBSet(i, "N") & "," & DBSet(FecVenci, "F") & "," & DBSet(ImpVenci, "N") & "),"
            ColCobros.Add PartFijaSQL & "1," & DBSet(FecVenci, "F") & "," & DBSet(ImpVenci, "N") & ")"
            
            
            'Resto Vencimientos
            '--------------------------------------------------------------------
            For I = 2 To rsVenci!numerove
                FecVenci = DateAdd("d", DBLet(rsVenci!restoven, "N"), FecVenci)
                    
                'IMPORTE Resto de Vendimientos
                ImpVenci = Round(TotalFac / rsVenci.Fields(0).Value, 2)
                
                'CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(i, "N") & "," & DBSet(FecVenci, "F") & "," & DBSet(ImpVenci, "N") & "),"
                ColCobros.Add PartFijaSQL & I & "," & DBSet(FecVenci, "F") & "," & DBSet(ImpVenci, "N") & ")"
            Next I
        End If
    End If
    
    Set rsVenci = Nothing
    
    
    
    CargarCobrosSobreCollectionConSQLInsert = True
    Exit Function

eCargarCobros:
    MuestraError Err.Number, "Cargar Cobros auxiliar", Err.Description
End Function











Public Function BancoPropio() As String
Dim SQL As String
Dim RS As ADODB.Recordset

    BancoPropio = ""

    SQL = "select codmacta from bancos "
    
    If TotalRegistrosConsulta(SQL) = 1 Then
        Set RS = New ADODB.Recordset
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then BancoPropio = DBLet(RS!Codmacta, "T")
        Set RS = Nothing
    End If

End Function

