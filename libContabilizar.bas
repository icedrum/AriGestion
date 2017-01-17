Attribute VB_Name = "libContabilizar"
Option Explicit


















'******************************************************************************************
'******************************************************************************************
'******************************************************************************************
'
'
'   Creara un apunte a partir de un collection
'   col: 'codmacta | docum | codconce | ampliaci | imported|importeH |ctacontrpar|
'
'
'******************************************************************************************
'******************************************************************************************
'******************************************************************************************
Public Function CrearApunteDesdeColeccion(Fecha As Date, Observaciones As String, ByRef ColApuntes As Collection) As Boolean
Dim mC As ContadoresConta
Dim Actual As Boolean
Dim Cadena As String
Dim Aux As String
Dim K As Integer

    On CrearApunteDesdeColeccion GoTo eCrearApunteDesdeColeccion
    CrearApunteDesdeColeccion = False

    Set mC = New ContadoresConta
    Actual = (Fecha < DateAdd("yyyy", 1, vEmpresa.FechaInicioEjercicio))
    Cadena = ""
    If mC.ConseguirContador("0", Actual, True) = 1 Then Err.Raise 513, , "Error consiguiendo contador"
    
    'hcabapu(numdiari,fechaent,numasien,obsdiari,feccreacion,usucreacion,desdeaplicacion)
    Cadena = "INSERT INTO ariconta" & vParam.Numconta & ".hcabapu(numdiari,fechaent,numasien,obsdiari,feccreacion,usucreacion,desdeaplicacion) "
    Cadena = Cadena & " VALUES (1," & DBSet(Fecha, "F") & "," & mC.Contador & "," & DBSet(Observaciones, "T")
    Cadena = Cadena & "," & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'arigestion')"
    Conn.Execute Cadena
    
    'Las lineas
    'codmacta | docum | codconce | ampliaci | imported|importeH |ctacontrpar|
    'hlinapu(numdiari,fechaent,numasien,linliapu,codmacta,numdocum,codconce,ampconce,timporteD,timporteH,ctacontr)
    Cadena = ""
    For K = 1 To ColApuntes.Count
        Cadena = Cadena & ", (1," & DBSet(Fecha, "F") & "," & mC.Contador & "," & K & ","
        Cadena = Cadena & DBSet(RecuperaValor(ColApuntes.Item(K), 1), "T") & "," 'codmacta
        Cadena = Cadena & DBSet(RecuperaValor(ColApuntes.Item(K), 2), "T") & "," 'numdocum
        Cadena = Cadena & RecuperaValor(ColApuntes.Item(K), 3) & "," 'codconce
        Cadena = Cadena & DBSet(RecuperaValor(ColApuntes.Item(K), 4), "T") & "," 'ampconce
        Aux = RecuperaValor(ColApuntes.Item(K), 5) 'ImporteD
        If Aux = "" Then
            Aux = RecuperaValor(ColApuntes.Item(K), 6) 'ImporteD
            If Aux = "" Then Aux = "0"
            Aux = ImporteFormateado(Aux)
            Cadena = Cadena & "NULL," & DBSet(Aux, "N")
        Else
            Aux = ImporteFormateado(Aux)
            Cadena = Cadena & DBSet(Aux, "N") & ",NULL"
        End If
        Cadena = Cadena & "," & DBSet(RecuperaValor(ColApuntes.Item(K), 7), "T") & ")" 'contrapartida
        
    Next
    Aux = "INSERT INTO ariconta" & vParam.Numconta & ".hlinapu(numdiari,fechaent,numasien,linliapu,codmacta,numdocum,codconce,ampconce,timporteD,timporteH,ctacontr) VALUES "
    Cadena = Mid(Cadena, 2) 'quitamos la primera coma
    Aux = Aux & Cadena
    Conn.Execute Aux
    
    
    CrearApunteDesdeColeccion = True
    
    
eCrearApunteDesdeColeccion:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set mC = Nothing
End Function
