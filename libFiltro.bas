Attribute VB_Name = "libFiltro"
Option Explicit





Public Sub CargaVectoresFiltro(NumFiltros As Integer, Textox As String, ByRef cboFiltro_ As ComboBox)  'Vendran empipados
Dim I As Integer
    
    
    cboFiltro_.Clear
    cboFiltro_.AddItem "Sin filtro"
    For I = 1 To NumFiltros
        cboFiltro_.AddItem RecuperaValor(Textox, I)
    Next I
    
    
    
End Sub


'
Public Sub ValorFiltroPorDefecto(Leer As Boolean, NombreForm As String, Opcion As Integer, ByRef ColumnaOrden As Integer, ByRef OtroDatos As String, ByRef AscenDes As Boolean, ByRef CadenaParaGuardarFiltro As String)
Dim RN As ADODB.Recordset
Dim cad As String

    If Leer Then
        cad = "Select * from usuarios.usuariosvaloresdefecto WHERE"
        cad = cad & " aplicacion='ariconta' AND codusu=1" 'IRA codusu
        cad = cad & " AND formulario='" & NombreForm & "' AND opcion= "
        cad = cad & " " & Opcion
            
        Set RN = New ADODB.Recordset
        RN.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RN.EOF Then
            AscenDes = DBLet(RN!ascendescen, "N") = 1
            ColumnaOrden = DBLet(RN!Columna, "N")
            OtroDatos = DBLet(RN!otros, "T")
            CadenaParaGuardarFiltro = ColumnaOrden & "|" & AscenDes & "|" & OtroDatos  'La cadena para comprara al guardar
        Else
            AscenDes = True
            ColumnaOrden = 1
            CadenaParaGuardarFiltro = "-1"
        End If
        RN.Close
        
    Else
        'INSERT
        cad = ColumnaOrden & "|" & AscenDes & "|" & OtroDatos
        If cad <> CadenaParaGuardarFiltro Then
            'UPDATE /INSERT
            cad = "REPLACE usuarios.usuariosvaloresdefecto(aplicacion,codusu,formulario,opcion,columna,otros,ascendescen) VALUES ("
            cad = cad & "'ariconta',1,'" & NombreForm & "',"
            cad = cad & Opcion
            cad = cad & "," & ColumnaOrden & ",'" & OtroDatos & "'," & CStr(Abs(AscenDes)) & ")"
            Conn.Execute cad
            
        End If
    End If

End Sub


Public Function DevuelveFiltro(Prg As Integer, Aplicacion As String) As Integer
Dim SQL As String
    
    SQL = "select filtro from menus_usuarios where codigo = " & Prg
    SQL = SQL & " and aplicacion = " & DBSet(Aplicacion, "T") & " and codusu =" & vUsu.Id
    
    DevuelveFiltro = DevuelveValor(SQL)

End Function

