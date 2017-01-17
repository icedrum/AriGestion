VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLOG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const AccionesRegistradas = 8
'ACCIONES DISPONIBLES
'
'   1- Modificar factura
'   2- Borrar factura
'   3- Modificar linea fra
'   4- Eliminar linea fra

'   5. Insertar caja desde exp
'   6- Modificar caja
'   7.- Eliminar caja

'   8.- Expediente. Borrar cobro a cuenta

Private mvarDatosDescripcion As String
Private mvarDatosAntes As String   'Para grabar cosas antes y desp
Private mvarDatosDespues As String   'Para grabar cosas antes y desp
Dim SQL As String


Public Property Let DatosDescripcion(ByVal vData As String)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Nivel = 5
    mvarDatosDescripcion = vData
End Property


Public Property Get DatosDescripcion() As String
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Nivel
    DatosDescripcion = mvarDatosDescripcion
End Property


'--------------------------------------------------------------------------
'--------------------------------------------------------------------------
'--------------------------------------------------------------------------
'  Servira para tener dos variables globarles.
' Cargaremos los valores que habian antes (suceptibles de ser logeados, y los cargaremos despues.
' Si hay diferencia, grabo log
Public Property Let DatosAntes(ByVal vData As String)
    mvarDatosAntes = vData
End Property
Public Property Get DatosAntes() As String
    DatosAntes = mvarDatosAntes
End Property
Public Property Let DatosDespues(ByVal vData As String)
    mvarDatosDespues = vData
End Property
Public Property Get DatosDespues() As String
    DatosDespues = mvarDatosDespues
End Property







Public Function Insertar(Accion As Byte, ByRef ElUsuario As Usuario, Descripcion As String) As Boolean

    
    On Error GoTo EI
    Insertar = False
    
    SQL = "insert into `acciones` (`fecha`,`accion`,`usuario`,`pc`,`descripcion`) values ( "
    SQL = SQL & " now()," & Accion & " ,'" & DevNombreSQL(ElUsuario.Login) & "','"
    SQL = SQL & DevNombreSQL(ElUsuario.PC) & "',"
    If Descripcion = "" Then
        SQL = SQL & "NULL)"
    Else
        SQL = SQL & DBSet(Descripcion, "T") & ")"
    End If
    Conn.Execute SQL
    
    
    Insertar = True
    Exit Function
EI:
    MuestraError Err.Number, "Registro LOG de acciones"

End Function

Private Function DimeAccion(N As Integer) As String
Dim C As String


    C = ""
    Select Case N
    Case 1
        C = "Modificar factura"
    Case 1
        C = "Borrar factura"
    Case 3
        C = "Modificar linea factura"
    Case 4
        C = "Eliminar linea factura"
        
    Case 5
        C = "Insertar en caja"
    Case 6
        C = "Modificar caja"
    Case 7
        C = "Eliminar en caja"
    Case 8
        C = "Exp. eliminar pago"
    Case 9
        C = "Quitar gestion stock"
    
    End Select
    DimeAccion = C
End Function


Public Function DevuelveAcciones(ByRef ListaAcciones As Collection) As Boolean
Dim I As Integer

 On Error GoTo EDevuelveAcciones
    For I = 1 To AccionesRegistradas
        SQL = I & "|" & DimeAccion(I) & "|"
        ListaAcciones.Add SQL
    Next I
    
    DevuelveAcciones = True
    
    Exit Function
EDevuelveAcciones:
    MuestraError Err.Number
End Function



Public Sub VolcarAFichero2()
Dim F As Date
Dim Rs As ADODB.Recordset




    On Error GoTo EVolcar

    

    Set Rs = New ADODB.Recordset
    F = Now   'Por si acaso esta el reloj mal
    Rs.Open "select curdate()", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then F = DBLet(Rs.Fields(0), "F")
    Rs.Close
    
    
    
    
    'Ya tengo la fecha.
    F = DateAdd("yyyy", -1, F)  'Le quito 1 año
    F = CDate("01/" & Month(F) & "/" & Year(F))   'Pongo el primer dia del mes anterior
    'Le quito un dia
    F = DateAdd("d", -1, F)
    
    
    SQL = "Select * from acciones where fecha < '" & Format(F, "yyyy-mm-dd") & "' "
    SQL = "INSERT INTO acciones_old " & SQL
    Conn.Execute SQL
    Conn.Execute "DELETE from acciones where fecha <'" & Format(F, "yyyy-mm-dd") & "'"
        
    
    
    Set Rs = Nothing
    
    
    Exit Sub
EVolcar:
    SQL = "Error volcado datos acciones: " & vbCrLf & Err.Description
    SQL = SQL & vbCrLf & vbCrLf & "El programa continuará." & vbCrLf & vbCrLf & "Avise a soporte técnico."
    MsgBox SQL, vbExclamation
    Err.Clear
    Set Rs = Nothing
End Sub





'Dos metodos sobre la variable mvarDatosDescripcion
' Inicializar y añadir texto
Public Sub InicializarDatosDesc()
    mvarDatosDescripcion = ""
End Sub


Public Sub AnyadeTextoDatosDes(T As String)
    
    If mvarDatosDescripcion <> "" Then mvarDatosDescripcion = mvarDatosDescripcion & ":"
    mvarDatosDescripcion = mvarDatosDescripcion & T
End Sub
