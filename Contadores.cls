VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Contadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'variables locales que contienen valores de propiedad

Private mvarTipoContador As String 'copia local
Private mvarNombreRegistro As String 'copia local
Private mvarContador As Long  'copia local

Private Rs As ADODB.Recordset
Private SQL As String
Private Mens As String


Public Property Let Contador(ByVal vData As Long)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Nivel = 5
    mvarContador = vData
End Property


Public Property Get Contador() As Long
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Nivel
    Contador = mvarContador
End Property


Public Property Let NombreRegistro(ByVal vData As String)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Nombre= 5
    mvarNombreRegistro = vData
End Property


Public Property Get NombreRegistro() As String
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Nombre
    NombreRegistro = mvarNombreRegistro
End Property



Public Property Let TipoContador(ByVal vData As String)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Nombre= 5
    mvarTipoContador = vData
End Property


Public Property Get TipoContador() As String
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Nombre
    TipoContador = mvarTipoContador
End Property



Public Function Eliminar() As Byte
    
    Exit Function
Err1:
    Mens = "Error: " & Err.Number & " : " & Err.Description
    MsgBox Mens, vbExclamation
    Rs.Close
    Eliminar = 1
End Function


'///////////////////////////////////////////////////////////
'//
'//  Devuelve el contador que corresponda.
'//  Si es llamado desde dentro de un begin-trans
'//  entonces NO ejecuta las sentencias commit
Public Function ConseguirContador(LEtra As String, EjercicioActual As Boolean, DentroBeginTrans As Boolean) As Byte
Dim Ok As Boolean
Dim C1 As Long
Dim C2 As Long
    On Error GoTo Err1
    'Abrimos bloqueando
    SQL = "Select * from contadores WHERE TipoRegi = '" & LEtra & "' FOR UPDATE"
    ConseguirContador = 1
    Ok = False
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        mvarTipoContador = LEtra
        mvarNombreRegistro = Rs!nomregis
        C1 = Rs!Contado1
        C2 = Rs!Contado2
        Ok = True
    End If
    Rs.Close
    If Not Ok Then
        If Not DentroBeginTrans Then Conn.Execute "commit"
        Exit Function
    End If
    
    'Actualizamos el contador
    SQL = "UPDATE contadores set "
    If EjercicioActual Then
        C1 = C1 + 1
        mvarContador = C1
        'c1 = c1 + 1 'ANTES
        SQL = SQL & " contado1=" & C1
    Else
        C2 = C2 + 1
        mvarContador = C2
        SQL = SQL & " contado2=" & C2
    End If
    SQL = SQL & " WHERE TipoRegi = '" & LEtra & "'"
    Conn.Execute SQL
    
    'Desbloqueamos
    If Not DentroBeginTrans Then Conn.Execute "commit"
    ConseguirContador = 0
    Exit Function
Err1:
    Mens = "Error: " & Err.Number & " : " & Err.Description
    MsgBox Mens, vbExclamation
End Function



'//////////////////////////////////////////////////////////////////7
'//
'//     Si es el ultimo numero de la serie devuelve el anterior
Public Function DevolverContador(LEtra As String, EjercicioActual As Boolean, Contador As Long, Optional DentroDeTransaccion As Boolean) As Byte
Dim Ok As Boolean
Dim C1 As Long
Dim C2 As Long
    On Error GoTo Err1
    'Abrimos bloqueando
    SQL = "Select * from contadores WHERE TipoRegi = '" & LEtra & "' FOR UPDATE"
    'Esto esba asi antes, sin comentar
    'Conn.Execute "Set autocommit = 0"
    DevolverContador = 1
    Ok = False
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        mvarTipoContador = LEtra
        mvarNombreRegistro = Rs!nomregis
        C1 = Rs!Contado1
        C2 = Rs!Contado2
        Ok = True
    End If
    Rs.Close
    If Not Ok Then
        If Not DentroDeTransaccion Then Conn.Execute "commit"
        Exit Function
    End If

    Ok = False
    'Contador = Contador + 1
    If EjercicioActual Then
        Ok = (C1 = Contador)
    Else
        Ok = (C2 = Contador)
    End If

    If Ok Then
        'Actualizamos el contador
        SQL = "UPDATE contadores set "
        If EjercicioActual Then
            C1 = C1 - 1
            SQL = SQL & " contado1=" & C1
        Else
            C2 = C2 - 1
            SQL = SQL & " contado2=" & C2
        End If
        SQL = SQL & " WHERE TipoRegi = '" & LEtra & "'"
        Conn.Execute SQL
    End If

    'Desbloqueamos
    If Not DentroDeTransaccion Then Conn.Execute "commit"
    DevolverContador = 0
    
Err1:
    If Err.Number <> 0 Then
        Mens = "Error: " & Err.Number & " : " & Err.Description
        MsgBox Mens, vbExclamation
    End If
    If Not DentroDeTransaccion Then Conn.Execute "Set autocommit = 1"
End Function

Public Function YaExisteContador(EjercicioActual As Boolean, FechaFinEjercicio As Date, MostrarError As Boolean) As Boolean
Dim F1 As Date
    On Error GoTo EYaExisteContador
    YaExisteContador = False

    Set Rs = New ADODB.Recordset
    SQL = "select fechaent,numdiari from hlinapu where numasien= " & mvarContador
    If EjercicioActual Then
        F1 = DateAdd("yyyy", -1, FechaFinEjercicio)
        SQL = SQL & " AND fechaent > '" & Format(F1, FormatoFecha) & "'"    'mayor estricto
        SQL = SQL & " AND fechaent <= '" & Format(FechaFinEjercicio, FormatoFecha) & "'"
    Else
        F1 = DateAdd("yyyy", 1, FechaFinEjercicio)
        SQL = SQL & " AND fechaent > '" & Format(FechaFinEjercicio, FormatoFecha) & "'"    'mayor estricto
        SQL = SQL & " AND fechaent <= '" & Format(F1, FormatoFecha) & "'"
    End If
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Rs.EOF Then
        YaExisteContador = True
    Else
        'YA existe.
        SQL = mvarContador & " -  " & Rs!FechaEnt & " - " & Rs!NumDiari & vbCrLf & vbCrLf & "Error contadores"
        If MostrarError Then MsgBox "Ya existe numero de asiento:" & SQL, vbExclamation
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
EYaExisteContador:
    MuestraError Err.Number, Err.Description
    Set Rs = Nothing
End Function


Public Function Modificar() As Byte
 
    Exit Function
Err1:
    Mens = "Error: " & Err.Number & " : " & Err.Description
    MsgBox Mens, vbExclamation
    Rs.Close
    Modificar = 1
End Function


Public Function Siguiente() As Long

End Function


