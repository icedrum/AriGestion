VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Ctipoformapago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Variables locales que contienen valores de propiedad
Private mvartipoformapago As Integer 'copia local
Private mvardescformapago As String 'copia local
Private mvarsiglas As String 'copia local
Private mvarPagoBancario As Integer 'copia local
Private mvardiaricli As Integer 'copia local
Private mvarcondecli As Integer 'copia local
Private mvarconhacli As Integer 'copia local
Private mvarampdecli As Integer 'copia local
Private mvaramphacli As Integer 'copia local
Private mvarctrdecli As Integer 'copia local
Private mvarctrhacli As Integer 'copia local
Private mvardiaripro As Integer 'copia local
Private mvarcondepro As Integer 'copia local
Private mvarconhapro As Integer 'copia local
Private mvarampdepro As Integer 'copia local
Private mvaramphapro As Integer 'copia local
Private mvarctrdepro As Integer 'copia local
Private mvarctrhapro As Integer 'copia local



'Para cargar una cadena de texo auxliar.
'  Por ejemplo en la contabilizacion de remesas TAL/PAG
'   grabare la ampliacion al HABER y en descforpa tendre la del debe
Private mvarCadenaAuxiliar As String

Private SQL As String  'Para la sentencia
Private Mens As String    'Para los mensajes
Private Rs As ADODB.Recordset 'Para la base de datos
Private RC As Byte



'------------------------------------------------
'Propiedades del modulo Ctipoformapago
'------------------------------------------------


Public Property Let tipoformapago(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.tipoformapago = 5
     mvartipoformapago = vData
End Property

Public Property Get tipoformapago() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.tipoformapago
     tipoformapago = mvartipoformapago
End Property


Public Property Let descformapago(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.descformapago = 5
     mvardescformapago = vData
End Property

Public Property Get descformapago() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.descformapago
     descformapago = mvardescformapago
End Property


Public Property Let siglas(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.siglas = 5
     mvarsiglas = vData
End Property

Public Property Get siglas() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.siglas
     siglas = mvarsiglas
End Property


Public Property Let PagoBancario(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.PagoBancario = 5
     mvarPagoBancario = vData
End Property

Public Property Get PagoBancario() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.PagoBancario
     PagoBancario = mvarPagoBancario
End Property


Public Property Let diaricli(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.diaricli = 5
     mvardiaricli = vData
End Property

Public Property Get diaricli() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.diaricli
     diaricli = mvardiaricli
End Property


Public Property Let condecli(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.condecli = 5
     mvarcondecli = vData
End Property

Public Property Get condecli() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.condecli
     condecli = mvarcondecli
End Property


Public Property Let conhacli(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.conhacli = 5
     mvarconhacli = vData
End Property

Public Property Get conhacli() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.conhacli
     conhacli = mvarconhacli
End Property


Public Property Let ampdecli(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.ampdecli = 5
     mvarampdecli = vData
End Property

Public Property Get ampdecli() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.ampdecli
     ampdecli = mvarampdecli
End Property


Public Property Let amphacli(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.amphacli = 5
     mvaramphacli = vData
End Property

Public Property Get amphacli() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.amphacli
     amphacli = mvaramphacli
End Property


Public Property Let ctrdecli(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.ctrdecli = 5
     mvarctrdecli = vData
End Property

Public Property Get ctrdecli() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.ctrdecli
     ctrdecli = mvarctrdecli
End Property


Public Property Let ctrhacli(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.ctrhacli = 5
     mvarctrhacli = vData
End Property

Public Property Get ctrhacli() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.ctrhacli
     ctrhacli = mvarctrhacli
End Property


Public Property Let diaripro(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.diaripro = 5
     mvardiaripro = vData
End Property

Public Property Get diaripro() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.diaripro
     diaripro = mvardiaripro
End Property


Public Property Let condepro(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.condepro = 5
     mvarcondepro = vData
End Property

Public Property Get condepro() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.condepro
     condepro = mvarcondepro
End Property


Public Property Let conhapro(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.conhapro = 5
     mvarconhapro = vData
End Property

Public Property Get conhapro() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.conhapro
     conhapro = mvarconhapro
End Property


Public Property Let ampdepro(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.ampdepro = 5
     mvarampdepro = vData
End Property

Public Property Get ampdepro() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.ampdepro
     ampdepro = mvarampdepro
End Property


Public Property Let amphapro(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.amphapro = 5
     mvaramphapro = vData
End Property

Public Property Get amphapro() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.amphapro
     amphapro = mvaramphapro
End Property


Public Property Let ctrdepro(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.ctrdepro = 5
     mvarctrdepro = vData
End Property

Public Property Get ctrdepro() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.ctrdepro
     ctrdepro = mvarctrdepro
End Property


Public Property Let ctrhapro(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.ctrhapro = 5
     mvarctrhapro = vData
End Property

Public Property Get ctrhapro() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.ctrhapro
     ctrhapro = mvarctrhapro
End Property



Public Property Let CadenaAuxiliar(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.ctrhapro = 5
     mvarCadenaAuxiliar = vData
End Property

Public Property Get CadenaAuxiliar() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.ctrhapro
     CadenaAuxiliar = mvarCadenaAuxiliar
End Property

'------------------------------------------------
'FUNCIONES del modulo Ctipoformapago
'------------------------------------------------

'****   LEER  ***
Public Function Leer(vIdtipoformapago As Integer) As Byte
    On Error GoTo Err3

    SQL = "SELECT * FROM tipofpago"
    SQL = SQL & " WHERE "
    SQL = SQL & " tipoformapago = " & vIdtipoformapago
    Rs.Open SQL, Conn, , , adCmdText
    If Rs.EOF Then
        Leer = 1
        'ELSE
        Else
            mvartipoformapago = Rs!tipoformapago
            mvardescformapago = DBLet(Rs!descformapago)
            mvarsiglas = DBLet(Rs!siglas)
            'mvarPagoBancario = RS!PagoBancario
            mvardiaricli = Rs!diaricli
            mvarcondecli = Rs!condecli
            mvarconhacli = Rs!conhacli
            mvarampdecli = Rs!ampdecli
            mvaramphacli = Rs!amphacli
            mvarctrdecli = Rs!ctrdecli
            mvarctrhacli = Rs!ctrhacli
            mvardiaripro = Rs!diaripro
            mvarcondepro = Rs!condepro
            mvarconhapro = Rs!conhapro
            mvarampdepro = Rs!ampdepro
            mvaramphapro = Rs!amphapro
            mvarctrdepro = Rs!ctrdepro
            mvarctrhapro = Rs!ctrhapro
            Leer = 0
    End If

    Rs.Close
    Exit Function

Err3:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Leer = 1
End Function


'****   MODIFICAR  ***
Public Function Modificar() As Byte
    On Error GoTo Err4

    SQL = "SELECT * FROM stipoformapago"
    SQL = SQL & " WHERE "
    SQL = SQL & " tipoformapago = " & mvartipoformapago
    Rs.Open SQL, Conn, , , adCmdText
    If Rs.EOF Then
        Mens = "No existe el Registro en la Tabla stipoformapago con esos valores."
        MsgBox Mens, vbExclamation
        Modificar = 1
        'ELSE
        Else
            'Rs.Edit
            Rs!descformapago = mvardescformapago
            Rs!siglas = mvarsiglas
            Rs!PagoBancario = mvarPagoBancario
            Rs!diaricli = mvardiaricli
            Rs!condecli = mvarcondecli
            Rs!conhacli = mvarconhacli
            Rs!ampdecli = mvarampdecli
            Rs!amphacli = mvaramphacli
            Rs!ctrdecli = mvarctrdecli
            Rs!ctrhacli = mvarctrhacli
            Rs!diaripro = mvardiaripro
            Rs!condepro = mvarcondepro
            Rs!conhapro = mvarconhapro
            Rs!ampdepro = mvarampdepro
            Rs!amphapro = mvaramphapro
            Rs!ctrdepro = mvarctrdepro
            Rs!ctrhapro = mvarctrhapro
            Rs.Update
            Modificar = 0
    End If

    Rs.Close
    Exit Function

Err4:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Modificar = 1
End Function

Private Sub Class_Initialize()
    Set Rs = New ADODB.Recordset
    'Ponemos el recordset en modo:
    Rs.CursorType = adOpenKeyset
    Rs.LockType = adLockOptimistic
End Sub


Private Sub Class_Terminate()
    Set Rs = Nothing
End Sub
