VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cparametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Private mvarPathAyuda As String



Private mvarNumInstalacion As Integer

Private mvarSerconta As String
Private mvarUsuconta As String
Private mvarPasconta As String
Private mvarNumconta As Integer
Private mvarRaizcuotas As String
Private mvarRaiztasas As String
Private mvarIvaexento As Integer
Private mvarBanco As String

Private Sql As String  'Para la sentencia
Private Mens As String    'Para los mensajes
Private RS As Recordset
Private RC As Byte



'------------------------------------------------
'Propiedades del modulo Cparametros
'------------------------------------------------
Public Property Let Serconta(ByVal vData As String)
     mvarSerconta = vData
End Property

Public Property Get Serconta() As String
     Serconta = mvarSerconta
End Property
Public Property Let Usuconta(ByVal vData As String)
     mvarUsuconta = vData
End Property

Public Property Get Usuconta() As String
     Usuconta = mvarUsuconta
End Property

Public Property Let Pasconta(ByVal vData As String)
     mvarPasconta = vData
End Property

Public Property Get Pasconta() As String
     Pasconta = mvarPasconta
End Property



Public Property Let Numconta(ByVal vData As Integer)
     mvarNumconta = vData
End Property

Public Property Get Numconta() As Integer
     Numconta = mvarNumconta
End Property






Public Property Let Raizcuotas(ByVal vData As String)
     mvarRaizcuotas = vData
End Property

Public Property Get Raizcuotas() As String
     Raizcuotas = mvarRaizcuotas
End Property
Public Property Let Raiztasas(ByVal vData As String)
     mvarRaiztasas = vData
End Property

Public Property Get Raiztasas() As String
     Raiztasas = mvarRaiztasas
End Property



Public Property Let Ivaexento(ByVal vData As Integer)
     mvarIvaexento = vData
End Property

Public Property Get Ivaexento() As Integer
     Ivaexento = mvarIvaexento
End Property





'****   LEER  ***
Public Function leer() As Byte
    On Error GoTo Err3
    Sql = "SELECT * FROM parametros"
    Set RS = New ADODB.Recordset
    RS.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic
    If RS.EOF Then
        leer = 1
        'ELSE
        Else


            mvarSerconta = RS!Serconta
            mvarUsuconta = RS!Usuconta
            mvarPasconta = RS!Pasconta
            mvarNumconta = RS!Numconta
            mvarRaizcuotas = DBLet(RS!Raizcuotas, "T")
            mvarRaiztasas = DBLet(RS!Raiztasas, "T")
            mvarIvaexento = DBLet(RS!Ivaexento, "N")

            
            'mvarPathAyuda = DBLet(Rs!PathAyuda, "T")
           leer = 0
    End If

    RS.Close
    
    
    
    
    Set RS = Nothing
    Exit Function

Err3:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    On Error Resume Next
    CerrarRs RS
    leer = 1
    Err.Clear
End Function



Public Function BancoPropioFacturacionContabilidad() As String
    If mvarBanco = "" Then
        mvarBanco = DevuelveDesdeBD("codmacta", "ariconta" & vParam.Numconta & ".bancos", "1", "1 ORDER BY codmacta desc")
    End If
    BancoPropioFacturacionContabilidad = mvarBanco
End Function
