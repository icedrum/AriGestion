VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cempresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Variables locales que contienen valores de propiedad
Private mvarcodempre As String 'copia local
Private mvarnomempre As String 'copia local
Private mvarCifEmpre As String 'copia local

Private mvarDigitosUltimoNivel  As Integer
Private mvarFechaInicioEjercicio As Date
Private mvarFechaFinEjercicio As Date
Private mvarUltimoDiaPeriodoLiquidado As Date   'ultima fecha IVA liquidado
Private mvarFechaActivaConta As Date
Private mvarFechaIncioConta As Date
Private mvarFechaFinConta As Date


Private SQL As String  'Para la sentencia
Private Mens As String    'Para los mensajes
Private Rs As ADODB.Recordset
Private RC As Byte


'------------------------------------------------
'Propiedades del modulo Cempresa
'------------------------------------------------


Public Property Let codempre(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.codempre = 5
     mvarcodempre = vData
End Property

Public Property Get codempre() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.codempre
     codempre = mvarcodempre
End Property


Public Property Let nomempre(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.nomempre = 5
     mvarnomempre = vData
End Property

Public Property Get nomempre() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.nomempre
     nomempre = mvarnomempre
End Property

Public Property Let CifEmpre(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.nomempre = 5
     mvarCifEmpre = vData
End Property

Public Property Get CifEmpre() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.nomempre
     CifEmpre = mvarCifEmpre
End Property

Public Property Get DigitosUltimoNivel() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.nomempre
     DigitosUltimoNivel = mvarDigitosUltimoNivel
End Property

Public Property Get FechaInicioEjercicio() As Date
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.nomempre
     FechaInicioEjercicio = mvarFechaInicioEjercicio
End Property

Public Property Get FechaFinEjercicio() As Date
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.nomempre
     FechaFinEjercicio = mvarFechaFinEjercicio
End Property

Public Property Get FechaActivaConta() As Date
     FechaActivaConta = mvarFechaActivaConta
End Property



Public Property Get UltimoDiaPeriodoLiquidado() As Date
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.nomempre
     UltimoDiaPeriodoLiquidado = mvarUltimoDiaPeriodoLiquidado
End Property


Public Property Get FechaIncioConta() As Date
     FechaIncioConta = mvarFechaIncioConta
End Property
Public Property Get FechaFinConta() As Date
     FechaFinConta = mvarFechaFinConta
End Property

    



'------------------------------------------------
'FUNCIONES del modulo Cempresa
'------------------------------------------------






'****   LEER  ***
Public Function Leer2() As Byte
    On Error GoTo Err3
    SQL = "SELECT codigo ,nomempre,CifEmpre FROM  Empresas "
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Mens = ""
    If Rs.EOF Then
        Leer2 = 1
        
        'ELSE
        Else
            Mens = "Datos basicos"
            mvarcodempre = Rs!Codigo
            mvarnomempre = Rs!nomempre
            mvarCifEmpre = Rs!CifEmpre
            
            
           
          
            Leer2 = 0
    End If
    Rs.Close
    
    Set Rs = Nothing
    Exit Function

Err3:
    Mens = "Se ha producido un error. " & Mens & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Set Rs = Nothing
    Leer2 = 1
End Function




Public Function FijarDatosAriconta() As Boolean
 Dim Mes As Integer
 Dim Dias As Integer
 
    On Error GoTo eFijarDatosAriconta
    FijarDatosAriconta = False
    Set Rs = New ADODB.Recordset
    SQL = "SELECT  numnivel,numdigi1,numdigi2,numdigi3,numdigi4,numdigi5,numdigi6,numdigi7,numdigi8,numdigi9,numdigi10"
    SQL = SQL & " FROM ariconta" & vParam.Numconta & ".empresa "
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = Rs!numnivel
    mvarDigitosUltimoNivel = Rs.Fields(CInt(SQL))
    Rs.Close
    
    SQL = "Select periodos,perfactu,anofactu , fechaini,FechaActiva,fechafin  "
    SQL = SQL & " FROM ariconta" & vParam.Numconta & ".parametros "
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Rs!periodos = 0 Then
        'Trimestral
        Mes = Rs!perfactu * 3
    Else
        Mes = Rs!perfactu
    End If
    Dias = DiasMes(CByte(Mes), Rs!anofactu)
    mvarUltimoDiaPeriodoLiquidado = CDate(Dias & "/" & Format(Mes, "00") & "/" & Rs!anofactu)
    
    'A MANO,    FALTA#####
    
    'mvarFechaInicioEjercicio = Rs!fechaini   'deberiamos crear un parametro
    mvarFechaInicioEjercicio = "01/01/2018"
    mvarFechaFinEjercicio = DateAdd("yyyy", 1, mvarFechaInicioEjercicio)
    mvarFechaFinEjercicio = DateAdd("d", -1, mvarFechaFinEjercicio)
    mvarFechaActivaConta = mvarFechaInicioEjercicio
    If Not IsNull(Rs!FechaActiva) Then mvarFechaActivaConta = Rs!FechaActiva
    
    mvarFechaIncioConta = Rs!fechaini
    mvarFechaFinConta = Rs!fechafin
    
    
    Rs.Close
    FijarDatosAriconta = True
    
eFijarDatosAriconta:
    If Err.Number <> 0 Then MuestraError Err.Number, SQL
    Set Rs = Nothing
End Function


