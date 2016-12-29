VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Control2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'#Generador DAVID(Ariadna)



'-------------------------------------------------------
'-------------------------------------------------------
'Declaracion de las variables
'---- Servidor donde esta MYSQL
Private mvarODBC As String ' Variable local
Private mvarUltusu As String 'Variable local
Private mvarUltempre As String 'Variable local
Private mvarAncho1 As String 'Variable local
Private mvarAncho2 As String 'Variable local
Private mvarAncho3 As String 'Variable local
Private mvarUltaccesoBDs As String 'Variable local
Private mvarUltreferrem As String 'Variable local


Private mvarUsuarioBD As String 'Variable local
Private mvarPassworBD As String 'Variable local



'--------
'Private mvarPedirPasswd As Boolean

'Variables de este modulo
Dim NF As Integer
Dim Mens As String



'------------------------------------------------------
'------------------------------------------------------
'      Propiedades del modulo de control
'------------------------------------------------------
'------------------------------------------------------

'----
'---- ODBC donde esta MYSQL
'----
Public Property Let ODBC(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.SERVER = 5
     mvarODBC = vData
End Property

Public Property Get ODBC() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.SERVER
     ODBC = mvarODBC
End Property


'----
'---- Ultimo usuario
'----
Public Property Let UltUsu(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.SERVER = 5
     mvarUltusu = vData
End Property

Public Property Get UltUsu() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.SERVER
     UltUsu = mvarUltusu
End Property

'----
'---- Ultima empresa
'----
Public Property Let UltEmpre(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.SERVER = 5
     mvarUltempre = vData
End Property

Public Property Get UltEmpre() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.SERVER
     UltEmpre = mvarUltempre
End Property

'----
'---- Ancho 1
'----
Public Property Let Ancho1(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.User = 5
     mvarAncho1 = vData
End Property

Public Property Get Ancho1() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.User
     Ancho1 = mvarAncho1
End Property

'----
'---- Ancho 2
'----
Public Property Let Ancho2(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.User = 5
     mvarAncho2 = vData
End Property

Public Property Get Ancho2() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.User
     Ancho2 = mvarAncho2
End Property

'----
'---- Ancho 3
'----
Public Property Let Ancho3(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.User = 5
     mvarAncho3 = vData
End Property

Public Property Get Ancho3() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.User
     Ancho3 = mvarAncho3
End Property

'----
'---- Ultimo Acceso BDS en el mantenimiento de usuarios ( permitidas, no permitidas, todas )
'----
Public Property Let UltAccesoBDs(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.password = 5
     mvarUltaccesoBDs = vData
End Property

Public Property Get UltAccesoBDs() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.password
     UltAccesoBDs = mvarUltaccesoBDs
End Property


'----
'---- Ultima referencia de remesa utilizada en la grabacion de soporte
'----
Public Property Let UltReferRem(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.password = 5
     mvarUltreferrem = vData
End Property

Public Property Get UltReferRem() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.password
     UltReferRem = mvarUltreferrem
End Property



Public Property Get UsuarioBD() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.password
     UsuarioBD = mvarUsuarioBD
End Property


Public Property Let UsuarioBD(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.password
     mvarUsuarioBD = vData
End Property


Public Property Get PassworBD() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.password
     PassworBD = mvarPassworBD
End Property



Public Property Let PassworBD(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.password
     mvarPassworBD = vData
End Property





'//////////////////////////////////////////////////////
'             LEER  Y  GUARDAR
'//////////////////////////////////////////////////////
'
'------------------------------------------------------
'     LEER
'
Public Function Leer() As Byte
    '--------------------------------------------------
    '   Leer:
    '   Busca un fichero de nombre Config.ini y lee de él los valores
    '   de parametrización para la aplicación. El fichero se busca
    '   en el directorio desde el que se ejecuta la aplicación.
    '--------------------------------------------------
    Dim Fichero As String ' Nombre completo del fichero a leer
    Dim LinAnt As String
    Dim Lin As String
    Dim NomParam As String ' Para leer el nombre del parámetro
    Dim ValParam As String ' Para leer el valor del parámetro
    Dim nFile As Integer
    
    On Error GoTo Err1
    Fichero = App.Path & "\Control.dat"
    nFile = FreeFile
    Open Fichero For Input As #nFile   ' Abre el archivo.
    Do While Not EOF(nFile) ' Repite el bucle hasta el final del archivo.
        Line Input #nFile, Lin
        Select Case Lin
           Case "[ODBC]", "[ultusu]", "[ultempre]", "[ancho1]", "[ancho2]", "[ancho3]", "[ultaccesoBDs]", "[ultreferrem]", "[usuBD]", "[passBD]"
               
               ValParam = ""
               NomParam = Lin

           Case "<-------->"
               Select Case NomParam
                    Case "[ODBC]": mvarODBC = ValParam
                    Case "[ultusu]": mvarUltusu = ValParam
                    Case "[ultempre]": mvarUltempre = ValParam
                    Case "[ancho1]": mvarAncho1 = ValParam
                    Case "[ancho2]": mvarAncho2 = ValParam
                    Case "[ancho3]": mvarAncho3 = ValParam
                    Case "[ultaccesoBDs]": mvarUltaccesoBDs = ValParam
                    Case "[ultreferrem]": mvarUltreferrem = ValParam
                    Case "[passBD]": mvarPassworBD = ValParam
                    Case "[usuBD]": mvarUsuarioBD = ValParam
                    
                     
                    
                End Select
            Case Else
                If ValParam = "" Then
                    ValParam = Lin
                Else
                    ValParam = ValParam & vbCrLf & Lin
                End If
        End Select
    Loop
    Close #nFile    ' Cierra el archivo.
    
    
    'El password y el usuario van "encriptados". Los ponemos
    Desencripta (True)
    Desencripta (False)
    
    Leer = 0
    Exit Function
Err1:
   Leer = 1
End Function

'------------------------------------------------------
'     GUARDAR
Public Function Grabar() As Byte
    '--------------------------------------------------
    '   GuardaConfig:
    '   Guardamos en el fichero Config.ini los valores de los
    '   parámetros de la aplicación
    '--------------------------------------------------
    Dim Fichero As String ' Nombre completo del fichero a leer
    Dim nFile As Integer


    On Error GoTo Err1
    Fichero = App.Path & "\Control.dat"
    nFile = FreeFile
    
    Open Fichero For Output As #nFile   ' Abre el archivo.
    Print #nFile, "[ODBC]"
    Print #nFile, mvarODBC
    Print #nFile, "<-------->"
    Print #nFile, "[ultusu]"
    Print #nFile, mvarUltusu
    Print #nFile, "<-------->"
    Print #nFile, "[ultempre]"
    Print #nFile, mvarUltempre
    Print #nFile, "<-------->"
    Print #nFile, "[ancho1]"
    Print #nFile, mvarAncho1
    Print #nFile, "<-------->"
    Print #nFile, "[ancho2]"
    Print #nFile, mvarAncho2
    Print #nFile, "<-------->"
    Print #nFile, "[ancho3]"
    Print #nFile, mvarAncho3
    Print #nFile, "<-------->"
    Print #nFile, "[ultaccesoBDs]"
    Print #nFile, mvarUltaccesoBDs
    Print #nFile, "<-------->"
    Print #nFile, "[ultreferrem]"
    Print #nFile, mvarUltreferrem
    Print #nFile, "<-------->"
    Print #nFile, "[usuBD]"
    Print #nFile, Encripta(mvarUsuarioBD)
    Print #nFile, "<-------->"
    Print #nFile, "[passBD]"
    Print #nFile, Encripta(mvarPassworBD)
    Print #nFile, "<-------->"
    
    
    
' "[ODBC]", "[ultusu]", "[ultempre]", "[ancho1]", "[ancho2]", "[ancho3]", "[ultaccesoBDs]", "[ultreferrem]"
    Close #nFile
    Grabar = 0
    Exit Function
Err1:
    Mens = "Error " & Err.Number & " : " & Err.Description
    MsgBox Mens, vbCritical
    Grabar = 1
End Function













Private Sub Desencripta(ElPassword As Boolean)
Dim C As String
Dim I As Integer

If ElPassword Then
    C = mvarPassworBD
    Else
    C = mvarUsuarioBD
End If
Mens = ""
For I = 1 To Len(C)
    If (I Mod 2) <> 0 Then
        'Solo valen los caracteres impares
        NF = Asc(Mid(C, I, 1))
        NF = NF + 1
        Mens = Mens & Chr(NF)
    End If
Next I

If ElPassword Then
    mvarPassworBD = Mens
    Else
    mvarUsuarioBD = Mens
End If

End Sub



Private Function Encripta(ByRef C As String) As String
Dim I As Integer


Mens = ""
Randomize
For I = 1 To Len(C)
        'Solo valen los caracteres impares
        NF = Asc(Mid(C, I, 1))
        NF = NF - 1
        Mens = Mens & Chr(NF)
        'Y el aleatorio , entre 48 y 122  <- 74 ->
        NF = Int((74 * Rnd)) ' Genera valores aleatorios entre 0y74.
        NF = NF + 48
        Mens = Mens & Chr(NF)
Next I

Encripta = Mens

End Function

