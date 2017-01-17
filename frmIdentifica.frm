VERSION 5.00
Begin VB.Form frmIdentifica 
   BorderStyle     =   0  'None
   Caption         =   "Identificacion"
   ClientHeight    =   5640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9705
   ForeColor       =   &H8000000B&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   9705
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   4950
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "aritel"
      Top             =   5070
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "frmIdentifica.frx":0000
      Left            =   4950
      List            =   "frmIdentifica.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   4110
      Width           =   2685
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   4980
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   4170
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Versión 6.0.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   4350
      TabIndex        =   6
      Top             =   2340
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cargando ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00765341&
      Height          =   285
      Index           =   2
      Left            =   360
      TabIndex        =   5
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00765341&
      Height          =   375
      Index           =   1
      Left            =   4950
      TabIndex        =   4
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00765341&
      Height          =   375
      Index           =   0
      Left            =   4950
      TabIndex        =   3
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   5790
      Left            =   0
      Top             =   0
      Width           =   9750
   End
End
Attribute VB_Name = "frmIdentifica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PrimeraVez As Boolean
Dim T1 As Single
Dim CodPC As Long
Dim UsuarioOK As String

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub Combo1_LostFocus()
    Text1(0).Text = Combo1.Text
End Sub

Private Sub Form_Activate()

    If PrimeraVez Then
        PrimeraVez = False
        If Not vEmpresa Is Nothing Then Exit Sub
        espera 0.5
        Me.Refresh
        
        Set vControl = New Control2
        If vControl.Leer = 1 Then
            
            vControl.ODBC = "Arigestion"
            vControl.UltUsu = "root"
            vControl.UltEmpre = "arigestion1"
            vControl.Ancho1 = 4360
            vControl.Ancho2 = 1399
            vControl.Ancho3 = 3000
            vControl.UltAccesoBDs = 0
            vControl.UltReferRem = 0
            vControl.PassworBD = "aritel"
            vControl.UsuarioBD = vControl.UltUsu
            vControl.Grabar

            End
            Exit Sub
        End If
        
         'Abrimos conexion para comprobar el usuario
         'Luego, en funcion del nivel de usuario que tenga cerraremos la conexion
         'y la abriremos con usuario-codigo ajustado a su nivel
        If AbrirConexion("") = False Then
             MsgBox "La aplicación no puede continuar sin acceso a los datos. ", vbCritical
             End
        End If
         
        'Gestionar el nombredel PC para la asignacion de PC en el entorno de red
        CodPC = GestionaPC2
        CadenaDesdeOtroForm = ""
         
         'Leemos el ultimo usuario conectado
        Text1(0).Text = vControl.UltUsu
         
         CargaCombo
         PosicionarCombo2 Combo1, Text1(0)
         
'         If CodPC > 0 Then
'            If ActualizarVersion Then
'                Set Conn = Nothing
'               Unload Me
'               End
'               Exit Sub
'            End If
'        End If
         
     
        If vParam Is Nothing Then
            PonerVisible True
             
            '??????quitar
    '        If Combo1.Text = "root" Then
    '            Text1(1).Text = "aritel"
    '            Exit Sub
    '        End If
             
             
            If Text1(0).Text <> "" Then
                Text1(1).SetFocus
            Else
                Text1(0).SetFocus
            End If
            
        End If
        
    End If
    Screen.MousePointer = vbDefault
End Sub



Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    UsuarioOK = ""
    PonerVisible False
    T1 = Timer
    Text1(0).Text = ""
    Text1(1).Text = ""
    Combo1.ListIndex = -1
    
    Label1(3).Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
    Label1(2).Caption = "Cargando datos de usuarios"
    
    PrimeraVez = True
    CargaImagen
    Me.Height = 5625 '5535
    Me.Width = 9705 ' 7935
    
'    '?????????????? QUITAR ESTO
'    If Combo1.Text = "root" Then Text1(1).Text = "aritel"
    
End Sub


Private Sub CargaImagen()
    On Error Resume Next
    Me.Image1 = LoadPicture(App.Path & "\arifon6.dat")
    If Err.Number <> 0 Then
        MsgBox Err.Description & vbCrLf & vbCrLf & "Error cargando", vbCritical

        Set Conn = Nothing
        End
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
'    NumeroEmpresaMemorizar False
    If UsuarioOK <> "" Then
            If vControl.UltUsu <> UsuarioOK Then
                vControl.UltUsu = UsuarioOK 'Text1(0).Text
                vControl.Grabar
            End If
        
    End If
End Sub


Private Sub Text1_GotFocus(Index As Integer)
    With Text1(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).Text = Trim(Text1(Index).Text)
    
    'Comprobamos si los dos estan con datos
    If Text1(0).Text <> "" And Text1(1).Text <> "" Then
        'Probar conexion usuario
        Validar
    End If
        
    
End Sub

Public Sub pLabel(TEXTO As String)

    Me.Label1(2).Caption = TEXTO
    Label1(2).Refresh
    espera 0.05
End Sub


Private Sub Validar()
Dim NuevoUsu As Usuario
Dim Ok As Byte

    'Validaremos el usuario y despues el password
    pLabel "Creando"
    Set vUsu = New Usuario
    
    
    
    If vUsu.Leer(Text1(0).Text) = 0 Then
        If vUsu.Nivel < 0 Then
            'NO tiene autorizacion de ningun nivel. Es menos 1
            Ok = 3
        Else
            'Con exito
            If vUsu.PasswdPROPIO = Text1(1).Text Then
                Ok = 0
            Else
                Ok = 1
            End If
        End If
    Else
        Ok = 2
    End If
    
    If Ok <> 0 Then
        If Ok = 3 Then
            MsgBox "Usuario sin autorizacion.", vbExclamation
        Else
            MsgBox "Usuario-Clave Incorrecto", vbExclamation
        End If
        
        Text1(1).Text = ""
        If Ok = 2 Then
            Text1(0).SetFocus
        Else
            Text1(1).SetFocus
        End If
    Else
        'OK
        If vEmpresa Is Nothing Then
            UsuarioCorrecto
            Load frmLabels
            Unload Me
        End If
    End If

End Sub


Private Sub UsuarioCorrecto()
Dim Sql As String
Dim PrimeraBD As String
        Screen.MousePointer = vbHourglass
        CadenaDesdeOtroForm = "OK"
        Label1(2).Caption = "Leyendo ."  'Si tarda pondremos texto aquin
        UsuarioOK = Text1(0).Text
        PonerVisible False
        Me.Refresh
        espera 0.1
        Me.Refresh

        Screen.MousePointer = vbHourglass
        
        
        pLabel "Conectando BD"
        HacerAccionesBD
        



       
       '++
       CadenaDesdeOtroForm = vControl.UltEmpre 'ultima empresa
       vUsu.CadenaConexion = vControl.UltEmpre
       '++
       
       If CadenaDesdeOtroForm = "" Then
            'No ha seleccionado nonguna empresa
            Set Conn = Nothing
            End
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass

        

        ' antes de cerrar la conexion cojo de usuarios.empresasariconta la primera que encuentre
        Sql = "select min(codempre) from usuarios.empresasarigestion  "
        PrimeraBD = DevuelveValor(Sql)


        'Cerramos la conexion
        Conn.Close
        pLabel "Abriendo " & CadenaDesdeOtroForm
        
        If AbrirConexion(CadenaDesdeOtroForm, True) = False Then
            CadenaDesdeOtroForm = PrimeraBD
            If AbrirConexion(CadenaDesdeOtroForm) = False Then
                End
            End If
        End If
        
        Screen.MousePointer = vbHourglass
        pLabel "Leyendo parametros"
        LeerEmpresaParametros
        

        RevisarIntroduccion = 0


        OtrasAcciones

        

End Sub


Private Sub HacerAccionesBD()
Dim Sql As String
    
    T1 = Timer
    
    'Limpiamos datos blanace
    CadenaDesdeOtroForm = " WHERE codusu = " & vUsu.Codigo
'    Conn.Execute "Delete from zBloqueos " & CadenaDesdeOtroForm
'    Conn.Execute "Delete from tmpconextcab " & CadenaDesdeOtroForm
'    Conn.Execute "Delete from tmpactualizar " & CadenaDesdeOtroForm
'    Conn.Execute "Delete from usuarios.ztesoreriacomun  " & CadenaDesdeOtroForm
'    Conn.Execute "Delete from usuarios.ztmpfaclin " & CadenaDesdeOtroForm
'
    CadenaDesdeOtroForm = ""

    Me.Refresh
    T1 = Timer - T1
    If T1 < 1 Then espera 0.4
    
    DoEvents
    espera 0.2
End Sub


Private Sub PonerVisible(Visible As Boolean)
    Label1(2).Visible = Not Visible  'Cargando
    Text1(0).Visible = Visible
    Text1(1).Visible = Visible
    Label1(0).Visible = Visible
    Label1(1).Visible = Visible
    Combo1.Visible = Visible
End Sub


'Lo que haremos aqui es ver, o guardar, el ultimo numero de empresa
'a la que ha entrado, y el usuario
Private Sub NumeroEmpresaMemorizar(Leer As Boolean)
Dim NF As Integer
Dim Cad As String
On Error GoTo ENumeroEmpresaMemorizar


    Cad = App.Path & "\ultusu.dat"
    
    
    
    If Leer Then
        If Dir(Cad) <> "" Then
            NF = FreeFile
            Open Cad For Input As #NF
            Line Input #NF, Cad
            Close #NF
            Cad = Trim(Cad)
                
                
                'El primer pipe es el usuario
                Text1(0).Text = Cad
        End If
    Else 'Escribir
        NF = FreeFile
        Open Cad For Output As #NF
        Cad = Text1(0).Text
        Print #NF, Cad
        Close #NF
    End If
ENumeroEmpresaMemorizar:
    Err.Clear
End Sub


'Private Function ActualizarVersion() As Boolean
'Dim Version As Integer
'    ActualizarVersion = 0
'    If Dir(App.Path & "\Actualizar.exe", vbArchive) <> "" Then
'        Set miRsAux = New ADODB.Recordset
'        Version = HayQueActualizar
'        If Version > 0 Then
'            CadenaDesdeOtroForm = "Estan disponibles actualizaciones para instalarse en esta maquina. ¿Desea continuar?"
'            If MsgBox(CadenaDesdeOtroForm, vbQuestion + vbYesNo) = vbYes Then
'                'LANZAMOS EL actualizador
'                CadenaDesdeOtroForm = App.Path & "\Actualizar.exe "
'                '       Parametros
'                '       applicacion    version   PC
'                CadenaDesdeOtroForm = CadenaDesdeOtroForm & " CONTA " & Version & " " & CodPC
'                Shell CadenaDesdeOtroForm, vbNormalNoFocus
'                ActualizarVersion = True
'            End If
'        End If
'        Set miRsAux = Nothing
'        CadenaDesdeOtroForm = ""
'    End If
'End Function
'
'
''Private Function HayQueActualizar() As Integer
'Dim v As Integer
'    On Error GoTo EA
'    HayQueActualizar = 0
'
'    CadenaDesdeOtroForm = "Select max(ver) from yVersion where app='CONTA'"
'    miRsAux.Open CadenaDesdeOtroForm, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    v = 0
'    If Not miRsAux.EOF Then v = DBLet(miRsAux.Fields(0), "N")
'    miRsAux.Close
'    If v = 0 Then Exit Function
'
'
'    'YA TENGO LA ULTIMA VERSION disponible. Voy a ver cual tengo
'    CadenaDesdeOtroForm = DevuelveDesdeBD("Conta", "PCs", "codpc", CStr(CodPC), "N")
'    If CadenaDesdeOtroForm = "" Then CadenaDesdeOtroForm = 0
'    NumRegElim = Val(CadenaDesdeOtroForm)
'    If v > NumRegElim Then
'        'OK esta desactualizado.
'        'Veo cual es la version qe hay que lanzar.
'        HayQueActualizar = NumRegElim + 1
'    End If
'
'
'    Exit Function
'EA:
'    Err.Clear
'    Err.Clear
'    Set miRsAux = Nothing
'End Function



Private Sub CargaCombo()
Dim miRsAux As ADODB.Recordset

    Combo1.Clear
    'Conceptos
    Set miRsAux = New ADODB.Recordset
    
    miRsAux.Open "Select * from usuarios.usuarios where nivelariges <> -1 order by login", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    
    While Not miRsAux.EOF
        Combo1.AddItem miRsAux!Login
        Combo1.ItemData(Combo1.NewIndex) = miRsAux!codusu
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
End Sub





Private Sub OtrasAcciones()
On Error Resume Next

    FormatoFecha = "yyyy-mm-dd"
    formatoFechaVer = "dd/mm/yyyy"
    FormatoImporte = "#,###,###,##0.00"
    FormatoPrecio = "#,###0.000"
    FormatoDec10d2 = "##,###,##0.00"
    FormatoPorcen = "##0.00"
    
    '++
    teclaBuscar = 43

    DireccionAyuda = "http://help-ariconta.ariadnasw.com/"     'falta cambiar

    'Borramos uno de los archivos temporales
    If Dir(App.Path & "\ErrActua.txt") <> "" Then Kill App.Path & "\ErrActua.txt"
    
    
    'Borramos tmp bloqueos
    'Borramos temporal
    CadenaDesdeOtroForm = OtrosPCsContraContabiliad(True)
    NumRegElim = Len(CadenaDesdeOtroForm)
    If NumRegElim = 0 Then
        CadenaDesdeOtroForm = ""
    Else
        CadenaDesdeOtroForm = " WHERE codusu = " & vUsu.Codigo
    End If
    Conn.Execute "Delete from zBloqueos " & CadenaDesdeOtroForm
    CadenaDesdeOtroForm = ""
    NumRegElim = 0
    
    
End Sub
