VERSION 5.00
Begin VB.Form frmEstadisticaCli 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10065
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   10065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameConcepto 
      Caption         =   "Selección"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   19
      Top             =   0
      Width           =   9795
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmEstadisticaClie.frx":0000
         Left            =   7920
         List            =   "frmEstadisticaClie.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   3000
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmEstadisticaClie.frx":0023
         Left            =   2160
         List            =   "frmEstadisticaClie.frx":0030
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtFecha 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "imgConcepto"
         Top             =   3000
         Width           =   1305
      End
      Begin VB.TextBox txtFecha 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   4680
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "imgConcepto"
         Top             =   3000
         Width           =   1305
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmEstadisticaClie.frx":003E
         Left            =   8640
         List            =   "frmEstadisticaClie.frx":004B
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtCliente 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "imgConcepto"
         Top             =   2040
         Width           =   1065
      End
      Begin VB.TextBox txtDescCliente 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   24
         Tag             =   "imgConcepto"
         Top             =   1560
         Width           =   4425
      End
      Begin VB.TextBox txtDescCliente 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   23
         Tag             =   "imgConcepto"
         Top             =   2040
         Width           =   4425
      End
      Begin VB.TextBox txtCliente 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "imgConcepto"
         Top             =   1560
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo factura"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   8
         Left            =   6480
         TabIndex        =   35
         Top             =   3000
         Width           =   1290
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo estdística"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   2
         Left            =   240
         TabIndex        =   32
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "F. factura"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   7
         Left            =   120
         TabIndex        =   30
         Top             =   3000
         Width           =   1050
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1920
         Picture         =   "frmEstadisticaClie.frx":0059
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   1200
         TabIndex        =   29
         Top             =   3030
         Width           =   615
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   4320
         Picture         =   "frmEstadisticaClie.frx":00E4
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   3600
         TabIndex        =   28
         Top             =   3030
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   1
         Left            =   7920
         TabIndex        =   27
         Top             =   1560
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   26
         Top             =   1560
         Width           =   765
      End
      Begin VB.Image imgCli 
         Height          =   360
         Index           =   1
         Left            =   1800
         Top             =   2040
         Width           =   360
      End
      Begin VB.Image imgCli 
         Height          =   360
         Index           =   0
         Left            =   1800
         Top             =   1560
         Width           =   360
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1200
         TabIndex        =   25
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   1200
         TabIndex        =   22
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblAsiento 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2550
         TabIndex        =   21
         Top             =   1440
         Width           =   4095
      End
      Begin VB.Label lblAsiento 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2550
         TabIndex        =   20
         Top             =   990
         Width           =   4095
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8610
      TabIndex        =   7
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdAccion 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   7050
      TabIndex        =   5
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Frame FrameTipoSalida 
      Caption         =   "Tipo de salida"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   8
      Top             =   3660
      Width           =   9795
      Begin VB.CommandButton PushButtonImpr 
         Caption         =   "Propiedades"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7080
         TabIndex        =   18
         Top             =   600
         Width           =   1515
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   9240
         TabIndex        =   17
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   9240
         TabIndex        =   16
         Top             =   1200
         Width           =   255
      End
      Begin VB.TextBox txtTipoSalida 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1680
         Width           =   7305
      End
      Begin VB.TextBox txtTipoSalida 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1200
         Width           =   7305
      End
      Begin VB.TextBox txtTipoSalida 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   600
         Width           =   5025
      End
      Begin VB.OptionButton optTipoSal 
         Caption         =   "eMail"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   12
         Top             =   2160
         Width           =   975
      End
      Begin VB.OptionButton optTipoSal 
         Caption         =   "PDF"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   1720
         Width           =   975
      End
      Begin VB.OptionButton optTipoSal 
         Caption         =   "Archivo csv"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   1515
      End
      Begin VB.OptionButton optTipoSal 
         Caption         =   "Impresora"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Label lblIndi 
      Caption         =   "Hasta"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2280
      TabIndex        =   33
      Top             =   6600
      Width           =   3735
   End
End
Attribute VB_Name = "frmEstadisticaCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ***********************************************************************************************************
' ***********************************************************************************************************
' ***********************************************************************************************************
'
'  3 espacios
'       -Los desde hasta,
'       -las opciones / ordenacion
'       -el tipo salida
'
' ***********************************************************************************************************
' ***********************************************************************************************************
' ***********************************************************************************************************
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Dim PrimeraVez As String
Dim Cad As String
Dim RT As ADODB.Recordset




Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then Unload Me
End Sub
Private Sub Form_Load()
    PrimeraVez = True

    Me.Icon = frmppal.Icon
        
    'Otras opciones
    Me.Caption = "Estadistica clientes"

    Me.Combo2.Clear
    Combo2.AddItem "Cobros pendientes"
    Combo2.AddItem "Resumen facturas"
    Combo2.AddItem "Detalle facturas"
    Combo2.ListIndex = 0
    'Combo2.AddItem "Detalle facturas"
    
    
    PrimeraVez = True
     
    Me.imgCli(0).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Me.imgCli(1).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
    lblIndi.Caption = ""
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub frmF_Selec(vFecha As Date)
    Cad = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgCli_Click(Index As Integer)
    CadenaDesdeOtroForm = ""
    frmcolClientesBusqueda.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
        txtCliente(Index).Text = CadenaDesdeOtroForm
        txtCliente_LostFocus Index
        If Index = 0 Then
            PonFoco txtCliente(1)
        Else
            PonerFocoCmb Combo1
        End If
    End If

End Sub

Private Sub imgFecha_Click(Index As Integer)

    Set frmF = New frmCal
    frmF.Fecha = Now
    txtFecha(0).Tag = Index
    If txtFecha(Index).Text <> "" Then frmF.Fecha = CDate(txtFecha(Index).Text)
    Cad = ""
    frmF.Show vbModal
    Set frmF = Nothing
    If Cad <> "" Then
        txtFecha(Index) = Cad
        PonFoco txtFecha(Index)
    End If
End Sub

Private Sub imgppal_Click(Index As Integer)

End Sub


Private Sub optTipoSal_Click(Index As Integer)
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), Index
End Sub


Private Sub PushButton2_Click(Index As Integer)
'    'FILTROS
'    If Index = 0 Then
'        frmppal.cd1.Filter = "*.csv|*.csv"
'
'    Else
'        frmppal.cd1.Filter = "*.pdf|*.pdf"
'    End If
'    frmppal.cd1.InitDir = App.Path & "\Exportar" 'PathSalida
'    frmppal.cd1.FilterIndex = 1
'    frmppal.cd1.ShowSave
'    If frmppal.cd1.FileTitle <> "" Then
'        If Dir(frmppal.cd1.FileName, vbArchive) <> "" Then
'            If MsgBox("El archivo ya existe. Reemplazar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
'        End If
'        txtTipoSalida(Index + 1).Text = frmppal.cd1.FileName
'    End If
End Sub

Private Sub PushButtonImpr_Click()
 '   frmppal.cd1.ShowPrinter
    PonerDatosPorDefectoImpresion Me, True
End Sub



Private Sub LanzaFormAyuda(Nombre As String, indice As Integer)
    Select Case Nombre
    Case "imgFecha"
        'imgFec_Click Indice
    End Select
    
End Sub




Private Sub cmdAccion_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset
    
    HacerAcciones Index
    
    lblIndi.Caption = ""
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
End Sub
Private Sub HacerAcciones(Index As Integer)
Dim B As Boolean
Dim C As String
   
    Conn.Execute "Delete from tmpcomun1 where codusu =" & vUsu.Codigo
    
    
    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    
    InicializarVbles True
    
    
    If Not PonerDesdeHasta("clientes.codclien", "N", Me.txtCliente(0), Me.txtDescCliente(0), Me.txtCliente(1), Me.txtDescCliente(1), "pDH1=""Cliente ") Then Exit Sub
    
    
    
    I = InStr(1, cadParam, "pDH1=""")
    If I > 0 Then
        'NO existe
        J = InStr(I, cadParam, "|")
        If J = 0 Then Err.Raise 513, , "Imposible situar parametros DesdeHsta"
    
        Msg = Mid(cadParam, J + 1)
        
        Cad = Mid(cadParam, I + 6, J - I - 7) '6 +1 (la comilla)
        cadParam = Mid(cadParam, 1, I - 1)
        cadParam = cadParam & Msg
    Else
        Cad = ""
    End If

    
    
    
    If Me.Combo1.ListIndex > 0 Then
        'Ha seleccionado o socio o no socio
        'Para llevar cadparam
        C = "{clientes.essocio} =" & IIf(Combo1.ListIndex = 1, "1", "0")
        If Not AnyadirAFormula(cadFormula, C) Then Exit Sub
           
        'Añadimos si es socio o no
        Cad = Trim(Cad & "          Socio: " & IIf(Combo1.ListIndex = 1, "Si", "No"))
        
    End If
    
    If Me.Combo3.ListIndex > 0 Then
           
        'Añadimos si es tipo ventas  en el desde hasta
        Cad = Trim(Cad & "          Tipo: " & Combo3.Text)
        
    End If
    
    
    cadParam = cadParam & "pDH1=""" & Cad & """|"
        
        
        
        
    lblIndi.Caption = "Leyendo registros BD"
    lblIndi.Refresh
        
    If cadselect = "" Then cadselect = " true "
    C = "Select codclien FROM clientes where " & cadselect
    NumRegElim = 0
    
    miRsAux.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    C = ""
    Msg = ""
    While Not miRsAux.EOF
        C = C & ", (" & vUsu.Codigo & "," & miRsAux!CodClien & ")"
        NumRegElim = NumRegElim + 1
        
        If (NumRegElim Mod 100) = 0 Then
            C = Mid(C, 2)
            lblIndi.Caption = "Leyendo registros BD..." & NumRegElim
            lblIndi.Refresh
            C = "INSERT INTO tmpcomun1(codusu,codigo) VALUES " & C
            Conn.Execute C
            C = ""
        End If
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If C <> "" Then
        lblIndi.Caption = "Registros leidos" & NumRegElim
        lblIndi.Refresh
        C = "INSERT INTO tmpcomun1(codusu,codigo) VALUES " & Mid(C, 2)
        Conn.Execute C
    End If
        
    If NumRegElim = 0 Then
        MsgBox "No existen datos", vbExclamation
        Exit Sub
    End If
        
    
    
    
    Select Case Combo2.ListIndex
    Case 2
        B = DetalleFactura
    Case 1
        B = FacturasResumen
    Case Else
        'Cobros pendientes
        B = CobrosPendientes
        
    End Select
        
    If Not B Then Exit Sub
    
    'Fechas alta
    Msg = ""
    ValorAnterior = ""
    If Me.txtFecha(0).Text <> "" Then
        If cadselect <> "" Then
            cadselect = cadselect & " AND "
            cadFormula = cadFormula & " AND "
        End If
        If ValorAnterior = "" Then
            Msg = "Fecha factura: "
            ValorAnterior = "1"
        End If
        
        Msg = Msg & " desde " & Me.txtFecha(0).Text
        cadselect = cadselect & " clientes.fechaltaaso >= " & DBSet(txtFecha(0).Text, "F")
        cadFormula = cadFormula & " {clientes.fechaltaaso} >= cdate(" & Format(txtFecha(0).Text, "yyyy,mm,dd") & ")"
    End If
    If Me.txtFecha(1).Text <> "" Then
        If cadselect <> "" Then
            cadselect = cadselect & " AND "
            cadFormula = cadFormula & " AND "
        End If
        If ValorAnterior = "" Then
            Msg = "Fecha factura: "
            ValorAnterior = "1"
        End If
        
        Msg = Msg & " hasta " & Me.txtFecha(1).Text
        cadselect = cadselect & " clientes.fechaltaaso <= " & DBSet(txtFecha(1).Text, "F")
        cadFormula = cadFormula & " {clientes.fechaltaaso} <= cdate(" & Format(txtFecha(1).Text, "yyyy,mm,dd") & ")"
    End If
    
    
    cadParam = cadParam & "pDH2=""" & Msg & """|"
    numParam = numParam + 1
    
    cadFormula = "{tmpcomun.codusu} = " & vUsu.Codigo
    
    
    If optTipoSal(1).Value Then
        'EXPORTAR A CSV
        AccionesCSV
    
    Else
        'Tanto a pdf,imprimiir, preevisualizar como email van COntral Crystal
    
        If optTipoSal(2).Value Or optTipoSal(3).Value Then
            ExportarPDF = True 'generaremos el pdf
        Else
            ExportarPDF = False
        End If
        SoloImprimir = False
        If Index = 0 Then SoloImprimir = True 'ha pulsado impirmir
        
        AccionesCrystal
    End If
    
    
End Sub





Private Sub AccionesCSV()
Dim SQL As String

'    'Monto el SQL
    
    SQL = "select codclien,nomclien,domclien,codposta,pobclien,proclien,essocio,nifclien,fechaltaaso,fechaltaact,"
    SQL = SQL & " licencia,matricula,telefono,telmovil,maiclien,clientes.iban,"
    SQL = SQL & " clientes.codforpa,nomforpa from clientes,ariconta" & vParam.Numconta & ".formapago where clientes.codforpa=formapago.codforpa"
    If cadselect <> "" Then
        SQL = SQL & " AND " & cadselect
    End If
    GeneraFicheroCSV SQL, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal()
Dim indRPT As String
Dim nomDocu As String
    
    vMostrarTree = False
    conSubRPT = False
        
    
    
    'If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    Select Case Combo2.ListIndex
    Case 2
        nomDocu = "rEstaFacturasDeta.rpt"
    Case 1
        nomDocu = "rEstaFacturas.rpt"
    Case Else
        'Cbros pdtes
        nomDocu = "rEstaCobrosPdte.rpt"
    End Select
    cadNomRPT = nomDocu


    
    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then CopiarFicheroASalida False, txtTipoSalida(2).Text, False
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 2
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
End Sub





Private Sub txtCliente_GotFocus(Index As Integer)
    ConseguirFoco txtCliente(Index), 3
End Sub

Private Sub txtCliente_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = teclaBuscar Then
        KeyAscii = 0
        imgCli_Click Index
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub txtCliente_LostFocus(Index As Integer)
    InformeTxtLostFocus txtCliente(Index), Me.txtDescCliente(Index), False
End Sub

Private Sub txtFecha_GotFocus(Index As Integer)
    ConseguirFoco txtFecha(Index), 3
End Sub

Private Sub txtFecha_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        KeyAscii = 0
        imgppal_Click Index
    Else
        KEYpress KeyAscii
    End If
End Sub


Private Sub txtfecha_LostFocus(Index As Integer)
    If txtFecha(Index).Text <> "" Then
        If Not PonerFormatoFecha(txtFecha(Index)) Then txtFecha(Index).Text = ""
    End If
End Sub


Private Function CobrosPendientes() As Boolean
Dim C As String
Dim Impor As Currency

    On Error GoTo eCobrosPendientes
    
    CobrosPendientes = False
            
    'Cargo todos los cobros pdtes
    Set RT = New ADODB.Recordset
    lblIndi.Caption = "Cobros pendientes"
    lblIndi.Refresh
    
    C = "DELETE FROM tmpcomun WHERE codusu =" & vUsu.Codigo
    Conn.Execute C
    
    C = "select * from ariconta" & vParam.Numconta & ".cobros where (impvenci + coalesce(gastos,0) -coalesce(impcobro,0))<0"
    If txtFecha(0).Text <> "" Then C = C & " AND fecfactu >=" & DBSet(txtFecha(0), "F")
    If txtFecha(1).Text <> "" Then C = C & " AND fecfactu <=" & DBSet(txtFecha(1), "F")
    
    If Me.Combo3.ListIndex > 0 Then C = C & " AND impvenci " & IIf(Combo3.ListIndex = 1, ">", "<") & " 0"
    
    C = C & " ORDER BY codmacta,fecfactu"
    
    RT.Open C, Conn, adOpenKeyset, adLockReadOnly
            
    If RT.EOF Then
        RT.Close
        Exit Function
    End If
        
    
    C = "SELECT * from tmpcomun1 where codusu =" & vUsu.Codigo & " ORDER BY  codigo"
    miRsAux.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    C = ""
    Msg = ""
    MsgErr = "insert into `tmpcomun` (`codusu`,`codigo`,`codigo1`,codigo2 ,`texto1`,`importe1`,`importe2`,`fecha1`,`fecha2`) VALUES "
    NumRegElim = 0
    While Not miRsAux.EOF
        lblIndi.Caption = "Datos  " & miRsAux!Codigo
        lblIndi.Refresh
    
        For I = 1 To 2
            C = DevuelveCuentaContableCliente(I = 1, CStr(miRsAux!Codigo))
            Cad = "codmacta = '" & C & "'"
            RT.Find Cad, , adSearchForward, 1
            C = ""
            While Not RT.EOF
                
                NumRegElim = NumRegElim + 1
                '                                 clien      forpa     idvto    total      pdte       fecfac   fecvenci
                'insert into `tmpcomun` (`codusu`,`codigo`,`codigo1`,`texto1`,`importe1`,`importe2`,`fecha1`,`fecha2`)
                C = C & ", (" & vUsu.Codigo & "," & NumRegElim & "," & miRsAux!Codigo & "," & RT!Codforpa & ",'" & RT!numSerie & Format(RT!NumFactu, "000000") & "',"
                Impor = RT!ImpVenci + DBLet(RT!Gastos, "N")
                C = C & DBSet(Impor, "N") & ","
                Impor = Impor - DBLet(RT!impcobro, "N")
                C = C & DBSet(Impor, "N") & "," & DBSet(RT!Fecfactu, "F") & "," & DBSet(RT!FecVenci, "F") & ")"
                RT.MoveNext
                RT.Find Cad, , adSearchForward
            Wend
            If C <> "" Then Msg = Msg & C
            RT.MoveFirst
        Next I
        miRsAux.MoveNext
        
        If Len(Msg) > 10000 Then
            Msg = Mid(Msg, 2)
            C = MsgErr & Msg
            Conn.Execute C
            Msg = ""
        End If
    Wend
    miRsAux.Close
        
    If Msg <> "" Then
        Msg = Mid(Msg, 2)
        C = MsgErr & Msg
        Conn.Execute C
        Msg = ""
    End If

    If NumRegElim = 0 Then
        MsgBox "Ningun datos generado", vbExclamation
        Exit Function
    End If

    CobrosPendientes = True
    
eCobrosPendientes:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set RT = Nothing
End Function








Private Function FacturasResumen() As Boolean
Dim C As String
Dim Impor As Currency
Dim Aux As String

    On Error GoTo eFacturasResumen
    
    FacturasResumen = False
            
    'Cargo todos los cobros pdtes
    Set RT = New ADODB.Recordset
    lblIndi.Caption = "Leyendo fra resumen"
    lblIndi.Refresh
    
    C = "DELETE FROM tmpcomun WHERE codusu =" & vUsu.Codigo
    Conn.Execute C
    
    
    
    
    C = "Select codclien,numserie,count(*),sum(totfaccl) FROM factcli where 1=1"
    If txtFecha(0).Text <> "" Then C = C & " AND fecfactu >=" & DBSet(txtFecha(0), "F")
    If txtFecha(1).Text <> "" Then C = C & " AND fecfactu <=" & DBSet(txtFecha(1), "F")
    If Me.Combo3.ListIndex > 0 Then C = C & " AND totfaccl " & IIf(Combo3.ListIndex = 1, ">", "<") & " 0"
    C = C & " AND codclien in ( SELECT codigo FROM tmpcomun1 WHERE codusu =" & vUsu.Codigo & ")"
    C = C & " GROUP BY codclien,numserie ORDER BY 1,2"
    
    
    miRsAux.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        miRsAux.Close
        MsgBox "Ningun dato generado", vbExclamation
        Exit Function
    End If
    
    
    C = ""
    Msg = ""
    MsgErr = "INSERT INTO tmpcomun(codusu,codigo,codigo1,texto1,texto2,texto3,texto4,texto5,importe1) VALUES "
    NumRegElim = -1
    While Not miRsAux.EOF
        'If miRsAux!CodClien = 33378 Then St op
        lblIndi.Caption = "Cliente " & miRsAux!CodClien
        lblIndi.Refresh
        
        
        If miRsAux!CodClien <> NumRegElim Then
            
            If NumRegElim >= 0 Then CadenaFrasClie C
            
            NumRegElim = miRsAux!CodClien
            
            
            If Len(Msg) > 10000 Then
                Msg = Mid(Msg, 2)
                C = MsgErr & Msg
                Conn.Execute C
                Msg = ""
            End If
                
                
            'CUO|FEX|FDI|FAC|FRT|
            C = "cedar"
                
        End If
        
        I = InStr(1, "CUO|FEX|FDI|FAC|FRT|", miRsAux!numSerie)
        If I = 0 Then Err.Raise 513, , "Tipo factura no tratado: " & miRsAux!numSerie
        
        I = ((I - 1) \ 4) + 1
        If IsNull(miRsAux.Fields(3)) Then
            Aux = "0"
        Else
            Aux = CStr(miRsAux.Fields(3))
        End If
        
        Aux = "'" & Format(miRsAux.Fields(2), "0000") & Aux & "',"
        C = Replace(C, Mid("cedar", CInt(I), 1), Aux)
        
        
        
        
        
        miRsAux.MoveNext
        
    Wend
    miRsAux.Close
        
    'El ultimo siempre hay que hacerlo. Hay un EOF antes de empezar
    CadenaFrasClie C
    
    If Msg <> "" Then
        Msg = Mid(Msg, 2)
        C = MsgErr & Msg
        Conn.Execute C
        Msg = ""
    End If

    If NumRegElim = 0 Then
        MsgBox "Ningun datos generado", vbExclamation
        Exit Function
    End If

    FacturasResumen = True
    
eFacturasResumen:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set RT = Nothing
End Function

Private Sub CadenaFrasClie(ByRef C As String)

        For J = 1 To 5
            C = Replace(C, Mid("cedar", CInt(J), 1), "'00000',")
        Next
        
        'MsgErr = "INSERT INTO tmpcomun(codusu,codigo,codigo1,texto1,texto2,texto3,texto4,texto5,importe1,importe2) VALUES "
        C = ", (" & vUsu.Codigo & "," & NumRegElim & "," & NumRegElim & "," & C & "null)"
        Msg = Msg & C
    
End Sub


Private Function DetalleFactura() As Boolean
Dim C As String

    On Error GoTo eDetalleFactura
    
    DetalleFactura = False
            
    'Cargo todos los cobros pdtes
    
    lblIndi.Caption = "Leyendo fra resumen"
    lblIndi.Refresh
    
    C = "DELETE FROM tmpcomun WHERE codusu =" & vUsu.Codigo
    Conn.Execute C
    
    
    
    
    C = "Select numserie,numfactu,fecfactu,codclien FROM factcli where 1=1"
    If txtFecha(0).Text <> "" Then C = C & " AND fecfactu >=" & DBSet(txtFecha(0), "F")
    If txtFecha(1).Text <> "" Then C = C & " AND fecfactu <=" & DBSet(txtFecha(1), "F")
    If Me.Combo3.ListIndex > 0 Then C = C & " AND totfaccl " & IIf(Combo3.ListIndex = 1, ">", "<") & " 0"
    C = C & " AND codclien in ( SELECT codigo FROM tmpcomun1 WHERE codusu =" & vUsu.Codigo & ") order by codclien"
    miRsAux.Open C, Conn, adOpenForwardOnly, adLockPessimistic
    C = ""
    NumRegElim = 0
    While Not miRsAux.EOF
        NumRegElim = NumRegElim + 1
        lblIndi.Caption = "Cliente " & miRsAux!CodClien
        lblIndi.Refresh
        
        'codusu,codigo,texto1,codigo1,fecha1
        C = C & ", (" & vUsu.Codigo & "," & NumRegElim & ",'" & miRsAux!numSerie & "'," & miRsAux!NumFactu & "," & DBSet(miRsAux!Fecfactu, "F") & ")"
        If (NumRegElim Mod 50) = 0 Then
            'codusu,codigo,texto1,codigo1,fecha1
            C = Mid(C, 2)
            C = "INSERT INTO tmpcomun(codusu,codigo,texto1,codigo1,fecha1) VALUES " & C
            Conn.Execute C
            C = ""
        End If
        miRsAux.MoveNext
    Wend
    If C <> "" Then
        C = Mid(C, 2)
        C = "INSERT INTO tmpcomun(codusu,codigo,texto1,codigo1,fecha1) VALUES " & C
        Conn.Execute C
    End If
    If NumRegElim = 0 Then
        MsgBox "Ningun registro", vbExclamation
    Else
        DetalleFactura = True
    End If
    Exit Function
eDetalleFactura:
    MuestraError Err.Number, Err.Description
End Function
