VERSION 5.00
Begin VB.Form frmClientesList 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11385
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   11385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameConcepto 
      Caption         =   "Selecci�n"
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
      TabIndex        =   26
      Top             =   0
      Width           =   10995
      Begin VB.TextBox txtClienteL 
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
         Left            =   4560
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "imgConcepto"
         Top             =   1920
         Width           =   1185
      End
      Begin VB.TextBox txtClienteL 
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
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "imgConcepto"
         Top             =   1920
         Width           =   1185
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
         Index           =   3
         Left            =   4560
         MaxLength       =   10
         TabIndex        =   7
         Tag             =   "imgConcepto"
         Top             =   3120
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
         Index           =   2
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   6
         Tag             =   "imgConcepto"
         Top             =   3120
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
         Index           =   0
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "imgConcepto"
         Top             =   2550
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
         Left            =   4560
         MaxLength       =   10
         TabIndex        =   5
         Tag             =   "imgConcepto"
         Top             =   2550
         Width           =   1305
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Contacto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   8520
         TabIndex        =   9
         Top             =   1200
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Tesoreria"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   8520
         TabIndex        =   10
         Top             =   1560
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Datos basicos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   8520
         TabIndex        =   8
         Top             =   840
         Value           =   -1  'True
         Width           =   1695
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
         ItemData        =   "frmClientesList.frx":0000
         Left            =   9000
         List            =   "frmClientesList.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2400
         Width           =   1215
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
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "imgConcepto"
         Top             =   1200
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
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   31
         Tag             =   "imgConcepto"
         Top             =   720
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
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   30
         Tag             =   "imgConcepto"
         Top             =   1200
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
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "imgConcepto"
         Top             =   720
         Width           =   1065
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
         Index           =   13
         Left            =   3480
         TabIndex        =   44
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
         Index           =   12
         Left            =   960
         TabIndex        =   43
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Licencia"
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
         Index           =   11
         Left            =   240
         TabIndex        =   42
         Top             =   1680
         Width           =   885
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   4080
         Picture         =   "frmClientesList.frx":001B
         Top             =   3180
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
         Index           =   10
         Left            =   3480
         TabIndex        =   41
         Top             =   3210
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Baja"
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
         Index           =   9
         Left            =   240
         TabIndex        =   40
         Top             =   3180
         Width           =   495
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   1680
         Picture         =   "frmClientesList.frx":00A6
         Top             =   3180
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
         Index           =   8
         Left            =   960
         TabIndex        =   39
         Top             =   3210
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Alta"
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
         Left            =   240
         TabIndex        =   38
         Top             =   2610
         Width           =   435
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1680
         Picture         =   "frmClientesList.frx":0131
         Top             =   2610
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
         Left            =   960
         TabIndex        =   37
         Top             =   2640
         Width           =   615
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   4080
         Picture         =   "frmClientesList.frx":01BC
         Top             =   2610
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
         Left            =   3480
         TabIndex        =   36
         Top             =   2633
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
         Left            =   8160
         TabIndex        =   35
         Top             =   2400
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Datos a mostrar"
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
         Left            =   7800
         TabIndex        =   34
         Top             =   480
         Width           =   1755
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
         Index           =   2
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Width           =   765
      End
      Begin VB.Image imgCli 
         Height          =   360
         Index           =   1
         Left            =   1680
         Top             =   1200
         Width           =   360
      End
      Begin VB.Image imgCli 
         Height          =   360
         Index           =   0
         Left            =   1680
         Top             =   720
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
         Left            =   960
         TabIndex        =   32
         Top             =   1200
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
         Left            =   960
         TabIndex        =   29
         Top             =   720
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
         TabIndex        =   28
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
         TabIndex        =   27
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
      Left            =   9330
      TabIndex        =   14
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
      Left            =   7770
      TabIndex        =   12
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
      TabIndex        =   13
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
      TabIndex        =   15
      Top             =   3660
      Width           =   10995
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
         TabIndex        =   25
         Top             =   600
         Width           =   1515
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   9240
         TabIndex        =   24
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   9240
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   600
         Value           =   -1  'True
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmClientesList"
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
    Me.Caption = "Listado de clientes"

    
    PrimeraVez = True
     
    Me.imgCli(0).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Me.imgCli(1).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
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
 Dim C As String
   Dim Lic As String
    
    
    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    
    InicializarVbles True
    
    If Not PonerDesdeHasta("clientes.codclien", "N", Me.txtCliente(0), Me.txtDescCliente(0), Me.txtCliente(1), Me.txtDescCliente(1), "pDH1=""Cliente ") Then Exit Sub
    If Me.Combo1.ListIndex > 0 Then
        'Ha seleccionado o socio o no socio
        C = "{clientes.essocio} =" & IIf(Combo1.ListIndex = 1, "1", "0")
        If Not AnyadirAFormula(cadFormula, C) Then Exit Sub
        C = " clientes.essocio =" & IIf(Combo1.ListIndex = 1, "1", "0")
        If Not AnyadirAFormula(cadselect, C) Then Exit Sub
        C = ""
        
        I = InStr(1, cadParam, "pDH1=""")
        If I > 0 Then
            'NO existe
            J = InStr(I, cadParam, "|")
            If J = 0 Then Err.Raise 513, , "Imposible situar parametros DesdeHsta"
        
            Msg = Mid(cadParam, J + 1)
            
            C = Mid(cadParam, I + 6, J - I - 7) '6 +1 (la comilla)
            cadParam = Mid(cadParam, 1, I - 1)
            cadParam = cadParam & Msg
        Else
            
        End If
        
        'A�adimos si es socio o no
        C = Trim(C & "          Socio: " & IIf(Combo1.ListIndex = 1, "Si", "No"))
        cadParam = cadParam & "pDH1=""" & C & """|"
    End If
    
    Lic = ""
    If Me.txtClienteL(0).Text <> "" Or txtClienteL(1).Text <> "" Then
        
        Lic = "Licencia "
        If Me.txtClienteL(0).Text <> "" Then
            C = "{clientes.licencia} >= " & DBSet(Me.txtClienteL(0).Text, "T")
            If Not AnyadirAFormula(cadFormula, C) Then Exit Sub
            C = "(clientes.licencia) >= " & DBSet(Me.txtClienteL(0).Text, "T")
            If Not AnyadirAFormula(cadselect, C) Then Exit Sub
            Lic = Lic & " desde " & Me.txtClienteL(0).Text
        End If
        If Me.txtClienteL(1).Text <> "" Then
            C = "{clientes.licencia} <= " & DBSet(Me.txtClienteL(1).Text, "T")
            If Not AnyadirAFormula(cadFormula, C) Then Exit Sub
            C = "(clientes.licencia) <= " & DBSet(Me.txtClienteL(1).Text, "T")
            If Not AnyadirAFormula(cadselect, C) Then Exit Sub
            Lic = Lic & " hasta " & Me.txtClienteL(1).Text
        End If
        
        
        C = ""

    End If
    
    
    'Fechas alta
    Msg = ""
    ValorAnterior = ""
    If Me.txtFecha(0).Text <> "" Then
        If cadselect <> "" Then
            cadselect = cadselect & " AND "
            cadFormula = cadFormula & " AND "
        End If
        If ValorAnterior = "" Then
            Msg = "Fecha alta: "
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
            Msg = "Fecha alta: "
            ValorAnterior = "1"
        End If
        
        Msg = Msg & " hasta " & Me.txtFecha(1).Text
        cadselect = cadselect & " clientes.fechaltaaso <= " & DBSet(txtFecha(1).Text, "F")
        cadFormula = cadFormula & " {clientes.fechaltaaso} <= cdate(" & Format(txtFecha(1).Text, "yyyy,mm,dd") & ")"
    End If
    'Fecha baja -----------------------------------------------------------
    ValorAnterior = ""
    If Me.txtFecha(2).Text <> "" Then
        If cadselect <> "" Then
            cadselect = cadselect & " AND "
            cadFormula = cadFormula & " AND "
        End If
        If ValorAnterior = "" Then
            Msg = "Fecha baja: "
            ValorAnterior = "1"
        End If
        
        Msg = Msg & " desde " & Me.txtFecha(2).Text
        cadselect = cadselect & " clientes.fechabajact >= " & DBSet(txtFecha(2).Text, "F")
        cadFormula = cadFormula & " {clientes.fechabajact} >= cdate(" & Format(txtFecha(2).Text, "yyyy,mm,dd") & ")"
    End If
    If Me.txtFecha(3).Text <> "" Then
        If cadselect <> "" Then
            cadselect = cadselect & " AND "
            cadFormula = cadFormula & " AND "
        End If
        If ValorAnterior = "" Then
            Msg = "Fecha baja: "
            ValorAnterior = "1"
        End If
        
        Msg = Msg & " hasta " & Me.txtFecha(3).Text
        cadselect = cadselect & " clientes.fechabajact <= " & DBSet(txtFecha(3).Text, "F")
        cadFormula = cadFormula & " {clientes.fechabajact} <= cdate(" & Format(txtFecha(3).Text, "yyyy,mm,dd") & ")"
    End If
    
    Msg = Trim(Msg & "      " & Lic)
    cadParam = cadParam & "pDH2=""" & Msg & """|"
    numParam = numParam + 1
    
    
    
    
    
    
    
    
    
    
    
    
    If Not HayRegParaInforme("clientes", cadselect) Then Exit Sub
    
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
Dim Sql As String

'    'Monto el SQL
    
    Sql = "select codclien,nomclien,domclien,codposta,pobclien,proclien,essocio,nifclien,fechaltaaso,fechaltaact,"
    Sql = Sql & " licencia,matricula,telefono,telmovil,maiclien,clientes.iban,"
    Sql = Sql & " clientes.codforpa,nomforpa from clientes,ariconta" & vParam.Numconta & ".formapago where clientes.codforpa=formapago.codforpa"
    If cadselect <> "" Then
        Sql = Sql & " AND " & cadselect
    End If
    GeneraFicheroCSV Sql, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal()
Dim indRPT As String
Dim nomDocu As String
    
    vMostrarTree = False
    conSubRPT = False
        
    
    
    'If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    If Me.Option1(2).Value Then
        nomDocu = "rClientesTesoreria.rpt"
    ElseIf Me.Option1(1).Value Then
        nomDocu = "rClientesContac.rpt"
    Else
        nomDocu = "rClientes.rpt"
    End If
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

Private Sub txtClienteL_GotFocus(Index As Integer)
ConseguirFoco txtFecha(Index), 3

End Sub

Private Sub txtClienteL_KeyPress(Index As Integer, KeyAscii As Integer)
KEYpress KeyAscii
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

