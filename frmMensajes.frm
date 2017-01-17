VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMensajes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16440
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMensajes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10545
   ScaleWidth      =   16440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCompraTasas 
      Height          =   3615
      Left            =   2160
      TabIndex        =   110
      Top             =   840
      Visible         =   0   'False
      Width           =   8775
      Begin VB.CommandButton cmdCompraTasas 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   1
         Left            =   5400
         TabIndex        =   113
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CommandButton cmdCompraTasas 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   6960
         TabIndex        =   114
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox txtTasas 
         Height          =   360
         Index           =   1
         Left            =   2880
         MaxLength       =   30
         TabIndex        =   112
         Top             =   2400
         Width           =   1035
      End
      Begin VB.TextBox txtTasas 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   0
         Left            =   2880
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   117
         Top             =   1800
         Width           =   5355
      End
      Begin VB.TextBox txtFecha 
         Height          =   360
         Index           =   6
         Left            =   2880
         MaxLength       =   30
         TabIndex        =   111
         Text            =   "commor"
         Top             =   1200
         Width           =   2475
      End
      Begin VB.Label Label1 
         Caption         =   "Cantidad"
         Height          =   240
         Index           =   25
         Left            =   240
         TabIndex        =   119
         ToolTipText     =   "Fecha alta asociado"
         Top             =   2400
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "Concepto"
         Height          =   240
         Index           =   24
         Left            =   240
         TabIndex        =   118
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1800
         Width           =   2115
      End
      Begin VB.Label Label7 
         Caption         =   "Compra de tasas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   7
         Left            =   2520
         TabIndex        =   116
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha / hora compra"
         Height          =   240
         Index           =   23
         Left            =   240
         TabIndex        =   115
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1200
         Width           =   2115
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   6
         Left            =   2520
         Picture         =   "frmMensajes.frx":000C
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1200
         Width           =   240
      End
   End
   Begin VB.Frame FrameCierreCaja 
      Height          =   5535
      Left            =   4680
      TabIndex        =   82
      Top             =   240
      Visible         =   0   'False
      Width           =   6495
      Begin VB.TextBox txtCaja 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   7
         Left            =   2400
         MaxLength       =   30
         TabIndex        =   84
         Top             =   3840
         Width           =   1515
      End
      Begin VB.TextBox txtCaja 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   6
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   88
         Top             =   3240
         Width           =   1515
      End
      Begin VB.TextBox txtCaja 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   5
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   92
         Top             =   2280
         Width           =   1515
      End
      Begin VB.TextBox txtFecha 
         Height          =   360
         Index           =   4
         Left            =   2400
         MaxLength       =   30
         TabIndex        =   83
         Text            =   "commor"
         Top             =   1200
         Width           =   2475
      End
      Begin VB.CommandButton cmdCierreCaja 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   1
         Left            =   3360
         TabIndex        =   85
         Top             =   4800
         Width           =   1335
      End
      Begin VB.CommandButton cmdCierreCaja 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   0
         Left            =   4920
         TabIndex        =   86
         Top             =   4800
         Width           =   1335
      End
      Begin VB.TextBox txtCaja 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   4
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   87
         Top             =   1680
         Width           =   2475
      End
      Begin VB.Label Label1 
         Caption         =   "Queda en caja"
         Height          =   240
         Index           =   17
         Left            =   600
         TabIndex        =   95
         ToolTipText     =   "Fecha alta asociado"
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         X1              =   480
         X2              =   6240
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Label Label1 
         Caption         =   "€ a entregar"
         Height          =   240
         Index           =   14
         Left            =   600
         TabIndex        =   94
         ToolTipText     =   "Fecha alta asociado"
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Importe actual"
         Height          =   240
         Index           =   16
         Left            =   600
         TabIndex        =   93
         ToolTipText     =   "Fecha alta asociado"
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha cierre"
         Height          =   240
         Index           =   15
         Left            =   600
         TabIndex        =   91
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1200
         Width           =   1395
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   4
         Left            =   2040
         Picture         =   "frmMensajes.frx":0097
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label7 
         Caption         =   "Cierre de caja"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   5
         Left            =   2040
         TabIndex        =   90
         Top             =   360
         Width           =   3150
      End
      Begin VB.Label Label1 
         Caption         =   "Usuario"
         Height          =   240
         Index           =   13
         Left            =   600
         TabIndex        =   89
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1800
         Width           =   1395
      End
   End
   Begin VB.Frame FrameCaja 
      Height          =   3735
      Left            =   1800
      TabIndex        =   69
      Top             =   0
      Visible         =   0   'False
      Width           =   8055
      Begin VB.TextBox txtCaja 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   3
         Left            =   5160
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   80
         Top             =   1080
         Width           =   2475
      End
      Begin VB.CheckBox chkCaja 
         Caption         =   "Salida"
         Height          =   240
         Left            =   3360
         TabIndex        =   72
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox txtCaja 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   2
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   71
         Top             =   2520
         Width           =   1515
      End
      Begin VB.TextBox txtCaja 
         Height          =   360
         Index           =   1
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   70
         Top             =   1800
         Width           =   6195
      End
      Begin VB.TextBox txtCaja 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   0
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   75
         Top             =   1080
         Width           =   2595
      End
      Begin VB.CommandButton cmdCaja 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   5040
         TabIndex        =   73
         Top             =   3120
         Width           =   1335
      End
      Begin VB.CommandButton cmdCaja 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   6480
         TabIndex        =   74
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Usuario"
         Height          =   240
         Index           =   12
         Left            =   4320
         TabIndex        =   81
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1080
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Importe"
         Height          =   240
         Index           =   11
         Left            =   240
         TabIndex        =   79
         ToolTipText     =   "Fecha alta asociado"
         Top             =   2520
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Concepto"
         Height          =   240
         Index           =   10
         Left            =   240
         TabIndex        =   78
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1800
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   240
         Index           =   9
         Left            =   240
         TabIndex        =   77
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1080
         Width           =   1395
      End
      Begin VB.Label Label7 
         Caption         =   "Movimiento de caja"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   375
         Index           =   4
         Left            =   2400
         TabIndex        =   76
         Top             =   240
         Width           =   3150
      End
   End
   Begin VB.Frame FrameFraPrevision 
      Height          =   4215
      Left            =   240
      TabIndex        =   43
      Top             =   0
      Visible         =   0   'False
      Width           =   6855
      Begin VB.TextBox Text1 
         Height          =   1575
         Left            =   2040
         TabIndex        =   45
         Text            =   "Text1"
         Top             =   1680
         Width           =   4575
      End
      Begin VB.CommandButton cmdFacturarPrevision 
         Caption         =   "&Facturar"
         Height          =   375
         Index           =   1
         Left            =   3960
         TabIndex        =   46
         Top             =   3600
         Width           =   1215
      End
      Begin VB.CommandButton cmdFacturarPrevision 
         Caption         =   "Salir"
         Height          =   375
         Index           =   0
         Left            =   5400
         TabIndex        =   47
         Top             =   3600
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   360
         Index           =   1
         Left            =   2040
         MaxLength       =   30
         TabIndex        =   44
         Text            =   "commor"
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "Ampliación"
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   50
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1680
         Width           =   1395
      End
      Begin VB.Label Label7 
         Caption         =   "Facturar periódicas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   49
         Top             =   360
         Width           =   3150
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   1
         Left            =   1680
         Picture         =   "frmMensajes.frx":0122
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha factura"
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   48
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1080
         Width           =   1395
      End
   End
   Begin VB.Frame FramePedirFechaFactura 
      Height          =   2415
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton cmdFacturarExp 
         Caption         =   "Facturar"
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   40
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmdFacturarExp 
         Caption         =   "Salir"
         Height          =   375
         Index           =   0
         Left            =   3000
         TabIndex        =   41
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   360
         Index           =   0
         Left            =   2400
         MaxLength       =   30
         TabIndex        =   39
         Tag             =   "Fec. alta|F|N|||clientes|fechaltaaso|dd/mm/yyyy||"
         Text            =   "commor"
         Top             =   960
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha factura"
         Height          =   240
         Index           =   2
         Left            =   600
         TabIndex        =   42
         ToolTipText     =   "Fecha alta asociado"
         Top             =   960
         Width           =   1395
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   0
         Left            =   2040
         Picture         =   "frmMensajes.frx":01AD
         ToolTipText     =   "Fecha alta asociado"
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label7 
         Caption         =   "Facturar expediente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   38
         Top             =   240
         Width           =   3150
      End
   End
   Begin VB.Frame FrameErrorRestore 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   5775
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   4215
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   7435
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
      End
      Begin VB.Label Label29 
         Caption         =   "Cambio caracteres recupera backup"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   4935
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   4920
         Picture         =   "frmMensajes.frx":0238
         ToolTipText     =   "Quitar seleccion"
         Top             =   720
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   4920
         Picture         =   "frmMensajes.frx":0382
         ToolTipText     =   "Todos"
         Top             =   1080
         Width           =   240
      End
   End
   Begin VB.Frame FrameShowProcess 
      Height          =   6720
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   10950
      Begin VB.CommandButton CmdRegresar 
         Caption         =   "Salir"
         Height          =   375
         Left            =   9240
         TabIndex        =   30
         Top             =   6120
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   4905
         Left            =   225
         TabIndex        =   31
         Top             =   1005
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   8652
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label53 
         Caption         =   "Usuarios conectados a"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   390
         Width           =   10185
      End
   End
   Begin VB.Frame FrameInformeBBDD 
      Height          =   6720
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   10950
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Salir"
         Height          =   375
         Left            =   9240
         TabIndex        =   20
         Top             =   6120
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   4905
         Left            =   225
         TabIndex        =   21
         Top             =   1005
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   8652
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8370
         TabIndex        =   28
         Top             =   660
         Width           =   795
      End
      Begin VB.Label Label49 
         Alignment       =   1  'Right Justify
         Caption         =   "Porcentaje"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9240
         TabIndex        =   27
         Top             =   660
         Width           =   1005
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         TabIndex        =   26
         Top             =   660
         Width           =   795
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         Caption         =   "Porcentaje"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5910
         TabIndex        =   25
         Top             =   660
         Width           =   1005
      End
      Begin VB.Label Label48 
         Alignment       =   1  'Right Justify
         Caption         =   "Ejercicio Siguiente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6960
         TabIndex        =   24
         Top             =   300
         Width           =   3435
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         Caption         =   "Ejercicio Actual"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3210
         TabIndex        =   23
         Top             =   300
         Width           =   3705
      End
      Begin VB.Label Label46 
         Caption         =   "Concepto"
         Height          =   255
         Left            =   270
         TabIndex        =   22
         Top             =   660
         Width           =   2355
      End
   End
   Begin VB.Frame FrameBloqueoEmpresas 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7455
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   11415
      Begin VB.CommandButton cmdBlEmp 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   5520
         TabIndex        =   15
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton cmdBlEmp 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   5520
         TabIndex        =   14
         Top             =   3000
         Width           =   375
      End
      Begin VB.CommandButton cmdBlEmp 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   5520
         TabIndex        =   13
         Top             =   1920
         Width           =   375
      End
      Begin VB.CommandButton cmdBlEmp 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   5520
         TabIndex        =   10
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton cmdBloqEmpre 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   9840
         TabIndex        =   8
         Top             =   6840
         Width           =   1335
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   5775
         Index           =   0
         Left            =   210
         TabIndex        =   6
         Top             =   840
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   10186
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Empresa"
            Object.Width           =   5644
         EndProperty
      End
      Begin VB.CommandButton cmdBloqEmpre 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   8400
         TabIndex        =   5
         Top             =   6840
         Width           =   1335
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   5775
         Index           =   1
         Left            =   6240
         TabIndex        =   7
         Top             =   840
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   10186
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Empresa"
            Object.Width           =   5644
         EndProperty
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         Caption         =   "Bloqueadas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   10050
         TabIndex        =   12
         Top             =   600
         Width           =   1125
      End
      Begin VB.Label Label41 
         Caption         =   "Permitidas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Bloqueo de empresas por usuario"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   495
         Index           =   2
         Left            =   2880
         TabIndex        =   9
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Frame FramePagostasas 
      Height          =   8775
      Left            =   0
      TabIndex        =   51
      Top             =   0
      Visible         =   0   'False
      Width           =   16335
      Begin MSComctlLib.ListView ListView1 
         Height          =   5535
         Left            =   240
         TabIndex        =   68
         Top             =   2280
         Width           =   15855
         _ExtentX        =   27966
         _ExtentY        =   9763
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "NºExp"
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Lin"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cod.Cli"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Licencia"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Cliente"
            Object.Width           =   8009
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Cod."
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Concepto"
            Object.Width           =   5715
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Importe"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton cmdVerdatos 
         Caption         =   "Cargar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   11640
         TabIndex        =   67
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtCliente 
         Height          =   360
         Index           =   1
         Left            =   5160
         MaxLength       =   30
         TabIndex        =   65
         Top             =   1680
         Width           =   1035
      End
      Begin VB.TextBox txtClienteDes 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   1
         Left            =   6240
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   64
         Top             =   1680
         Width           =   3675
      End
      Begin VB.TextBox txtClienteDes 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   0
         Left            =   6240
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   62
         Top             =   1260
         Width           =   3675
      End
      Begin VB.TextBox txtCliente 
         Height          =   360
         Index           =   0
         Left            =   5160
         MaxLength       =   30
         TabIndex        =   61
         Top             =   1260
         Width           =   1035
      End
      Begin VB.TextBox txtFecha 
         Height          =   360
         Index           =   3
         Left            =   1920
         MaxLength       =   30
         TabIndex        =   58
         Top             =   1680
         Width           =   1515
      End
      Begin VB.CommandButton cmdpagoTasas 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   13320
         TabIndex        =   55
         Top             =   8040
         Width           =   1335
      End
      Begin VB.CommandButton cmdpagoTasas 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   14760
         TabIndex        =   54
         Top             =   8040
         Width           =   1335
      End
      Begin VB.TextBox txtFecha 
         Height          =   360
         Index           =   2
         Left            =   1920
         MaxLength       =   30
         TabIndex        =   52
         Top             =   1260
         Width           =   1515
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   3
         Left            =   240
         Picture         =   "frmMensajes.frx":04CC
         Top             =   8160
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   2
         Left            =   600
         Picture         =   "frmMensajes.frx":0616
         Top             =   8160
         Width           =   240
      End
      Begin VB.Image imgCli 
         Height          =   240
         Index           =   1
         Left            =   4920
         Top             =   1740
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   240
         Index           =   8
         Left            =   4200
         TabIndex        =   66
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1740
         Width           =   570
      End
      Begin VB.Image imgCli 
         Height          =   240
         Index           =   0
         Left            =   4920
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   240
         Index           =   7
         Left            =   4200
         TabIndex        =   63
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1320
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   6
         Left            =   3600
         TabIndex        =   60
         ToolTipText     =   "Fecha alta asociado"
         Top             =   960
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   240
         Index           =   5
         Left            =   720
         TabIndex        =   59
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1740
         Width           =   570
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   3
         Left            =   1560
         Picture         =   "frmMensajes.frx":0760
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1740
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   240
         Index           =   4
         Left            =   720
         TabIndex        =   57
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1320
         Width           =   600
      End
      Begin VB.Label Label7 
         Caption         =   "Pago tasas administrativas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   495
         Index           =   3
         Left            =   240
         TabIndex        =   56
         Top             =   360
         Width           =   5610
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha expediente"
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   3
         Left            =   240
         TabIndex        =   53
         ToolTipText     =   "Fecha alta asociado"
         Top             =   960
         Width           =   1755
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   2
         Left            =   1560
         Picture         =   "frmMensajes.frx":07EB
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1320
         Width           =   240
      End
   End
   Begin VB.Frame FrameCobros 
      Height          =   6720
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   13410
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   12000
         TabIndex        =   34
         Top             =   6000
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView5 
         Height          =   4905
         Left            =   225
         TabIndex        =   35
         Top             =   1005
         Width           =   13035
         _ExtentX        =   22992
         _ExtentY        =   8652
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label52 
         Caption         =   "Cobros de la factura "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   240
         TabIndex        =   36
         Top             =   390
         Width           =   10185
      End
   End
   Begin VB.Frame FramepagoFra 
      Height          =   3855
      Left            =   3360
      TabIndex        =   96
      Top             =   120
      Visible         =   0   'False
      Width           =   7335
      Begin VB.TextBox txtCaja 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   11
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   108
         Top             =   2280
         Width           =   1755
      End
      Begin VB.TextBox txtCaja 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   10
         Left            =   5520
         MaxLength       =   30
         TabIndex        =   98
         Top             =   2280
         Width           =   1515
      End
      Begin VB.CommandButton cmdCobroFactura 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   1
         Left            =   4200
         TabIndex        =   99
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox txtCaja 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   9
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   104
         Top             =   1680
         Width           =   1755
      End
      Begin VB.TextBox txtCaja 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   8
         Left            =   5520
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   103
         Top             =   1680
         Width           =   1515
      End
      Begin VB.TextBox txtFecha 
         Height          =   360
         Index           =   5
         Left            =   2640
         MaxLength       =   30
         TabIndex        =   97
         Text            =   "commor"
         Top             =   1080
         Width           =   2475
      End
      Begin VB.CommandButton cmdCobroFactura 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   0
         Left            =   5760
         TabIndex        =   100
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Usuario"
         Height          =   240
         Index           =   22
         Left            =   240
         TabIndex        =   109
         ToolTipText     =   "Fecha alta asociado"
         Top             =   2280
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Importe cobrado"
         Height          =   240
         Index           =   21
         Left            =   3480
         TabIndex        =   107
         ToolTipText     =   "Fecha alta asociado"
         Top             =   2400
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "Factura"
         Height          =   240
         Index           =   20
         Left            =   240
         TabIndex        =   106
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1680
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Importe pendiente"
         Height          =   240
         Index           =   19
         Left            =   3480
         TabIndex        =   105
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   5
         Left            =   2280
         Picture         =   "frmMensajes.frx":0876
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha / hora pago"
         Height          =   240
         Index           =   18
         Left            =   240
         TabIndex        =   102
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1080
         Width           =   1995
      End
      Begin VB.Label Label7 
         Caption         =   "Pago factura por caja"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   6
         Left            =   1800
         TabIndex        =   101
         Top             =   360
         Width           =   4095
      End
   End
   Begin VB.Label Label11 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   1800
      TabIndex        =   3
      Top             =   600
      Width           =   120
   End
   Begin VB.Line Line4 
      Index           =   2
      X1              =   2040
      X2              =   3960
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label10 
      Caption         =   "años de vida"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   2520
      TabIndex        =   2
      Top             =   840
      Width           =   1065
   End
   Begin VB.Label Label10 
      Caption         =   "Valor adquisición"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   2400
      TabIndex        =   1
      Top             =   480
      Width           =   1425
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   405
      Index           =   2
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "frmMensajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
     '1.- Facturar expediente
     '2.- Facturas periodicas
     '3.- pagos tasas administrativas
     '4.- Caja infreso/gastos
     '5.- Cierre caja
     '6.- Cobro por caja de factura pendiente
     '7.- Compra tasas
     
     
     '22- Ver empresas bloquedas
     '25- Informe de base de datos
     '26- Show processlist
     
     '27- Cobros de la factura
    
         
    
    
Public Parametros As String
    


Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1

'recepcion de talon/pagare
Public Importe As Currency
Public Codigo As String
Public Tipo As String
Public FecCobro As String
Public FecVenci As String
Public Banco As String
Public Referencia As String


Private PrimeraVez As Boolean

Dim I As Integer
Dim SQL As String
Dim Rs As Recordset
Dim ItmX As ListItem
Dim Errores As String
Dim NE As Integer
Dim Ok As Integer

Dim CampoOrden As String
Dim Orden As Boolean


Private Sub PasarUnaEmpresaBloqueada(ABLoquedas As Boolean, indice As Integer)
Dim Origen As Integer
Dim Destino As Integer
Dim IT
    If ABLoquedas Then
        Origen = 0
        Destino = 1
        NE = 2
    Else
        Origen = 1
        Destino = 0
        NE = 1 'icono
    End If
    
    SQL = ListView2(Origen).ListItems(indice).Key
    Set IT = ListView2(Destino).ListItems.Add(, SQL)
    IT.SmallIcon = NE
    IT.Text = ListView2(Origen).ListItems(indice).Text
    IT.SubItems(1) = ListView2(Origen).ListItems(indice).SubItems(1)

    'Borramos en origen
    ListView2(Origen).ListItems.Remove indice
End Sub

Private Sub chkCaja_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdBloqEmpre_Click(Index As Integer)
    If Index = 0 Then
        SQL = "DELETE FROM usuarios.usuarioempresasariconta WHERE codusu =" & Parametros
        Conn.Execute SQL
        SQL = ""
        For I = 1 To ListView2(1).ListItems.Count
            SQL = SQL & ", (" & Parametros & "," & Val(Mid(ListView2(1).ListItems(I).Key, 2)) & ")"
        Next I
        If SQL <> "" Then
            'Quitmos la primera coma
            SQL = Mid(SQL, 2)
            SQL = "INSERT INTO usuarios.usuarioempresasariconta(codusu,codempre) VALUES " & SQL
            If Not EjecutaSQL(SQL) Then MsgBox "Se han producido errores insertando datos", vbExclamation
        End If
    End If
    Unload Me
End Sub


Private Sub cmdCaja_Click(Index As Integer)
    CadenaDesdeOtroForm = ""
    If Index = 0 Then
        Msg = ""
        For I = 0 To 3   'los 4 primeros es apunte de caja
            Me.txtCaja(I).Text = Trim(Me.txtCaja(I).Text)
            If Me.txtCaja(I).Text = "" Then Msg = "N"
        Next
        If Msg <> "" Then
            MsgBox "Campos obligatorios", vbExclamation
            Exit Sub
        End If
        If ImporteFormateado(txtCaja(2).Text) < 0 Then
            MsgBox "Importe debe sera mayor que cero", vbExclamation
            Exit Sub
        End If
        If Parametros <> "" Then
            'MODIFICAR
            Msg = "UPDATE caja set importe= " & DBSet(txtCaja(2).Text, "N") & ", tipomovi=" & Abs(Me.chkCaja.Value)
            Msg = Msg & ", ampliacion =" & DBSet(txtCaja(1).Text, "T")
            Msg = Msg & " WHERE usuario = " & DBSet(txtCaja(3).Text, "T") & " AND feccaja=" & DBSet(txtCaja(0).Text, "FH")
        Else
            'INSERTAR
            Msg = "INSERT INTO caja(usuario,feccaja,tipomovi,importe,ampliacion) VALUES (" & DBSet(txtCaja(3).Text, "T")
            Msg = Msg & "," & DBSet(txtCaja(0).Text, "FH") & "," & Abs(Me.chkCaja.Value)
            Msg = Msg & "," & DBSet(txtCaja(2).Text, "N") & "," & DBSet(txtCaja(1).Text, "T") & ")"
            
        End If
        If Not Ejecuta(Msg) Then Exit Sub
        
        CadenaDesdeOtroForm = "OK"
    End If
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    CadenaDesdeOtroForm = ""
    Unload Me
End Sub

Private Sub CmdCancelBancoRem_Click()
    CadenaDesdeOtroForm = ""
    Unload Me
End Sub

Private Sub CmdCancelTalPag_Click()
    CadenaDesdeOtroForm = ""
    Unload Me
End Sub

Private Sub cmdCanIconos_Click()
    Reorganizar = False
    Unload Me
End Sub





Private Sub cmdCierreCaja_Click(Index As Integer)
Dim Hora As Date

    CadenaDesdeOtroForm = ""
    If Index = 1 Then
        SQL = ""
        If Me.txtCaja(6).Text = "" Then SQL = "- Indique importe de cierre"
        If Me.txtCaja(7).Text = "" Then SQL = "- Indique importe de cierre"
        If txtFecha(4).Text = "" Then SQL = SQL & vbCrLf & "- Indique fecha cierre"
        
        If SQL <> "" Then
            MsgBox SQL, vbExclamation
            Exit Sub
        End If
        If Not FechaFacturaOK(CDate(txtFecha(4).Text)) Then Exit Sub

        'A partir de la fecha, vamos a ver la hora de cierre de caja.
        'Primero. NO pueden haber movimientos posteriores al cierre
        SQL = DevuelveDesdeBD("max(feccaja)", "caja", "usuario", txtCaja(4).Text, "T")
        
        If SQL <> "" Then
            Hora = CDate(SQL)
            If CDate(Hora) > CDate(txtFecha(4).Text) Then
                MsgBox "Hay moviemientos posteriores en la caja. (" & Hora & ")", vbExclamation
                Exit Sub
            End If
    
        End If
            
            
        If ImporteFormateado(txtCaja(6).Text) > ImporteFormateado(txtCaja(5).Text) Then
            MsgBox "Importe cierre mayor que el importe actual en caja  ", vbExclamation
            Exit Sub
        End If
        
            
        SQL = "Va a cerrar la caja" & vbCrLf & vbCrLf & "Usuario: " & txtCaja(4).Text & vbCrLf
        SQL = SQL & "Fecha cierre: " & txtFecha(4).Text & vbCrLf & vbCrLf
        SQL = SQL & "Importe a entregar: " & txtCaja(6).Text & "€ " & vbCrLf & vbCrLf
        SQL = SQL & "Importe queda en caja: " & txtCaja(7).Text & "€ " & vbCrLf & vbCrLf
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        CadenaDesdeOtroForm = txtFecha(4).Text & "|" & txtCaja(6).Text & "|" & txtCaja(7).Text & "|"
        
    End If
    Unload Me
End Sub

Private Sub cmdCobroFactura_Click(Index As Integer)

    If Index = 1 Then
        SQL = ""
        If Me.txtCaja(10).Text = "" Then SQL = "- Indique importe de cobro"
        If txtFecha(5).Text = "" Then SQL = SQL & vbCrLf & "- Indique fecha cobro"
        
        If SQL <> "" Then
            MsgBox SQL, vbExclamation
            Exit Sub
        End If
        If Not FechaFacturaOK(CDate(txtFecha(5).Text)) Then Exit Sub

        'A partir de la fecha, vamos a ver la hora de cierre de caja.
        'Primero. NO pueden haber movimientos posteriores al cierre
        SQL = DevuelveDesdeBD("max(feccaja)", "caja", "usuario", txtCaja(11).Text, "T")
        
        If SQL <> "" Then
            'Hora = CDate(SQL)
            'If CDate(Hora) > CDate(txtFecha(5).Text) Then
            If CDate(SQL) > CDate(txtFecha(5).Text) Then
                MsgBox "Hay moviemientos posteriores en la caja. (" & SQL & ")", vbExclamation
                Exit Sub
            End If
    
        End If
            

        'DE momento NO admito cobros parciales
        If txtCaja(8).Text <> txtCaja(10).Text Then
            MsgBox "No aceptados cobors a cuenta", vbExclamation
            Exit Sub
        End If


        If Not HacerCobroFactura Then Exit Sub

    End If
    Unload Me

End Sub

Private Sub cmdCompraTasas_Click(Index As Integer)
    If Index = 1 Then
         If txtFecha(6).Text = "" Then Exit Sub
         If txtTasas(1).Text = "" Then Exit Sub
         If ImporteFormateado(txtTasas(1).Text) < 0 Then
            MsgBox "importe no puede ser negativo", vbExclamation
            Exit Sub
         End If
            
         CadenaDesdeOtroForm = txtFecha(6).Text & "|" & txtTasas(1).Text & "|"
         
    End If
    Unload Me
    
End Sub

Private Sub cmdFacturarExp_Click(Index As Integer)
    If Index = 1 Then
        If txtFecha(0).Text = "" Then Exit Sub
        If Not FechaFacturaOK(CDate(txtFecha(0).Text)) Then Exit Sub
        
        'Es un expediente. Veremos si hay una factura > o esta con fecha menor
        CadenaDesdeOtroForm = txtFecha(0).Text & "|"
        
    End If
    Unload Me
End Sub

Private Sub cmdFacturarPrevision_Click(Index As Integer)
    
    If Index = 1 Then
        If txtFecha(1).Text = "" Then Exit Sub
        
        If Not FechaFacturaOK(CDate(txtFecha(1).Text)) Then Exit Sub
        
        
        If MsgBox("¿Facturar?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        
        CadenaDesdeOtroForm = txtFecha(1).Text & "|" & Text1.Text
    
    End If
    Unload Me
End Sub

Private Sub cmdpagoTasas_Click(Index As Integer)
    If Index = 0 Then
        Importe = 0
        J = 0
        For I = 1 To ListView1.ListItems.Count
            If ListView1.ListItems(I).Checked Then
                Importe = Importe + ImporteFormateado(ListView1.ListItems(I).SubItems(8))
                J = J + 1
            End If
        Next
        If J = 0 Then
            MsgBox "No hay valores seleccionados", vbExclamation
            Exit Sub
        End If
                
        SQL = "Total lineas: " & J & vbCrLf
        SQL = SQL & "Importe:    " & Format(Importe, FormatoImporte)
        SQL = SQL & vbCrLf & vbCrLf & "¿Continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
                
        Conn.BeginTrans
        If GenerarTasas Then
            Conn.CommitTrans
            'CadenaDesdeOtroForm = "OK"  dentro de la funcion ya esta
        Else
            Conn.RollbackTrans
            CadenaDesdeOtroForm = ""
            Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub cmdRegresar_Click()
    Unload Me
End Sub

Private Sub cmdVerObservaciones_Click()
    Unload Me
End Sub


Private Sub SQL_Exepdientes(Valor As String)
       SQL = "select l.numexped,l.anoexped,fecexped,licencia,nomclien ,pagado ,nomconce,e.codclien,l.importe"
        SQL = SQL & " ,l.codconce ,l.tiporegi,l.numlinea"
        SQL = SQL & " from expedientes e,expedientes_lineas l,clientes"
        SQL = SQL & " Where pagado=" & Valor & " and e.tiporegi = L.tiporegi And e.numexped = L.numexped And e.anoexped = L.anoexped And e.CodClien = Clientes.CodClien"
        
        If Valor = "0" Then
            'Desde hasta
            If Me.txtCliente(0).Text <> "" Then SQL = SQL & " AND e.codclien >= " & txtCliente(0).Text
            If Me.txtCliente(1).Text <> "" Then SQL = SQL & " AND e.codclien <= " & txtCliente(1).Text
              
            If Me.txtFecha(2).Text <> "" Then SQL = SQL & " AND e.fecexped >= " & DBSet(txtFecha(2).Text, "F")
            If Me.txtFecha(3).Text <> "" Then SQL = SQL & " AND e.fecexped <= " & DBSet(txtFecha(3).Text, "F")
        End If
        
End Sub

Private Sub cmdVerdatos_Click()
    Screen.MousePointer = vbHourglass
    ListView1.ListItems.Clear
    Set miRsAux = New ADODB.Recordset
    I = 0
    For J = 1 To 2
        SQL = ""
        If J = 1 And Parametros <> "" Then SQL_Exepdientes Parametros
        If J = 2 Then SQL_Exepdientes 0
        
        
        If SQL <> "" Then
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                I = I + 1
                ListView1.ListItems.Add , , Format(miRsAux!numexped, "0000")
                ListView1.ListItems(I).SubItems(1) = Format(miRsAux!fecexped, "yy/mm/dd")
                ListView1.ListItems(I).SubItems(2) = Format(miRsAux!numlinea, "00")
                ListView1.ListItems(I).SubItems(3) = Format(miRsAux!CodClien, "0000")
                ListView1.ListItems(I).SubItems(4) = DBLet(miRsAux!licencia, "T")
                ListView1.ListItems(I).SubItems(5) = miRsAux!NomClien
                ListView1.ListItems(I).SubItems(6) = Format(miRsAux!codconce, "000")
                ListView1.ListItems(I).SubItems(7) = miRsAux!nomconce
                ListView1.ListItems(I).SubItems(8) = Format(miRsAux!Importe, FormatoImporte)
                ListView1.ListItems(I).Tag = miRsAux!anoexped
                If J = 1 Then ListView1.ListItems(I).Checked = True
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            DoEvents
        End If
    Next
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case Opcion
        Case 1
            txtFecha(0).Text = Format(Now, "dd/mm/yyyy")
        Case 2
            
            txtFecha(1).Text = Format(Parametros, "dd/mm/yyyy")
            Text1.Text = ""
        
        Case 3
            If Parametros <> "" Then cmdVerdatos_Click
        Case 4
            PonFoco txtCaja(1)
        
        Case 5
            'Cierre caja
            Me.txtFecha(4).Text = Format(Now, "dd/mm/yyyy hh:nn:ss")
            txtCaja(4).Text = RecuperaValor(Parametros, 1)
            txtCaja(5).Text = RecuperaValor(Parametros, 2)
        
        Case 6
            Me.txtFecha(5).Text = Format(Now, "dd/mm/yyyy hh:nn:ss")
            txtCaja(8).Text = RecuperaValor(Parametros, 5)   'pendeinte
            txtCaja(9).Text = RecuperaValor(Parametros, 2)   'factura
            txtCaja(10).Text = RecuperaValor(Parametros, 5)  'cobrado
            txtCaja(11).Text = RecuperaValor(Parametros, 1)  'usuar
            PonFoco txtCaja(10)
        
        Case 7
            Me.txtTasas(0).Text = RecuperaValor(Parametros, 1)   'factura
            Me.txtFecha(6).Text = Format(Now, "dd/mm/yyyy hh:nn:ss")
            Me.txtTasas(1).Text = ""  '
            
            
        Case 22
            cargaempresasbloquedas
            
        Case 25
            CargaInformeBBDD
        
        Case 26
            CargaShowProcessList
        
        Case 27
            CargaCobrosFactura
       End Select
    End If
    Screen.MousePointer = vbDefault
End Sub

'++
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And Opcion = 23 And KeyCode = vbKeyEscape Then Unload Me
End Sub
'++


Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub CargaIMG(indice As Integer)
    Me.imgCli(indice).Picture = frmppal.imgIcoForms.ListImages(1).Picture
End Sub


Private Sub Form_Load()
Dim W, H

    PrimeraVez = True
    Me.Icon = frmppal.Icon
    Me.FrameBloqueoEmpresas.Visible = False
    Me.FrameInformeBBDD.Visible = False
    Me.FrameShowProcess.Visible = False
    FrameCobros.Visible = False
    FramePedirFechaFactura.Visible = False
    FrameFraPrevision.Visible = False
    FramePagostasas.Visible = False
    FrameCaja.Visible = False
    FrameCompraTasas.Visible = False
    Select Case Opcion
    Case 1
        Caption = "Facturar"
        W = Me.FramePedirFechaFactura.Width
        H = Me.FramePedirFechaFactura.Height + 300
        FramePedirFechaFactura.Visible = True
        Me.cmdFacturarExp(0).Cancel = True
    Case 2
        Caption = "Facturar"
        W = Me.FrameFraPrevision.Width
        H = Me.FrameFraPrevision.Height + 300
        FrameFraPrevision.top = 0
        FrameFraPrevision.Left = 30
        FrameFraPrevision.Visible = True
        Me.cmdFacturarPrevision(0).Cancel = True
    
    Case 3
        
        Caption = "Pago tasas"
        W = Me.FramePagostasas.Width
        H = Me.FramePagostasas.Height + 300
        FramePagostasas.Visible = True
        FramePagostasas.Left = 0
        CargaIMG 0
        CargaIMG 1
    
    Case 4
         
        Caption = "Caja"
        W = Me.FrameCaja.Width
        H = Me.FrameCaja.Height + 300
        FrameCaja.Visible = True
        FrameCaja.Left = 0
        Me.txtCaja(3).Text = RecuperaValor(Parametros, 1)
        Msg = Trim(RecuperaValor(Parametros, 2)) 'La fecha
        If Msg <> "" Then
            'Esta modifiacando
            txtCaja(0).Text = Msg
            txtCaja(1).Text = RecuperaValor(Parametros, 3)
            txtCaja(2).Text = RecuperaValor(Parametros, 4)
            Me.chkCaja.Value = IIf(RecuperaValor(Parametros, 5) = "1", 1, 0)
        Else
            txtCaja(0).Text = Format(Now, "dd/mm/yyyy hh:nn")
            Parametros = ""
        End If
        
        
        
    Case 5
        'Cierre caja
        FrameCierreCaja.Left = 0
        FrameCierreCaja.top = 0
        Me.FrameCierreCaja.Visible = True
        Caption = "CAJA"
        W = Me.FrameCierreCaja.Width
        H = Me.FrameCierreCaja.Height + 300
        cmdCierreCaja(0).Cancel = True
        
    Case 6
        
         FramepagoFra.Left = 0
        FramepagoFra.top = 0
        Me.FramepagoFra.Visible = True
        Caption = "CAJA"
        W = Me.FramepagoFra.Width
        H = Me.FramepagoFra.Height + 300
        cmdCobroFactura(0).Cancel = True
        
        
    Case 7
        'Cmpra tasas
        FrameCompraTasas.Left = 0
        FrameCompraTasas.top = 0
        Me.FrameCompraTasas.Visible = True
        Caption = "TASAS"
        W = Me.FrameCompraTasas.Width
        H = Me.FrameCompraTasas.Height + 300
        cmdCompraTasas(0).Cancel = True
        txtCaja(1).Text = RecuperaValor(Parametros, 3)
        txtCaja(2).Text = RecuperaValor(Parametros, 4)
        
        
    Case 22
        Me.FrameBloqueoEmpresas.Visible = True
        Caption = "Bloqueo empresas"
        W = Me.FrameBloqueoEmpresas.Width
        H = Me.FrameBloqueoEmpresas.Height + 300
        'Como cuando venga por esta opcion, viene llamado desde el manteusu
        Me.ListView2(0).SmallIcons = frmMantenusu.ImageList1
        Me.ListView2(1).SmallIcons = frmMantenusu.ImageList1
        Me.cmdBloqEmpre(1).Cancel = True
        

        
    Case 25 ' informe de base de datos
        Me.Caption = "Información de Base de Datos"
        Me.FrameInformeBBDD.Visible = True
        W = Me.FrameInformeBBDD.Width
        H = Me.FrameInformeBBDD.Height + 300
                
    Case 26 ' show process list
        Me.Caption = "Información de Procesos del Sistema"
        Me.FrameShowProcess.Visible = True
        W = Me.FrameShowProcess.Width
        H = Me.FrameShowProcess.Height + 300
        
        Label53.Caption = Label53.Caption & " Ariconta" & vEmpresa.codempre & " (" & vEmpresa.nomempre & ")"

        
    Case 27
        '27- Cobros de la factura
        Me.Caption = "Facturas de Cliente"
        Label52.Caption = "Cobros de la Factura " & RecuperaValor(Parametros, 1) & "-" & Format(RecuperaValor(Parametros, 2), "0000000") & " de fecha " & RecuperaValor(Parametros, 3)
        Me.FrameCobros.Visible = True
        W = Me.FrameCobros.Width
        H = Me.FrameCobros.Height + 300
 
    
    End Select
    
    Me.Width = W + 120
    Me.Height = H + 120
End Sub





Private Sub VerEmresasProhibidas(ByRef VarProhibidas As String)

On Error GoTo EVerEmresasProhibidas
    VarProhibidas = "|"
    SQL = "Select codempre from Usuarios.usuarioempresasariconta WHERE codusu = " & (vUsu.Codigo Mod 1000)
    SQL = SQL & " order by codempre"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs.EOF
          VarProhibidas = VarProhibidas & Rs!codempre & "|"
          Rs.MoveNext
    Wend
    Rs.Close
    Exit Sub
EVerEmresasProhibidas:
    MuestraError Err.Number, Err.Description & vbCrLf & " Consulte soporte técnico"
    Set Rs = Nothing
End Sub







Private Sub frmF_Selec(vFecha As Date)
    SQL = vFecha
End Sub

Private Sub imgCheck_Click(Index As Integer)
    If Index < 2 Then
        For NE = 1 To TreeView1.Nodes.Count
            TreeView1.Nodes(NE).Checked = Index = 1
        Next
    Else
        For I = 1 To ListView1.ListItems.Count
            ListView1.ListItems(I).Checked = Index = 2
        Next
    End If
End Sub






Private Sub LeerCadenaFicheroTexto()
On Error GoTo ELeerCadenaFicheroTexto
    'Son dos lineas. La primaera indica k campo y la segunda el valor
    Line Input #I, SQL
    Line Input #I, SQL
    Exit Sub
ELeerCadenaFicheroTexto:
    SQL = ""
    Err.Clear
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Function ValorSQL(ByRef C As String) As String
    If C = "" Then
        ValorSQL = "NULL"
    Else
        ValorSQL = "'" & C & "'"
    End If
End Function
Private Function EjecutaSQL2(SQL As String) As Boolean
    EjecutaSQL2 = False
    On Error Resume Next
    Conn.Execute SQL
    If Err.Number <> 0 Then
        AnyadeErrores "SQL: " & SQL, Err.Description
        Err.Clear
    Else
        EjecutaSQL2 = True
    End If
End Function


Private Sub AnyadeErrores(L1 As String, L2 As String)
    NE = NE + 1
    Errores = Errores & "-----------------------------" & vbCrLf
    Errores = Errores & L1 & vbCrLf
    Errores = Errores & L2 & vbCrLf


End Sub

Private Sub ImprimeFichero()
Dim NF As Integer
    On Error GoTo EImprimeFichero
    NF = FreeFile
    Open App.Path & "\errimpdat.txt" For Output As #NF
    Print #NF, Errores
    Close (NF)
    Shell "notepad.exe " & App.Path & "\errimpdat.txt", vbMaximizedFocus
    Exit Sub
EImprimeFichero:
    MsgBox Err.Description & vbCrLf, vbCritical
    Err.Clear
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
                PonerFocoBtn Me.cmdpagoTasas(0)
            End If
            CadenaDesdeOtroForm = ""
        End If

End Sub

Private Sub imgppal_Click(Index As Integer)
    Set frmF = New frmCal
    frmF.Fecha = Now
    SQL = ""
    If Me.txtFecha(Index).Text <> "" Then frmF.Fecha = txtFecha(Index).Text
    frmF.Show vbModal
    If SQL <> "" Then
        txtFecha(Index).Text = SQL
        If Index = 4 Or Index = 5 Or Index = 6 Then txtFecha(Index).Text = txtFecha(Index).Text & " " & Format(Now, "hh:mm:ss")
    End If
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If ListView1.Sorted = False Then
        ListView1.SortKey = 1
        ListView1.SortOrder = lvwAscending
        ListView1.Sorted = True
    Else
        If Me.ListView1.SortKey = ColumnHeader.Index - 1 Then
            If ListView1.SortOrder = lvwAscending Then
                ListView1.SortOrder = lvwDescending
            Else
                ListView1.SortOrder = lvwAscending
            End If
        Else
            ListView1.SortKey = ColumnHeader.Index - 1
            ListView1.SortOrder = lvwAscending
        End If
    End If
End Sub

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim N As Node
    'Si es padre
    If Node.Parent Is Nothing Then
        If Node.Children > 0 Then
            Set N = Node.Child
            Do
                N.Checked = Node.Checked
                Set N = N.Next
            Loop Until N Is Nothing
        End If
    End If
End Sub

'-----------------------------------------------------------------------------------
'
'




Private Sub EncabezadoPieFact(Pie As Boolean, ByVal Text As String, REG As Integer)
    If Pie Then
        Text = "[/" & Text & "]" & REG
    Else
        Text = "[" & Text & "]"
    End If
    Print #NE, Text
End Sub


Private Sub InsertaEnTmpCta()
On Error Resume Next
    
    Conn.Execute "INSERT INTO tmpcierre1 (codusu, cta) VALUES (" & vUsu.Codigo & ",'" & Rs.Fields(0) & "')"
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub EjecutarSQL()
    On Error Resume Next
    
    Conn.Execute SQL
    If Err.Number <> 0 Then
        If Conn.Errors(0).Number = 1062 Then
            Err.Clear
        Else
            'MuestraError Err.Number, Err.Description
        End If
        Err.Clear
    End If
End Sub





Private Sub cargaempresasbloquedas()
Dim IT As ListItem
    On Error GoTo Ecargaempresasbloquedas
    Set Rs = New ADODB.Recordset
    SQL = "select empresasariconta.codempre,nomempre,nomresum,usuarioempresasariconta.codempre bloqueada from usuarios.empresasariconta left join usuarios.usuarioempresasariconta on "
    SQL = SQL & " empresasariconta.codempre = usuarioempresasariconta.codempre And (usuarioempresasariconta.codusu = " & Parametros & " Or codusu Is Null)"
    '[Monica] solo ariconta
    SQL = SQL & " WHERE conta like 'ariconta%' "
    SQL = SQL & " ORDER BY empresasariconta.codempre"
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs.EOF
        Errores = Format(Rs!codempre, "00000")
        SQL = "C" & Errores
        
        If IsNull(Rs!bloqueada) Then
            'Va al list de la derecha
            Set IT = ListView2(0).ListItems.Add(, SQL)
            IT.SmallIcon = 1
        Else
            Set IT = ListView2(1).ListItems.Add(, SQL)
            IT.SmallIcon = 2
        End If
        IT.Text = Errores
        IT.SubItems(1) = Rs!nomempre
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    Errores = ""
    Exit Sub
Ecargaempresasbloquedas:
    MuestraError Err.Number, Err.Description
    Me.cmdBloqEmpre(0).Enabled = False
    Errores = ""
    Set Rs = Nothing
End Sub












Private Sub CargaInformeBBDD()
Dim IT As ListItem
Dim TotalArray  As Long
    On Error GoTo ECargaInformeBBDD
    
    Set Rs = New ADODB.Recordset
    
    SQL = "select * from tmpinfbbdd where codusu = " & vUsu.Codigo & " order by posicion "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    ListView3.ColumnHeaders.Clear
    ListView3.ColumnHeaders.Add , , "CONCEPTO", 3500.0631
    ListView3.ColumnHeaders.Add , , "count ACTUAL", 2250.2522, 1
    ListView3.ColumnHeaders.Add , , "porcen ACTUAL", 1000.2522, 1
    ListView3.ColumnHeaders.Add , , "count siguiente", 2250.2522, 1
    ListView3.ColumnHeaders.Add , , "porcen siguiente", 1000.2522, 1
    
    
    
    
    TotalArray = 0
    While Not Rs.EOF
        Set IT = ListView3.ListItems.Add
        
        IT.Text = UCase(DBLet(Rs!Concepto, "T"))
        
        If DBLet(Rs!posicion, "N") > 2 Then
            IT.SubItems(1) = Format(DBLet(Rs!nactual, "N"), "###,###,###,##0")
            IT.SubItems(2) = Format(DBLet(Rs!Poractual, "N"), "##0.00") & "%"
            IT.SubItems(3) = Format(DBLet(Rs!nsiguiente, "N"), "###,###,###,##0")
            IT.SubItems(4) = Format(DBLet(Rs!Porsiguiente, "N"), "##0.00") & "%"
        Else
            IT.SubItems(1) = Format(DBLet(Rs!nactual, "N"), "###,###,###,##0")
            IT.SubItems(3) = Format(DBLet(Rs!nsiguiente, "N"), "###,###,###,##0")
        End If
        
        Rs.MoveNext
    Wend
    
    Rs.Close
    Exit Sub
    
ECargaInformeBBDD:
    MuestraError Err.Number, Err.Description
    Errores = ""
    Set Rs = Nothing
End Sub


Private Sub CargaShowProcessList()
Dim IT As ListItem
Dim TotalArray  As Long
Dim SERVER As String
Dim EquipoConBD As Boolean
Dim Cad As String
Dim Equipo As String

    On Error GoTo ECargaShowProcessList
    
    Set Rs = New ADODB.Recordset
    
    ListView4.ColumnHeaders.Clear
    
    ListView4.ColumnHeaders.Add , , "ID", 1500.0631
    ListView4.ColumnHeaders.Add , , "User", 2250.2522, 1
    ListView4.ColumnHeaders.Add , , "Host", 3000.2522, 1
    ListView4.ColumnHeaders.Add , , "Tiempo espera", 3050.2522, 1
    
    
    Set Rs = New ADODB.Recordset
    
    SERVER = Mid(Conn.ConnectionString, InStr(LCase(Conn.ConnectionString), "server=") + 7)
    SERVER = Mid(SERVER, 1, InStr(1, SERVER, ";"))
    
    EquipoConBD = (UCase(vUsu.PC) = UCase(SERVER)) Or (LCase(SERVER) = "localhost")
    
    Cad = "show full processlist"
    Rs.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not Rs.EOF
        If Not IsNull(Rs.Fields(3)) Then
            If InStr(1, Rs.Fields(3), "arigestion") <> 0 Then
                If UCase(Rs.Fields(3)) = UCase(vUsu.CadenaConexion) Then
                    Equipo = Rs.Fields(2)
                    'Primero quitamos los dos puntos del puerto
                    NumRegElim = InStr(1, Equipo, ":")
                    If NumRegElim > 0 Then Equipo = Mid(Equipo, 1, NumRegElim - 1)
                    
                    'El punto del dominio
                    NumRegElim = InStr(1, Equipo, ".")
                    If NumRegElim > 0 Then Equipo = Mid(Equipo, 1, NumRegElim - 1)
                    
                    Equipo = UCase(Equipo)
                    
                    
                    Set IT = ListView4.ListItems.Add
                    
                    IT.Text = Rs.Fields(0)
                    IT.SubItems(1) = Rs.Fields(1)
                    IT.SubItems(2) = Equipo
                    
                    'tiempo de espera
                    Dim FechaAnt As Date
                    FechaAnt = DateAdd("s", Rs.Fields(5), Now)
                    IT.SubItems(3) = Format((Now - FechaAnt), "hh:mm:ss")
                End If
            End If
        End If
        
        'Siguiente
        Rs.MoveNext
    Wend
    NumRegElim = 0
    Rs.Close
    Set Rs = Nothing
    
    Exit Sub
    
ECargaShowProcessList:
    MuestraError Err.Number, Err.Description
    Errores = ""
    Set Rs = Nothing
End Sub





Private Sub CargaCobrosFactura()
Dim IT As ListItem
Dim TotalArray  As Long
Dim SERVER As String
Dim EquipoConBD As Boolean
Dim Cad As String
Dim Equipo As String

    On Error GoTo ECargaCobrosFactura
    
    Set Rs = New ADODB.Recordset
    
    ListView5.ColumnHeaders.Clear
    
    ListView5.ColumnHeaders.Add , , "Ord.", 790
    ListView5.ColumnHeaders.Add , , "Forma de Pago", 3100.2522
    ListView5.ColumnHeaders.Add , , "Fecha Vto", 1450.2522
    ListView5.ColumnHeaders.Add , , "Importe Vto", 1500.2522, 1
    ListView5.ColumnHeaders.Add , , "Gastos", 1550.2522, 1
    ListView5.ColumnHeaders.Add , , "F.Ult.Cobro", 1450.2522
    ListView5.ColumnHeaders.Add , , "Imp.Pagado", 1550.2522, 1
    ListView5.ColumnHeaders.Add , , "Pendiente", 1550.2522, 1
    
    Set Rs = New ADODB.Recordset
    
    ListView5.SmallIcons = frmppal.ImgListComun
    
    Cad = "select numorden, formapago.nomforpa, fecvenci, impvenci, gastos, fecultco, impcobro, (coalesce(impvenci,0) + coalesce(gastos,0) - coalesce(impcobro,0)) pendiente, cobros.ctabanc1  "
    Cad = Cad & " from (ariconta" & vParam.Numconta & ".cobros left join ariconta" & vParam.Numconta & ".formapago on cobros.codforpa = formapago.codforpa) "
    Cad = Cad & " where cobros.numserie = " & DBSet(RecuperaValor(Parametros, 1), "T")
    Cad = Cad & " and cobros.numfactu = " & DBSet(RecuperaValor(Parametros, 2), "N")
    Cad = Cad & " and cobros.fecfactu = " & DBSet(RecuperaValor(Parametros, 3), "F")
    Cad = Cad & " order by numorden "
    
    Rs.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not Rs.EOF
                    
        Set IT = ListView5.ListItems.Add
        
        If CobroContabilizado(RecuperaValor(Parametros, 1), RecuperaValor(Parametros, 2), RecuperaValor(Parametros, 3), DBLet(Rs.Fields(0))) Then IT.SmallIcon = 18
        
        IT.Text = DBLet(Rs.Fields(0))
        IT.SubItems(1) = DBLet(Rs.Fields(1))
        IT.SubItems(2) = DBLet(Rs.Fields(2))
        IT.SubItems(3) = Format(DBLet(Rs.Fields(3)), "###,###,##0.00")
        
        
        IT.SubItems(4) = " " & Format(DBLet(Rs.Fields(4)), "###,###,##0.00")
        IT.SubItems(5) = " " & DBLet(Rs.Fields(5))
        IT.SubItems(6) = " " & Format(DBLet(Rs.Fields(6)), "###,###,##0.00")
        IT.SubItems(7) = Format(DBLet(Rs.Fields(7)), "###,###,##0.00")
        
        'Siguiente
        Rs.MoveNext
    Wend
    NumRegElim = 0
    Rs.Close
    
    
    'Pudieron dar entregas a cuenta
    'Si es factura de expediente
    '
    If RecuperaValor(Parametros, 1) = "FEX" Then
        Cad = "Select numexped,fecexped from factcli "
        Cad = Cad & " where factcli.numserie = " & DBSet(RecuperaValor(Parametros, 1), "T")
        Cad = Cad & " and factcli.numfactu = " & DBSet(RecuperaValor(Parametros, 2), "N")
        Cad = Cad & " and factcli.fecfactu = " & DBSet(RecuperaValor(Parametros, 3), "F")
        Rs.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        Cad = ""
        If Not Rs.EOF Then
            If Not IsNull(Rs!numexped) Then Cad = "tiporegi = 0 AND numdocum = " & Rs!numexped & " AND anoexped =" & Year(Rs!fecexped)
        
        End If
        Rs.Close
        
        If Cad <> "" Then
            Cad = "Select * from caja where " & Cad
            Rs.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
            Cad = ""
            While Not Rs.EOF
                            
                Set IT = ListView5.ListItems.Add
                
                'numserie   numdocum anoexped importe ampliacion
                IT.Text = DBLet(Rs!numserie)
                IT.Bold = True
                IT.ForeColor = vbRed
                IT.SubItems(1) = "Pago a cuenta caja:" & Rs!Usuario
                IT.SubItems(2) = Format(Rs!Feccaja, "dd/mm/yyyy")
                IT.SubItems(3) = Format(DBLet(Rs!Importe), "###,###,##0.00")
                
                
                IT.SubItems(4) = " "
                IT.SubItems(5) = " "
                IT.SubItems(6) = " " & Format(DBLet(Rs!Importe), "###,###,##0.00")
                IT.SubItems(7) = Format(0, "###,###,##0.00")
                
                'Siguiente
                Rs.MoveNext
            Wend
            Rs.Close
        End If
    End If
    
    
    Set Rs = Nothing
    
    Exit Sub
    
ECargaCobrosFactura:
    MuestraError Err.Number, Err.Description
    Errores = ""
    Set Rs = Nothing
End Sub


Private Function CobroContabilizado(Serie As String, FACTURA As String, Fecha As String, Orden As String) As Boolean
Dim SQL As String

    SQL = "select * from ariconta" & vParam.Numconta & ".hlinapu where numserie = " & DBSet(Serie, "T") & " and numfaccl = " & DBSet(FACTURA, "N") & " and fecfactu = " & DBSet(Fecha, "F") & " and numorden = " & DBSet(Orden, "N")
    CobroContabilizado = (TotalRegistrosConsulta(SQL) <> 0)

End Function



Private Sub txtCaja_GotFocus(Index As Integer)
    ConseguirFoco txtCaja(Index), 4
End Sub

Private Sub txtCaja_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        KeyAscii = 0
        imgCli_Click Index
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub txtCaja_LostFocus(Index As Integer)

    'If Index = 2 Then        If Not PonerFormatoFechaHora(txtCaja(Index)) Then txtCaja(Index).Text = ""
    
    If Index = 2 Or Index = 7 Or Index = 10 Then
        Importe = 0
        If Not PonerFormatoDecimal(txtCaja(Index), 1) Then txtCaja(Index).Text = ""
           
         If Index = 7 Then
            
            Importe = ImporteFormateado(txtCaja(5).Text) - ImporteFormateado(txtCaja(7).Text)
            txtCaja(6).Text = Format(Importe, FormatoImporte)
        End If
    End If
End Sub

Private Sub txtCliente_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
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
    SQL = ""
    If PonerFormatoEntero(txtCliente(Index)) Then

        SQL = DevuelveDesdeBD("nomclien", "clientes", "codclien", txtCliente(Index).Text, "N")
        If SQL = "" Then SQL = "No existe el cliente"
    End If
    Me.txtClienteDes(Index).Text = SQL
End Sub

Private Sub txtFecha_GotFocus(Index As Integer)
    ConseguirFoco txtFecha(Index), 4
End Sub

Private Sub txtFecha_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtfecha_LostFocus(Index As Integer)
Dim B As Boolean
    If txtFecha(Index).Text = "" Then Exit Sub
    B = True
    If Index = 4 Or Index = 5 Or Index = 6 Then
        If Not EsFechaHoraOK(txtFecha(Index)) Then B = False
    Else
        If Not EsFechaOK(txtFecha(Index)) Then B = False
    End If
    If Not B Then
        txtFecha(Index).Text = ""
        MsgBox "Fecha incorrecta", vbExclamation
        PonFoco txtFecha(Index)
        Exit Sub
   
    End If

End Sub


Private Sub txtTasas_GotFocus(Index As Integer)
    ConseguirFoco txtTasas(Index), 4
End Sub

Private Sub txtTasas_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtTasas_LostFocus(Index As Integer)
    If txtTasas(Index).Text = "" Then Exit Sub
   
    If Index = 1 Then
        If Not PonerFormatoEntero(txtTasas(Index)) Then txtTasas(Index).Text = ""
    End If
End Sub


'Cargar tasas administrativas
Private Function GenerarTasas() As Boolean
    
    GenerarTasas = False
    On Error GoTo eGenerarTasas
      
    If Parametros <> "" Then
        SQL = "UPDATE expedientes_lineas set  codsitua=0, pagado =0 where "
        SQL = SQL & " pagado =" & Parametros
        Conn.Execute SQL
    
        Codigo = Parametros
        
        SQL = "UPDATE gestadministrativa SET llevados = " & J & ", importe=" & DBSet(Importe, "N")
        SQL = SQL & " WHERE id =" & Codigo
        CadenaDesdeOtroForm = "OK"
    Else
        Codigo = DevuelveDesdeBD("max(id)", "gestadministrativa", "1", "1")
        Codigo = Val(Codigo) + 1
        CadenaDesdeOtroForm = Codigo
        SQL = "INSERT INTO gestadministrativa(id,usuario,fechacreacion,llevados,importe) VALUES (" & Codigo
        SQL = SQL & "," & DBSet(vUsu.Login, "T") & "," & DBSet(Now, "FH") & "," & J & "," & DBSet(Importe, "N") & ")"
    End If
    Conn.Execute SQL
    
    J = 0
    Importe = 0
    For I = 1 To ListView1.ListItems.Count
        If Me.ListView1.ListItems(I).Checked Then
            SQL = "UPDATE expedientes_lineas set  codsitua=1, pagado =" & Codigo & " where "
            SQL = SQL & " tiporegi='0' AND numexped =" & ListView1.ListItems(I).Text
            SQL = SQL & " AND anoexped=" & ListView1.ListItems(I).Tag
            SQL = SQL & " AND numlinea=" & ListView1.ListItems(I).SubItems(2)
            Conn.Execute SQL
        End If
    Next
    
    GenerarTasas = True
    Exit Function
eGenerarTasas:
 MuestraError Err.Number
CadenaDesdeOtroForm = ""
End Function




'------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------
'
'       CObros sobre facturas en ARIMONEY
'
'
'------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------
Private Function HacerCobroFactura() As Boolean

    
    'Transaccion
    Screen.MousePointer = vbHourglass
    HacerCobroFactura = False
    Conn.BeginTrans
    If RealizarCobro Then
        Conn.CommitTrans
        HacerCobroFactura = True
    Else
        Conn.RollbackTrans
    End If
    Screen.MousePointer = vbDefault
End Function

Private Function RealizarCobro() As Boolean
Dim Cobrado As Currency

    On Error GoTo eRealizarCobro
    RealizarCobro = False
    Set miRsAux = New ADODB.Recordset
    
    Errores = "Separando datos factura"
    CampoOrden = RecuperaValor(Parametros, 3)
    CampoOrden = Replace(CampoOrden, ":", "|")
    
    SQL = " numserie = " & DBSet(RecuperaValor(CampoOrden, 1), "T") & " AND numfactu =" & RecuperaValor(CampoOrden, 2)
    SQL = SQL & " AND numorden=" & RecuperaValor(CampoOrden, 3) & " AND fecfactu =" & DBSet(RecuperaValor(Parametros, 4), "F")
    Codigo = "SELECT impvenci,gastos,fecultco,impcobro,numserie,numfactu,fecfactu,numorden FROM ariconta" & vParam.Numconta & ".cobros WHERE " & SQL
    miRsAux.Open Codigo, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'NO PUEDE SER EOF
    Cobrado = DBLet(miRsAux!impcobro, "N")
    Importe = miRsAux!ImpVenci + DBLet(miRsAux!gastos, "N")
    
    
    If Cobrado + ImporteFormateado(Me.txtCaja(10).Text) > Importe Then Err.Raise 513, "Cobrado mas de lo que tiene pendiente"
    
    Importe = ImporteFormateado(txtCaja(10).Text)  'lo que vamos a cobrar ahora
    Cobrado = Cobrado + ImporteFormateado(Me.txtCaja(10).Text)   'lo que habia mas lo de ahora
    
    Codigo = "UPDATE ariconta" & vParam.Numconta & ".cobros SET "
    Codigo = Codigo & " fecultco = " & DBSet(txtFecha(5), "F") & ", impcobro = " & DBSet(Cobrado, "N")
    Codigo = Codigo & " WHERE " & SQL
    Conn.Execute Codigo
    miRsAux.Close
    
    Errores = "Obteniendo tipo registro"
    SQL = "Select * from contadores where serfactur =" & DBSet(RecuperaValor(CampoOrden, 1), "T")
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    
    'Ya hemos actualizado en cobros. Ahora metemos en caja
    Codigo = "INSERT INTO caja(usuario,feccaja,tipomovi,tiporegi,numserie,numdocum,anoexped,numlinea,importe,ampliacion) VALUES ("
    Codigo = Codigo & DBSet(txtCaja(11).Text, "T") & "," & DBSet(txtFecha(5).Text, "FH") & ",0," & DBSet(miRsAux!tiporegi, "N")
    SQL = RecuperaValor(Parametros, 4)
    SQL = Year(CDate(SQL))
    Codigo = Codigo & "," & DBSet(miRsAux!serfactur, "T") & "," & DBSet(RecuperaValor(CampoOrden, 2), "T") & "," & DBSet(SQL, "N")
    SQL = RecuperaValor(Parametros, 6)
    Codigo = Codigo & "," & DBSet(RecuperaValor(CampoOrden, 3), "N") & "," & DBSet(Importe, "T") & "," & DBSet(SQL, "T") & ")"
    Conn.Execute Codigo
    
    miRsAux.Close
    RealizarCobro = True
    
eRealizarCobro:
    If Err.Number <> 0 Then MuestraError Err.Number, Errores & vbCrLf & Err.Description
    Set miRsAux = Nothing
End Function

