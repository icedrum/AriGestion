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
   Begin VB.Frame FramepagoFra 
      Height          =   4215
      Left            =   3360
      TabIndex        =   98
      Top             =   120
      Visible         =   0   'False
      Width           =   8055
      Begin VB.CommandButton cmdTarjetaCredito 
         Height          =   495
         Left            =   240
         Picture         =   "frmMensajes.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   180
         ToolTipText     =   "TARJETA DE CREDITO"
         Top             =   3360
         Width           =   615
      End
      Begin VB.CheckBox chkQuitarGastos 
         Caption         =   "Quitar gastos"
         Height          =   240
         Left            =   1800
         TabIndex        =   173
         Top             =   2880
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtCaja 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   12
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   171
         Top             =   2340
         Width           =   1515
      End
      Begin VB.TextBox txtCaja 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   11
         Left            =   6120
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   110
         Top             =   1080
         Width           =   1755
      End
      Begin VB.TextBox txtCaja 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   10
         Left            =   6360
         MaxLength       =   30
         TabIndex        =   100
         Top             =   2340
         Width           =   1515
      End
      Begin VB.CommandButton cmdCobroFactura 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   1
         Left            =   4920
         TabIndex        =   101
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox txtCaja 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   9
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   106
         Top             =   1680
         Width           =   1755
      End
      Begin VB.TextBox txtCaja 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   8
         Left            =   6360
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   105
         Top             =   1680
         Width           =   1515
      End
      Begin VB.TextBox txtFecha 
         Height          =   360
         Index           =   5
         Left            =   2640
         MaxLength       =   30
         TabIndex        =   99
         Text            =   "commor"
         Top             =   1080
         Width           =   2475
      End
      Begin VB.CommandButton cmdCobroFactura 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   0
         Left            =   6480
         TabIndex        =   102
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Lleva gastos"
         Height          =   240
         Index           =   38
         Left            =   240
         TabIndex        =   172
         ToolTipText     =   "Fecha alta asociado"
         Top             =   2400
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Usuario"
         Height          =   240
         Index           =   22
         Left            =   5280
         TabIndex        =   111
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Importe cobrado"
         Height          =   240
         Index           =   21
         Left            =   4080
         TabIndex        =   109
         ToolTipText     =   "Fecha alta asociado"
         Top             =   2400
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "Factura"
         Height          =   240
         Index           =   20
         Left            =   240
         TabIndex        =   108
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1680
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Importe pendiente"
         Height          =   240
         Index           =   19
         Left            =   4080
         TabIndex        =   107
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   5
         Left            =   2280
         Picture         =   "frmMensajes.frx":0316
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha / hora pago"
         Height          =   240
         Index           =   18
         Left            =   240
         TabIndex        =   104
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
         TabIndex        =   103
         Top             =   360
         Width           =   4095
      End
   End
   Begin VB.Frame FrameCaja 
      Height          =   4935
      Left            =   3480
      TabIndex        =   69
      Top             =   1800
      Visible         =   0   'False
      Width           =   8055
      Begin VB.ComboBox cboConceptosCaja 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   71
         Top             =   2280
         Width           =   6015
      End
      Begin VB.OptionButton optConceCaja 
         Caption         =   "Introduccion manual cuenta"
         Height          =   375
         Index           =   1
         Left            =   4200
         TabIndex        =   170
         Top             =   1740
         Width           =   3375
      End
      Begin VB.OptionButton optConceCaja 
         Caption         =   "Conceptos caja"
         Height          =   375
         Index           =   0
         Left            =   1560
         TabIndex        =   169
         Top             =   1740
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.TextBox txtCodmacta 
         Height          =   360
         Index           =   0
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   72
         Top             =   2280
         Width           =   1875
      End
      Begin VB.TextBox txtCodmactaDes 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   0
         Left            =   3600
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   140
         Top             =   2280
         Width           =   4035
      End
      Begin VB.TextBox txtCaja 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   3
         Left            =   5160
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   82
         Top             =   1080
         Width           =   2475
      End
      Begin VB.CheckBox chkCaja 
         Caption         =   "Salida"
         Height          =   240
         Left            =   3360
         TabIndex        =   75
         Top             =   3840
         Width           =   1095
      End
      Begin VB.TextBox txtCaja 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   2
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   74
         Top             =   3720
         Width           =   1395
      End
      Begin VB.TextBox txtCaja 
         Height          =   360
         Index           =   1
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   73
         Top             =   3000
         Width           =   6075
      End
      Begin VB.TextBox txtCaja 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   0
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   70
         Top             =   1080
         Width           =   2475
      End
      Begin VB.CommandButton cmdCaja 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   5040
         TabIndex        =   76
         Top             =   4320
         Width           =   1335
      End
      Begin VB.CommandButton cmdCaja 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   6480
         TabIndex        =   77
         Top             =   4320
         Width           =   1335
      End
      Begin VB.Image imgFecCaja 
         Height          =   240
         Index           =   0
         Left            =   1200
         Picture         =   "frmMensajes.frx":03A1
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image imgCodmacta 
         Height          =   240
         Index           =   0
         Left            =   1320
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Usuario"
         Height          =   240
         Index           =   12
         Left            =   4320
         TabIndex        =   83
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1080
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Importe"
         Height          =   240
         Index           =   11
         Left            =   240
         TabIndex        =   81
         ToolTipText     =   "Fecha alta asociado"
         Top             =   3720
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Concepto"
         Height          =   240
         Index           =   10
         Left            =   240
         TabIndex        =   80
         ToolTipText     =   "Fecha alta asociado"
         Top             =   3000
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   240
         Index           =   9
         Left            =   240
         TabIndex        =   79
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
         TabIndex        =   78
         Top             =   240
         Width           =   3150
      End
      Begin VB.Label Label1 
         Caption         =   "Cta conta."
         Height          =   240
         Index           =   31
         Left            =   240
         TabIndex        =   141
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1800
         Width           =   1395
      End
   End
   Begin VB.Frame FrameCompensacion 
      Height          =   5175
      Left            =   2280
      TabIndex        =   174
      Top             =   0
      Visible         =   0   'False
      Width           =   8895
      Begin VB.CommandButton cmdCompensa 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   5640
         TabIndex        =   178
         Top             =   4680
         Width           =   1335
      End
      Begin VB.CommandButton cmdCompensa 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   7200
         TabIndex        =   176
         Top             =   4680
         Width           =   1335
      End
      Begin MSComctlLib.ListView ListView8 
         Height          =   2895
         Left            =   240
         TabIndex        =   177
         Top             =   1440
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   5106
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tipo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Numero"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Observaciones"
            Object.Width           =   5539
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Total"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha expediente"
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   39
         Left            =   240
         TabIndex        =   179
         ToolTipText     =   "Fecha alta asociado"
         Top             =   960
         Width           =   7995
      End
      Begin VB.Label Label7 
         Caption         =   "Compensar  cobros / abonos"
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
         Index           =   12
         Left            =   240
         TabIndex        =   175
         Top             =   360
         Width           =   7890
      End
   End
   Begin VB.Frame FrameAbono 
      Height          =   6615
      Left            =   2880
      TabIndex        =   148
      Top             =   960
      Visible         =   0   'False
      Width           =   8775
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1200
         TabIndex        =   150
         Text            =   "Text2"
         Top             =   1680
         Width           =   7095
      End
      Begin VB.TextBox txtCliente 
         Height          =   360
         Index           =   2
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   151
         Top             =   2160
         Width           =   1035
      End
      Begin VB.TextBox txtClienteDes 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   2
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   156
         Top             =   2160
         Width           =   5955
      End
      Begin VB.TextBox txtFecha 
         Height          =   360
         Index           =   10
         Left            =   2160
         MaxLength       =   30
         TabIndex        =   149
         Top             =   1200
         Width           =   1515
      End
      Begin VB.CommandButton cmdAbono 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   5760
         TabIndex        =   153
         Top             =   6000
         Width           =   1335
      End
      Begin VB.CommandButton cmdAbono 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   7200
         TabIndex        =   152
         Top             =   6000
         Width           =   1335
      End
      Begin MSComctlLib.ListView ListView6 
         Height          =   2895
         Left            =   240
         TabIndex        =   158
         Top             =   2880
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   5106
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tipo"
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Numero"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Observaciones"
            Object.Width           =   3069
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Total"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Motivo"
         Height          =   240
         Index           =   35
         Left            =   240
         TabIndex        =   159
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1680
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   240
         Index           =   34
         Left            =   240
         TabIndex        =   157
         ToolTipText     =   "Fecha alta asociado"
         Top             =   2220
         Width           =   675
      End
      Begin VB.Image imgCli 
         Height          =   240
         Index           =   2
         Left            =   960
         Top             =   2220
         Width           =   240
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   10
         Left            =   1800
         Picture         =   "frmMensajes.frx":042C
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label7 
         Caption         =   "Factura de abono"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Index           =   10
         Left            =   2520
         TabIndex        =   155
         Top             =   360
         Width           =   3525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha factura"
         Height          =   240
         Index           =   33
         Left            =   240
         TabIndex        =   154
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1200
         Width           =   1395
      End
   End
   Begin VB.Frame FramePagostasas 
      Height          =   8775
      Left            =   0
      TabIndex        =   51
      Top             =   0
      Visible         =   0   'False
      Width           =   16335
      Begin VB.Frame FramePreguntaTasa 
         BorderStyle     =   0  'None
         Height          =   3975
         Left            =   4560
         TabIndex        =   133
         Top             =   2880
         Visible         =   0   'False
         Width           =   6375
         Begin VB.ComboBox cboBanco 
            Height          =   360
            Index           =   1
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   168
            Top             =   2520
            Width           =   4335
         End
         Begin VB.CommandButton cmdConfirmaTipoPagoTasa 
            Caption         =   "Cancelar"
            Height          =   495
            Index           =   0
            Left            =   4440
            TabIndex        =   138
            Top             =   3240
            Width           =   1455
         End
         Begin VB.CommandButton cmdConfirmaTipoPagoTasa 
            Caption         =   "Aceptar"
            Height          =   495
            Index           =   1
            Left            =   2760
            TabIndex        =   137
            Top             =   3240
            Width           =   1455
         End
         Begin VB.TextBox txtInforPagoTasa 
            Height          =   1455
            Left            =   360
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   136
            Text            =   "frmMensajes.frx":04B7
            Top             =   360
            Width           =   5535
         End
         Begin VB.OptionButton optPagoTasas 
            Caption         =   "Banco"
            Height          =   255
            Index           =   1
            Left            =   4200
            TabIndex        =   135
            Top             =   2040
            Width           =   1695
         End
         Begin VB.OptionButton optPagoTasas 
            Caption         =   "Caja"
            Height          =   255
            Index           =   0
            Left            =   2640
            TabIndex        =   134
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Banco"
            Height          =   240
            Index           =   37
            Left            =   360
            TabIndex        =   167
            ToolTipText     =   "Fecha alta asociado"
            Top             =   2565
            Width           =   600
         End
         Begin VB.Shape Shape1 
            BorderWidth     =   3
            Height          =   3735
            Left            =   120
            Top             =   120
            Width           =   6135
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Realizar pago por :"
            Height          =   240
            Index           =   30
            Left            =   360
            TabIndex        =   139
            ToolTipText     =   "Fecha alta asociado"
            Top             =   2040
            Width           =   1860
         End
      End
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
         Picture         =   "frmMensajes.frx":04BD
         Top             =   8160
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   2
         Left            =   600
         Picture         =   "frmMensajes.frx":0607
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
         Picture         =   "frmMensajes.frx":0751
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
         Picture         =   "frmMensajes.frx":07DC
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1320
         Width           =   240
      End
   End
   Begin VB.Frame FrameContabTasasAdm 
      Height          =   3255
      Left            =   3120
      TabIndex        =   142
      Top             =   4320
      Visible         =   0   'False
      Width           =   7815
      Begin VB.ComboBox cboBanco 
         Height          =   360
         Index           =   0
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   166
         Top             =   1800
         Width           =   4335
      End
      Begin VB.CommandButton cmdContabTaasAdm 
         Caption         =   "Contabilizar"
         Height          =   480
         Index           =   1
         Left            =   3960
         TabIndex        =   144
         Top             =   2520
         Width           =   1575
      End
      Begin VB.CommandButton cmdContabTaasAdm 
         Caption         =   "Cancelar"
         Height          =   480
         Index           =   0
         Left            =   5760
         TabIndex        =   145
         Top             =   2520
         Width           =   1575
      End
      Begin VB.TextBox txtFecha 
         Height          =   360
         Index           =   9
         Left            =   3240
         MaxLength       =   30
         TabIndex        =   143
         Top             =   1140
         Width           =   1515
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Banco"
         Height          =   240
         Index           =   36
         Left            =   2160
         TabIndex        =   165
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1800
         Width           =   600
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   9
         Left            =   2880
         Picture         =   "frmMensajes.frx":0867
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label7 
         Caption         =   "Contabilizar pago tasas administrativas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   495
         Index           =   9
         Left            =   360
         TabIndex        =   147
         Top             =   240
         Width           =   7170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   240
         Index           =   32
         Left            =   2160
         TabIndex        =   146
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1200
         Width           =   600
      End
   End
   Begin VB.Frame FrameImpHcoCierres 
      Height          =   6015
      Left            =   360
      TabIndex        =   160
      Top             =   2640
      Visible         =   0   'False
      Width           =   8295
      Begin VB.CommandButton cmdImprCierre 
         Caption         =   "Imprimir"
         Height          =   375
         Index           =   0
         Left            =   4920
         TabIndex        =   164
         Top             =   5280
         Width           =   1335
      End
      Begin VB.CommandButton cmdImprCierre 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   6480
         TabIndex        =   161
         Top             =   5280
         Width           =   1335
      End
      Begin MSComctlLib.ListView ListView7 
         Height          =   3615
         Left            =   120
         TabIndex        =   162
         Top             =   1440
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   6376
         View            =   3
         LabelEdit       =   1
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
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Importe"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "FechaCierreAnterior"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label7 
         Caption         =   "C"
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
         Index           =   11
         Left            =   240
         TabIndex        =   163
         Top             =   360
         Width           =   7170
      End
   End
   Begin VB.Frame FrameContabilizarFras 
      Height          =   3375
      Left            =   4080
      TabIndex        =   122
      Top             =   2160
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CommandButton cmdContabilizar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   1
         Left            =   3480
         TabIndex        =   125
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CommandButton cmdContabilizar 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   0
         Left            =   5040
         TabIndex        =   126
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox txtFecha 
         Height          =   360
         Index           =   8
         Left            =   4680
         MaxLength       =   30
         TabIndex        =   124
         Text            =   "commor"
         Top             =   1320
         Width           =   1515
      End
      Begin VB.TextBox txtFecha 
         Height          =   360
         Index           =   7
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   123
         Text            =   "commor"
         Top             =   1320
         Width           =   1515
      End
      Begin VB.Label Label1 
         Height          =   240
         Index           =   29
         Left            =   360
         TabIndex        =   131
         ToolTipText     =   "Fecha alta asociado"
         Top             =   2040
         Width           =   5760
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha factura"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   28
         Left            =   240
         TabIndex        =   130
         ToolTipText     =   "Fecha alta asociado"
         Top             =   960
         Width           =   1875
      End
      Begin VB.Label Label7 
         Caption         =   "Traspaso facturas contabilidad"
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
         Index           =   8
         Left            =   960
         TabIndex        =   129
         Top             =   240
         Width           =   4935
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   240
         Index           =   27
         Left            =   3720
         TabIndex        =   128
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1320
         Width           =   570
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   8
         Left            =   4440
         Picture         =   "frmMensajes.frx":08F2
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   240
         Index           =   26
         Left            =   240
         TabIndex        =   127
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1320
         Width           =   600
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   7
         Left            =   1080
         Picture         =   "frmMensajes.frx":097D
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1320
         Width           =   240
      End
   End
   Begin VB.Frame FrameCompraTasas 
      Height          =   3615
      Left            =   4200
      TabIndex        =   112
      Top             =   240
      Visible         =   0   'False
      Width           =   8775
      Begin VB.CommandButton cmdCompraTasas 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   1
         Left            =   5400
         TabIndex        =   115
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CommandButton cmdCompraTasas 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   6960
         TabIndex        =   116
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox txtTasas 
         Height          =   360
         Index           =   1
         Left            =   2880
         MaxLength       =   30
         TabIndex        =   114
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
         TabIndex        =   119
         Top             =   1800
         Width           =   5355
      End
      Begin VB.TextBox txtFecha 
         Height          =   360
         Index           =   6
         Left            =   2880
         MaxLength       =   30
         TabIndex        =   113
         Text            =   "commor"
         Top             =   1200
         Width           =   2475
      End
      Begin VB.Label Label1 
         Caption         =   "Cantidad"
         Height          =   240
         Index           =   25
         Left            =   240
         TabIndex        =   121
         ToolTipText     =   "Fecha alta asociado"
         Top             =   2400
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "Concepto"
         Height          =   240
         Index           =   24
         Left            =   240
         TabIndex        =   120
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
         TabIndex        =   118
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha / hora compra"
         Height          =   240
         Index           =   23
         Left            =   240
         TabIndex        =   117
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1200
         Width           =   2115
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   6
         Left            =   2520
         Picture         =   "frmMensajes.frx":0A08
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1200
         Width           =   240
      End
   End
   Begin VB.Frame FrameCierreCaja 
      Height          =   5535
      Left            =   4680
      TabIndex        =   84
      Top             =   240
      Visible         =   0   'False
      Width           =   6495
      Begin VB.TextBox txtCaja 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   7
         Left            =   2400
         MaxLength       =   30
         TabIndex        =   86
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
         TabIndex        =   90
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
         TabIndex        =   94
         Top             =   2280
         Width           =   1515
      End
      Begin VB.TextBox txtFecha 
         Height          =   360
         Index           =   4
         Left            =   2400
         MaxLength       =   30
         TabIndex        =   85
         Text            =   "commor"
         Top             =   1200
         Width           =   2475
      End
      Begin VB.CommandButton cmdCierreCaja 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   1
         Left            =   3360
         TabIndex        =   87
         Top             =   4800
         Width           =   1335
      End
      Begin VB.CommandButton cmdCierreCaja 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   0
         Left            =   4920
         TabIndex        =   88
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
         TabIndex        =   89
         Top             =   1680
         Width           =   2475
      End
      Begin VB.Label Label1 
         Caption         =   "Queda en caja"
         Height          =   240
         Index           =   17
         Left            =   600
         TabIndex        =   97
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
         TabIndex        =   96
         ToolTipText     =   "Fecha alta asociado"
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Importe actual"
         Height          =   240
         Index           =   16
         Left            =   600
         TabIndex        =   95
         ToolTipText     =   "Fecha alta asociado"
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha cierre"
         Height          =   240
         Index           =   15
         Left            =   600
         TabIndex        =   93
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1200
         Width           =   1395
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   4
         Left            =   2040
         Picture         =   "frmMensajes.frx":0A93
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
         TabIndex        =   92
         Top             =   360
         Width           =   3150
      End
      Begin VB.Label Label1 
         Caption         =   "Usuario"
         Height          =   240
         Index           =   13
         Left            =   600
         TabIndex        =   91
         ToolTipText     =   "Fecha alta asociado"
         Top             =   1800
         Width           =   1395
      End
   End
   Begin VB.Frame FrameCobros 
      Height          =   6720
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   13410
      Begin VB.CommandButton cmdEfectuarCobro 
         Caption         =   "Efectuar cobro"
         Height          =   495
         Left            =   360
         TabIndex        =   132
         Top             =   6000
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   495
         Left            =   12000
         TabIndex        =   34
         Top             =   6000
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView5 
         Height          =   4770
         Left            =   225
         TabIndex        =   35
         Top             =   1005
         Width           =   13035
         _ExtentX        =   22992
         _ExtentY        =   8414
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
         Top             =   480
         Width           =   10185
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
         Picture         =   "frmMensajes.frx":0B1E
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
         Picture         =   "frmMensajes.frx":0BA9
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
         Picture         =   "frmMensajes.frx":0C34
         ToolTipText     =   "Quitar seleccion"
         Top             =   720
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   4920
         Picture         =   "frmMensajes.frx":0D7E
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
     '1.-  Facturar expediente
     '2.-  Facturas periodicas
     '3.-  pagos tasas administrativas
     '4.-  Caja infreso/gastos
     '5.-  Cierre caja
     '6.-  Cobro por caja de factura pendiente
     '7.-  Compra tasas
     '8.-  Contabilziar facturas
     '9.-  Contabilizar tasas administrativas
     '10.- Hco cierres de caja
     
     
     
     '22- Ver empresas bloquedas
     '25- Informe de base de datos
     '26- Show processlist
     
     '27- Cobros de la factura
    
     '28- Factura abono
     '29- Factura ABONO desde mto cliente. Ya viene el numero de factura, cliente....
    
    
    
    
Public Parametros As String
    


Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmCta As frmBasico
Attribute frmCta.VB_VarHelpID = -1


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
Dim Sql As String
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
    
    Sql = ListView2(Origen).ListItems(indice).Key
    Set IT = ListView2(Destino).ListItems.Add(, Sql)
    IT.SmallIcon = NE
    IT.Text = ListView2(Origen).ListItems(indice).Text
    IT.SubItems(1) = ListView2(Origen).ListItems(indice).SubItems(1)

    'Borramos en origen
    ListView2(Origen).ListItems.Remove indice
End Sub

Private Sub Check1_Click()

End Sub

Private Sub chkCaja_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub chkQuitarGastos_Click()
    If chkQuitarGastos.Value Then
       
       txtCaja(10).Text = CCur(txtCaja(8).Tag) - CCur(txtCaja(12).Tag)
    Else
        txtCaja(10).Text = txtCaja(8).Text
    End If
End Sub

Private Sub cmdAbono_Click(Index As Integer)
Dim N As Integer

    CadenaDesdeOtroForm = ""
    If Index = 0 Then
        Codigo = "OK"
        If ListView6.ListItems.Count = 0 Then Codigo = ""
        If Me.txtCliente(2).Text = "" Then Codigo = ""
        If Me.txtFecha(10).Text = "" Then Exit Sub
        If Codigo = "" Then Exit Sub
        
        Codigo = ""
        For I = 1 To ListView6.ListItems.Count
            If ListView6.ListItems(I).Checked Then
                Codigo = Codigo & "X"
                N = I
            End If
        Next I
        If Len(Codigo) <> 1 Then
            MsgBox "Seleccione una unica factura para realizar el abono", vbExclamation
            Exit Sub
        End If
        
        If Not FechaFacturaOK(CDate(txtFecha(10).Text)) Then Exit Sub
                
        Codigo = "RectSer = " & DBSet(ListView6.ListItems(N).Text, "T") & " And RectNum = " & ListView6.ListItems(N).SubItems(2) & " and  RectFecha ="
        Codigo = Codigo & DBSet(ListView6.ListItems(N).SubItems(1), "F") & " AND 1"
        Codigo = DevuelveDesdeBD("concat('FRT ',numfactu, '         Por ', usuario ,' el ',fecha)", "factcli", Codigo, "1")
        
        If Codigo <> "" Then
            Codigo = "Factura fue rectificada/abonada anteriormente: " & vbCrLf & Codigo
            Codigo = Codigo & vbCrLf & "¿Continuar de igual modo?"
            If MsgBox(Codigo, vbQuestion + vbYesNoCancel + vbDefaultButton3) <> vbYes Then Exit Sub
        End If
        
        'Hacemos la pregunta
        Codigo = ListView6.ListItems(N).Text & " " & ListView6.ListItems(N).SubItems(2) & " fecha: " & ListView6.ListItems(N).SubItems(1)
        Codigo = Codigo & " de " & ListView6.ListItems(N).SubItems(4) & " €" & vbCrLf
        If ListView6.ListItems(N).Text = "CUO" Then Codigo = Codigo & vbCrLf & vbCrLf & "         *** CUOTA ***   " & vbCrLf & vbCrLf
        Codigo = "Va a realizar el abono de la factura : " & vbCrLf & Space(10) & Codigo & "¿Continuar?"
        If MsgBox(Codigo, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        
        Screen.MousePointer = vbHourglass
        Conn.BeginTrans
        J = 0
        If CrearFacturaAbono Then
            J = 1
            Conn.CommitTrans
            
        Else
            CadenaDesdeOtroForm = ""
            Conn.RollbackTrans
        End If
        Screen.MousePointer = vbDefault
        If J = 0 Then Exit Sub
        
    End If
    Unload Me
End Sub

Private Sub cmdBloqEmpre_Click(Index As Integer)
    If Index = 0 Then
        Sql = "DELETE FROM usuarios.usuarioempresasariconta WHERE codusu =" & Parametros
        Conn.Execute Sql
        Sql = ""
        For I = 1 To ListView2(1).ListItems.Count
            Sql = Sql & ", (" & Parametros & "," & Val(Mid(ListView2(1).ListItems(I).Key, 2)) & ")"
        Next I
        If Sql <> "" Then
            'Quitmos la primera coma
            Sql = Mid(Sql, 2)
            Sql = "INSERT INTO usuarios.usuarioempresasariconta(codusu,codempre) VALUES " & Sql
            If Not EjecutaSQL(Sql) Then MsgBox "Se han producido errores insertando datos", vbExclamation
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
        
        If Me.optConceCaja(1).Value Then
            If Me.txtCodmacta(0).Text = "" Then
                MsgBox "Debe especifiar la cuenta contable", vbExclamation
                Exit Sub
            End If
        Else
            'Pondremos en la cuenta y la contrapartida el el valor desde caja_conceptos
            I = cboConceptosCaja.ItemData(cboConceptosCaja.ListIndex)
            Sql = "caja_conceptos inner join ariconta" & vParam.Numconta & ".cuentas on caja_conceptos.codmacta=cuentas.codmacta"
            cadParam = "nommacta"
            Sql = DevuelveDesdeBD("caja_conceptos.codmacta", Sql, "codconcec", CStr(I), "N", cadParam)
            If Sql = "" Then
                MsgBox "Error obteniendo cuenta contable del concepto", vbExclamation
                Exit Sub
            End If
            Me.txtCodmacta(0).Text = Sql
            txtCodmactaDes(0).Text = cadParam
            
        End If
        
        
        
        If Me.txtCodmacta(0).Text = "" Xor Me.txtCodmactaDes(0).Text = "" Then
            MsgBox "Error en cuenta contable", vbExclamation
            Exit Sub
        End If
        
        If ImporteFormateado(txtCaja(2).Text) < 0 Then
            MsgBox "Importe debe sera mayor que cero", vbExclamation
            Exit Sub
        End If
        
        If Parametros = "" Then
            'NUEVO
            Sql = "usuario = " & DBSet(txtCaja(3).Text, "T") & " AND 1 "
            Sql = DevuelveDesdeBD("feccaja", "caja_param", Sql, " 1 ORDER BY feccaja DESC", "N")
            If Sql <> "" Then
                If CDate(Sql) > CDate(txtCaja(0).Text) Then
                    MsgBox "La caja esta cerrada para esta fecha", vbExclamation
                    Exit Sub
                End If
            End If
        End If
        
        If Parametros <> "" Then
            'MODIFICAR
            Msg = "UPDATE caja set importe= " & DBSet(txtCaja(2).Text, "N") & ", tipomovi=" & Abs(Me.chkCaja.Value)
            Msg = Msg & ", ampliacion =" & DBSet(txtCaja(1).Text, "T") & ", codmacta ="
            If Me.txtCodmacta(0).Text = "" Then
                Msg = Msg & "NULL"
            Else
                Msg = Msg & DBSet(Me.txtCodmacta(0).Text, "T")
            End If
            
             Msg = Msg & ", codconceC ="
            If Me.optConceCaja(0).Value Then
                Msg = Msg & cboConceptosCaja.ItemData(cboConceptosCaja.ListIndex)
            Else
                Msg = Msg & "-1"
            End If
            
            
            Msg = Msg & " WHERE usuario = " & DBSet(txtCaja(3).Text, "T") & " AND feccaja=" & DBSet(txtCaja(0).Text, "FH")
            
        Else
        
        
          
            'INSERTAR
            Msg = "INSERT INTO caja(usuario,feccaja,tipomovi,importe,ampliacion,codmacta,codconceC) VALUES (" & DBSet(txtCaja(3).Text, "T")
            Msg = Msg & "," & DBSet(txtCaja(0).Text, "FH") & "," & Abs(Me.chkCaja.Value)
            Msg = Msg & "," & DBSet(txtCaja(2).Text, "N") & "," & DBSet(txtCaja(1).Text, "T") & ","
            If Me.txtCodmacta(0).Text = "" Then
                Msg = Msg & "NULL"
            Else
                Msg = Msg & DBSet(Me.txtCodmacta(0).Text, "T")
            End If
            'codnce
            Msg = Msg & ","
            If Me.optConceCaja(0).Value Then
                Msg = Msg & cboConceptosCaja.ItemData(cboConceptosCaja.ListIndex)
            Else
                Msg = Msg & "-1"
            End If
            Msg = Msg & ")"
        End If
        If Not Ejecuta(Msg) Then Exit Sub
        
        'Si es miodificar. LOG
        If Parametros <> "" Then
            
            
            'MODIFICAR
            Msg = "concat(if(tipomovi=1,'Salida','Entrada'),'  €',importe,'  ',ampliacion,'   Cta:',coalesce(codmacta,''),'|')"
            Msg = DevuelveDesdeBD(Msg, "caja", "usuario = " & DBSet(txtCaja(3).Text, "T") & " AND feccaja=" & DBSet(txtCaja(0).Text, "FH") & " AND 1", "1")
            Msg = "Anterior: " & Msg & vbCrLf & "ACTUAL  : "
            Msg = Msg & IIf(Me.chkCaja.Value, "Salida", "Entrada") & "  €" & txtCaja(2).Text
            Msg = Msg & "   " & txtCaja(1).Text & "   Cta: " & Me.txtCodmacta(0).Text
            vLog.Insertar 6, vUsu, Msg
            
        End If
        
        CadenaDesdeOtroForm = "OK"
    End If
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub






Private Sub cmdCierreCaja_Click(Index As Integer)
Dim Hora As Date

    CadenaDesdeOtroForm = ""
    If Index = 1 Then
        Sql = ""
        If Me.txtCaja(6).Text = "" Then Sql = "- Indique importe de cierre"
        If Me.txtCaja(7).Text = "" Then Sql = "- Indique importe de cierre"
        If txtFecha(4).Text = "" Then Sql = Sql & vbCrLf & "- Indique fecha cierre"
        
        If Sql <> "" Then
            MsgBox Sql, vbExclamation
            Exit Sub
        End If
        If Not FechaFacturaOK(CDate(txtFecha(4).Text)) Then Exit Sub

        'A partir de la fecha, vamos a ver la hora de cierre de caja.
        'Primero. NO pueden haber movimientos posteriores al cierre
        Sql = DevuelveDesdeBD("max(feccaja)", "caja", "usuario", txtCaja(4).Text, "T")
        
        If Sql <> "" Then
            Hora = CDate(Sql)
            If CDate(Hora) > CDate(txtFecha(4).Text) Then
                MsgBox "Hay moviemientos posteriores en la caja. (" & Hora & ")", vbExclamation
                Exit Sub
            End If
    
        End If
            
            
        If ImporteFormateado(txtCaja(6).Text) > ImporteFormateado(txtCaja(5).Text) Then
            MsgBox "Importe cierre mayor que el importe actual en caja  ", vbExclamation
            Exit Sub
        End If
        
        If ImporteFormateado(txtCaja(6).Text) < 0 Then
            MsgBox "Importe a entregar NEGATIVO", vbExclamation
            
            
            If MsgBox("¿CONTINUAR?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        End If
        
            
        Sql = "Va a cerrar la caja" & vbCrLf & vbCrLf & "Usuario: " & txtCaja(4).Text & vbCrLf
        Sql = Sql & "Fecha cierre: " & txtFecha(4).Text & vbCrLf & vbCrLf
        Sql = Sql & "Importe a entregar: " & txtCaja(6).Text & "€ " & vbCrLf & vbCrLf
        Sql = Sql & "Importe queda en caja: " & txtCaja(7).Text & "€ " & vbCrLf & vbCrLf
        If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        CadenaDesdeOtroForm = txtFecha(4).Text & "|" & txtCaja(6).Text & "|" & txtCaja(7).Text & "|"
        
    End If
    Unload Me
End Sub


Private Sub cmdCobroFactura_Click(Index As Integer)
    BotonCobroFactura Index, False
End Sub

Private Sub BotonCobroFactura(Index As Integer, Credito As Boolean)

    If Index = 1 Then
        Sql = ""
        If Me.txtCaja(10).Text = "" Then Sql = "- Indique importe de cobro"
        If txtFecha(5).Text = "" Then Sql = Sql & vbCrLf & "- Indique fecha cobro"
        
        If Sql <> "" Then
            MsgBox Sql, vbExclamation
            Exit Sub
        End If
        
        Sql = ""
        If CDate(txtFecha(5).Text) < vEmpresa.FechaInicioEjercicio Then
            Sql = "Menor inicio ejercicios"
        Else
            If CDate(txtFecha(5).Text) > DateAdd("yyyy", 1, vEmpresa.FechaFinEjercicio) Then
                Sql = "Menor fin ejercicios"
            Else
                If CDate(txtFecha(5).Text) < vEmpresa.FechaActivaConta Then Sql = "Menor fecha activa contabilidad"
                
            End If
        End If
        
        If Sql <> "" Then
            MsgBox Sql, vbExclamation
            Exit Sub
        End If

        'A partir de la fecha, vamos a ver la hora de cierre de caja.
        'Primero. NO pueden haber movimientos posteriores al cierre
        Sql = DevuelveDesdeBD("max(feccaja)", "caja", "usuario", txtCaja(11).Text, "T")
        
        If Sql <> "" Then
            'Hora = CDate(SQL)
            'If CDate(Hora) > CDate(txtFecha(5).Text) Then
            If CDate(Sql) > CDate(txtFecha(5).Text) Then
                MsgBox "Hay moviemientos posteriores en la caja. (" & Sql & ")", vbExclamation
                Exit Sub
            End If
    
        End If
            

        'DE momento NO admito cobros parciales
        Importe = ImporteFormateado(txtCaja(8).Text)
        If Me.chkQuitarGastos.Value Then Importe = CCur(txtCaja(8).Tag) - CCur(txtCaja(12).Tag)
        
        Sql = Format(Importe, FormatoImporte)
        
        
        If Sql <> txtCaja(10).Text Then
            MsgBox "No aceptados cobros a cuenta", vbExclamation
            Exit Sub
        End If
        Sql = ""
        If Credito Then
            Sql = String(40, "*") & vbCrLf & vbCrLf
            Sql = Sql & Space(15) & "Tarjeta de crédito" & vbCrLf & vbCrLf & Sql
        End If
        Sql = Sql & "¿Realizar cobro?"
        If MsgBox(Sql, vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then Exit Sub
        
        
        If Not HacerCobroFactura(Credito) Then Exit Sub
            
            
        'Vamos a imprimir la factura pq lleva la fecha de pago puesta
        Sql = " {factcli.numserie}  = " & DBSet(RecuperaValor(CampoOrden, 1), "T") & " AND {factcli.numfactu} =" & RecuperaValor(CampoOrden, 2)
        Sql = Sql & "  AND {factcli.fecfactu} = Date(" & Format(RecuperaValor(Parametros, 4), "yyyy,mm,dd") & ") "
        
        cadNomRPT = "rFactura.rpt"
        cadFormula = Sql
        cadParam = ""
        numParam = 0
        conSubRPT = False
                
        ImprimeGeneral
                
    End If
    Unload Me

End Sub

Private Sub cmdCompensa_Click(Index As Integer)
Dim B As Boolean
    If Index = 0 Then
        Sql = ""
        Importe = 0
        For I = 1 To Me.ListView8.ListItems.Count
            If Me.ListView8.ListItems(I).Checked Then
                Importe = ImporteFormateado(ListView8.ListItems(I).SubItems(4)) + Importe
                Sql = Sql & "X"
            End If
        Next
        
        
        If Len(Sql) <= 1 Then
            MsgBox "Seleccione al menos dos vencimientos", vbExclamation
            Exit Sub
        End If
        
        If Importe <> 0 Then
            MsgBox "La compensacion no da como resultado CERO", vbExclamation
            Exit Sub
        End If
        
        
        'OK vamos a realizar la compensacion
        If MsgBox("Realizar la compensacion?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        
        Screen.MousePointer = vbHourglass
        Conn.BeginTrans
        B = RealizaCompensacion
        Screen.MousePointer = vbDefault
        If B Then
            Conn.CommitTrans
            
        Else
            Conn.RollbackTrans
            Exit Sub
        End If
        
        
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

Private Sub cmdConfirmaTipoPagoTasa_Click(Index As Integer)
    If Index = 0 Then
        'Ha cancelado
        PonerFramesPagoTasas False
        
    Else
        
        'Aceptar. Vamos adelante
                
        Conn.BeginTrans
        If GenerarTasas Then
            Conn.CommitTrans
            'CadenaDesdeOtroForm = "OK"  dentro de la funcion ya esta
            Unload Me
        Else
            Conn.RollbackTrans
            CadenaDesdeOtroForm = ""
            PonerFramesPagoTasas False
        End If
        
        
        
    End If
        
End Sub


Private Sub PonerFramesPagoTasas(Pregunta As Boolean)
    
    Me.FramePreguntaTasa.Visible = Pregunta
    cmdpagoTasas(0).Enabled = Not Pregunta
    cmdpagoTasas(1).Enabled = Not Pregunta
    cmdVerdatos.Enabled = Not Pregunta

    If Pregunta Then
        cmdConfirmaTipoPagoTasa(0).Cancel = True
    Else
        cmdpagoTasas(1).Cancel = True
    End If
    
    
End Sub


Private Sub cmdContabilizar_Click(Index As Integer)
Dim Aux As String

    If Index = 1 Then
        InicializarVbles False
        
        'Primero comprobar si hay alguna factura en el intervalo
        If Not PonerDesdeHasta("factcli.fecfactu", "FEC", Me.txtFecha(7), Nothing, txtFecha(8), Nothing, "pDH=""") Then Exit Sub
        
        If cadselect = "" Then cadselect = " 1= 1"
        Sql = cadselect & "  AND intconta=0 "
        Aux = Sql
        Sql = DevuelveDesdeBD("count(*)", "factcli", Sql & " AND 0", "0")
        If Val(Sql) = 0 Then
            MsgBox "Ninguna factura para contabilizar", vbExclamation
            Exit Sub
        End If
        
        
        'Un par de comprobaciones basicas
        'Conceptos. Exsiten todos
        Sql = "from factcli,factcli_lineas ,conceptos where"
        Sql = Sql & " factcli_lineas.numserie = factcli.numserie and factcli_lineas.numfactu = factcli.numfactu and"
        Sql = Sql & " factcli_lineas.Fecfactu = factcli.Fecfactu And factcli_lineas.codconce = conceptos.codconce"
        Sql = Sql & " AND " & Aux
        Sql = Sql & " and codmacta is null order by 1"
        Set miRsAux = New ADODB.Recordset
        
        miRsAux.Open "select distinct factcli_lineas.codconce,conceptos.nomconce " & Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Sql = ""
        I = 0
        While Not miRsAux.EOF
            Sql = Sql & Format(miRsAux!codconce, "0000") & " " & miRsAux!nomconce & vbCrLf
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        If Sql <> "" Then
            MsgBox "Conceptos sin asiganar cuenta contabilidad: " & vbCrLf & Sql, vbExclamation
            Set miRsAux = Nothing
            Exit Sub
        End If
            
        'En "aux" esta el D/H fecha
        Sql = " select distinct factcli.numserie,factcli.numfactu,factcli.fecfactu  from factcli,factcli_lineas ,conceptos where"
        Sql = Sql & " factcli_lineas.numserie = factcli.numserie and factcli_lineas.numfactu = factcli.numfactu and"
        Sql = Sql & " factcli_lineas.Fecfactu = factcli.Fecfactu And factcli_lineas.codconce = conceptos.codconce"
        Sql = Sql & " AND " & Aux
        Sql = Sql & "  ORDER BY 1,2,3"
        miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        I = 0
        While Not miRsAux.EOF
            I = I + 1
            miRsAux.MoveNext
        Wend
        
      
        miRsAux.Close
        Set miRsAux = Nothing
        If I = 0 Then
            MsgBox "Ninguna factura a contabilizar", vbExclamation
            Exit Sub
        End If
        
        Aux = I
        If I > 1 Then Aux = "las " & I
        Aux = "Va a contabilizar " & Aux & " facturas. ¿Continuar?"
        
        If MsgBox(Aux, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
            
        
        
        If BloqueoManual("CONTABILIZAR", "1") Then
            Me.cmdContabilizar(0).Enabled = False
            Me.cmdContabilizar(1).Enabled = False
            Me.Refresh
            DoEvents
            Label1(29).Caption = "Inicio proceso"
            Label1(29).Refresh
        
        
            Screen.MousePointer = vbHourglass
            
            PasarFactura cadselect, Label1(29)
                
            Screen.MousePointer = vbDefault
            DesBloqueoManual "CONTABILIZAR"
            Me.cmdContabilizar(0).Enabled = True
            Me.cmdContabilizar(1).Enabled = True
            Label1(29).Caption = ""
        End If

    End If
    Unload Me
End Sub

Private Sub cmdContabTaasAdm_Click(Index As Integer)
    If Index = 1 Then
        
        If txtFecha(9).Text = "" Then
            MsgBox "Indique fecha contabilizacion", vbExclamation
            PonFoco txtFecha(9)
            Exit Sub
        End If
        If Not FechaFacturaOK(CDate(txtFecha(9).Text)) Then Exit Sub
        
        If MsgBox("Desea contabilizar con fecha: " & txtFecha(9).Text & "?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        
        CadenaDesdeOtroForm = txtFecha(9).Text
        If Me.cboBanco(0).Visible Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & "|" & Me.cboBanco(0).ListIndex + 1
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & "|"
    End If
    Unload Me
End Sub


Private Sub CargaComboBancos(indice As Integer)
    
    cboBanco(indice).Clear
    cboBanco(indice).AddItem DevuelveDesdeBD("nommacta", "ariconta" & vParam.Numconta & ".cuentas", "codmacta", vParam.CtaBanco, "T")
    If vParam.CtaBanco2 <> "" Then
        cboBanco(indice).AddItem DevuelveDesdeBD("nommacta", "ariconta" & vParam.Numconta & ".cuentas", "codmacta", vParam.CtaBanco2, "T")
    End If
    cboBanco(indice).ListIndex = 0
End Sub


Private Sub cmdEfectuarCobro_Click()
Dim Cad As String
    If ListView5.ListItems.Count = 0 Then Exit Sub
    If ListView5.SelectedItem Is Nothing Then Exit Sub
    If ListView5.SelectedItem.Tag = "-1" Then
        MsgBox "No se puede realizar el cobro sobre el seleccionado", vbExclamation
        Exit Sub
    End If
    
    
    Cad = "select numorden, formapago.nomforpa, fecvenci, impvenci, gastos, fecultco, impcobro, (coalesce(impvenci,0) + coalesce(gastos,0) - coalesce(impcobro,0)) pendiente, cobros.ctabanc1  "
    Cad = Cad & " ,tipforpa,formapago.codforpa,numserie,numfactu,fecfactu,numorden ,codmacta,nomclien"
    Cad = Cad & " from (ariconta" & vParam.Numconta & ".cobros left join ariconta" & vParam.Numconta & ".formapago on cobros.codforpa = formapago.codforpa) "
    Cad = Cad & " where cobros.numserie = " & DBSet(RecuperaValor(Parametros, 1), "T")
    Cad = Cad & " and cobros.numfactu = " & DBSet(RecuperaValor(Parametros, 2), "N")
    Cad = Cad & " and cobros.fecfactu = " & DBSet(RecuperaValor(Parametros, 3), "F")
    Cad = Cad & " and cobros.numorden=" & ListView5.SelectedItem.Text
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If miRsAux.EOF Then
        'Error
        Set miRsAux = Nothing
    Else
        
        frmTESParciales.Vto = miRsAux!numSerie & "|" & miRsAux!NumFactu & "|" & miRsAux!Fecfactu & "|" & miRsAux!numorden & "|"
        Cad = ""
        If Not IsNull(miRsAux!Gastos) Then Cad = Cad & Format(miRsAux!Gastos, FormatoImporte)
        Cad = Cad & "|"
        If Not IsNull(miRsAux!impcobro) Then Cad = Cad & Format(miRsAux!impcobro, FormatoImporte)
        
        frmTESParciales.Importes = Format(miRsAux!ImpVenci, FormatoImporte) & Cad & "|"
        Cad = DevuelveDesdeBD("nommacta", "ariconta" & vParam.Numconta & ".cuentas", "codmacta", miRsAux!ctabanc1, "T")
        frmTESParciales.Cta = miRsAux!codmacta & "|" & DBLet(miRsAux!NomClien, "T") & "|" & miRsAux!ctabanc1 & "|" & Cad & "|"
        frmTESParciales.FormaPago = CInt(Mid(ListView5.SelectedItem.ListSubItems(1).Key, 2))
        Set miRsAux = Nothing
        frmTESParciales.Show vbModal
        CargaCobrosFactura
    End If
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

Private Sub cmdImprCierre_Click(Index As Integer)
    If Index = 1 Then
        Unload Me
    Else
        If ListView7.ListItems.Count = 0 Then Exit Sub
        If ListView7.SelectedItem Is Nothing Then Exit Sub
        ImprimirCajaDelDia
        
    End If
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
                
        'Por banco o caja
        Sql = "1"
        If Parametros <> "" Then
            Sql = DevuelveDesdeBD("pagoPorBanco", "gestadministrativa", "id", Parametros)
            
            If Sql = "1" Then
                'Pago por banco
                Sql = DevuelveDesdeBD("QueBanco", "gestadministrativa", "id", Parametros)
                Me.cboBanco(1).ListIndex = 0
                If Val(Sql) > 1 Then Me.cboBanco(1).ListIndex = 1
                Sql = "1"
            End If
        End If
        Me.optPagoTasas(CInt(Sql)).Value = True
                
        Sql = "Total lineas: " & Chr(9) & J & vbCrLf
        Sql = Sql & "Importe:   " & Chr(9) & Format(Importe, FormatoImporte)
        
            
        Me.txtInforPagoTasa.Text = Sql
        PonerFramesPagoTasas True
        Exit Sub
        
    End If
    Unload Me
End Sub

Private Sub cmdRegresar_Click()
    Unload Me
End Sub
 


Private Sub SQL_Exepdientes(Valor As String)
       Sql = "select l.numexped,l.anoexped,fecexped,licencia,nomclien ,pagado ,conceptos.nomconce,e.codclien,l.importe"
        Sql = Sql & " ,l.codconce ,l.tiporegi,l.numlinea"
        Sql = Sql & " from expedientes e,expedientes_lineas l,clientes,conceptos"
        Sql = Sql & " Where pagado=" & Valor & " and e.tiporegi = L.tiporegi And e.numexped = L.numexped And e.anoexped = L.anoexped And e.CodClien = Clientes.CodClien"
        'Solo las de gestion amdinistrativa
        Sql = Sql & " AND conceptos.codconce=l.codconce and conceptos.gestionadm = 1"

         
        If Valor = "0" Then
            'Desde hasta
            If Me.txtCliente(0).Text <> "" Then Sql = Sql & " AND e.codclien >= " & txtCliente(0).Text
            If Me.txtCliente(1).Text <> "" Then Sql = Sql & " AND e.codclien <= " & txtCliente(1).Text
              
            If Me.txtFecha(2).Text <> "" Then Sql = Sql & " AND e.fecexped >= " & DBSet(txtFecha(2).Text, "F")
            If Me.txtFecha(3).Text <> "" Then Sql = Sql & " AND e.fecexped <= " & DBSet(txtFecha(3).Text, "F")
        End If
        
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdTarjetaCredito_Click()
    BotonCobroFactura 1, True
End Sub

Private Sub cmdVerdatos_Click()
    Screen.MousePointer = vbHourglass
    ListView1.ListItems.Clear
    Set miRsAux = New ADODB.Recordset
    I = 0
    For J = 1 To 2
        Sql = ""
        If J = 1 And Parametros <> "" Then SQL_Exepdientes Parametros
        If J = 2 Then SQL_Exepdientes 0
        
        
        If Sql <> "" Then
            miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
        
            If Parametros = "" Then
                PonFoco txtCaja(0)
            Else
                PonFoco txtCaja(1)
            End If
        
        Case 5
            'Cierre caja
            Me.txtFecha(4).Text = Format(Now, "dd/mm/yyyy hh:nn:ss")
            txtCaja(4).Text = RecuperaValor(Parametros, 1)
            txtCaja(5).Text = RecuperaValor(Parametros, 2)
        
        Case 6
            
        
            Me.txtFecha(5).Text = Format(Now, "dd/mm/yyyy hh:nn:ss")
            txtCaja(8).Text = RecuperaValor(Parametros, 5)   'pendeinte
            txtCaja(8).Tag = ImporteFormateado(txtCaja(8).Text)
            txtCaja(9).Text = RecuperaValor(Parametros, 2)   'factura
            txtCaja(10).Text = RecuperaValor(Parametros, 5)  'cobrado
            txtCaja(11).Text = RecuperaValor(Parametros, 1)  'usuar
            PonFoco txtCaja(10)
            PonerGastosVto
        
            
        
        
        Case 7
            Me.txtTasas(0).Text = RecuperaValor(Parametros, 1)   'factura
            Me.txtFecha(6).Text = Format(Now, "dd/mm/yyyy hh:nn:ss")
            Me.txtTasas(1).Text = ""  '
            
        Case 9
            
            
        Case 10
             cargaCierresDeCaja

        Case 22
            cargaempresasbloquedas
            
        Case 25
            CargaInformeBBDD
        
        Case 26
            CargaShowProcessList
        
        Case 27
            CargaCobrosFactura
            
        Case 28, 29
        
            txtFecha(10).Text = Format(Now, "dd/mm/yyyy")
            
            If Opcion = 29 Then

                
                imgCli(2).Enabled = False
                txtCliente(2).Enabled = False
                txtCliente(2).Text = Format(RecuperaValor(CadenaDesdeOtroForm, 1), "0000")
                Me.txtClienteDes(2).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
                CargaFacturasCliente True
            End If
            PonFoco txtFecha(10)
            CadenaDesdeOtroForm = ""
       Case 30
            CargaCobrosAbonos
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
Private Sub CargaIMGCta(indice As Integer)
    imgCodmacta(indice).Picture = frmppal.imgIcoForms.ListImages(1).Picture
End Sub

Private Sub Form_Load()
Dim W As Long
Dim H As Long

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
    FrameContabilizarFras.Visible = False
    FrameAbono.Visible = False
    FrameImpHcoCierres.Visible = False
    FrameCompensacion.Visible = False
    Limpiar Me
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
        CargaComboBancos 1
    Case 4
        CargarComboConceptos
        PonerVisibleConceptosCaja True
        
        Caption = "Caja"
        PonerFrameVisible Me.FrameCaja, W, H
        Me.txtCaja(3).Text = RecuperaValor(Parametros, 1)
        Msg = Trim(RecuperaValor(Parametros, 2)) 'La fecha
        
        BloquearTxt txtCaja(0), Msg <> ""
        imgFecCaja(0).Visible = Msg = ""
        If Msg <> "" Then
            'Esta modifiacando
            txtCaja(0).Text = Msg
            txtCaja(1).Text = RecuperaValor(Parametros, 3)
            txtCaja(2).Text = RecuperaValor(Parametros, 4)
            Me.chkCaja.Value = IIf(RecuperaValor(Parametros, 5) = "1", 1, 0)
            Me.txtCodmacta(0).Text = RecuperaValor(Parametros, 6)
            Me.txtCodmactaDes(0).Text = RecuperaValor(Parametros, 7)
            Sql = RecuperaValor(Parametros, 8)
            If Sql <> "-1" Then
                
                'Conceptos caja
                For I = 0 To Me.cboConceptosCaja.ListCount - 1
                    If cboConceptosCaja.ItemData(I) = Sql Then
                        cboConceptosCaja.ListIndex = I
                        Exit For
                    End If
                Next
                If I >= Me.cboConceptosCaja.ListCount Then MsgBox "No se ha encontrado el concepto de caja: " & Sql, vbExclamation
            Else
                PonerVisibleConceptosCaja False
                Me.optConceCaja(1).Value = True
            End If
        Else
            txtCaja(0).Text = Format(Now, "dd/mm/yyyy hh:nn")
            Parametros = ""
        End If
        
        CargaIMGCta 0
        
    Case 5
        'Cierre caja
        PonerFrameVisible FrameCierreCaja, W, H
        Caption = "CAJA"
        cmdCierreCaja(0).Cancel = True
        
    Case 6
        
        PonerFrameVisible FramepagoFra, W, H
        Caption = "CAJA"
        cmdCobroFactura(0).Cancel = True
        
        
    Case 7
        'Cmpra tasas
        PonerFrameVisible FrameCompraTasas, W, H
        Caption = "TASAS"
        cmdCompraTasas(0).Cancel = True
        txtCaja(1).Text = RecuperaValor(Parametros, 3)
        txtCaja(2).Text = RecuperaValor(Parametros, 4)
        
            
    Case 8
        'Paso facturas contabilidad
        PonerFrameVisible FrameContabilizarFras, W, H
        Caption = "Contabilizar"
        cmdContabilizar(0).Cancel = True
   
    Case 9
        PonerFrameVisible FrameContabTasasAdm, W, H
        Caption = "Contabilizar"
        cmdContabTaasAdm(0).Cancel = True
        If CadenaDesdeOtroForm <> "NO" Then CargaComboBancos 0
        Me.cboBanco(0).Visible = CadenaDesdeOtroForm <> "NO"
        Label1(36).Visible = CadenaDesdeOtroForm <> "NO"
        If CadenaDesdeOtroForm <> "NO" Then cboBanco(0).ListIndex = Val(CadenaDesdeOtroForm) - 1
        CadenaDesdeOtroForm = ""
    Case 10
        PonerFrameVisible FrameImpHcoCierres, W, H
        Caption = "Historico caja"
        Label7(11).Caption = "Cierres caja usuario: " & Parametros
   
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
        
        PonerFrameVisible FrameCobros, W, H
        
    Case 28, 29
        
        Me.Caption = "Facturas de abono"
        PonerFrameVisible FrameAbono, W, H
        CargaIMG 2
        
        cmdAbono(1).Cancel = True
    
    Case 30
        Caption = "Compensa"
        PonerFrameVisible FrameCompensacion, W, H
        Me.cmdCompensa(1).Cancel = True
        
        Me.Label1(39).Caption = RecuperaValor(Parametros, 1) & " - " & RecuperaValor(Parametros, 2)
        
    End Select
    
    Me.Width = W + 120
    Me.Height = H + 120
End Sub


Private Sub PonerFrameVisible(QueFrame As Frame, ByRef Wi As Long, ByRef He As Long)
    QueFrame.top = -60
    QueFrame.Left = 0
    QueFrame.Visible = True
    Wi = QueFrame.Width
    He = QueFrame.Height + 300
 
End Sub





Private Sub VerEmresasProhibidas(ByRef VarProhibidas As String)

On Error GoTo EVerEmresasProhibidas
    VarProhibidas = "|"
    Sql = "Select codempre from Usuarios.usuarioempresasariconta WHERE codusu = " & (vUsu.Codigo Mod 1000)
    Sql = Sql & " order by codempre"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
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







Private Sub frmCta_DatoSeleccionado(CadenaSeleccion As String)
    Sql = CadenaSeleccion
End Sub

Private Sub frmF_Selec(vFecha As Date)
    Sql = vFecha
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
    Line Input #I, Sql
    Line Input #I, Sql
    Exit Sub
ELeerCadenaFicheroTexto:
    Sql = ""
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
Private Function EjecutaSQL2(Sql As String) As Boolean
    EjecutaSQL2 = False
    On Error Resume Next
    Conn.Execute Sql
    If Err.Number <> 0 Then
        AnyadeErrores "SQL: " & Sql, Err.Description
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

Private Sub imgCodmacta_Click(Index As Integer)
    Set frmCta = New frmBasico
    Sql = ""
    AyudaCtasContabilidad frmCta, , "codmacta like '4%'"
    If Sql <> "" Then
        
        Me.txtCodmacta(Index).Text = RecuperaValor(Sql, 1)
        Me.txtCodmactaDes(Index).Text = RecuperaValor(Sql, 2)
    End If
    Set frmCta = Nothing
End Sub

Private Sub imgFecCaja_Click(Index As Integer)


    Set frmF = New frmCal
    frmF.Fecha = Now
    Sql = ""
    If txtCaja(Index).Text <> "" Then frmF.Fecha = txtCaja(Index).Text
    frmF.Show vbModal
    If Sql <> "" Then
        txtCaja(Index).Text = Sql
        txtCaja(Index).Text = txtCaja(Index).Text & " " & Format(Now, "hh:mm:ss")
    End If
End Sub

Private Sub imgppal_Click(Index As Integer)
    Set frmF = New frmCal
    frmF.Fecha = Now
    Sql = ""
    If Me.txtFecha(Index).Text <> "" Then frmF.Fecha = txtFecha(Index).Text
    frmF.Show vbModal
    If Sql <> "" Then
        txtFecha(Index).Text = Sql
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

Private Sub optConceCaja_Click(Index As Integer)
    PonerVisibleConceptosCaja Index = 0
End Sub

Private Sub PonerVisibleConceptosCaja(Visible As Boolean)
    cboConceptosCaja.Visible = Visible
    cboConceptosCaja.Left = txtCodmacta(0).Left
    imgCodmacta(0).Visible = Not Visible
    txtCodmacta(0).Visible = Not Visible
    txtCodmactaDes(0).Visible = Not Visible
End Sub


Private Sub optPagoTasas_Click(Index As Integer)
    Me.cboBanco(1).Visible = Index = 1
    Label1(37).Visible = Index = 1
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
    
    Conn.Execute Sql
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
    Sql = "select empresasariconta.codempre,nomempre,nomresum,usuarioempresasariconta.codempre bloqueada from usuarios.empresasariconta left join usuarios.usuarioempresasariconta on "
    Sql = Sql & " empresasariconta.codempre = usuarioempresasariconta.codempre And (usuarioempresasariconta.codusu = " & Parametros & " Or codusu Is Null)"
    '[Monica] solo ariconta
    Sql = Sql & " WHERE conta like 'ariconta%' "
    Sql = Sql & " ORDER BY empresasariconta.codempre"
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs.EOF
        Errores = Format(Rs!codempre, "00000")
        Sql = "C" & Errores
        
        If IsNull(Rs!bloqueada) Then
            'Va al list de la derecha
            Set IT = ListView2(0).ListItems.Add(, Sql)
            IT.SmallIcon = 1
        Else
            Set IT = ListView2(1).ListItems.Add(, Sql)
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
    
    Sql = "select * from tmpinfbbdd where codusu = " & vUsu.Codigo & " order by posicion "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
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
Dim MostramosBoton As Boolean
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
    ListView5.ListItems.Clear
    
    Cad = "select numorden, formapago.nomforpa, fecvenci, impvenci, gastos, fecultco, impcobro, (coalesce(impvenci,0) + coalesce(gastos,0) - coalesce(impcobro,0)) pendiente, cobros.ctabanc1  "
    Cad = Cad & " ,tipforpa,formapago.codforpa "
    Cad = Cad & " from (ariconta" & vParam.Numconta & ".cobros left join ariconta" & vParam.Numconta & ".formapago on cobros.codforpa = formapago.codforpa) "
    Cad = Cad & " where cobros.numserie = " & DBSet(RecuperaValor(Parametros, 1), "T")
    Cad = Cad & " and cobros.numfactu = " & DBSet(RecuperaValor(Parametros, 2), "N")
    Cad = Cad & " and cobros.fecfactu = " & DBSet(RecuperaValor(Parametros, 3), "F")
    Cad = Cad & " order by numorden "
    
    Rs.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    MostramosBoton = False
    While Not Rs.EOF
                    
        Set IT = ListView5.ListItems.Add
        
        If CobroContabilizado(RecuperaValor(Parametros, 1), RecuperaValor(Parametros, 2), RecuperaValor(Parametros, 3), DBLet(Rs.Fields(0))) Then IT.SmallIcon = 18
        
        IT.Text = DBLet(Rs.Fields(0))
        IT.SubItems(1) = DBLet(Rs.Fields(1))
        
        IT.ListSubItems(1).Key = "F" & Rs!TipForpa
        IT.SubItems(2) = DBLet(Rs.Fields(2))
        IT.SubItems(3) = Format(DBLet(Rs.Fields(3)), "###,###,##0.00")
        
        
        IT.SubItems(4) = " " & Format(DBLet(Rs.Fields(4)), "###,###,##0.00")
        IT.SubItems(5) = " " & DBLet(Rs.Fields(5))
        IT.SubItems(6) = " " & Format(DBLet(Rs.Fields(6)), "###,###,##0.00")
        IT.SubItems(7) = Format(DBLet(Rs.Fields(7)), "###,###,##0.00")
        
        
'        0   "EFECTIVO"
'        1   "TRANSFERENCIA"
'        2   "TALON"
'        3   "PAGARE"
'        4   "RECIBO BANCARIO"
'        5   "CONFIRMING"
'        6   "TARJETA DE CREDITO"
        
        IT.Tag = -1
        If Rs!pendiente > 0 Then
            'Hay pendiente, igual puede efectuar el cobro
          '  If Rs!TipForpa <> 4 And Rs!TipForpa <> 0 Then
          '      MostramosBoton = True
          '      IT.Tag = Rs!TipForpa
          '  End If
        End If
        
        
        'Siguiente
        Rs.MoveNext
    Wend
    NumRegElim = 0
    Rs.Close
    Me.cmdEfectuarCobro.Visible = MostramosBoton
    
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
                IT.Text = DBLet(Rs!numSerie)
                IT.Bold = True
                IT.ForeColor = vbRed
                IT.SubItems(1) = "Pago a cuenta caja:" & Rs!Usuario
                IT.SubItems(2) = Format(Rs!feccaja, "dd/mm/yyyy")
                IT.SubItems(3) = Format(DBLet(Rs!Importe), "###,###,##0.00")
                
                
                IT.SubItems(4) = " "
                IT.SubItems(5) = " "
                IT.SubItems(6) = " " & Format(DBLet(Rs!Importe), "###,###,##0.00")
                IT.SubItems(7) = Format(0, "###,###,##0.00")
                IT.Tag = -1   'no se puede efectuar el cobro
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
Dim Sql As String

    Sql = "select * from ariconta" & vParam.Numconta & ".hlinapu where numserie = " & DBSet(Serie, "T") & " and numfaccl = " & DBSet(FACTURA, "N") & " and fecfactu = " & DBSet(Fecha, "F") & " and numorden = " & DBSet(Orden, "N")
    CobroContabilizado = (TotalRegistrosConsulta(Sql) <> 0)

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
    Else
        If Index = 0 And txtCaja(Index).Text <> "" Then
            If Not EsFechaHoraOK(txtCaja(Index)) Then
                MsgBox "Fecha incorrecta", vbExclamation
                txtCaja(Index).Text = ""
                PonFoco txtCaja(Index)
            End If
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
    Sql = ""
    If PonerFormatoEntero(txtCliente(Index)) Then

        Sql = DevuelveDesdeBD("nomclien", "clientes", "codclien", txtCliente(Index).Text, "N")
        If Sql = "" Then Sql = "No existe el cliente"
    End If
    Me.txtClienteDes(Index).Text = Sql
    If Index = 2 Then CargaFacturasCliente Sql <> ""
        
End Sub



Private Sub txtCodmacta_GotFocus(Index As Integer)
     ConseguirFoco txtFecha(Index), 4
End Sub

Private Sub txtCodmacta_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCodmacta_LostFocus(Index As Integer)
    CuentaCorrectaUltimoNivelTextBox txtCodmacta(Index), txtCodmactaDes(Index)
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
        Sql = "UPDATE expedientes_lineas set  codsitua=0, pagado =0 where "
        Sql = Sql & " pagado =" & Parametros
        Conn.Execute Sql
    
        Codigo = Parametros
        
        Sql = "UPDATE gestadministrativa SET llevados = " & J & ", importe=" & DBSet(Importe, "N")
        Sql = Sql & ", pagoPorBanco =" & IIf(Me.optPagoTasas(1).Value, 1, 0)
        Sql = Sql & ", queBanco= 0"
        If Me.optPagoTasas(1).Value Then Sql = Sql & Me.cboBanco(1).ListIndex + 1
        
        Sql = Sql & " WHERE id =" & Codigo
        CadenaDesdeOtroForm = "OK"
    Else
        Codigo = DevuelveDesdeBD("max(id)", "gestadministrativa", "1", "1")
        Codigo = Val(Codigo) + 1
        CadenaDesdeOtroForm = Codigo
        Sql = "INSERT INTO gestadministrativa(id,usuario,fechacreacion,llevados,importe,pagoPorBanco,quebanco) VALUES (" & Codigo
        Sql = Sql & "," & DBSet(vUsu.Login, "T") & "," & DBSet(Now, "FH") & "," & J & "," & DBSet(Importe, "N")
        Sql = Sql & "," & IIf(Me.optPagoTasas(1).Value, 1, 0) & ","
        If optPagoTasas(1).Value Then
            Sql = Sql & cboBanco(1).ListIndex + 1 & ")"    'Sera el quebanco 1.Ppal  2 Segundo en parametros
        Else
            Sql = Sql & "0)"   'CAJA
        End If
    End If
    Conn.Execute Sql
    
    J = 0
    Importe = 0
    For I = 1 To ListView1.ListItems.Count
        If Me.ListView1.ListItems(I).Checked Then
            Sql = "UPDATE expedientes_lineas set  codsitua=1, pagado =" & Codigo & " where "
            Sql = Sql & " tiporegi='0' AND numexped =" & ListView1.ListItems(I).Text
            Sql = Sql & " AND anoexped=" & ListView1.ListItems(I).Tag
            Sql = Sql & " AND numlinea=" & ListView1.ListItems(I).SubItems(2)
            Conn.Execute Sql
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
Private Function HacerCobroFactura(Credito As Boolean) As Boolean

    
    'Transaccion
    Screen.MousePointer = vbHourglass
    HacerCobroFactura = False
    Conn.BeginTrans
    If RealizarCobro(Credito) Then   'Credito: TARJETA CREDITO
        Conn.CommitTrans
        HacerCobroFactura = True
    Else
        Conn.RollbackTrans
    End If
    Screen.MousePointer = vbDefault
End Function

Private Function RealizarCobro(Credito As Boolean) As Boolean
Dim Cobrado As Currency

    On Error GoTo eRealizarCobro
    RealizarCobro = False
    Set miRsAux = New ADODB.Recordset
    
    Errores = "Separando datos factura"
    CampoOrden = RecuperaValor(Parametros, 3)
    CampoOrden = Replace(CampoOrden, ":", "|")
    
    Sql = " numserie = " & DBSet(RecuperaValor(CampoOrden, 1), "T") & " AND numfactu =" & RecuperaValor(CampoOrden, 2)
    Sql = Sql & " AND numorden=" & RecuperaValor(CampoOrden, 3) & " AND fecfactu =" & DBSet(RecuperaValor(Parametros, 4), "F")
    Codigo = "SELECT impvenci,gastos,fecultco,impcobro,numserie,numfactu,fecfactu,numorden FROM ariconta" & vParam.Numconta & ".cobros WHERE " & Sql
    miRsAux.Open Codigo, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'NO PUEDE SER EOF
    Cobrado = DBLet(miRsAux!impcobro, "N")
    Importe = miRsAux!ImpVenci
    If Not (Me.chkQuitarGastos.Value = 1) Then Importe = Importe + DBLet(miRsAux!Gastos, "N")
    
    
    If Cobrado + ImporteFormateado(Me.txtCaja(10).Text) > Importe Then Err.Raise 513, "Cobrado mas de lo que tiene pendiente"
    
    Importe = ImporteFormateado(txtCaja(10).Text)  'lo que vamos a cobrar ahora
    Cobrado = Cobrado + ImporteFormateado(Me.txtCaja(10).Text)   'lo que habia mas lo de ahora
     
    Codigo = "UPDATE ariconta" & vParam.Numconta & ".cobros SET "
    Codigo = Codigo & " fecultco = " & DBSet(txtFecha(5), "F") & ", impcobro = " & DBSet(Cobrado, "N")
    'Si ponia a NULL los gastos
    If Me.chkQuitarGastos.Value = 1 Then Codigo = Codigo & " ,gastos = NULL"
    
    'observaciones
    Referencia = " Caja: " & txtCaja(11).Text & "     Fecha/hora: " & txtFecha(5).Text & IIf(Credito, "     TARJETA CREDITO", "")
    Referencia = Replace(Referencia, "'", "")
    Codigo = Codigo & " , observa=trim(concat(coalesce(observa,''),' " & Referencia & "'))"
    
    Codigo = Codigo & " WHERE " & Sql
    Conn.Execute Codigo
    miRsAux.Close
    
    Errores = "Obteniendo tipo registro"
    Sql = "Select * from contadores where serfactur =" & DBSet(RecuperaValor(CampoOrden, 1), "T")
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    
    'Ya hemos actualizado en cobros. Ahora metemos en caja
    Codigo = "INSERT INTO caja(usuario,feccaja,tipomovi,tiporegi,numserie,numdocum,anoexped,numlinea,importe,ampliacion,codconcec) VALUES ("
    Codigo = Codigo & DBSet(txtCaja(11).Text, "T") & "," & DBSet(txtFecha(5).Text, "FH") & ",0," & DBSet(miRsAux!TipoRegi, "N")
    Sql = RecuperaValor(Parametros, 4)
    Sql = Year(CDate(Sql))
    Codigo = Codigo & "," & DBSet(miRsAux!serfactur, "T") & "," & DBSet(RecuperaValor(CampoOrden, 2), "T") & "," & DBSet(Sql, "N")
    Sql = RecuperaValor(Parametros, 6)
    Codigo = Codigo & "," & DBSet(RecuperaValor(CampoOrden, 3), "N") & "," & DBSet(Importe, "N") & "," & DBSet(Sql, "T") & ",-1)"
    Conn.Execute Codigo
    
    
    
    
    
    
    'Si es tarjeta de credito, entonces hacemos el apunte de salida, un segundo mas tarde
    If Credito Then
        Codigo = "INSERT INTO caja(usuario,feccaja,tipomovi,tiporegi,numserie,numdocum,anoexped,numlinea,importe,ampliacion,codmacta,codconcec) VALUES ("
        Codigo = Codigo & DBSet(txtCaja(11).Text, "T") & "," & DBSet(DateAdd("s", 1, CDate(txtFecha(5).Text)), "FH") & ",1,null,null,null,null,1"
        Sql = RecuperaValor(Parametros, 2)
        Sql = "TAR.CRED:" & Sql & " " & RecuperaValor(Parametros, 6)
        Sql = Mid(Sql, 1, 50)
        Codigo = Codigo & "," & DBSet(Importe, "N") & "," & DBSet(Sql, "T")
        
        Sql = DevuelveDesdeBD("codmacta", "caja_conceptos", "CodConceC", vParam.CajaConceptoTarjetaCredito)
        If Sql = "" Then Err.Raise 513, , "Error obteniendo concepot TARJETA: " & vParam.CajaConceptoTarjetaCredito
        Codigo = Codigo & "," & DBSet(Sql, "T") & "," & vParam.CajaConceptoTarjetaCredito & ")"
        Conn.Execute Codigo
        
    End If
    
    
    
    
    
    
    
    
    miRsAux.Close
    
    'Marzo 18
    'Updateamos el campo `FecPago` de factcli
    
    Sql = " numserie = " & DBSet(RecuperaValor(CampoOrden, 1), "T") & " AND numfactu =" & RecuperaValor(CampoOrden, 2)
    Sql = Sql & " AND fecfactu =" & DBSet(RecuperaValor(Parametros, 4), "F")
    Codigo = "UPDATE factcli SET FecPago = " & DBSet(txtFecha(5).Text, "FH") & " WHERE " & Sql
    Conn.Execute Codigo
    
    
    If Me.chkQuitarGastos.Value = 1 Then
        Codigo = "Factura: " & txtCaja(9).Text & vbCrLf
        Codigo = Codigo & "Pendiente: " & txtCaja(8).Text & " (Gastos " & txtCaja(12).Text & ")"
        Codigo = Codigo & vbCrLf & "Pagado: " & Me.txtCaja(10).Text & IIf(Credito, "     TARJETA CREDITO", "")
        vLog.Insertar 11, vUsu, Codigo
    
    End If
    
    RealizarCobro = True
    
eRealizarCobro:
    If Err.Number <> 0 Then MuestraError Err.Number, Errores & vbCrLf & Err.Description
    Set miRsAux = Nothing
End Function


Private Sub CargaFacturasCliente(Enlaza As Boolean)
    
    Me.ListView6.ListItems.Clear
    If Not Enlaza Then Exit Sub
    
    Set miRsAux = New ADODB.Recordset
    Codigo = "select numserie,numfactu,fecfactu,observa,totfaccl from factcli WHERE CODCLIEN=" & Me.txtCliente(2).Text & " AND numserie <>'FRT'"
    Codigo = Codigo & " AND RectSer is null"
    
    'Estamos en mto clientes. Vamos a rectificar una factura /abono
    If Me.Opcion = 29 Then
        
        'Ejmplo parametros: FAC|1|29/01/2018|
        Codigo = Codigo & " AND numserie =" & DBSet(RecuperaValor(Parametros, 1), "T")
        Codigo = Codigo & " AND numfactu =" & DBSet(RecuperaValor(Parametros, 2), "N")
        Codigo = Codigo & " AND fecfactu =" & DBSet(RecuperaValor(Parametros, 3), "F")
        
    End If
    Codigo = Codigo & " order by fecfactu desc, numserie,numfactu "
    miRsAux.Open Codigo, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    While Not miRsAux.EOF
        I = I + 1
        ListView6.ListItems.Add , , Format(miRsAux!numSerie, "0000")
        ListView6.ListItems(I).SubItems(1) = Format(miRsAux!Fecfactu, "dd/mm/yyyy")
        ListView6.ListItems(I).SubItems(2) = Format(miRsAux!NumFactu, "000000")
        ListView6.ListItems(I).SubItems(3) = DBLet(miRsAux!observa, "T")
        ListView6.ListItems(I).SubItems(4) = Format(miRsAux!totfaccl, FormatoImporte)
           
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If Opcion = 29 Then
        
        If I > 0 Then
            ListView6.ListItems(1).Selected = True
            ListView6.ListItems(1).Checked = True
        End If
    End If
    Set miRsAux = Nothing
End Sub




Private Function CrearFacturaAbono() As Boolean
Dim Obser As String
Dim EsCuota As Boolean
Dim txtaux As String

On Error GoTo eCrearFacturaAbono

    CrearFacturaAbono = False
    Sql = ""
    For I = 1 To ListView6.ListItems.Count
        If ListView6.ListItems(I).Checked Then Exit For
    Next
    CampoOrden = "numserie =" & DBSet(ListView6.ListItems(I).Text, "T") & " AND numfactu =" & ListView6.ListItems(I).SubItems(2)
    CampoOrden = CampoOrden & " AND fecfactu  =" & DBSet(ListView6.ListItems(I).SubItems(1), "F")
                
                    
    'LAS CUOTAS no generan FRT, generan CUO con importe negativo
    If ListView6.ListItems(I).Text = "CUO" Then
        Referencia = "CUO"
        txtaux = "Abono de la cuota: "
    Else
        txtaux = "Rectifica a la factura: "
        Referencia = "FRT"
    End If
    Sql = "numfactu"
    Codigo = DevuelveDesdeBD("numfactu", "contadores", "serfactur", Referencia, "T")
    If Codigo = "" Then Err.Raise 513, , "No existe contador abonos"
    Codigo = Val(Codigo) + 1
    CadenaDesdeOtroForm = Referencia & "|" & Codigo & "|" & txtFecha(10).Text & "|"
       
    'Cabecera
    
    Obser = Trim(Text2.Text)
    If Obser <> "" Then Obser = Obser & vbCrLf
    
    Obser = Obser & txtaux
    Obser = Obser & ListView6.ListItems(I).Text & ListView6.ListItems(I).SubItems(2)
    Obser = Obser & " de fecha " & ListView6.ListItems(I).SubItems(1)
    
    'numserie,numfactu,fecfactu,codclien,codforpa,numexped,fecexped,observa,totbases,totbasesret,totivas,totrecargo,totfaccl,retfaccl,
    'trefaccl,cuereten,tiporeten,intconta,usuario,fecha
    Sql = "INSERT INTO factcli(numserie,numfactu,fecfactu,codclien,codforpa,numexped,fecexped,observa,totbases,totbasesret,totivas,totrecargo,totfaccl,retfaccl,trefaccl,cuereten,tiporeten,intconta,usuario,fecha,RectSer ,RectNum  ,RectFecha) "
    Sql = Sql & "select " & DBSet(Referencia, "T") & "," & Codigo & "," & DBSet(txtFecha(10).Text, "F") & ",codclien,codforpa,"
    Sql = Sql & " null as numexped,null as fecexped," & DBSet(Obser, "T") & " observa,"
    Sql = Sql & "-totbases,-totbasesret,-totivas,-totrecargo,-totfaccl,retfaccl,trefaccl,cuereten,tiporeten,0 intconta,"
    Sql = Sql & DBSet(vUsu.Login, "T") & " usuario," & DBSet(Now, "FH") & " fecha  ,"
    Sql = Sql & DBSet(ListView6.ListItems(I).Text, "T") & "," & DBSet(ListView6.ListItems(I).SubItems(2), "N") & "," & DBSet(ListView6.ListItems(I).SubItems(1), "F")
    Sql = Sql & " from factcli WHERE " & CampoOrden
    Conn.Execute Sql
    
    EsCuota = False
    If ListView6.ListItems(I).Text = "CUO" Then EsCuota = True
    
    
    'Lineas
    Sql = "INSERT INTO factcli_lineas(numserie,numfactu,fecfactu,numlinea,codconce,nomconce,ampliaci,cantidad,precio,importe,codigiva,porciva,porcrec,impoiva,imporec,aplicret)"
    Sql = Sql & " select " & DBSet(Referencia, "T") & "," & Codigo & "," & DBSet(txtFecha(10).Text, "F")
    Sql = Sql & ",numlinea,codconce,nomconce,ampliaci,-cantidad,precio,-importe,codigiva,porciva,porcrec,-impoiva,-imporec,aplicret"
    Sql = Sql & " FROM factcli_lineas WHERE " & CampoOrden
    Conn.Execute Sql
    
    'Totales
    Sql = "INSERT INTO factcli_totales(numserie,numfactu,fecfactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)"
    Sql = Sql & "select " & DBSet(Referencia, "T") & "," & Codigo & "," & DBSet(txtFecha(10).Text, "F")
    Sql = Sql & ",numlinea,-baseimpo,codigiva,porciva,porcrec,-impoiva,-imporec"
    Sql = Sql & " FROM factcli_totales WHERE " & CampoOrden
    Conn.Execute Sql
    
    'Vencimientos
    'se generan en InsertarEnTesoreria
    
    'Tesoreria
    espera 0.3
    
    Set miRsAux = New ADODB.Recordset
    CampoOrden = "numserie = " & DBSet(Referencia, "T") & " AND numfactu = " & Codigo & " AND fecfactu = " & DBSet(txtFecha(10).Text, "F")
    Sql = " from factcli,clientes where factcli.codclien=clientes.codclien AND " & CampoOrden
    Sql = "licencia,PobClien ,codposta ,ProClien ,NIFClien ,codpais ,IBAN ,totfaccl " & Sql
    Sql = "SELECT factcli.codclien, factcli.codforpa ,numserie,NumFactu ,FecFactu ,NomClien ,DomClien," & Sql
    miRsAux.Open Sql, Conn, adOpenKeyset, adCmdText
    If Not InsertarEnTesoreria(EsCuota, miRsAux, vParam.CtaBanco, "", Msg, 0) Then Err.Raise 513, , "Creando cobro en tesoreria" & vbCrLf & Msg
    
    miRsAux.Close
    Set miRsAux = Nothing
    
    
    Sql = "UPDATE contadores set numfactu=" & Codigo
    Sql = Sql & " WHERE serfactur = " & DBSet(Referencia, "T")
    Ejecuta Sql
    
    
    
    CrearFacturaAbono = True
    
    Exit Function
eCrearFacturaAbono:
    MuestraError Err.Number, , Err.Description
    Set miRsAux = Nothing
End Function



Private Sub cargaCierresDeCaja()
Dim Anterior As String

    Me.ListView7.ListItems.Clear

    
    Set miRsAux = New ADODB.Recordset
    Codigo = "SELECT * from caja_param where usuario=" & DBSet(Parametros, "T") & " order by feccaja desc"
    miRsAux.Open Codigo, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    While Not miRsAux.EOF
        I = I + 1
        ListView7.ListItems.Add , , Format(miRsAux!feccaja, "dd/mm/yyyy hh:mm")
        ListView7.ListItems(I).SubItems(1) = Format(miRsAux!Importe, FormatoImporte)
        ListView7.ListItems(I).SubItems(2) = " "
        ListView7.ListItems(I).Tag = miRsAux!feccaja
               
        miRsAux.MoveNext
        If Not miRsAux.EOF Then ListView7.ListItems(I).SubItems(2) = miRsAux!feccaja
        
    Wend
    miRsAux.Close
    Set miRsAux = Nothing


End Sub

Private Sub ImprimirCajaDelDia()
Dim Aux As String
Dim F As Date
Dim F2 As Date


    InicializarVbles True
    '{caja.feccaja} >= DateTime (2016, 12, 31, 12, 20, 10)
    cadNomRPT = "rCajaCierreHco.rpt"
    
    
    F = CDate(ListView7.SelectedItem.Tag)
    
    'rpt
    cadFormula = Year(F) & "," & Month(F) & "," & Day(F) & "," & Hour(F) & "," & Minute(F) & "," & Second(F)
    cadFormula = "({caja.feccaja} <=  DateTime (" & cadFormula & "))"
    'select
    Msg = "feccaja < " & DBSet(F, "FH")
    Msg = Msg & " AND usuario=" & DBSet(Parametros, "T") & " AND 1 "
    Aux = "feccaja"
    Msg = DevuelveDesdeBD("importe", "caja_param", Msg, "1 ORDER BY  feccaja DESC", "N", Aux)
    If Msg = "" Then Msg = "0": Aux = ""
    Importe = CCur(Msg)
    'Le sumo un segundo a la fecha para el listado , ya que no puedo poner >  y pone>=
    If Aux <> "" Then
        F2 = CDate(Aux)
        Aux = Year(F2) & "," & Month(F2) & "," & Day(F2) & "," & Hour(F2) & "," & Minute(F2) & "," & Second(F2)
        cadFormula = cadFormula & " AND ({caja.feccaja} >  DateTime (" & Aux & " ))"
        Aux = Format(F2, "dd/mm/yyyy hh:mm:ss")
    End If
    
    cadParam = cadParam & "|ImporteIncial=" & TransformaComasPuntos(CCur(Importe)) & "|"
    cadParam = cadParam & "|UltimoCierre= """ & Aux & """|"
    numParam = numParam + 2
    cadFormula = cadFormula & " AND {caja.usuario}= """ & Parametros & """"
    ImprimeGeneral
    
End Sub



Private Sub CargarComboConceptos()
    Me.cboConceptosCaja.Clear
    Sql = "caja_conceptos inner join ariconta" & vParam.Numconta & ".cuentas on caja_conceptos.codmacta=cuentas.codmacta "
    CargaComboTabla cboConceptosCaja, "codconcec", "nomconcec", Sql, " ORDER BY codconcec"
    If cboConceptosCaja.ListCount > 0 Then cboConceptosCaja.ListIndex = 0
End Sub




Private Sub PonerGastosVto()
On Error GoTo ePonerGastosVto
    Set miRsAux = New ADODB.Recordset
    txtCaja(12).Text = ""
    chkQuitarGastos.Visible = False
    Errores = "Separando datos factura"
    CampoOrden = RecuperaValor(Parametros, 3)
    CampoOrden = Replace(CampoOrden, ":", "|")
    
    Sql = " numserie = " & DBSet(RecuperaValor(CampoOrden, 1), "T") & " AND numfactu =" & RecuperaValor(CampoOrden, 2)
    Sql = Sql & " AND numorden=" & RecuperaValor(CampoOrden, 3) & " AND fecfactu =" & DBSet(RecuperaValor(Parametros, 4), "F")
    Codigo = "SELECT gastos FROM ariconta" & vParam.Numconta & ".cobros WHERE " & Sql
    miRsAux.Open Codigo, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux!Gastos) Then
            If miRsAux!Gastos > 0 Then
                txtCaja(12).Text = Format(miRsAux!Gastos, FormatoImporte)
                txtCaja(12).ForeColor = vbRed
                txtCaja(12).FontBold = True
                txtCaja(12).Tag = miRsAux!Gastos
                ' de momento , hay un btono para quitar los gastos       chkQuitarGastos.Visible = True
            End If
        End If
    End If
    miRsAux.Close
    
ePonerGastosVto:
    If Err.Number <> 0 Then MuestraError Err.Number, Errores & vbCrLf & Err.Description
    Set miRsAux = Nothing
    
End Sub



Private Sub CargaCobrosAbonos()
Dim Cad As String
On Error GoTo eCargaCobrosAbonos

    
    Me.ListView8.ListItems.Clear
    Sql = RecuperaValor(Parametros, 1)
    Set miRsAux = New ADODB.Recordset
    Cad = "select if(now()>fecvenci,2,0) vencido "
    Cad = Cad & " ,numorden,fecfactu,fecvenci,ImpVenci + coalesce(gastos, 0) - coalesce(impcobro, 0) as pendiente ,nomforpa,"
    Cad = Cad & "numserie,numfactu,tipoformapago,cobros.codforpa, if(coalesce(gastos,0)>0,2,0) importancia,codmacta,tipforpa "
    
    Cad = Cad & " from ariconta" & vParam.Numconta & ".cobros,ariconta" & vParam.Numconta & ".formapago "
    Cad = Cad & " ,ariconta" & vParam.Numconta & ".tipofpago "
    Cad = Cad & " where codmacta IN ('" & DevuelveCuentaContableCliente(True, Sql)
    Cad = Cad & " ','" & DevuelveCuentaContableCliente(False, Sql) & "')"
    Cad = Cad & " AND cobros.codforpa=formapago.codforpa and tipofpago.tipoformapago = formapago.tipforpa AND "
    Cad = Cad & " ImpVenci + coalesce(gastos, 0) - coalesce(impcobro, 0) <> 0"
    Cad = Cad & " ORDER BY fecvenci desc"
    
    
    
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    I = 0
    While Not miRsAux.EOF
        I = I + 1
        ListView8.ListItems.Add , , miRsAux!numSerie
        ListView8.ListItems(I).SubItems(1) = Format(miRsAux!Fecfactu, "dd/mm/yyyy")
        ListView8.ListItems(I).SubItems(2) = Format(miRsAux!NumFactu, "0000000")
        ListView8.ListItems(I).SubItems(3) = miRsAux!nomforpa
        ListView8.ListItems(I).SubItems(4) = Format(miRsAux!pendiente, FormatoImporte)
        ListView8.ListItems(I).Tag = miRsAux!numorden & "|" & miRsAux!TipForpa & "|" & miRsAux!codmacta & "|"
               
        
        miRsAux.MoveNext
    Wend
    
    
    miRsAux.Close
    
eCargaCobrosAbonos:
    If Err.Number <> 0 Then MuestraError Err.Number, Errores & vbCrLf & Err.Description
    Set miRsAux = Nothing
    
End Sub




Private Function RealizaCompensacion() As Boolean
Dim ColApu As Collection
Dim FP As Ctipoformapago

On Error GoTo eRealizaCompensacion
    RealizaCompensacion = False


    Set ColApu = New Collection
    
    'select condecli,conhacli from formapago,tipofpago where formapago.tipforpa=tipoformapago
   
    For I = 1 To Me.ListView8.ListItems.Count
        

        'codmacta | docum | codconce | ampliaci | imported|importeH |ctacontrpar|
        ' numseri| numfaccl fecfac numorden forpa
          
        Codigo = ""
        With ListView8.ListItems(I)
            If .Checked Then
                NE = RecuperaValor(.Tag, 2) 'Forma de pago
                Ok = 0
                If FP Is Nothing Then
                    Set FP = New Ctipoformapago
                    Ok = 1
                Else
                    If FP.tipoformapago <> NE Then Ok = 1
                End If
                
                If Ok = 1 Then
                    If FP.Leer(NE) = 1 Then Err.Raise 513, , "Tipo forma pago incorrecto."
                End If
                    
        
                
                Codigo = RecuperaValor(.Tag, 3) & "|" & .Text & .SubItems(2) & "|"
                If CCur(.SubItems(4)) < 0 Then
                    'Abono, va al debe
                    Codigo = Codigo & FP.condecli & "|Compen. gestion " & .Text & .SubItems(2) & "|" & Abs(.SubItems(4)) & "|"
                Else
                    'Al haber
                    Codigo = Codigo & FP.conhacli & "|Compen gestion " & .Text & .SubItems(2) & "||" & .SubItems(4)
                End If
                '                   ctra  y datos del cobro
                Codigo = Codigo & "||" & .Text & "|" & .SubItems(2) & "|" & .SubItems(1) & "|"
        
            End If
        End With
        
        If Codigo <> "" Then ColApu.Add Codigo

    Next I

    Sql = "Compensacion cliente gestion: " & Label1(39).Caption
    Sql = Sql & "Usuario gestion:" & vUsu.Login & "   Vtos: " & Format(ColApu.Count, "00")
    If Not CrearApunteDesdeColeccion(Now, Sql, ColApu) Then Err.Raise 513, , "Crear apunte en contabilidad"

    
    
    
    
    'Los cobros los damos como cobrados
    Set ColApu = Nothing
    Set ColApu = New Collection
    Referencia = ""
    For I = 1 To Me.ListView8.ListItems.Count
    
        With ListView8.ListItems(I)
            If .Checked Then
                'fecultco  impcobro situacion  codusu observa
                Sql = "UPDATE ariconta" & vParam.Numconta & ".cobros SET fecultco = " & DBSet(Now, "F")
                Sql = Sql & ", impcobro =" & DBSet(.SubItems(4), "N")
                Sql = Sql & ", situacion=1, codusu = " & vUsu.id & ", Observa =@#@# "
                Sql = Sql & " WHERE numserie =" & DBSet(.Text, "T")
                Sql = Sql & " AND numfactu =" & DBSet(.SubItems(2), "N")
                Sql = Sql & " AND fecfactu =" & DBSet(.SubItems(1), "F")
                Sql = Sql & " AND numorden =" & DBSet(RecuperaValor(.Tag, 1), "N")
                ColApu.Add Sql
                
                
                Referencia = Referencia & .Text & .SubItems(2) & " de " & .SubItems(1) & "-" & RecuperaValor(.Tag, 1) & "     " & .SubItems(4) & "€" & vbCrLf
                

            End If
        End With
    Next I

    Sql = DBSet(Referencia, "T")
    For I = 1 To ColApu.Count
        
        Codigo = ColApu.Item(I)
        Codigo = Replace(Codigo, "@#@#", Sql)
        Conn.Execute Codigo
    Next


    'LOG
    Sql = "Cliente: " & Label1(39).Caption & vbCrLf
    Sql = Sql & Referencia
    vLog.Insertar 12, vUsu, Sql
    
    
    Referencia = ""
    Codigo = ""
    
    
    
    RealizaCompensacion = True
    
    
    Exit Function
eRealizaCompensacion:
    MuestraError Err.Number, Err.Description
    Set ColApu = Nothing
    Referencia = ""
    Codigo = ""
End Function
