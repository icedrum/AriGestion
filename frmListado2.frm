VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListado2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameLog 
      Height          =   6615
      Left            =   120
      TabIndex        =   53
      Top             =   0
      Width           =   6135
      Begin VB.OptionButton optLog 
         Caption         =   "Trabajador"
         Height          =   195
         Index           =   1
         Left            =   1560
         TabIndex        =   50
         Top             =   6120
         Width           =   1095
      End
      Begin VB.OptionButton optLog 
         Caption         =   "Fecha"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   49
         Top             =   6120
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton cmdAcciones 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   2
         Left            =   3600
         TabIndex        =   51
         Top             =   6000
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3735
         Index           =   0
         Left            =   120
         TabIndex        =   47
         Top             =   1920
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   6588
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   2
         Left            =   4800
         TabIndex        =   52
         Top             =   6000
         Width           =   1095
      End
      Begin VB.TextBox txtFecha 
         Height          =   315
         Index           =   4
         Left            =   3480
         TabIndex        =   46
         Text            =   "Text1"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtFecha 
         Height          =   315
         Index           =   3
         Left            =   960
         TabIndex        =   45
         Text            =   "Text1"
         Top             =   1080
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3735
         Index           =   1
         Left            =   3720
         TabIndex        =   48
         Top             =   1920
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   6588
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "1800"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   720
         Picture         =   "frmListado2.frx":0000
         ToolTipText     =   "Quitar al haber"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   1080
         Picture         =   "frmListado2.frx":014A
         ToolTipText     =   "Puntear al haber"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   2
         Left            =   4680
         Picture         =   "frmListado2.frx":0294
         ToolTipText     =   "Quitar al haber"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   3
         Left            =   4920
         Picture         =   "frmListado2.frx":03DE
         ToolTipText     =   "Puntear al haber"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   8
         Left            =   3720
         TabIndex        =   59
         Top             =   1680
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Acción"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   58
         Top             =   1680
         Width           =   555
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   57
         Top             =   1155
         Width           =   615
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   4
         Left            =   3240
         Picture         =   "frmListado2.frx":0528
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   56
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   55
         Top             =   1155
         Width           =   615
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   720
         Picture         =   "frmListado2.frx":05B3
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Impresion LOG"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   1
         Left            =   840
         TabIndex        =   54
         Top             =   360
         Width           =   3615
      End
   End
   Begin VB.Frame Fr340 
      Height          =   4935
      Left            =   -720
      TabIndex        =   0
      Top             =   30
      Width           =   8775
      Begin VB.Frame FrameTick 
         Caption         =   "Frame1"
         Height          =   1095
         Left            =   5040
         TabIndex        =   20
         Top             =   2280
         Width           =   3615
         Begin VB.CommandButton cmdAsignarLetraCta340 
            Cancel          =   -1  'True
            Caption         =   "Asig."
            Height          =   255
            Left            =   2520
            TabIndex        =   9
            ToolTipText     =   "Fijar tickets para empresa"
            Top             =   360
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   2
            Left            =   1920
            MaxLength       =   10
            TabIndex        =   8
            Text            =   "cta2 ticket"
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   1
            Left            =   360
            MaxLength       =   10
            TabIndex        =   7
            Text            =   "cta1 ticket"
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   0
            Left            =   1800
            MaxLength       =   3
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Cuentas"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   360
            TabIndex        =   22
            Top             =   360
            Width           =   1410
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tickets"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   21
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.CheckBox chk340 
         Caption         =   "Fra. proveedores. Grabar misma fecha"
         Height          =   255
         Index           =   1
         Left            =   5040
         TabIndex        =   11
         Top             =   4080
         Width           =   3375
      End
      Begin VB.CheckBox chk340 
         Caption         =   "No añadir facturas régimen especial agrario"
         Height          =   255
         Index           =   0
         Left            =   5040
         TabIndex        =   10
         Top             =   3600
         Width           =   3375
      End
      Begin VB.ComboBox cmbPeriodo 
         Height          =   315
         Index           =   1
         ItemData        =   "frmListado2.frx":063E
         Left            =   5040
         List            =   "frmListado2.frx":0648
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox txtAno 
         Height          =   315
         Index           =   0
         Left            =   7320
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdAcciones 
         Caption         =   "Generar"
         Height          =   375
         Index           =   0
         Left            =   6120
         TabIndex        =   12
         Top             =   4440
         Width           =   1095
      End
      Begin VB.ComboBox cmbPeriodo 
         Height          =   315
         Index           =   0
         ItemData        =   "frmListado2.frx":065F
         Left            =   5040
         List            =   "frmListado2.frx":0661
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1080
         Width           =   1935
      End
      Begin VB.ListBox List2 
         Height          =   3375
         ItemData        =   "frmListado2.frx":0663
         Left            =   360
         List            =   "frmListado2.frx":0665
         TabIndex        =   14
         Top             =   960
         Width           =   4215
      End
      Begin VB.TextBox txtFecha 
         Height          =   315
         Index           =   0
         Left            =   5040
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   7440
         TabIndex        =   13
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "La opcion de tickets se ira pidiendo para cada empresa a declarar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5040
         TabIndex        =   23
         Top             =   2280
         Width           =   3375
      End
      Begin VB.Label lbl340 
         Caption         =   "Label1"
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   4560
         Width           =   5535
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Destino"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   2
         Left            =   5040
         TabIndex        =   18
         Top             =   1560
         Width           =   645
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   5640
         Picture         =   "frmListado2.frx":0667
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   5040
         TabIndex        =   17
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Periodo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   5040
         TabIndex        =   16
         Top             =   840
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Empresas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   33
         Left            =   360
         TabIndex        =   15
         Top             =   720
         Width           =   825
      End
      Begin VB.Image Image5 
         Height          =   240
         Left            =   1440
         Picture         =   "frmListado2.frx":06F2
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Modelo 340"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   9
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame FrameConsum 
      Height          =   2295
      Left            =   1200
      TabIndex        =   73
      Top             =   4200
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CommandButton cmdAcciones 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   4
         Left            =   3960
         TabIndex        =   78
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtCD1 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   75
         Top             =   1200
         Width           =   6015
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   4
         Left            =   5280
         TabIndex        =   74
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Fichero"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   77
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Importar factura CONSUM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   3
         Left            =   840
         TabIndex        =   76
         Top             =   360
         Width           =   4695
      End
      Begin VB.Image ImgCd1 
         Height          =   240
         Index           =   0
         Left            =   1080
         Picture         =   "frmListado2.frx":10F4
         Top             =   960
         Width           =   240
      End
   End
   Begin VB.Frame FrameMemoriaPagos 
      Height          =   3735
      Left            =   120
      TabIndex        =   60
      Top             =   0
      Width           =   5415
      Begin VB.TextBox txtAno 
         Height          =   315
         Index           =   1
         Left            =   1080
         TabIndex        =   64
         Text            =   "Text1"
         Top             =   1920
         Width           =   495
      End
      Begin VB.CheckBox chkMemoria 
         Caption         =   "Cualquier concepto"
         Height          =   255
         Left            =   2880
         TabIndex        =   65
         Top             =   2040
         Width           =   1935
      End
      Begin VB.CommandButton cmdAcciones 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   3
         Left            =   2640
         TabIndex        =   66
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox txtFecha 
         Height          =   315
         Index           =   6
         Left            =   3480
         TabIndex        =   63
         Text            =   "Text1"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   3960
         TabIndex        =   67
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox txtFecha 
         Height          =   315
         Index           =   5
         Left            =   1080
         TabIndex        =   62
         Text            =   "Text1"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Plazo"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   72
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   71
         Top             =   2760
         Width           =   5145
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   6
         Left            =   3240
         Picture         =   "frmListado2.frx":1AF6
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   3
         Left            =   2640
         TabIndex        =   70
         Top             =   1395
         Width           =   615
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   5
         Left            =   840
         Picture         =   "frmListado2.frx":1B81
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   69
         Top             =   1395
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha factua"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   9
         Left            =   240
         TabIndex        =   68
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Memoria pagos proveedor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   61
         Top             =   360
         Width           =   4695
      End
   End
   Begin VB.Frame FrameTraspasoBancoSecciones 
      Height          =   5295
      Left            =   120
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   5415
      Begin VB.TextBox txtDatosDestino 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   1320
         TabIndex        =   43
         Text            =   "Text5"
         Top             =   3720
         Width           =   3855
      End
      Begin VB.TextBox txtDatosDestino 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   41
         Text            =   "Text5"
         Top             =   3360
         Width           =   3855
      End
      Begin VB.CommandButton cmdAcciones 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   1
         Left            =   2880
         TabIndex        =   28
         Top             =   4800
         Width           =   1095
      End
      Begin VB.TextBox txtDatosDestino 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   38
         Text            =   "Text5"
         Top             =   3000
         Width           =   3855
      End
      Begin VB.TextBox txtFecha 
         Height          =   315
         Index           =   2
         Left            =   3480
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   315
         Index           =   1
         Left            =   1080
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2280
         TabIndex        =   34
         Text            =   "Text5"
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   27
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   4080
         TabIndex        =   29
         Top             =   4800
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Diario"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   44
         Top             =   3720
         Width           =   585
      End
      Begin VB.Image ImgDatosDestino 
         Height          =   240
         Index           =   2
         Left            =   1080
         Picture         =   "frmListado2.frx":1C0C
         Top             =   3720
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Concepto"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   42
         Top             =   3360
         Width           =   825
      End
      Begin VB.Image ImgDatosDestino 
         Height          =   240
         Index           =   1
         Left            =   1080
         Picture         =   "frmListado2.frx":260E
         Top             =   3360
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Cuenta"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   40
         Top             =   3000
         Width           =   585
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   39
         Top             =   4800
         Width           =   2385
      End
      Begin VB.Image ImgDatosDestino 
         Height          =   240
         Index           =   0
         Left            =   1080
         Picture         =   "frmListado2.frx":3010
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Destino"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   37
         Top             =   2640
         Width           =   645
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   3120
         Picture         =   "frmListado2.frx":3A12
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   840
         Picture         =   "frmListado2.frx":3A9D
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image ImgAyuda 
         Height          =   240
         Index           =   0
         Left            =   4920
         Picture         =   "frmListado2.frx":3B28
         ToolTipText     =   "Acerca del 347"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image ImgCta 
         Height          =   240
         Index           =   0
         Left            =   840
         Picture         =   "frmListado2.frx":452A
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   72
         Left            =   240
         TabIndex        =   36
         Top             =   2040
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta origen"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   119
         Left            =   240
         TabIndex        =   35
         Top             =   1680
         Width           =   1185
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   14
         Left            =   2520
         TabIndex        =   33
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   15
         Left            =   240
         TabIndex        =   32
         Top             =   1155
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   58
         Left            =   240
         TabIndex        =   31
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Traspaso cuenta banco"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   0
         Left            =   840
         TabIndex        =   30
         Top             =   360
         Width           =   3615
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   8400
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmListado2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public opcion As Byte
    '0 .-  Listado consultas extractos, listado de MAYOR
    '1 .-  Traspaso cuenta banco entre secciones
    '2 .-  Impresion LOG acciones
    '3 .-  Memoria pagos a proveedores
    '4 .-  Pedir fichero importacion CONSUM

Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCta As frmColCtas
Attribute frmCta.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1


Dim Rs As ADODB.Recordset
Dim i As Integer
Dim Cad As String

Dim V340()   'Llevara un str
             'indicara si cada empresa a declarr tiene
             'los tickets como letra de serie o como cuenta
             'en los campos 2 y 3 llevara si es serie la serie
             ' y si es cta las cuentas 1 y dos






Private Sub chk340_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub



Private Sub chkMemoria_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmbPeriodo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAcciones_Click(Index As Integer)
    Select Case Index
    Case 0
        '340
        Hacer340
        
    Case 1
        'Traspaso
        TraspasoEntreCtas
    Case 2
        ListadoLog
        
    Case 3
        'memoria de pagos a proveedor
        MemoriaPagosProv
    Case 4
        
        CadenaDesdeOtroForm = txtCD1(0).Text
        If CadenaDesdeOtroForm = "" Then
            MsgBox "Seleccione el fichero para importar", vbExclamation
            Exit Sub
        End If
        Unload Me
    End Select
End Sub

Private Sub cmdAsignarLetraCta340_Click()
    If Text1(0).Visible Then
        'LLeva Letra serie
        Cad = "1|" & Trim(Text1(0).Text) & "||"
    Else
        'Lleva ctas
        Cad = "0|" & Trim(Text1(1).Text) & "|" & Trim(Text1(2).Text) & "|"
    End If
    For i = 0 To List2.ListCount - 1
        If List2.Selected(i) Then
            V340(i) = Cad
            Exit Sub
        End If
    Next
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    Unload Me
End Sub


Private Sub Form_Load()
        
    Me.Icon = frmPpal.Icon
    Limpiar Me
    Me.Fr340.Visible = False
    FrameTraspasoBancoSecciones.Visible = False
    FrameLog.Visible = False
    FrameMemoriaPagos.Visible = False
    FrameConsum.Visible = False
    Select Case opcion
    Case 0
        '340
        lbl340.Caption = ""
        Caption = "Generar 340"
        Me.txtFecha(0).Text = Format(Now, "dd/mm/yyyy")
        PonerFrameVisible Fr340
        PonerEmpresaSeleccionEmpresa
        PonerPeriodoPresentacion340
        cmbPeriodo(1).ListIndex = 0
        
        Me.cmdAcciones(0).Enabled = vUsu.Nivel < 3
        
        'Si los tiockets van por letra de serie o van por una cta contable
        Poner340SerieCuentas vParam.TicketsEn340LetraSerie
        ReDim V340(0)
        V340(0) = Abs(vParam.TicketsEn340LetraSerie) & "|||"
        
    Case 1
        Me.Caption = "Traspaso"
        PonerFrameVisible FrameTraspasoBancoSecciones
        Label3(0).Caption = "" 'indicador
        txtFecha(1).Text = Format(vParam.fechaini, "dd/mm/yyyy")
        txtFecha(2).Text = Format(Now - 1, "dd/mm/yyyy")
    Case 2
        Me.Caption = "Imprimir log"
        PonerFrameVisible FrameLog
        CargaListLog
    Case 3
        PonerFrameVisible FrameMemoriaPagos
        Label3(4).Caption = "" 'indicador
        
    Case 4
        PonerFrameVisible Me.FrameConsum
        Me.Caption = "Consum"

    End Select
    
    cmdCancelar(opcion).Cancel = True
End Sub

Private Sub Poner340SerieCuentas(TicketsEnLetraSerie As Boolean)
        Text1(0).Visible = TicketsEnLetraSerie
        Text1(1).Visible = Not TicketsEnLetraSerie
        Text1(2).Visible = Not TicketsEnLetraSerie
        If TicketsEnLetraSerie Then
            Label4(4).Caption = "Letra de serie"
        Else
            Label4(4).Caption = "Cuentas "
        End If
End Sub


Private Sub PonerFrameVisible(kFrame As Frame)
    kFrame.Top = 0
    kFrame.Left = 30
    Me.Height = kFrame.Height + 360
    Me.Width = kFrame.Width + 120
    kFrame.Visible = True
End Sub



Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Cad = CadenaDevuelta
End Sub

Private Sub frmC_Selec(vFecha As Date)
    txtFecha(i).Text = Format(vFecha, "dd/mm/yyyy")
End Sub



Private Sub frmCta_DatoSeleccionado(CadenaSeleccion As String)
    txtCta(i).Text = RecuperaValor(CadenaSeleccion, 1)
    DtxtCta(i).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub Image5_Click()
    'Seleccionar empresa para
    ConsolidadoEmpresas List2
    cmdAsignarLetraCta340.Visible = List2.ListCount > 1
    If List2.ListCount > 1 Then
            ReDim V340(List2.ListCount)
            For i = 0 To List2.ListCount - 1
                Cad = List2.ItemData(i)
                
                'Valor de letra serie
                Cad = DevuelveDesdeBD("TicketsEn340LetraSerie", "conta" & Cad & ".parametros", "1", "1", "N")
                
                V340(i) = Cad & "|||"
                    
            Next
            List2.Selected(0) = True
    Else
        'Probamos a borrar
        ReDim V340(0)
        Cad = List2.ItemData(0)
        Cad = DevuelveDesdeBD("TicketsEn340LetraSerie", "conta" & Cad & ".parametros", "1", "1", "N")
        V340(0) = Cad & "|||"
        Poner340SerieCuentas Cad = "1"
        
    End If
End Sub

Private Sub ImgAyuda_Click(Index As Integer)
  If Index = 0 Then
        Cad = String(40, "*") & vbCrLf
        Cad = Cad & vbCrLf & vbCrLf & "Generará un apunte,con le fecha solicitada, "
        Cad = Cad & vbCrLf & "en la contabilidad de parámetros."
        Cad = Cad & vbCrLf & "Llevará todos las lineas que no esten trasapasadas contra  "
        Cad = Cad & vbCrLf & "un UNICO apunte de la cuenta destino"
        
    End If
    MsgBox Cad, vbInformation
End Sub

Private Sub ImgCd1_Click(Index As Integer)


    cd1.FileName = ""
    cd1.InitDir = "c:\"
    cd1.CancelError = False
    If Index = 0 Then
        cd1.Filter = "RTF (*.rtf)|*.rtf|DAT (*.dat)|*.dat"
        cd1.FilterIndex = 0
    End If
    cd1.ShowOpen
    
    Screen.MousePointer = vbDefault
    If cd1.FileName = "" Then Exit Sub
    
    txtCD1(Index).Text = cd1.FileName
    
End Sub

Private Sub imgCheck_Click(Index As Integer)
    NumRegElim = 0
    If Index > 1 Then NumRegElim = 1
    
    Cad = "0"
    If (Index Mod 2) = 1 Then Cad = "1"
    For i = 1 To Me.ListView1(NumRegElim).ListItems.Count
        ListView1(NumRegElim).ListItems(i).Checked = Cad = "1"
    Next
End Sub

Private Sub imgcta_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    Set frmCta = New frmColCtas
    i = Index
    frmCta.DatosADevolverBusqueda = "0|1"
    frmCta.ConfigurarBalances = 3
    frmCta.Show vbModal
    Set frmCta = Nothing
End Sub



Private Sub ImgDatosDestino_Click(Index As Integer)
    Set frmB = New frmBuscaGrid
    
    Select Case Index
    Case 0
        Cad = "Cuenta|codmacta|T|20·"
        Cad = Cad & "Descripción|nommacta|T|70·"
        frmB.vTabla = "cuentas"
        frmB.vSQL = "apudirec='S'"
    Case 1
        Cad = "Código|codconce|T|20·"
        Cad = Cad & "Descrición|nomconce|T|70·"
        frmB.vTabla = "conceptos"
        frmB.vSQL = "tipoconce=3" 'decide en asiento
    Case 2
        Cad = "Código|numdiari|T|20·"
        Cad = Cad & "Descrición|desdiari|T|70·"
        frmB.vTabla = "tiposdiario"
        frmB.vSQL = ""
    End Select
    frmB.vCampos = Cad
    frmB.vTabla = "conta" & vParam.TraspasCtasBanco & "." & frmB.vTabla
    frmB.vDevuelve = "0|1|"
    frmB.vTitulo = Me.Label3(Index + 1).Caption & " contab. DESTINO"
    frmB.vSelElem = 0
    Cad = ""
    frmB.Show vbModal
    Set frmB = Nothing
    If Cad <> "" Then
        Cad = RecuperaValor(Cad, 1) & " - " & RecuperaValor(Cad, 2)
        txtDatosDestino(Index).Text = Cad
    End If
        
End Sub

Private Sub imgFecha_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    Set frmC = New frmCal
    frmC.Fecha = Now
    If txtFecha(Index).Text <> "" Then frmC.Fecha = CDate(txtFecha(Index).Text)
    i = Index
    frmC.Show vbModal
    Set frmC = Nothing

End Sub





Private Sub List2_Click()
Dim C As String
    If List2.ListCount <= 1 Then Exit Sub
    Screen.MousePointer = vbHourglass
    For i = 0 To List2.ListCount - 1
        If List2.Selected(i) Then
            Cad = V340(i)
            C = RecuperaValor(Cad, 1)
            Poner340SerieCuentas C <> "0"
            If C = "1" Then
                Text1(0).Text = RecuperaValor(Cad, 2)
            Else
                Text1(1).Text = RecuperaValor(Cad, 2)
                Text1(2).Text = RecuperaValor(Cad, 3)
            End If
            
        End If
    Next
    Screen.MousePointer = vbDefault
End Sub



Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_LostFocus(Index As Integer)

    Text1(Index).Text = Trim(Text1(Index).Text)
    
    If Index > 0 Then
        If Text1(Index).Text = "" Then Exit Sub
        CadenaDesdeOtroForm = (Text1(Index).Text)
        If CuentaCorrectaUltimoNivelSIN(CadenaDesdeOtroForm, Cad) Then
            Text1(Index).Text = CadenaDesdeOtroForm
        Else
            MsgBox Cad, vbExclamation
            Text1(Index).Text = ""
        End If
        CadenaDesdeOtroForm = ""
     End If
End Sub

Private Sub txtAno_GotFocus(Index As Integer)
    PonFoco txtAno(Index)
End Sub

Private Sub txtAno_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
     If KeyCode = 112 Then HacerF1
End Sub

Private Sub txtAno_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAno_LostFocus(Index As Integer)
    txtAno(Index).Text = Trim(txtAno(Index))
    If txtAno(Index) = "" Then Exit Sub
    If Not IsNumeric(txtAno(Index)) Then
        MsgBox "Número incorrecto: " & txtAno(Index), vbExclamation
        txtAno(Index).Text = ""
        txtAno(Index).SetFocus
    End If
End Sub

Private Sub txtCta_GotFocus(Index As Integer)
    PonFoco txtCta(Index)
End Sub

Private Sub txtCta_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 112 Then
        HacerF1
    Else
        If KeyCode = 107 Or KeyCode = 187 Then
            KeyCode = 0
            txtCta(Index).Text = ""
            imgcta_Click Index
        End If
    End If
End Sub

Private Sub txtCta_KeyPress(Index As Integer, KeyAscii As Integer)
    
    KEYpress KeyAscii
End Sub

Private Sub txtCta_LostFocus(Index As Integer)
Dim Cta As String
Dim B As Byte
Dim Hasta As Integer   'Cuando en cuenta pongo un desde, para poner el hasta

    txtCta(Index).Text = Trim(txtCta(Index).Text)
    If txtCta(Index).Text = "" Then
        DtxtCta(Index).Text = ""
        Exit Sub
    End If
    
    If Not IsNumeric(txtCta(Index).Text) Then
        If InStr(1, txtCta(Index).Text, "+") = 0 Then MsgBox "La cuenta debe ser numérica: " & txtCta(Index).Text, vbExclamation
        txtCta(Index).Text = ""
        DtxtCta(Index).Text = ""
        Exit Sub
    End If
  
        'DE ULTIMO NIVEL
        Cta = (txtCta(Index).Text)
        If CuentaCorrectaUltimoNivel(Cta, Cad) Then
            txtCta(Index).Text = Cta
            DtxtCta(Index).Text = Cad
        Else
            MsgBox Cad, vbExclamation
            txtCta(Index).Text = ""
            DtxtCta(Index).Text = ""
            txtCta(Index).SetFocus
        End If

End Sub

Private Sub txtFecha_GotFocus(Index As Integer)
    PonFoco txtFecha(Index)
End Sub

Private Sub txtFecha_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then HacerF1
End Sub

Private Sub txtfecha_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtfecha_LostFocus(Index As Integer)
    txtFecha(Index).Text = Trim(txtFecha(Index))
    If txtFecha(Index) = "" Then Exit Sub
    If Not EsFechaOK(txtFecha(Index)) Then
        MsgBox "Fecha incorrecta: " & txtFecha(Index), vbExclamation
        txtFecha(Index).Text = ""
        txtFecha(Index).SetFocus
    End If
End Sub



Private Sub HacerF1()
    Select Case opcion
    Case 0
        
    Case 1
         cmdAcciones_Click 1
         
    End Select
End Sub


Private Sub PonerEmpresaSeleccionEmpresa()

    
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from Usuarios.empresasariconta where codempre < 100", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    NumRegElim = 0
    While Not Rs.EOF
        NumRegElim = NumRegElim + 1
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
    'Si solo hay una empresa deshabilitamos el boton, dependiendo de vOpcion
    Select Case opcion
    Case 0
        Image5.Visible = (NumRegElim > 1)
        List2.AddItem vEmpresa.nomempre
        List2.ItemData(0) = vEmpresa.codempre
    
    End Select
    
        
End Sub



Private Sub ConsolidadoEmpresas(ByRef L As ListBox)
Dim i As Integer
    CadenaDesdeOtroForm = ""
    frmMensajes.opcion = 4
    frmMensajes.Show vbModal
    Text1(0).Text = "": Text1(1).Text = "": Text1(2).Text = ""
    If CadenaDesdeOtroForm <> "" Then
        NumRegElim = RecuperaValor(CadenaDesdeOtroForm, 1)
        L.Clear
        If NumRegElim = 0 Then Exit Sub
        For i = 1 To NumRegElim
            L.AddItem RecuperaValor(CadenaDesdeOtroForm, i + 1)
        Next i
        For i = 0 To NumRegElim - 1
            L.ItemData(i) = RecuperaValor(CadenaDesdeOtroForm, NumRegElim + i + 2)
        Next i
    End If
End Sub



Private Sub PonerPeriodoPresentacion340()

    cmbPeriodo(0).Clear
    If vParam.periodos = 0 Then
        'Liquidacion TRIMESTRAL
        For i = 1 To 4
            If i = 1 Or i = 3 Then
                CadenaDesdeOtroForm = "er"
            Else
                CadenaDesdeOtroForm = "º  "
            End If
            CadenaDesdeOtroForm = i & CadenaDesdeOtroForm & " "
            Me.cmbPeriodo(0).AddItem CadenaDesdeOtroForm & "      periodo"
        Next i
    Else
        'Liquidacion MENSUAL
        For i = 1 To 12
            CadenaDesdeOtroForm = MonthName(i)
            CadenaDesdeOtroForm = UCase(Mid(CadenaDesdeOtroForm, 1, 1)) & LCase(Mid(CadenaDesdeOtroForm, 2))
            Me.cmbPeriodo(0).AddItem CadenaDesdeOtroForm
        Next
    End If
    
    
    'Leeremos ultimo valor liquidaco
    
    txtAno(0).Text = vParam.anofactu
    i = vParam.perfactu + 1
    If vParam.periodos = 0 Then
        NumRegElim = 4
    Else
        NumRegElim = 12
    End If
        
    If i > NumRegElim Then
            i = 1
            txtAno(0).Text = vParam.anofactu + 1
    End If
    Me.cmbPeriodo(0).ListIndex = i - 1
     
    
    CadenaDesdeOtroForm = ""
End Sub




Private Sub Hacer340()
Dim UltimoPeriodoLiquidacion As Boolean
Dim C2 As String
'Dim Tickets As String
    'Comprobaciones
    If List2.ListCount = 0 Then
        MsgBox "Seleccione una empresa", vbExclamation
        Exit Sub
    End If
    If Me.cmbPeriodo(0).ListIndex < 0 Or txtAno(0).Text = "" Then
        MsgBox "Seleccione un periodo/año", vbExclamation
        Exit Sub
    End If
    
    UltimoPeriodoLiquidacion = False
    If cmbPeriodo(0).ListIndex = cmbPeriodo(0).ListCount - 1 Then UltimoPeriodoLiquidacion = True
    
    
    'Tickets
    If List2.ListCount = 1 Then
        Cad = V340(0)
        Cad = RecuperaValor(Cad, 1)
        
        If Cad = "1" Then
            Cad = Cad & "|" & Text1(0).Text & "||"
        Else
            'Tickets = Text1(1).Text & "|" & Text1(2).Text & "|"
            Cad = Cad & "|" & Text1(1).Text & "|" & Text1(2).Text & "|"
        End If
        V340(0) = Cad
        
        
    Else
        
        'Esta consolidadno varias. Vere para cuales NO ha ndicado el valor de letra serie
                i = 0
                Cad = ""
                Do
                    C2 = V340(i)
                    C2 = RecuperaValor(C2, 1)
                    NumRegElim = Val(C2)
                    If NumRegElim = 1 Then
                        'Letra serie
                        C2 = V340(i)
                        C2 = RecuperaValor(C2, 2)
                        If C2 <> "" Then C2 = "Serie " & C2
                    Else
                        'Cuenta
                        C2 = V340(i)
                        C2 = Mid(C2, 3)
                        C2 = Trim(Replace(C2, "|", " "))
                        If C2 <> "" Then C2 = "Ctas: " & C2
                    End If
                    If C2 = "" Then
                        'NO ha indicado nada para esta empresa
                        C2 = Space(15) & "-" & List2.List(i)
                        If Cad <> "" Then Cad = Cad & vbCrLf
                        Cad = Cad & C2
                    End If
                   
                    i = i + 1
                Loop Until i = Me.List2.ListCount
                If Cad <> "" Then
                    Cad = "Las siguientes empresas no se le han asignado valor para los tickets: " & vbCrLf & vbCrLf & Cad
                    Cad = Cad & vbCrLf & vbCrLf & "         ¿Continuar?"
                    If MsgBox(Cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
                End If
        
    End If
    Screen.MousePointer = vbHourglass
    If Modelo340(Me.List2, CInt(txtAno(0).Text), cmbPeriodo(0).ListIndex + 1, Cad, lbl340, False, Me.cmbPeriodo(1).ListIndex = 1, V340(), UltimoPeriodoLiquidacion) Then
        lbl340.Caption = ""
        'AHora podremos o generar fichero, o bien imprimirlo
        If Me.cmbPeriodo(1).ListIndex = 0 Then
            
            'Estableceremos el nombre de empresa
            
                Cad = "Empresas= """
                i = 0
                Do
                    C2 = V340(i)
                    C2 = RecuperaValor(C2, 1)
                    NumRegElim = Val(C2)
                    If NumRegElim = 1 Then
                        'Letra serie
                        C2 = V340(i)
                        C2 = RecuperaValor(C2, 2)
                        If C2 <> "" Then C2 = "Serie " & C2
                    Else
                        'Cuenta
                        C2 = V340(i)
                        C2 = Mid(C2, 3)
                        C2 = Trim(Replace(C2, "|", " "))
                        If C2 <> "" Then C2 = "Ctas: " & C2
                    End If
                    If i > 0 Then Cad = Cad & """ + chr(13) + """
                    Cad = Cad & List2.List(i)
                    If C2 <> "" Then Cad = Cad & " (" & C2 & ")"
                    i = i + 1
                Loop Until i = Me.List2.ListCount
                
            
            Cad = Cad & """|"
            
            'Diciembre 2012. Pongo el peridodo en el rpt
            Cad = Cad & "Periodo= ""Periodo: " & cmbPeriodo(0).ListIndex + 1 & "/" & CInt(txtAno(0).Text) & """|"
            
            'Borrador
            With frmImprimir
                .OtrosParametros = Cad
                .NumeroParametros = 3
                .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
                .SoloImprimir = False
                .opcion = 93
                .Show vbModal
            End With
            
        Else
                                                    
               'Adelante
    
            '¡Anño periodo. Variable que se le pasa al mod340
            '
        Cad = Format(Me.txtAno(0).Text, "0000") & Format(Me.cmbPeriodo(0).ListIndex + 1, "00")
        If vParam.periodos = 0 Then
            'TRIMESTRAL
            Cad = Cad & Me.cmbPeriodo(0).ListIndex + 1 & "T"
        Else
            'MENSUAL
            Cad = Cad & Format(Me.cmbPeriodo(0).ListIndex + 1, "00")
        End If
                                                    
                                                    
                                                    
                                                    'Guardar como
            If GeneraFichero340(True, Cad, False) Then
                'INSERTO EL LOG
                If CuardarComo340 Then InsertaLog340
                    
                    
                
            End If
        End If
    End If
    Screen.MousePointer = vbDefault
    lbl340.Caption = ""
    
    
End Sub

Private Sub InsertaLog340()
Dim C2 As String
    Cad = ""
    i = 0
    Do
        C2 = V340(i)
        C2 = RecuperaValor(C2, 1)
        NumRegElim = Val(C2)
        If NumRegElim = 1 Then
            'Letra serie
            C2 = V340(i)
            C2 = RecuperaValor(C2, 2)
            If C2 <> "" Then C2 = "Serie " & C2
        Else
            'Cuenta
            C2 = V340(i)
            C2 = Mid(C2, 3)
            C2 = Trim(Replace(C2, "|", " "))
            If C2 <> "" Then C2 = "Ctas: " & C2
        End If
        If i > 0 Then Cad = Cad & vbCrLf
        Cad = Cad & List2.List(i)
        If C2 <> "" Then Cad = Cad & " (" & C2 & ")"
        i = i + 1
    Loop Until i = Me.List2.ListCount
    
    vLog.Insertar 15, vUsu, Cad
    
    
    
    'DICIMEBRE 2012
    'Diciembre 2012
    'Pagos en efectivo
    'Para guardarme un LOG de pagos declardaos
    'Ya que si luego modifican un apunte ...  perderiamos datos realmente.
    'ASi, con este log me que declaramos de efectivo
    Cad = Format(Me.txtAno(0).Text, "0000") & "-"
    If vParam.periodos = 0 Then
        'TRIMESTRAL
        Cad = Cad & Me.cmbPeriodo(0).ListIndex + 1 & "T"
    Else
        'MENSUAL
        Cad = Cad & Format(Me.cmbPeriodo(0).ListIndex + 1, "00")
    End If
                       
    Cad = " SELECT  now() fecha, codusu,'" & Cad & "',nifdeclarado,razosoci,fechaexp,base,totiva  "
    Cad = Cad & " FROM usuarios.z340 where codusu =" & vUsu.Codigo & " and clavelibro='Z'"
    
    Cad = "INSERT INTO slog340 " & Cad
    If Not EjecutaSQL(Cad) Then MsgBox "Error insertando LOG. Consulte soporte técnico", vbExclamation
    
    
    
End Sub

Private Function CuardarComo340() As Boolean
    On Error GoTo ECopiarFichero347
    
    CuardarComo340 = False
    cd1.CancelError = True
    cd1.InitDir = Mid(App.Path, 1, 3)
    cd1.ShowSave
        
    Cad = App.Path & "\tmp340.dat"
    
    If cd1.FileTitle <> "" Then
        If Dir(cd1.FileName, vbArchive) <> "" Then
            If MsgBox("Ya existe: " & cd1.FileName & vbCrLf & "¿Sobreescribir?", vbQuestion + vbYesNo) = vbNo Then Exit Function
        End If
        FileCopy Cad, cd1.FileName
        MsgBox Space(20) & "Copia efectuada correctamente" & Space(20), vbInformation
        CuardarComo340 = True
    End If
    Exit Function
ECopiarFichero347:
    If Err.Number <> 32755 Then MuestraError Err.Number, "Copiar fichero"
    
End Function



'---------------------------------------------------------------------
'----------------------------------------------------------------------
Private Sub TraspasoEntreCtas()
Dim F As Date
Dim B  As Boolean
Dim AUx As String

    'Comprobaciones
    Cad = ""
    If Me.txtCta(0).Text = "" Then Cad = Cad & vbCrLf & "  -Cuenta origen"
    If Me.txtFecha(1).Text = "" Then Cad = Cad & vbCrLf & "  -Fecha inicio"
    If Me.txtFecha(2).Text = "" Then Cad = Cad & vbCrLf & "  -Fecha fin"
    For i = 1 To 3
        If Me.txtDatosDestino(i - 1) = "" Then Cad = Cad & vbCrLf & "  -" & Me.Label3(i).Caption & " destino"
    Next i
   
    If Cad <> "" Then
        MsgBox "Campos requeridos: " & vbCrLf & Cad, vbExclamation
        Exit Sub
    End If
    
    
    
    'Vamos a ver si existe la cuenta
    i = InStr(1, txtDatosDestino(0).Text, "-")
    If i = 0 Then Cad = Cad & vbCrLf & "-Error obteniendo cta destino"""
    If Cad <> "" Then
        MsgBox Cad, vbExclamation
        Exit Sub
    End If
    
    Cad = Trim(Mid(txtDatosDestino(0).Text, 1, i - 1))
    
    'Veamos si existe la cuenta contable en destino
    AUx = DevuelveDesdeBD("codmacta", "conta" & vParam.TraspasCtasBanco & ".cuentas", "apudirec='S' AND codmacta ", txtCta(0).Text, "T")
    If AUx = "" Then
        AUx = "Cuenta no origen no existe en contabilidad destino"
    Else
        If AUx = Cad Then
            AUx = "Cuenta origen y destino es la misma"
        Else
            AUx = ""
        End If
    End If
    If AUx <> "" Then
        MsgBox AUx, vbExclamation
        Exit Sub
    End If
    
    
    
    
    
    Cad = DevuelveDesdeBD("concat(if(fecbloq is null,'',fecbloq) ,""|"",codmacta,""|"")", "conta" & vParam.TraspasCtasBanco & ".cuentas", "apudirec='S' AND codmacta ", Cad, "T")
    If Cad = "" Then
        Cad = "Cuenta destino NO encontrada"
    Else
    
        'Mismo ultmonivel
        If Len(RecuperaValor(Cad, 2)) <> Len(Me.txtCta(0).Text) Then
            Cad = "Distinto longitud para las cuentas de ultimo nivel"
    
        Else
            'Bloqueo
            Cad = RecuperaValor(Cad, 1)
            If Cad <> "" Then
                F = CDate(Cad)
                If CDate(Me.txtFecha(3).Text) >= F Then
                    'Cad = "Cuenta bloqueada desde " & F
                Else
                    Cad = ""
                End If
                
            End If
        End If
    End If
    If Cad <> "" Then
        MsgBox Cad, vbExclamation
        Exit Sub
    End If
    
    
    Cad = DevuelveDesdeBD("fechaini", "conta" & vParam.TraspasCtasBanco & ".parametros", "1", "1")
    If Cad = "" Then
        Cad = "Error obteniendo ejercicios empresa destino"
    Else
        F = CDate(Cad)
        If CDate(Me.txtFecha(1).Text) < F Then
            Cad = "Fecha ejercicios cerrados"
        Else
            F = DateAdd("yyyy", 1, F)
            F = DateAdd("d", -1, F)
           
                
            F = DateAdd("yyyy", 1, F)
            If CDate(Me.txtFecha(2).Text) > F Then
                Cad = "Fechas fuera de ejercicio"
            Else
                Cad = "" 'Para que siga
            End If
        End If
    End If
    If Cad <> "" Then
        MsgBox Cad, vbExclamation
        Exit Sub
    End If



    'OK
    'Llegados aqui es que podemos traspasar
    Set miRsAux = New ADODB.Recordset
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""


    B = HacerTraspasoCuentas()
    
    'reestablezco
    txtDatosDestino(0).Tag = ""
    Me.txtDatosDestino(1).Tag = 0
    Me.txtDatosDestino(2).Tag = 0
    


    If B Then
        i = 1
        Do
            Label3(0).Caption = "Guardando proceso " & i
            Label3(0).Refresh
            
            If Len(CadenaDesdeOtroForm) > 220 Then
                NumRegElim = InStr(185, CadenaDesdeOtroForm, vbCrLf)
                If NumRegElim > 0 Then
                    Cad = Mid(CadenaDesdeOtroForm, 1, NumRegElim - 1)
                    CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, NumRegElim + 2)
                Else
                    Cad = CadenaDesdeOtroForm
                    CadenaDesdeOtroForm = ""
                End If
            Else
                Cad = CadenaDesdeOtroForm
                CadenaDesdeOtroForm = ""
            End If
            Cad = "[TR.BANCO] " & i & vbCrLf & Cad
            vLog.Insertar 19, vUsu, Cad
            If CadenaDesdeOtroForm <> "" Then espera 1
            i = i + 1
        Loop Until CadenaDesdeOtroForm = ""
        MsgBox "Proceso finalizado con exito", vbInformation
        Unload Me
    End If
    CadenaDesdeOtroForm = ""
    Label3(0).Caption = ""
    Screen.MousePointer = vbDefault
End Sub


Private Function HacerTraspasoCuentas() As Boolean
Dim F As Date

    On Error GoTo eCadenaDesdeOtroForm

    HacerTraspasoCuentas = False
    Label3(0).Caption = "Obteniendo registros"
    Label3(0).Refresh
    
    Cad = "from hlinapu where traspasado=0 "
    Cad = Cad & " AND codmacta = '" & Me.txtCta(0).Text & "'"
    If Me.txtFecha(1).Text <> "" Then Cad = Cad & " AND fechaent >= '" & Format(Me.txtFecha(1).Text, FormatoFecha) & "'"
    If Me.txtFecha(2).Text <> "" Then Cad = Cad & " AND fechaent <= '" & Format(Me.txtFecha(2).Text, FormatoFecha) & "'"
    
    Set Rs = New ADODB.Recordset
    
    Rs.Open "Select fechaent,count(*) " & Cad & " GROUP BY fechaent", Conn, adOpenKeyset, adLockPessimistic, adCmdText
    If Rs.EOF Then
    
        MsgBox "Ningun dato a traspasar", vbExclamation
    Else
        i = 0
        Cad = ""
        While Not Rs.EOF
            Cad = Cad & "X"
           i = i + DBLet(Rs.Fields(1), "N")
           Rs.MoveNext
        Wend
        Rs.MoveFirst
        Cad = Len(Cad)
        Cad = "Apuntes: " & Cad & ".     "
        Cad = Cad & i & " lineas para traspasar.  ¿Continuar?"
        If MsgBox(Cad, vbQuestion + vbYesNoCancel) = vbYes Then
        
            Cad = DevuelveDesdeBD("fechafin", "conta" & vParam.TraspasCtasBanco & ".parametros", "1", "1")
            F = CDate(Cad)
        
            While Not Rs.EOF
                Label3(0).Caption = "Fecha: " & Rs.Fields(0) & " - Lin: " & Rs.Fields(1)
                Label3(0).Refresh
                DoEvents
                espera 0.2
                HacerTraspasoCuentas2 Rs.Fields(0), Rs.Fields(1), Rs.Fields(0) <= F
                
                Rs.MoveNext
            Wend
            HacerTraspasoCuentas = True
        End If
    End If
    Rs.Close
 
    
    
eCadenaDesdeOtroForm:

    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set Rs = Nothing
    Label3(0).Caption = ""
End Function

Private Function HacerTraspasoCuentas2(Fecha As Date, Cuantos As Integer, EnActual As Boolean) As Boolean
Dim Importe As Currency
Dim AUx As String
Dim ApunteYContraApunte As Boolean  'Para cada linea de apunte generara un contrapartida.
                                    'De momento siempre TRUE

    ApunteYContraApunte = True

    'Para las observaciones
    AUx = "Agrupacion cuenta banco desde " & vEmpresa.nomresum
    AUx = AUx & "   Fecha: " & Now & vbCrLf & vUsu.Nombre
    AUx = AUx & "      Lineas: " & Cuantos
    
    'Ok vamos p'alla.
    '1,. Ponermos en el tag de concepto, y de diario el NUMERO
    'DIARIO
    i = InStr(1, Me.txtDatosDestino(2).Text, "-")
    Me.txtDatosDestino(2).Tag = Val(Mid(Me.txtDatosDestino(2), 1, i - 1)) 'i no sera 0
    'Concepto
    i = InStr(1, Me.txtDatosDestino(1).Text, "-")
    Me.txtDatosDestino(1).Tag = Val(Mid(Me.txtDatosDestino(1), 1, i - 1)) 'i no sera 0
    'Contrapartida
    i = InStr(1, Me.txtDatosDestino(0).Text, "-")
    Me.txtDatosDestino(0).Tag = Trim(Mid(Me.txtDatosDestino(0), 1, i - 1)) 'i no sera 0
    
    'Vamos a por el contador
    If EnActual Then
        Cad = "1"
    Else
        Cad = "2"
    End If
    Cad = "Select contado" & Cad & " FROM conta" & vParam.TraspasCtasBanco & ".contadores WHERE tiporegi ='0'" 'apuntes
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'NO puede ser EOF ni NULL
    NumRegElim = miRsAux.Fields(0) + 1
    miRsAux.Close
    
    
    If EnActual Then
        Cad = "1"
    Else
        Cad = "2"
    End If
    Cad = "UPDATE  conta" & vParam.TraspasCtasBanco & ".contadores SET contado" & Cad
    Cad = Cad & "= " & NumRegElim & " WHERE  tiporegi ='0'"
    Conn.Execute Cad
    
    
    
    
    'Ahora crearemos la cabecera de apunte
    Cad = "INSERT INTO conta" & vParam.TraspasCtasBanco & ".cabapu(numdiari,numasien,fechaent,bloqactu,obsdiari) VALUES ("
    Cad = Cad & txtDatosDestino(2).Tag & "," & NumRegElim & ",'" & Format(Fecha, FormatoFecha)
    Cad = Cad & "',1,'" & DevNombreSQL(AUx) & "')" 'bloqueado y observaciones
    Conn.Execute Cad
    
    'Para el LOG
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & Fecha & " " & Format(NumRegElim, "0000") & "  " & Format(Cuantos, "00")
    
    'Lineas de apuntes
    ReDim V340(0)
    Cad = "Select numdocum,ampconce,timporteD,timporteH,idcontab,ctacontr,nommacta,hlinapu.codmacta"
    Cad = Cad & "  from hlinapu left join cuentas on  hlinapu.ctacontr=cuentas.codmacta "
    Cad = Cad & " where  traspasado=0 AND hlinapu.codmacta = '" & Me.txtCta(0).Text & "'"
    
    Cad = Cad & " AND fechaent = '" & Format(Fecha, FormatoFecha) & "'"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    
    i = 1
    If ApunteYContraApunte Then i = 0 'para que empiece en el 1
    
    Importe = 0
    While Not miRsAux.EOF
        i = i + 1
        'numdiari,fechaent,numasien,linliapu,codmacta,numdocum,codconce,ampconce,
        'timporteD,timporteH,codccost,ctacontr,idcontab,punteada
        V340(0) = miRsAux!Ampconce
        If Len(miRsAux!Ampconce) < 25 Then
            If Not IsNull(miRsAux!nommacta) Then V340(0) = Mid(V340(0) & " " & miRsAux!nommacta, 1, 30)
        End If
        V340(0) = ",'" & DevNombreSQL(CStr(V340(0))) & "'"
        'codconce,amconce
        V340(0) = "," & txtDatosDestino(1).Tag & V340(0)
        V340(0) = i & ",'" & miRsAux!codmacta & "','" & DevNombreSQL(DBLet(miRsAux!Numdocum, "T")) & "'" & V340(0) & ","
        'Meto, por delante, numdiari,fechaent, numasien
        V340(0) = txtDatosDestino(2).Tag & ",'" & Format(Fecha, FormatoFecha) & "'," & NumRegElim & "," & V340(0)
        If IsNull(miRsAux!timported) Then
            Importe = Importe - DBLet(miRsAux!timporteH)
            V340(0) = V340(0) & "NULL," & TransformaComasPuntos(CStr(DBLet(miRsAux!timporteH, "N")))
        Else
            Importe = Importe + DBLet(miRsAux!timported)
            V340(0) = V340(0) & TransformaComasPuntos(CStr(miRsAux!timported)) & ",NULL"
        End If
        V340(0) = V340(0) & ",NULL,'" & txtDatosDestino(0).Tag & "','contab',0)"
        Cad = Cad & ", (" & V340(0)
        
        If ApunteYContraApunte Then
            'Generamos la contrapartida
            i = i + 1
            V340(0) = DevNombreSQL(DBLet(miRsAux!Ampconce, "T"))
            V340(0) = ",'" & V340(0) & "'"
            
            'codconce,amconce
            V340(0) = "," & txtDatosDestino(1).Tag & V340(0)
            V340(0) = i & ",'" & txtDatosDestino(0).Tag & "','" & DevNombreSQL(DBLet(miRsAux!Numdocum, "T")) & "'" & V340(0) & ","
            'Meto, por delante, numdiari,fechaent, numasien
            V340(0) = txtDatosDestino(2).Tag & ",'" & Format(Fecha, FormatoFecha) & "'," & NumRegElim & "," & V340(0)
            If Importe > 0 Then
                V340(0) = V340(0) & "NULL," & TransformaComasPuntos(CStr(Importe))
            Else
                Importe = Abs(Importe)
                V340(0) = V340(0) & TransformaComasPuntos(CStr(Importe)) & ",NULL"
            End If
            V340(0) = V340(0) & ",NULL,'" & miRsAux!codmacta & "','contab',0)"
            Cad = Cad & ", (" & V340(0)
            
            Importe = 0  'para que no lo acumule
        End If
        
        miRsAux.MoveNext
        
       
    Wend
    miRsAux.Close
    
    
    
     If V340(0) <> "" Then
            Label3(0).Caption = "Insertando datos " & Fecha
            Label3(0).Refresh
            
            Cad = Mid(Cad, 2)
            V340(0) = "INSERT INTO conta" & vParam.TraspasCtasBanco & ".linapu("
            V340(0) = V340(0) & "numdiari,fechaent,numasien,linliapu,codmacta,numdocum,codconce,ampconce,timporteD,timporteH,codccost,ctacontr,idcontab,punteada) VALUES "
            Cad = V340(0) & Cad
            Conn.Execute Cad
            Cad = ""
        End If
    
    
    'Metemos el importe del totoal
    ''El total. Sra la linea 1 (la he reservado)
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & vbCrLf
  
    If Not ApunteYContraApunte Then
        'QUiere decir que al final de todo meteremos el apunte contrapartida de todos
        If Importe <> 0 Then
            i = 1
            V340(0) = Mid("Trasp. cta. banco " & vEmpresa.nomresum, 1, 30)
            V340(0) = ",'" & V340(0) & "'"
            
            'codconce,amconce
            V340(0) = "," & txtDatosDestino(1).Tag & V340(0)
            V340(0) = i & ",'" & txtDatosDestino(0).Tag & "','TR.BANCO'" & V340(0) & ","
            'Meto, por delante, numdiari,fechaent, numasien
            V340(0) = txtDatosDestino(2).Tag & ",'" & Format(Fecha, FormatoFecha) & "'," & NumRegElim & "," & V340(0)
            If Importe > 0 Then
                V340(0) = V340(0) & "NULL," & TransformaComasPuntos(CStr(Importe))
            Else
                Importe = Abs(Importe)
                V340(0) = V340(0) & TransformaComasPuntos(CStr(Importe)) & ",NULL"
            End If
            V340(0) = V340(0) & ",NULL,NULL,'contab',0)"
            Cad = Cad & ", (" & V340(0)
            
        
        End If
    End If
        
    If Cad <> "" Then
            Cad = Mid(Cad, 2)
            V340(0) = "INSERT INTO conta" & vParam.TraspasCtasBanco & ".linapu("
            V340(0) = V340(0) & "numdiari,fechaent,numasien,linliapu,codmacta,numdocum,codconce,ampconce,timporteD,timporteH,codccost,ctacontr,idcontab,punteada) VALUES "
            Cad = V340(0) & Cad
            Conn.Execute Cad
    End If
    
    
    Cad = "UPDATE hlinapu set traspasado=1  where traspasado=0 "
    Cad = Cad & " AND codmacta = '" & Me.txtCta(0).Text & "'"
    Cad = Cad & " AND fechaent = '" & Format(Fecha, FormatoFecha) & "'"
    Conn.Execute Cad
    
    
    Cad = "UPDATE conta" & vParam.TraspasCtasBanco & ".cabapu SET bloqactu =0 "
    Cad = Cad & " WHERE numdiari = " & txtDatosDestino(2).Tag
    Cad = Cad & " AND numasien = " & NumRegElim & " AND fechaent ='" & Format(Fecha, FormatoFecha) & "'"
    
    Conn.Execute Cad
    
    HacerTraspasoCuentas2 = True
    
    
    
End Function


Private Sub CargaListLog()
Dim IT As ListItem
 'CadenaConsulta = "select slog.fecha,titulo,usuario,pc,descripcion from slog,tmppresu1 "
  '  CadenaConsulta = CadenaConsulta & " where tmppresu1.codusu=" & vUsu.Codigo & " and slog.accion=tmppresu1.codigo"
   

    Set miRsAux = New ADODB.Recordset
    Cad = "Select * from tmppresu1 where tmppresu1.codusu=" & vUsu.Codigo & " ORDER BY codigo"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = ListView1(0).ListItems.Add(, CStr("K" & miRsAux!Codigo), miRsAux!Titulo)
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    Cad = "Select usuario from slog group by 1 order by 1"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = ListView1(1).ListItems.Add(, miRsAux!Usuario, miRsAux!Usuario)
            
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
End Sub




Private Sub ListadoLog()
    ReDim V340(5)
    Cad = ""
    i = 0
    CadenaDesdeOtroForm = ""
    For NumRegElim = 1 To Me.ListView1(0).ListItems.Count
        If Me.ListView1(0).ListItems(NumRegElim).Checked Then
            i = i + 1
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & ", " & Me.ListView1(0).ListItems(NumRegElim).Text
            V340(3) = V340(3) & "," & Mid(ListView1(0).ListItems(NumRegElim).Key, 2)
        End If
    Next NumRegElim
    If i = 0 Then Cad = " - Accion"
    If i = Me.ListView1(0).ListItems.Count Then
        'VAN TODOS. No pongo ninguno
        CadenaDesdeOtroForm = ""
        V340(3) = ""
    Else
        'QUito la primeracom
        CadenaDesdeOtroForm = "Acciones: " & Trim(Mid(CadenaDesdeOtroForm, 2))
        V340(3) = Mid(V340(3), 2)
    End If
    V340(0) = CadenaDesdeOtroForm
    
    i = 0
    For NumRegElim = 1 To Me.ListView1(1).ListItems.Count
        If Me.ListView1(1).ListItems(NumRegElim).Checked Then
            i = i + 1
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & ", " & Me.ListView1(1).ListItems(NumRegElim).Text
            V340(4) = V340(4) & ",'" & DevNombreSQL(ListView1(1).ListItems(NumRegElim).Text) & "'"
        End If
    Next NumRegElim
    If i = 0 Then Cad = Cad & vbCrLf & " - Trabajador"
    If i = Me.ListView1(1).ListItems.Count Then
        'VAN TODOS. No pongo ninguno
        CadenaDesdeOtroForm = ""
        V340(4) = ""
    Else
        'QUito la primeracom
        CadenaDesdeOtroForm = "Trabajadores: " & Trim(Mid(CadenaDesdeOtroForm, 2))
        V340(4) = Mid(V340(4), 2)
    End If
    V340(1) = CadenaDesdeOtroForm
    If Cad <> "" Then
        Cad = "Debe seleccionar: " & vbCrLf & Cad
        MsgBox Cad, vbExclamation
        Exit Sub
    End If
    
    
    Cad = ""
    If txtFecha(3).Text <> "" Then Cad = "desde " & txtFecha(3).Text
    If txtFecha(3).Text <> "" Then Cad = Cad & "      hasta " & txtFecha(4).Text
    If Cad <> "" Then Cad = "Fechas: " & Trim(Cad)
    V340(5) = Cad
        
    

    
    Screen.MousePointer = vbHourglass
    If ListadoAcciones Then
            'Borrador
            Cad = ""
            Cad = V340(0)
            If V340(5) <> "" Then
                If Cad <> "" Then Cad = Cad & "     "
                Cad = Cad & V340(5)
            End If
            Cad = "pdh1= """ + Cad + """|"
            
            Cad = Cad & "pdh2= """ + V340(1) + """|"
            Cad = Cad & "Emp= """ & vEmpresa.nomempre & """|"
            i = 100
            If Me.optLog(1).Value Then i = 101
                
            With frmImprimir
                .OtrosParametros = Cad
                .NumeroParametros = 4
                .FormulaSeleccion = "{zpendientes.codusu}=" & vUsu.Codigo
                .SoloImprimir = False
                .opcion = i
                .Show vbModal
            End With
    
    End If
    Screen.MousePointer = vbDefault
End Sub



Private Function ListadoAcciones() As Boolean
    
    On Error GoTo eListadoAcciones
    
    ListadoAcciones = False
    
    Conn.Execute "Delete from usuarios.zpendientes WHERE codusu = " & vUsu.Codigo
    
    '               secuenc  fecha      accion   trab      desc
    'z347carta(codusu,nif,otralineadir,razosoci,dirdatos,parrafo1)
    
    Cad = "select " & vUsu.Codigo & ",'A',right(concat(""0000000"",@rownum:=@rownum+1),7), slog.fecha,"
    Cad = Cad & "date_format(slog.fecha,'%d/%m/%Y %H:%i:%s'),titulo,usuario,descripcion"
    Cad = Cad & " from slog,tmppresu1,(SELECT @rownum:=0) r "
    Cad = Cad & " where tmppresu1.codusu=" & vUsu.Codigo & " and slog.accion=tmppresu1.codigo"
    If V340(3) <> "" Then Cad = Cad & " AND slog.accion IN (" & V340(3) & ")"
    If V340(4) <> "" Then Cad = Cad & " AND usuario IN (" & V340(4) & ")"
    If txtFecha(3).Text <> "" Then Cad = Cad & " AND  slog.fecha >= '" & Format(txtFecha(3).Text, FormatoFecha) & " 00:00:00'"
    If txtFecha(4).Text <> "" Then Cad = Cad & " AND  slog.fecha <= '" & Format(txtFecha(4).Text, FormatoFecha) & " 23:59:59'"
    
    Cad = "INSERT INTO usuarios.zpendientes(codusu,serie_cta,factura,fecha,nomforpa,Situacion,nombre,observa) " & Cad
    Conn.Execute Cad
    
    NumRegElim = 0
    Cad = DevuelveDesdeBD("count(*)", "Usuarios.zpendientes", "codusu", CStr(vUsu.Codigo))
    If Cad <> "" Then NumRegElim = Val(Cad)
    
    ListadoAcciones = NumRegElim > 0
    If NumRegElim = 0 Then MsgBox "Ningun dato devuelto", vbExclamation
    
    Exit Function
eListadoAcciones:
    MuestraError Err.Number, Err.Description
End Function


'**************************************


Private Sub MemoriaPagosProv()
    
    Cad = ""
    If Me.txtFecha(6).Text = "" Then Cad = "6"
    If Me.txtFecha(5).Text = "" Then Cad = "5"
    
    If Cad = "" Then
        If Year(CDate(Me.txtFecha(5).Text)) <> Year(CDate(Me.txtFecha(6).Text)) Then
            Cad = "Debe ser mismo año"
            
        Else
            If CDate(Me.txtFecha(5).Text) > CDate(Me.txtFecha(6).Text) Then Cad = "Fecha fin mayor que fecha inicio"
        End If
        If Cad <> "" Then PonleFoco txtFecha(6)
        
        If Me.txtAno(1).Text = "" Then
            Cad = Cad & vbCrLf & " -Dias plazo obligado"
            PonleFoco txtAno(1)
        End If
    Else
        PonleFoco txtFecha(CInt(Cad))
        Cad = "Fechas obligatorias"
        
    End If
    
    If Cad <> "" Then
        MsgBox Cad, vbExclamation
        Exit Sub
    End If
    
    
    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset
    Set Rs = New ADODB.Recordset
    
    ProcesoMemoriaPagosProveedor 0
    ProcesoMemoriaPagosProveedor 1 'Anterior
    
    Label3(4).Caption = ""
    Label3(4).Refresh
    
    If NumRegElim > 0 Then
            Cad = "ph1= ""Dias: " & txtAno(1).Text & "        Fechas: " & txtFecha(5).Text & " - " & txtFecha(6).Text & """|"
            
            'Borrador
            With frmImprimir
                .OtrosParametros = Cad
                .NumeroParametros = 2
                .FormulaSeleccion = "{z347.codusu}=" & vUsu.Codigo
                .SoloImprimir = False
                .opcion = 102
                .Show vbModal
            End With
   End If
    
    
    Set miRsAux = Nothing
    Set Rs = Nothing
    Screen.MousePointer = vbDefault
End Sub


'NumProceso.    0 - Actual
'               1 - Anterior

Private Function ProcesoMemoriaPagosProveedor(NumProceso As Byte) As Boolean
Dim ConceptosProvee As String
Dim F1 As Date
Dim F2 As Date
Dim J As Byte
Dim CtaTratada As String
Dim DatosGenerados As Boolean
Dim Importe As Currency
Dim AbiertoRS As Boolean
Dim TieneDatosRs As Boolean
    On Error GoTo eProcesoMemoriaPagosProveedor
    ProcesoMemoriaPagosProveedor = False
    DatosGenerados = False
    NumRegElim = 0
    Label3(4).Caption = "Obteniedo facturas [" & NumProceso & "]"
    Label3(4).Refresh
    
    If NumProceso = 0 Then
        Conn.Execute "Delete from tmp340 WHERE codusu = " & vUsu.Codigo
        Conn.Execute "Delete from usuarios.z340 WHERE codusu = " & vUsu.Codigo
        Conn.Execute "Delete from usuarios.z347 WHERE codusu = " & vUsu.Codigo
        'Para el saldo final
        Conn.Execute "DELETE from tmpconext where codusu= " & vUsu.Codigo
        
        
        
        
        
    End If
    Conn.Execute "Delete from tmpconext WHERE codusu = " & vUsu.Codigo
    
    ConceptosProvee = DevuelveDesdeBD("max(codigo)", "tmp340", "codusu", CStr(vUsu.Codigo))
    If ConceptosProvee = "" Then ConceptosProvee = "0"
    'Para
    F1 = CDate(txtFecha(5).Text)
    F2 = CDate(txtFecha(6).Text)
    
    If NumProceso = 1 Then
        F1 = DateAdd("yyyy", -1, F1)
        F2 = DateAdd("yyyy", -1, F2)
    End If
    Cad = " WHERE fecfacpr >= '" & Format(F1, FormatoFecha) & "'"
    Cad = Cad & " AND fecfacpr <= '" & Format(F2, FormatoFecha) & "'"
    Cad = Cad & " AND totfacpr >0"
    
    Cad = "Select " & vUsu.Codigo & ",@rownum:=@rownum+1 AS rownum," & Year(F1) & ",numregis,fecfacpr,fecrecpr,numfacpr,codmacta,0,0,totfacpr from cabfactprov,(SELECT @rownum:=" & ConceptosProvee & ") r  " & Cad
    Cad = "insert into tmp340(codusu,codigo,nifdeclarado ,nifrepresante,fechaexp,fechaop,rectifica,cp_intracom,numiva,totiva ,totalfac) " & Cad
    Conn.Execute Cad
    
    
    Cad = ""
    Cad = DevuelveDesdeBD("count(*)", "tmp340", "codusu", CStr(vUsu.Codigo))
    
    If Cad = "0" Or Cad = "" Then
        ProcesoMemoriaPagosProveedor = True 'Ha ido bien
        MsgBox "Ningun factura para ver el pago[" & NumProceso & "]", vbExclamation
        Exit Function
    End If
    
    
    
    'NUestro pago va al debe, con lo cual cogere todos los conceptos de la tesoreria
    ConceptosProvee = ""
    If chkMemoria.Value = 0 Then
        If vEmpresa.TieneTesoreria Then
            Cad = "select distinct(condepro) from stipoformapago"
            miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not miRsAux.EOF Then
                While Not miRsAux.EOF
                    ConceptosProvee = ConceptosProvee & ", " & miRsAux.Fields(0)
                    miRsAux.MoveNext
                Wend
                ConceptosProvee = Mid(ConceptosProvee, 2)
                ConceptosProvee = " codconce IN (" & ConceptosProvee & ")"
                
            End If
            miRsAux.Close
            
            
        End If
    End If
    If ConceptosProvee = "" Then ConceptosProvee = " codconce between 0 and 899"

     DoEvents
    
    
    
    
    'Ahora vamos a cruzar las facturas con la hlinapu buscando el pago de la factura
    'Habra dos procesos, 1 que busccara en funcion de numregis y otro en funcion de numfac
    'debeido al parametrp codinume
    
    'HAremos 3 pasadas
    'Primera buscaremos por el concepto en parametros(numregis o factura)
    '
    'y en la segunda buscaremos si en la ampliacion pone la factura
    
    AbiertoRS = False
    CtaTratada = ""
    For J = 1 To 3
        Label3(4).Caption = "[" & NumProceso & "] Cuadrando datos (" & J & " de 3)"
        Label3(4).Refresh
        DoEvents
    
        Cad = "Select tmp340.*,nommacta from tmp340,cuentas where codusu = " & vUsu.Codigo & " AND numiva =0" 'Por saber si se ha encontrado o no
        Cad = Cad & " AND nifdeclarado='" & Year(F1) & "' AND tmp340.cp_intracom =cuentas.codmacta  ORDER BY cp_intracom,Rectifica"
        
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        CtaTratada = ""
        
        While Not miRsAux.EOF
            
            
            
            If CtaTratada <> miRsAux!cp_intracom Then
                Label3(4).Caption = "Apuntes para " & miRsAux!nommacta
                Label3(4).Refresh
                DoEvents
                
                If AbiertoRS Then Rs.Close
                
                Cad = "SELECT codmacta,numdiari,fechaent,numasien,linliapu,numdocum,ampconce,coalesce(timported,0) importe FROM hlinapu "
                Cad = Cad & " WHERE fechaent >= '" & Format(F1, FormatoFecha) & "'"
                'Cad = Cad & " AND fechaent <= '" & Format(DateAdd("d", Val(Me.txtAno(1).Text), F2), FormatoFecha) & "'"
                Cad = Cad & " AND idcontab<>'FRAPRO' and " & ConceptosProvee
                Cad = Cad & " AND codmacta = '" & miRsAux!cp_intracom & "' and " & ConceptosProvee
    
                Rs.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
                CtaTratada = miRsAux!cp_intracom
                AbiertoRS = True
                TieneDatosRs = Not Rs.EOF
            
            
                'If CtaTratada = "400008001" And J = 3 Then Stop
            Else
                If TieneDatosRs Then Rs.MoveFirst
            
            End If
            
            Label3(4).Caption = NumProceso & " - (" & J & "/3) " & miRsAux!nommacta
            Label3(4).Refresh
            
            Cad = ""
            i = Abs(Rs.EOF)
            While i = 0
                
                
                If J = 1 Then
                    'Primera pasada
                    If vParam.CodiNume = 2 Then
                    
         
                        'Numero de factura
                        If DBLet(miRsAux!rectifica, "T") <> "" Then
                            If DBLet(miRsAux!rectifica, "T") = DBLet(Rs!Numdocum, "T") Then Cad = "OK"
                        End If
                    
                    Else
                        '****************************************
                        'registro
                        If IsNumeric(DBLet(Rs!Numdocum, "T")) Then
                             'nifrepre tiene nnumregis
                             If PonerValorBDnumregis = DBLet(Rs!Numdocum, "T") Then Cad = "OK"
                        End If
                         
                    End If
                                               
                ElseIf J = 2 Then
                    'Por si acaso fueran cambiados.
                    'Es decir en contabilidad van por numero de registro y en tesoreria va por numero de factura
                    If vParam.CodiNume = 1 Then  'va al reves que arriba
                        'Numero de factura
                        If DBLet(miRsAux!rectifica, "T") <> "" Then
                            If DBLet(miRsAux!rectifica, "T") = DBLet(Rs!Numdocum, "T") Then Cad = "OK"
                        End If
                    
                    
                    End If
                          
                Else
                    'A la desesperada
                    If DBLet(miRsAux!rectifica, "T") <> "" Then
                            If InStr(1, DBLet(Rs!Ampconce, "T"), DBLet(miRsAux!rectifica, "T")) > 0 Then
                                Cad = "OK"
                                
                            End If
                    End If
                End If
                                               
                If Cad <> "" Then
                        'OK. Aqui lo tenemos el pago
                        i = DateDiff("d", CDate(miRsAux!fechaexp), CDate(Rs!FechaEnt))
                        
                        Cad = "UPDATE tmp340 SET numiva=" & J
                        Cad = Cad & ", base=" & i
                        If Not IsNull(Rs!Importe) Then Cad = Cad & ", totiva=totiva + " & TransformaComasPuntos(CStr(Rs!Importe))
                        Cad = Cad & " WHERE codusu = " & vUsu.Codigo & " AND codigo = " & miRsAux!Codigo
                        Cad = Cad & " AND cp_intracom = '" & miRsAux!cp_intracom & "' AND rectifica = '" & DevNombreSQL(miRsAux!rectifica) & "'"
                        Conn.Execute Cad
                        i = 1
                  
                End If
                If i = 0 Then
                    Rs.MoveNext
                    i = Abs(Rs.EOF)
                    
                End If
            Wend
            
            
            miRsAux.MoveNext 'sig factura
        Wend
        miRsAux.Close
        If AbiertoRS Then
            Rs.Close
            AbiertoRS = False
        End If

        
    Next J  'siguiente pasada
    
   
    '
    
    
    
    
    '
    
    'Para los pagos que me queden por saldar,
    'vere el saldo de la cuenta a final del periodo
    If NumProceso = 1 Then
        J = 0
        If Year(vParam.fechaini) <> Year(vParam.fechafin) Then
            J = 1
            'Años partidos NO esta todavia
        Else
            F1 = CDate(txtFecha(5).Text)
            F2 = DateAdd("yyyy", 1, F1)
            F2 = DateAdd("d", -1, F2)
            
            If Format(F2, "dd/mm/yyyy") <> txtFecha(6).Text Then
                'No abarca un ejercicio. No haremos el cuadre de saldos
                J = 1
            End If
            
        End If
        
        If J = 0 Then
            'Haremos el cuadre de pagos por saldo.
            Label3(4).Caption = "Cuadre por saldo"
            Label3(4).Refresh
            
            Cad = "insert into tmpconext(codusu,numdiari,cta,saldo)"
            Cad = Cad & "Select codusu,nifdeclarado,cp_intracom,0 from tmp340 where codusu = " & vUsu.Codigo & " AND numiva =0 group by nifdeclarado,cp_intracom"
            Conn.Execute Cad
            
            F1 = CDate(txtFecha(5).Text)
            NumRegElim = Year(F1)
            For J = 0 To 1
                Label3(4).Caption = "Obtener saldos " & J
                Label3(4).Refresh
               
               
                Cad = "Select codmacta ,sum(impmesha-impmesde) from hsaldos where anopsald=" & NumRegElim - J
                Cad = Cad & " AND codmacta in (select cta  FROM tmpconext WHERE codusu = " & vUsu.Codigo
                Cad = Cad & " AND numdiari =" & NumRegElim - J & ") group by 1"
               
                miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not miRsAux.EOF
                    If miRsAux.Fields(1) <> 0 Then
                        Cad = "UPDATE tmpconext set saldo = " & TransformaComasPuntos(CStr(miRsAux.Fields(1)))
                        Cad = Cad & " WHERE codusu = " & vUsu.Codigo & " AND numdiari = " & NumRegElim - J
                        Cad = Cad & " AND cta = '" & miRsAux!codmacta & "'"
                        Conn.Execute Cad
                        
                    End If
                    miRsAux.MoveNext
                Wend
                miRsAux.Close
                
                
            Next
            
            
            'Qquit el ciere y la apertura
            Label3(4).Caption = "Quitar cierres..."
            Label3(4).Refresh
            F1 = CDate(txtFecha(5).Text)
            F1 = DateAdd("yyyy", -1, F1)
            
            DoEvents
            Cad = "select year(fechaent),codmacta,sum(coalesce(timporteh,0)-coalesce(timported,0)) from hlinapu"
            Cad = Cad & " where codconce=980 and fechaent>='" & Format(F1, FormatoFecha) & "'"
            Cad = Cad & " and codmacta in (select distinct(cta)  FROM tmpconext WHERE codusu = " & vUsu.Codigo & " ) group by 1,2"
            miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                If miRsAux.Fields(2) <> 0 Then
                    Cad = TransformaComasPuntos(CStr(miRsAux.Fields(2)))
                    If miRsAux.Fields(2) > 0 Then Cad = " + " & Cad
                
                    Cad = "UPDATE tmpconext set saldo = saldo  " & Cad
                    Cad = Cad & " WHERE codusu = " & vUsu.Codigo & " AND numdiari = " & miRsAux.Fields(0)
                    Cad = Cad & " AND cta = '" & miRsAux!codmacta & "'"
                    Conn.Execute Cad
                    
                End If
                miRsAux.MoveNext
            Wend
            miRsAux.Close
        End If
        
        
        
        
        Label3(4).Caption = "Cruzar pendientes con saldo"
        Label3(4).Refresh
    
        Cad = "select cta,numdiari,saldo from tmpconext where saldo >=0 and codusu = " & vUsu.Codigo
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            Importe = miRsAux!Saldo
            
            If Importe >= 0 Then
                'TODOS pagados
                Cad = "UPDATE tmp340 set totiva=totalfac, numiva=5 where  codusu= " & vUsu.Codigo & " AND cp_intracom='" & miRsAux!Cta
                Cad = Cad & "' AND nifdeclarado=" & miRsAux!NumDiari
                Conn.Execute Cad
            Else
            
            
            
            
            End If
            miRsAux.MoveNext
        Wend
        
        miRsAux.Close
        
        
        
        
        
    End If
    
    
    If NumProceso = 1 Then
    
        Label3(4).Caption = "Establecer nombres"
        Label3(4).Refresh
        
        Cad = "UPDATE tmp340,cuentas set dom_intracom = nommacta where codusu = " & vUsu.Codigo
        Cad = Cad & " AND tmp340.cp_intracom =cuentas.codmacta  "
        Conn.Execute Cad
        
        
        
        'Preparamos el rpt
        Cad = "select nifdeclarado from tmp340 where codusu = " & vUsu.Codigo & " GROUP BY 1 ORDER BY 1 desc"
        Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        ConceptosProvee = "Dentro del plazo máximo legal|Resto|Total pagos ejercicio|PMPE(dias)|Aplaza. que a fecha cierre superen plazo|"
        While Not Rs.EOF
            Cad = ""
            For NumRegElim = 1 To 5             '2 ult cifras año
                Cad = Cad & ", (" & vUsu.Codigo & "," & Right(Rs!nifdeclarado, 2) & "," & NumRegElim & ",0,'" & RecuperaValor(ConceptosProvee, CInt(NumRegElim)) & "')"
            Next
            Cad = "INSERT INTO usuarios.z347(codusu,cliprov,nif,importe,provincia) VALUES " & Mid(Cad, 2)
            Conn.Execute Cad
        
            Rs.MoveNext
        Wend
        Rs.Close
        
        
        'AHora relleno los importes
        Label3(4).Caption = "Totales"
        Label3(4).Refresh
        Cad = "select nifdeclarado,if(base>=" & txtAno(1).Text & ",'NO',''),sum(totiva) from tmp340 where codusu = " & vUsu.Codigo
        Cad = Cad & " AND numiva>0 group by 1,2"
        Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs.EOF
            Cad = "UPDATE usuarios.z347 SET importe = " & TransformaComasPuntos(CStr(Rs.Fields(2)))
            Cad = Cad & " WHERE codusu = " & vUsu.Codigo & " AND cliprov = '" & Right(Rs!nifdeclarado, 2) & "' AND nif ='"
            If Rs.Fields(1) = "" Then
                Cad = Cad & "1"
            Else
                Cad = Cad & "2" 'fuera de plazo
            End If
            Conn.Execute Cad & "'"
            
            Rs.MoveNext
        Wend
        Rs.Close
        
        
       
        DatosGenerados = True
        
        
        'Total pagos y PMPE
        Cad = "select cliprov,sum(importe) from usuarios.z347  WHERE codusu=" & vUsu.Codigo & " group by 1"
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            If miRsAux.Fields(1) <> 0 Then
                CtaTratada = TransformaComasPuntos(CStr(miRsAux.Fields(1)))
                Cad = "UPDATE usuarios.z347 SET importe=" & CtaTratada
                Cad = Cad & " WHERE codusu = " & vUsu.Codigo & " AND cliprov = '" & miRsAux!cliprov & "' AND nif='3'"
                Conn.Execute Cad
                
                'Vamos con los porcentajes
                
                Cad = "UPDATE usuarios.z347 SET razosoci=concat(round((importe/" & CtaTratada & ")*100,2),' %')"
                Cad = Cad & " WHERE codusu = " & vUsu.Codigo & " AND cliprov = '" & miRsAux!cliprov & "' AND nif in ('1','2')"
                Conn.Execute Cad
             End If
             miRsAux.MoveNext
             
        Wend
        miRsAux.Close
        
        Cad = "select nifdeclarado,sum(totalfac) from tmp340 where codusu=" & vUsu.Codigo & " and numiva>0 and base >" & Me.txtAno(1).Text & " group by 1"
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            CtaTratada = TransformaComasPuntos(CStr(miRsAux.Fields(1)))
            Cad = "select sum(totalfac*(base-" & Me.txtAno(1).Text & "))/" & CtaTratada & " from tmp340 where codusu="
            Cad = Cad & vUsu.Codigo & " and numiva>0 and base > " & Me.txtAno(1).Text & " and  nifdeclarado='" & miRsAux!nifdeclarado & "'"
            Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not Rs.EOF Then
                'Solo un registro
                CtaTratada = TransformaComasPuntos(CStr(Rs.Fields(0)))
                Cad = "UPDATE usuarios.z347 SET importe=" & CtaTratada
                Cad = Cad & " WHERE codusu = " & vUsu.Codigo & " AND cliprov = '" & Right(CStr(miRsAux!nifdeclarado), 2) & "' AND nif='4'"
                Conn.Execute Cad
                
            End If
            Rs.Close
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        
        
        NumRegElim = 0
        Cad = DevuelveDesdeBD("count(*)", "tmp340", "codusu", CStr(vUsu.Codigo))
        If Cad <> "" Then
            If Val(Cad) > 0 Then
                'LO metemos en la de informes
                
                Cad = "INSERT INTO usuarios.z340 select * from tmp340 WHERE codusu = " & vUsu.Codigo
                EjecutaSQL Cad
                    
                NumRegElim = 1 'para que lance el rpt
            End If
        End If
        
    End If
    
    
   
    
    Exit Function
eProcesoMemoriaPagosProveedor:
    MuestraError Err.Number, Err.Description
    NumRegElim = 0
End Function


Private Function PonerValorBDnumregis() As Long
    PonerValorBDnumregis = -1
    If Not IsNull(miRsAux!nifrepresante) Then
        If IsNumeric(miRsAux!nifrepresante) Then PonerValorBDnumregis = Val(miRsAux!nifrepresante)
    End If
End Function

