VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmVto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vencimientos"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6195
   Icon            =   "frmVto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBloqueo 
      Height          =   5535
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   6135
      Begin VB.Frame FrameEstaBloq 
         BorderStyle     =   0  'None
         Height          =   3615
         Left            =   240
         TabIndex        =   49
         Top             =   1080
         Width           =   5775
         Begin VB.Label Label5 
            Caption         =   "esta bloqueada para la generacion de vtos."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   50
            Top             =   120
            Width           =   4890
         End
      End
      Begin VB.TextBox txtLineasCobrosPagos 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   47
         Text            =   "Text2"
         Top             =   600
         Width           =   5535
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2895
         Left            =   240
         TabIndex        =   45
         Top             =   1680
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   5106
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Vto"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fec. Vto"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Importe"
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   4080
         TabIndex        =   44
         Top             =   4920
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "ya tiene vencimientos en la tesoreria:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   48
         Top             =   1200
         Width           =   3885
      End
      Begin VB.Label Label5 
         Caption         =   "La factura :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   46
         Top             =   240
         Width           =   1155
      End
   End
   Begin VB.Frame FrameIntroCli 
      Height          =   5535
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   6135
      Begin VB.Frame FramePROV 
         BorderStyle     =   0  'None
         Height          =   4815
         Left            =   120
         TabIndex        =   51
         Top             =   120
         Width           =   5895
         Begin VB.Frame FrameTapaCCC 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   495
            Left            =   120
            TabIndex        =   77
            Top             =   3360
            Width           =   5775
         End
         Begin VB.TextBox txtBanco 
            Height          =   285
            Index           =   9
            Left            =   1680
            MaxLength       =   4
            TabIndex        =   57
            Text            =   "Text2"
            Top             =   3480
            Width           =   735
         End
         Begin VB.TextBox txtBanco 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   4440
            MaxLength       =   10
            TabIndex        =   61
            Text            =   "Text2"
            Top             =   3480
            Width           =   1455
         End
         Begin VB.TextBox txtBanco 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   3840
            MaxLength       =   2
            TabIndex        =   60
            Text            =   "Text2"
            Top             =   3480
            Width           =   375
         End
         Begin VB.TextBox txtBanco 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   3240
            MaxLength       =   4
            TabIndex        =   59
            Text            =   "Text2"
            Top             =   3480
            Width           =   495
         End
         Begin VB.TextBox txtBanco 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   2520
            MaxLength       =   4
            TabIndex        =   58
            Text            =   "Text2"
            Top             =   3480
            Width           =   615
         End
         Begin VB.TextBox txtDescCta 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   1
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   67
            Text            =   "Text3"
            Top             =   2880
            Width           =   3135
         End
         Begin VB.TextBox txtCta 
            Height          =   285
            Index           =   1
            Left            =   1680
            TabIndex        =   56
            Text            =   "1234567890"
            Top             =   2880
            Width           =   1095
         End
         Begin VB.TextBox txtImporte 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   4680
            TabIndex        =   63
            Text            =   "Text2"
            Top             =   4320
            Width           =   1215
         End
         Begin VB.TextBox txtFecha 
            Height          =   285
            Index           =   1
            Left            =   2520
            TabIndex        =   62
            Text            =   "Text2"
            Top             =   4320
            Width           =   1215
         End
         Begin VB.TextBox txtPagoDesc 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   1
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   66
            Text            =   "Text2"
            Top             =   2160
            Width           =   3135
         End
         Begin VB.TextBox txtPago 
            Height          =   285
            Index           =   1
            Left            =   1680
            TabIndex        =   55
            Text            =   "Text2"
            Top             =   2160
            Width           =   735
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   5
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   54
            Text            =   "Text1"
            Top             =   1200
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   4
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   53
            Text            =   "Text1"
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   3
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   52
            Text            =   "Text1"
            Top             =   720
            Width           =   4695
         End
         Begin VB.Label Label1 
            Caption         =   "IBAN"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   76
            Top             =   3480
            Width           =   975
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   1
            Left            =   2280
            Picture         =   "frmVto.frx":000C
            Top             =   4320
            Width           =   240
         End
         Begin VB.Image Imgbanco 
            Height          =   240
            Index           =   1
            Left            =   1320
            Picture         =   "frmVto.frx":0097
            Stretch         =   -1  'True
            Top             =   2880
            Width           =   240
         End
         Begin VB.Image imgFPago 
            Height          =   240
            Index           =   1
            Left            =   1320
            Picture         =   "frmVto.frx":68E9
            Stretch         =   -1  'True
            Top             =   2160
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cta. banco"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   75
            Top             =   2880
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Importe"
            Height          =   255
            Index           =   1
            Left            =   3960
            TabIndex        =   74
            Top             =   4320
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha"
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   73
            Top             =   4320
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "1er Vencimiento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   11
            Left            =   240
            TabIndex        =   72
            Top             =   4320
            Width           =   1575
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   240
            X2              =   5880
            Y1              =   4080
            Y2              =   4080
         End
         Begin VB.Label Label2 
            Caption         =   "Vencimientos factura PROVEEDOR"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   345
            Index           =   0
            Left            =   600
            TabIndex        =   71
            Top             =   240
            Width           =   4935
         End
         Begin VB.Label Label1 
            Caption         =   "Forma pago"
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
            Index           =   13
            Left            =   240
            TabIndex        =   70
            Top             =   2160
            Width           =   1005
         End
         Begin VB.Line Line1 
            Index           =   3
            X1              =   360
            X2              =   5880
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Label Label1 
            Caption         =   "Factura"
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
            Index           =   14
            Left            =   240
            TabIndex        =   69
            Top             =   1200
            Width           =   660
         End
         Begin VB.Label Label1 
            Caption         =   "Proveedor"
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
            Index           =   15
            Left            =   240
            TabIndex        =   68
            Top             =   720
            Width           =   885
         End
      End
      Begin VB.TextBox txtBanco 
         Height          =   285
         Index           =   8
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   1
         Text            =   "Text2"
         Top             =   2280
         Width           =   615
      End
      Begin VB.Frame FrameTapaDpto 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         TabIndex        =   39
         Top             =   3600
         Width           =   5895
      End
      Begin VB.TextBox txtDptoDesc 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "Text2"
         Top             =   3720
         Width           =   3375
      End
      Begin VB.TextBox txtDpto 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   3720
         Width           =   735
      End
      Begin VB.TextBox txtAgenteDesc 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "Text2"
         Top             =   3240
         Width           =   3375
      End
      Begin VB.TextBox txtAgente 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   7
         Text            =   "Text2"
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox txtDescCta 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "Text3"
         Top             =   2760
         Width           =   3135
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   6
         Text            =   "1234567890"
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton cmdPrimer 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   4800
         TabIndex        =   65
         Top             =   5040
         Width           =   1095
      End
      Begin VB.CommandButton cmdPrimer 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   3480
         TabIndex        =   64
         Top             =   5040
         Width           =   1095
      End
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   4680
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   0
         Left            =   2520
         TabIndex        =   9
         Text            =   "Text2"
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox txtBanco 
         Height          =   285
         Index           =   2
         Left            =   3840
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtBanco 
         Height          =   285
         Index           =   1
         Left            =   3120
         MaxLength       =   4
         TabIndex        =   3
         Text            =   "Text2"
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox txtBanco 
         Height          =   285
         Index           =   0
         Left            =   2400
         MaxLength       =   4
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox txtBanco 
         Height          =   285
         Index           =   3
         Left            =   4320
         TabIndex        =   5
         Text            =   "Text2"
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox txtPagoDesc 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "Text2"
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox txtPago 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   0
         Text            =   "Text2"
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   720
         Width           =   4695
      End
      Begin VB.Label Label1 
         Caption         =   "Departamento"
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
         Index           =   7
         Left            =   240
         TabIndex        =   38
         Top             =   3720
         Width           =   1245
      End
      Begin VB.Image imgDpto 
         Height          =   240
         Index           =   0
         Left            =   1440
         Picture         =   "frmVto.frx":D13B
         Stretch         =   -1  'True
         Top             =   3720
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Agente"
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
         Index           =   6
         Left            =   240
         TabIndex        =   36
         Top             =   3240
         Width           =   885
      End
      Begin VB.Image ImgAgente 
         Height          =   240
         Index           =   0
         Left            =   1440
         Picture         =   "frmVto.frx":EE35
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   2280
         Picture         =   "frmVto.frx":10B2F
         Top             =   4560
         Width           =   240
      End
      Begin VB.Image Imgbanco 
         Height          =   240
         Index           =   0
         Left            =   1440
         Picture         =   "frmVto.frx":10C31
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image imgFPago 
         Height          =   240
         Index           =   0
         Left            =   1440
         Picture         =   "frmVto.frx":1292B
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cta. banco"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   33
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Importe"
         Height          =   255
         Index           =   0
         Left            =   3960
         TabIndex        =   32
         Top             =   4560
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   31
         Top             =   4560
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "1er Vencimiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   30
         Top             =   4320
         Width           =   1575
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   240
         X2              =   5880
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Label Label1 
         Caption         =   "IBAN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   29
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Vencimientos factura CLIENTE"
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
         Height          =   345
         Index           =   13
         Left            =   600
         TabIndex        =   28
         Top             =   120
         Width           =   4845
      End
      Begin VB.Label Label1 
         Caption         =   "Forma pago"
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
         Left            =   240
         TabIndex        =   26
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   360
         X2              =   5880
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Label1 
         Caption         =   "Factura"
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
         Index           =   1
         Left            =   360
         TabIndex        =   25
         Top             =   1080
         Width           =   660
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
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
         Index           =   0
         Left            =   360
         TabIndex        =   24
         Top             =   720
         Width           =   600
      End
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   3
      Left            =   2400
      TabIndex        =   41
      Tag             =   "Código concepto|N|N|0|900|conceptos|codconce|000|S|"
      Text            =   "Dat"
      Top             =   4800
      Width           =   800
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   3240
      MaxLength       =   30
      TabIndex        =   40
      Tag             =   "Denominación|T|N|||conceptos|nomconce|||"
      Text            =   "Dato2"
      Top             =   4800
      Width           =   1395
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3600
      TabIndex        =   16
      Top             =   5040
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4920
      TabIndex        =   17
      Top             =   5040
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   290
      Index           =   1
      Left            =   900
      MaxLength       =   30
      TabIndex        =   15
      Tag             =   "Denominación|T|N|||conceptos|nomconce|||"
      Text            =   "Dato2"
      Top             =   4800
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   290
      Index           =   0
      Left            =   60
      MaxLength       =   3
      TabIndex        =   14
      Tag             =   "Código concepto|N|N|0|900|conceptos|codconce|000|S|"
      Text            =   "Dat"
      Top             =   4800
      Width           =   800
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmVto.frx":14625
      Height          =   1725
      Left            =   60
      TabIndex        =   20
      Top             =   540
      Width           =   5970
      _ExtentX        =   10530
      _ExtentY        =   3043
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   120
      TabIndex        =   18
      Top             =   4920
      Width           =   2865
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   2550
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   375
      Left            =   3840
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Modificar Lineas"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   4560
         TabIndex        =   22
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar"
      Height          =   375
      Left            =   4920
      TabIndex        =   42
      Top             =   5040
      Width           =   1035
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Visible         =   0   'False
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmVto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public opcion As Byte       '0- CObros
                            '1- Pagos

Public Datos As String
'Aqui vendra, segun sea CLI, or pro, los datos a poner
' CLIENTE ->>   Serie|codfac  |anofac|fecfac|CLIENTE|
' PROVEEDOR     vacio|numregis|anofac|fecfac|PROVEEODOR|
'                   tanto CLIE como PRO seran los campos
Public Importe As Currency


Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCta As frmCuentasBancarias
Attribute frmCta.VB_VarHelpID = -1

Private CadenaConsulta As String
Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas
Dim Modo As Byte


'Dim Para el buscagrid. pq se llam desde varios sitios
Dim CampoDefecto As Byte
Dim CamposBuscaGrid As String   'Como siempre, empipado titulo|tabla|devuelve
Dim Campos As String
Dim Devuelve As String

Dim SQL As String
Dim SubImporte As Currency


'Si va a generar mas de un pago, utilizaremos temporalmente la tabla tmpfaclin
'Meteremos todos los pagos FEcha, importe en tmpfaclin. Cuando ordene generar pagos
' Cojeremos el conjunto de importes y los insertaremos en cobros pagos en funcion de
' de cliente/prov


'----------------------------------------------
'----------------------------------------------
'   Deshabilitamos todos los botones menos
'   el de salir
'   Ademas mostramos aceptar y cancelar
'   Modo 0->  Normal
'   Modo 1 -> Lineas  INSERTAR
'   Modo 2 -> Lineas MODIFICAR
'   Modo 3 -> Lineas BUSCAR
'----------------------------------------------
'----------------------------------------------

Private Sub PonerModo(vModo)
Dim B As Boolean
Modo = vModo

B = (Modo = 0)

txtaux(0).Visible = Not B
txtaux(1).Visible = Not B
txtaux(2).Visible = Not B
txtaux(3).Visible = Not B

mnOpciones.Enabled = B
Toolbar1.Buttons(1).Enabled = B
Toolbar1.Buttons(2).Enabled = B
Toolbar1.Buttons(8).Enabled = B
Toolbar1.Buttons(7).Enabled = B
Toolbar1.Buttons(6).Enabled = B

'Prueba


cmdAceptar.Visible = Not B
cmdCancelar.Visible = Not B
DataGrid1.Enabled = B

cmdGenerar.Visible = False
''Si estamo mod or insert
'If Modo = 2 Then
'   txtAux(0).BackColor = &H80000018
'   Else
'    txtAux(0).BackColor = &H80000005
'End If
'txtAux(0).Enabled = (Modo <> 2)
End Sub

Private Sub BotonAnyadir()
    Dim NumF As String
    Dim anc As Single
    
    'Obtenemos la siguiente numero de factura
    NumF = SugerirCodigoSiguiente
    
    If SubImporte = 0 Then
        MsgBox "La suma coincide con el total factura. No se puden insertar mas lineas", vbExclamation
        Exit Sub
    End If
        
    
    lblIndicador.Caption = "INSERTANDO"
    'Situamos el grid al final
    DataGrid1.AllowAddNew = True
    If Adodc1.Recordset.RecordCount > 0 Then
        DataGrid1.HoldFields
        Adodc1.Recordset.MoveLast
        DataGrid1.Row = DataGrid1.Row + 1
    End If
    
    
   
    If DataGrid1.Row < 0 Then
        anc = 770
        Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 545
    End If
    txtaux(0).Text = Format(NumF, "000")
    txtaux(1).Text = RecuperaValor(Datos, 6)
    txtaux(2).Text = Campos
    txtaux(3).Text = SubImporte
    
    
    
    
    LLamaLineas anc, 0
    
    
    'Ponemos el foco
    txtaux(3).SetFocus
    
'    If FormularioHijoModificado Then
'        CargaGrid
'        BotonAnyadir
'        Else
'            'cmdCancelar.SetFocus
'            If Not Adodc1.Recordset.EOF Then _
'                Adodc1.Recordset.MoveFirst
'    End If
End Sub



Private Sub BotonVerTodos()
    CargaGrid
End Sub

Private Sub BotonBuscar()
    CargaGrid
    'Buscar
    txtaux(0).Text = ""
    txtaux(1).Text = ""
    LLamaLineas DataGrid1.Top + 206, 2
    txtaux(0).SetFocus
End Sub

Private Sub BotonModificar()
    '---------
    'MODIFICAR
    '----------
    
    Dim anc As Single
    Dim i As Integer
    If Adodc1.Recordset.EOF Then Exit Sub
    If Adodc1.Recordset.RecordCount < 1 Then Exit Sub


    Screen.MousePointer = vbHourglass
    'Para que me indique cuanto queda
    SQL = SugerirCodigoSiguiente
    SubImporte = SubImporte + Adodc1.Recordset!Total
    Me.lblIndicador.Caption = "MODIFICAR"
    DeseleccionaGrid DataGrid1
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 320
        Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 545
    End If

    'Llamamos al form
    txtaux(0).Text = DataGrid1.Columns(0).Text
    txtaux(1).Text = DataGrid1.Columns(1).Text
    txtaux(2).Text = DataGrid1.Columns(2).Text
    txtaux(3).Text = TransformaComasPuntos(Adodc1.Recordset!Total)
    
    
    LLamaLineas anc, 1
   
   'Como es modificar
   txtaux(2).SetFocus
   
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
PonerModo xModo + 1
'Fijamos el ancho
txtaux(0).Top = alto
txtaux(1).Top = alto
txtaux(2).Top = alto
txtaux(3).Top = alto
End Sub




Private Sub BotonEliminar()

    On Error GoTo Error2
    'Ciertas comprobaciones
    If Adodc1.Recordset.EOF Then Exit Sub
    Campos = ""
    CampoDefecto = 0
    If Not Adodc1.Recordset.AbsolutePosition = Adodc1.Recordset.RecordCount Then
        Campos = "H"
        CampoDefecto = 1
        SQL = "Va a eliminar un registro que no es el ultimo. El programa renumerará los vencimientos." & vbCrLf & "¿Desea continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    '### a mano

    If Campos = "" Then
        SQL = "Seguro que desea eliminar el vencimiento:"
        SQL = SQL & vbCrLf & "Fecha: " & Format(Adodc1.Recordset.Fields(2), "dd/mm/yyyy")
        SQL = SQL & vbCrLf & "Importe: " & Adodc1.Recordset.Fields(3)
        'Hago la pregunta pq no la he hecho
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
    
        'Hay que eliminar
        
        SQL = "Delete from tmpfaclin where codigo=" & Adodc1.Recordset!Codigo
        SQL = SQL & " AND codusu = " & vUsu.Codigo
        Conn.Execute SQL
        
        
        'Renumero
        If CampoDefecto = 1 Then
            SQL = "UPDATE tmpfaclin Set codigo=codigo +1000 WHERE codusu = " & vUsu.Codigo
            Conn.Execute SQL
                
                
            Set miRsAux = New ADODB.Recordset
            SQL = "Select * from tmpfaclin where codusu = " & vUsu.Codigo & " ORDER By codigo "
            NumRegElim = 1
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            While Not miRsAux.EOF
                SQL = "UPDATE tmpfaclin set codigo = " & NumRegElim & " where codigo= " & miRsAux!Codigo & " AND codusu = " & vUsu.Codigo
                Conn.Execute SQL
                NumRegElim = NumRegElim + 1
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            Set miRsAux = Nothing
        
        End If
        
        CargaGrid
        Adodc1.Recordset.Cancel
        cmdGenerar.Visible = False
    Exit Sub
Error2:
        Screen.MousePointer = vbDefault
        MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub





Private Sub cmdAceptar_Click()


Select Case Modo
    Case 1
        If DatosOK2 Then
            '-----------------------------------------
                CargaGrid
                If SubImporte > 0 Then
                    BotonAnyadir
                Else
                    cmdCancelar_Click
                    cmdGenerar.Visible = SubImporte = 0
                End If
                
                
         End If
    Case 2
            'Modificar
            If DatosOK2 Then
                Me.lblIndicador.Caption = ""
                PonerModo 0
                CargaGrid
                cmdGenerar.Visible = SubImporte = 0
            End If
    Case 3
'        'HacerBusqueda
'        CadB = ObtenerBusqueda(Me)
'        If CadB <> "" Then
'            PonerModo 0
'            CargaGrid 'CadB
'        End If
    End Select


End Sub

Private Sub cmdCancelar_Click()
Select Case Modo
Case 1
    DataGrid1.AllowAddNew = False
    'CargaGrid
    If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveFirst
    
Case 3
    CargaGrid
End Select
PonerModo 0
lblIndicador.Caption = ""
DataGrid1.SetFocus
End Sub







Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdGenerar_Click()
    'Comprobamos que el importe es correcto
    SQL = SugerirCodigoSiguiente
    If SubImporte <> 0 Then
        MsgBox "El total vencimientos no suma el total factura", vbExclamation
        Exit Sub
    End If
    'Preguntamos si desea generar los vencimineots
    SQL = "¿Desea generar los vencimientos?"
    CampoDefecto = MsgBox(SQL, vbQuestion + vbYesNoCancel)
    If CampoDefecto = vbCancel Then Exit Sub
    
    
    'Si o no
    
    'Genero vtos
    If CampoDefecto = vbYes Then GenerarPagosCobros
        
    Unload Me
        
    
End Sub

Private Sub cmdPrimer_Click(Index As Integer)
    If Index = 1 Then
        '.....
        'SALIR
        Unload Me
        Exit Sub
    End If
    
    
    'ACeptamos.
    'Primeras comprobaciones
    If Not DatosOK Then Exit Sub
        
    
    
    'Si el importe escrito es igual que el total es que efectua un UNICO pago
    'Si no pasaremos al aprtado 2
    If Not InsertaPagoCobro(SubImporte, CDate(txtFecha(opcion).Text), 1, True) Then Exit Sub
        
    SubImporte = Importe - SubImporte
    
    If SubImporte <> 0 Then
        'Aun falta por pagar, luego mostramos el grid
        CargaGrid
        FrameIntroCli.Visible = False
        BotonAnyadir
    Else
    
        Campos = "¿Desea generar el vencimiento?"
        CampoDefecto = MsgBox(Campos, vbQuestion + vbYesNoCancel)
        
        
        If CampoDefecto = vbCancel Then
            BorraTmpFaclin
            Exit Sub
        End If
            
        
    
        If CampoDefecto = vbYes Then GenerarPagosCobros
        'Salimos
        Unload Me
    End If
    
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub



Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If opcion = 1 Then txtPago(1).SetFocus
        
End Sub


Private Sub Form_Load()
    
    Me.Icon = frmPpal.Icon

    Limpiar Me
    
    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1
        .Buttons(2).Image = 2
        .Buttons(6).Image = 3
        .Buttons(7).Image = 4
        .Buttons(8).Image = 5
        '.Buttons(10).Image = 10
        .Buttons(11).Image = 16
        .Buttons(12).Image = 15
        .Buttons(14).Image = 6
        .Buttons(15).Image = 7
        .Buttons(16).Image = 8
        .Buttons(17).Image = 9
    End With
    Me.Icon = frmPpal.Icon
    
    '## A mano
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)

    'Borramos el temporal
    BorraTmpFaclin
        
        
    'Comun para cli y prov
    
    
    If opcion = 0 Then
        
        'CLIENTES
        ' CLIENTE ->>   Serie|codfac  |anofac|fecfac|CLIENTE| ->cta |desc|
        CampoDefecto = 0
        'Ahora, si la cuenta tiene departamentos pondre enabled
        CadenaConsulta = DevuelveDesdeBD("Dpto", "Departamentos", "codmacta", RecuperaValor(Datos, 5), "T")
        FrameTapaDpto.Visible = CadenaConsulta = ""
        Me.txtdpto(0).Enabled = FrameTapaDpto.Visible
    Else
        'PROVEEDORES
        ' PROVEEDOR     vacio|numregis|anofac|fecfac|PROVEEODOR|
        CampoDefecto = 3
        FrameTapaCCC.Left = Label1(8).Left
        FrameTapaCCC.Visible = True
    End If
    txtFecha(opcion).Text = Format(Now, "dd/mm/yyyy")
    Me.txtImporte(opcion).Text = Format(Importe, FormatoImporte)
    'Datos factura en funcion cli/prov
    Text1(0 + CampoDefecto).Text = RecuperaValor(Datos, 5) & " - " & RecuperaValor(Datos, 6)
    Text1(1 + CampoDefecto).Text = RecuperaValor(Datos, 1) & " " & RecuperaValor(Datos, 2)
    Text1(2 + CampoDefecto).Text = RecuperaValor(Datos, 4)
    
    
    
    DespalzamientoVisible False
    PonerModo 0
    CadAncho = False
    PonerOpcionesMenu  'En funcion del usuario
    
    'Poner cuenta banco SI TIENE
    PonerCuentaBanco_Forpa
    
    'Cadena consulta
    CadenaConsulta = "Select codigo, Cliente,Fecha, Total from tmpfaclin where codusu =" & vUsu.Codigo
    lblIndicador.Caption = ""


    'Bloquea y Compruebo si ya existe o no
    If ComprobarExistencia Then
        FramePROV.Visible = opcion = 1
        cmdPrimer(1).Cancel = True
    Else
        cmdCerrar.Cancel = True
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    BloqueoManual False, "frmVto", ""
End Sub

Private Sub BorraTmpFaclin()
    SQL = "DELETE FROM tmpfaclin where codusu = " & vUsu.Codigo
    Conn.Execute SQL
End Sub

Private Sub LlamaBuscaGrid(Optional vSQL As String)

 ' CamposBuscaGrid : titulo|tabla|


           Screen.MousePointer = vbHourglass
           Set frmB = New frmBuscaGrid
           frmB.vCampos = Campos
           frmB.vTabla = RecuperaValor(CamposBuscaGrid, 2)
           frmB.vSQL = vSQL
           '###A mano
           frmB.vDevuelve = Devuelve
           frmB.vTitulo = RecuperaValor(CamposBuscaGrid, 1)
           frmB.vSelElem = CInt(CampoDefecto)
           '#
           frmB.Show vbModal
           Set frmB = Nothing
            
End Sub






Private Sub frmB_Selecionado(CadenaDevuelta As String)
    CadenaDesdeOtroForm = CadenaDevuelta
End Sub

Private Sub frmC_Selec(vFecha As Date)
    'txtFecha(0).Text = Format(vFecha, "dd/mm/yyyy")
    CadenaDesdeOtroForm = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCta_DatoSeleccionado(CadenaSeleccion As String)
    CadenaDesdeOtroForm = CadenaSeleccion
End Sub

Private Sub ImgAgente_Click(Index As Integer)
    CadenaDesdeOtroForm = ""
    'LlamaBuscaGrid()
    CamposBuscaGrid = "Agentes|Agentes|"
    'Los campos Desc|campo|N|10·"
    Campos = "Codigo|codigo|N|20·Nombre|nombre|T|70·"
    Devuelve = "0|1|"
    LlamaBuscaGrid
    If CadenaDesdeOtroForm <> "" Then
        Me.txtAgente(Index).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
        Me.txtAgenteDesc(Index).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
    End If
    CamposBuscaGrid = ""
End Sub

Private Sub imgDpto_Click(Index As Integer)
    CadenaDesdeOtroForm = ""
    'LlamaBuscaGrid()
    CamposBuscaGrid = "Departamentos. (" & RecuperaValor(Datos, 5) & ")" & "|departamentos|"
    'Los campos Desc|campo|N|10·"
    Campos = "Codigo|dpto|N|20·Descripcion|descripcion|T|70·"
    Devuelve = "0|1|"
    LlamaBuscaGrid "codmacta = '" & RecuperaValor(Datos, 5) & "'"
    If CadenaDesdeOtroForm <> "" Then
        Me.txtdpto(Index).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
        Me.txtDptoDesc(Index).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
    End If
    CamposBuscaGrid = ""
End Sub

Private Sub imgFecha_Click(Index As Integer)
    CadenaDesdeOtroForm = ""
    Screen.MousePointer = vbHourglass
    Set frmC = New frmCal
    frmC.Fecha = Now
    If txtFecha(Index).Text <> "" Then frmC.Fecha = CDate(txtFecha(Index).Text)
    frmC.Show vbModal
    Set frmC = Nothing
    If CadenaDesdeOtroForm <> "" Then
        Me.txtFecha(Index).Text = Format(CadenaDesdeOtroForm, "dd/mm/yyyy")
        PonleFoco Me.txtImporte(Index)
    End If
End Sub

Private Sub imgFPago_Click(Index As Integer)
    CadenaDesdeOtroForm = ""
    'LlamaBuscaGrid()
    CamposBuscaGrid = "Formas de Pago|formapago|"
    'Los campos Desc|campo|N|10·"
    Campos = "Codigo|codforpa|N|20·Descripcion|nomforpa|T|70·"
    Devuelve = "0|1|"
    LlamaBuscaGrid
    If CadenaDesdeOtroForm <> "" Then
        Me.txtPago(Index).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
        
        
        
        txtPago_LostFocus Index
        'Me.txtPagoDesc(Index).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
        If Index = 0 Then
            PonFoco txtBanco(Index)
        Else
            PonFoco txtcta(1)
        End If
    End If
    CamposBuscaGrid = ""
End Sub


Private Sub Imgbanco_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    Set frmCta = New frmCuentasBancarias
    frmCta.DatosADevolverBusqueda = "0|1|"
    frmCta.Show vbModal
    Set frmCta = Nothing
    If CadenaDesdeOtroForm <> "" Then
        txtcta(Index).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
        txtDescCta(Index).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
        If opcion = 0 Then txtAgente(Index).SetFocus
    End If
End Sub



Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
    BotonModificar
End Sub

Private Sub mnNuevo_Click()
BotonAnyadir
End Sub

Private Sub mnSalir_Click()
Screen.MousePointer = vbHourglass
Unload Me
End Sub

Private Sub mnVerTodos_Click()
BotonVerTodos
End Sub



'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
'### A mano
'Esto es para que cuando pincha en siguiente le sugerimos
'Se puede comentar todo y asi no hace nada ni da error
'El SQL es propio de cada tabla
Private Function SugerirCodigoSiguiente() As String
    Dim F As Date
    Dim Rs As ADODB.Recordset
    
    SQL = "Select codigo,total,fecha from tmpfaclin where codusu = " & vUsu.Codigo & " ORDER BY codigo"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, , , adCmdText
    SQL = "1"
    SubImporte = 0
    F = Now
    If Not Rs.EOF Then
        While Not Rs.EOF
            If Not IsNull(Rs.Fields(0)) Then
                SQL = CStr(Rs.Fields(0) + 1)
                SubImporte = SubImporte + Rs.Fields(1)
                F = Rs!Fecha
            End If
        
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    Campos = Format(DateAdd("m", 1, F), "dd/mm/yyyy")
    SubImporte = Importe - SubImporte
    SugerirCodigoSiguiente = SQL
End Function

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
        BotonBuscar
Case 2
        BotonVerTodos
Case 6
        BotonAnyadir
Case 7
        BotonModificar
Case 8
        BotonEliminar
Case 11
        'Ha ido bien
        If InformeConceptos(CadenaConsulta) Then
            frmImprimir.opcion = 0
            frmImprimir.FormulaSeleccion = "{ado.codusu} = " & vUsu.Codigo
            frmImprimir.NumeroParametros = 0
            frmImprimir.Show vbModal
        End If
Case 12
        Unload Me
Case Else

End Select
End Sub


Private Sub DespalzamientoVisible(bol As Boolean)
    Dim i
    For i = 14 To 17
        Toolbar1.Buttons(i).Visible = bol
    Next i
End Sub

Private Sub CargaGrid()
    Dim J As Integer
    Dim TotalAncho As Integer
    Dim i As Integer
    
    Adodc1.ConnectionString = Conn
    Adodc1.RecordSource = CadenaConsulta
    Adodc1.CursorType = adOpenDynamic
    Adodc1.LockType = adLockOptimistic
    Adodc1.Refresh
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 290
    
    
    'Nombre producto
    i = 0
        DataGrid1.Columns(i).Caption = "Vto."
        DataGrid1.Columns(i).Width = 500
        DataGrid1.Columns(i).NumberFormat = "000"
        
    
    'Leemos del vector en 2
    i = 1
        DataGrid1.Columns(i).Caption = "Cuenta"
        DataGrid1.Columns(i).Width = 2800
        TotalAncho = TotalAncho + DataGrid1.Columns(i).Width
    
        DataGrid1.Columns(2).Width = 1000
        DataGrid1.Columns(2).NumberFormat = "dd/mm/yyyy"
    i = 3
        DataGrid1.Columns(i).Caption = "importe"
        DataGrid1.Columns(i).Width = 1200
        DataGrid1.Columns(i).NumberFormat = FormatoImporte
        DataGrid1.Columns(i).Alignment = dbgRight
        
            
        
        'Fiajamos el cadancho
    If Not CadAncho Then
        'La primera vez fijamos el ancho y alto de  los txtaux
        For i = 0 To 3
            txtaux(i).Width = DataGrid1.Columns(i).Width - 48
        Next i
        txtaux(0).Left = DataGrid1.Left + 340
        For i = 0 To 2
            txtaux(i + 1).Left = txtaux(i).Left + txtaux(i).Width + 48
        Next i
        CadAncho = True
    End If
    'Habilitamos modificar y eliminar
    If vUsu.Nivel < 2 Then
        Toolbar1.Buttons(7).Enabled = Not Adodc1.Recordset.EOF
        Toolbar1.Buttons(8).Enabled = Not Adodc1.Recordset.EOF
    End If
End Sub


Private Sub txtAgente_GotFocus(Index As Integer)
    PonFoco txtAgente(Index)
End Sub

Private Sub txtAgente_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub txtAgente_LostFocus(Index As Integer)
    txtAgente(Index).Text = Trim(txtAgente(Index).Text)
    Devuelve = ""
    CampoDefecto = 0
    If txtAgente(Index).Text <> "" Then
        If Not IsNumeric(txtAgente(Index).Text) Then
            MsgBox "Código agente debe ser numérico", vbExclamation
            CampoDefecto = 1
        Else
            Campos = DevuelveDesdeBD("nombre", "agentes", "codigo", txtAgente(Index).Text, "N")
            If Campos <> "" Then
                Devuelve = Campos
            Else
                MsgBox "No existe el agente : " & txtAgente(Index).Text, vbExclamation
                CampoDefecto = 1
            End If
        End If
    End If
    'txtPago(Index).Text = ""
    Me.txtAgenteDesc(Index).Text = Devuelve
    If Devuelve = "" Then txtAgente(Index).Text = ""
    If CampoDefecto <> 0 Then txtAgente(Index).SetFocus
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    PonFoco txtaux(Index)
    
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAux_LostFocus(Index As Integer)

txtaux(Index).Text = Trim(txtaux(Index).Text)
If txtaux(Index).Text = "" Then Exit Sub
If Modo = 3 Then Exit Sub 'Busquedas
If Index = 0 Then
    If Not IsNumeric(txtaux(0).Text) Then
        MsgBox "Código concepto tiene que ser numérico", vbExclamation
        Exit Sub
    End If
    txtaux(0).Text = Format(txtaux(0).Text, "000")
End If
End Sub


Private Function DatosOK2() As Boolean
Dim Im As Currency


    DatosOK2 = False
    
    If Me.txtaux(2).Text = "" Or txtaux(3).Text = "" Then
        MsgBox "Campos en blanco", vbExclamation
        Exit Function
    End If
    
    If Not EsFechaOK(txtaux(2)) Then
        MsgBox "Fecha incorrecta", vbExclamation
        Exit Function
    End If
    
    varFecOk = FechaCorrecta2(CDate(txtaux(2).Text))
    If varFecOk > 1 Then
        If varFecOk = 2 Then
            MsgBox varTxtFec, vbExclamation
            Exit Function
        Else
            SQL = "Fecha fuera de ejercicios. "
            If varFecOk = 3 Then
                'ANTERIOR A INICIO EJERCICIO
                Im = vbExclamation
            Else
                Im = vbQuestion + vbYesNo
                SQL = SQL & vbCrLf & "¿Desea continuar?"
            End If
            If MsgBox(SQL, CLng(Im)) <> vbYes Then Exit Function
        End If
    End If
        
    If Not IsNumeric(txtaux(3).Text) Then
        MsgBox "Campo numerico", vbExclamation
        Exit Function
    End If
    
    
    
    
    
    
    
    
    
    'Im = Round(CCur(TransformaPuntosComas(txtAux(3).Text)), 2)
    
    CadenaCurrency txtaux(3).Text, Im
    If SubImporte < Im Then
        MsgBox "Importe excede del total factura", vbExclamation
        Exit Function
    End If
    If Modo = 1 Then
        'INSERTO
        If Not InsertaPagoCobro(Im, CDate(txtaux(2).Text), CInt(txtaux(0).Text), True) Then Exit Function
    Else
        If Not InsertaPagoCobro(Im, CDate(txtaux(2).Text), CInt(txtaux(0).Text), False) Then Exit Function
    End If
    espera 0.2
    SubImporte = SubImporte - Im
        
    DatosOK2 = True
    
  
        
End Function

Private Function DatosOK() As Boolean

    DatosOK = False
    SQL = ""
    Campos = ""
    'Valores comunes para cobros/pagos
    If Me.txtPago(opcion).Text = "" Then SQL = "Forma de pago en blanco"
    If Me.txtcta(opcion).Text = "" Then SQL = "Cuenta banco en blanco"
    
    If txtImporte(opcion) = "" Then SQL = "Importe en blanco"
    If Me.txtFecha(opcion) = "" Then
        SQL = "Fecha en blanco"
    Else
    
        varFecOk = FechaCorrecta2(CDate(txtFecha(opcion)))
        
        If varFecOk > 1 Then
            If varFecOk = 2 Then
                SQL = varTxtFec
            Else
                
                If varFecOk = 4 Then
                    SQL = "Ejercicio no abierto todavia."
                    Campos = vbCrLf & " Continuar?"
                Else
                    SQL = "Fecha fuera de ejercicios"
                End If
            End If
        End If
    End If
    'Si ya da error salimos y asi que ponga los valores
    If SQL <> "" Then
        
        If Campos = "" Then
            MsgBox SQL, vbExclamation
            Exit Function
        Else
            If MsgBox(SQL & Campos, vbQuestion + vbYesNo) <> vbYes Then Exit Function
            SQL = ""
        End If
    End If
    
    'Solo cobros o pagos
    If opcion = 0 Then

        'Miramos la cuenta bancaria
        If Me.FrameTapaCCC.Tag <> "" Then
          '----------------------------------
            Campos = ""
            If SQL <> "" Then SQL = SQL & vbCrLf
            For CampoDefecto = 0 To 3
                txtBanco(CampoDefecto).Text = Trim(txtBanco(CampoDefecto).Text)
                If txtBanco(CampoDefecto).Text = "" Then Campos = Campos & "1"
                If CampoDefecto <> 2 Then SQL = SQL & txtBanco(CampoDefecto).Text
            Next CampoDefecto
            If Campos <> "" Then
                'Cuenta bancaria incorrecta
                If FrameTapaCCC.Tag = "4" Then
                    SQL = "Cuenta bancaria incorrecta"
                    Campos = "2"
                End If
            Else
                'Comprobaremos el CC
                Devuelve = SQL
                SQL = CodigoDeControl(SQL)
                If SQL <> txtBanco(2).Text Then
                    SQL = "CC incorrecto: " & txtBanco(2).Text & " -  " & SQL
                Else
                    SQL = ""
                End If
                
                
                'Compruebo EL IBAN
                'Meto el CC
                Devuelve = Mid(Devuelve, 1, 8) & Me.txtBanco(2).Text & Mid(Devuelve, 9)
                CamposBuscaGrid = ""
                If Me.txtBanco(8).Text <> "" Then CamposBuscaGrid = Mid(txtBanco(8).Text, 1, 2)
                    
                If DevuelveIBAN2(CamposBuscaGrid, Devuelve, Devuelve) Then
                    If Me.txtBanco(8).Text = "" Then
                        If MsgBox("Poner IBAN ?", vbQuestion + vbYesNo) = vbYes Then Me.txtBanco(8).Text = Devuelve
                    Else
                        If Mid(txtBanco(8).Text, 3) <> Devuelve Then
                            Devuelve = "Calculado : " & CamposBuscaGrid & Devuelve
                            Devuelve = "Introducido: " & Me.txtBanco(8).Text & vbCrLf & Devuelve & vbCrLf
                            Devuelve = "Error en codigo IBAN" & vbCrLf & Devuelve & "Continuar?"
                            If MsgBox(Devuelve, vbQuestion + vbYesNo) = vbNo Then SQL = SQL & vbCrLf & " Error en IBAN"
                        End If
                    End If
                End If
                CamposBuscaGrid = ""
                
                
                
                
                
            End If
        Else
            If FrameTapaCCC.Tag = "4" Then SQL = "Cuenta bancaria obligatoria"
        End If
        
        If Me.txtdpto(0).Enabled Then
            If Me.txtdpto(0).Text = "" Then SQL = SQL & vbCrLf & "Departamento en blanco"
        End If
        If Me.txtAgente(0).Text = "" Then SQL = SQL & vbCrLf & "Agente en blanco"
        
    Else
        'Pagos
        If Me.FrameTapaCCC.Tag <> "" Then
          '----------------------------------
            Campos = ""
            Devuelve = ""
            For CampoDefecto = 4 To 7
                txtBanco(CampoDefecto).Text = Trim(txtBanco(CampoDefecto).Text)
                If txtBanco(CampoDefecto).Text = "" Then
                    Campos = Campos & "1"
                Else
                    Devuelve = Devuelve & txtBanco(CampoDefecto).Text
                End If
            Next CampoDefecto
            If Len(Devuelve) <> 20 Then
                SQL = "Cuenta bancaria incorrecta"
                Campos = "2"
            
            
            Else
                'Compruebo EL IBAN
                'Meto el CC
               
                CamposBuscaGrid = ""
                If Me.txtBanco(9).Text <> "" Then CamposBuscaGrid = Mid(txtBanco(9).Text, 1, 2)
                    
                If DevuelveIBAN2(CamposBuscaGrid, Devuelve, Devuelve) Then
                    If Me.txtBanco(9).Text = "" Then
                        If MsgBox("Poner IBAN ?", vbQuestion + vbYesNo) = vbYes Then Me.txtBanco(9).Text = Devuelve
                    Else
                        If Mid(txtBanco(9).Text, 3) <> Devuelve Then
                            Devuelve = "Calculado : " & CamposBuscaGrid & Devuelve
                            Devuelve = "Introducido: " & Me.txtBanco(9).Text & vbCrLf & Devuelve & vbCrLf
                            Devuelve = "Error en codigo IBAN" & vbCrLf & Devuelve & "Continuar?"
                            If MsgBox(Devuelve, vbQuestion + vbYesNo) = vbNo Then SQL = SQL & vbCrLf & " Error en IBAN"
                        End If
                    End If
                End If
                CamposBuscaGrid = ""
            End If
            
            
            
        End If
    End If
    If SQL <> "" Then
        MsgBox SQL, vbExclamation
        Exit Function
    End If

    CadenaCurrency txtImporte(opcion).Text, SubImporte
    If SubImporte = 0 Then
        MsgBox "El importe no puede ser Cero", vbExclamation
        Exit Function
    End If
    
    If Importe - SubImporte < 0 Then
        MsgBox "Excede del total factura", vbExclamation
        Exit Function
    End If

    
    
    'Insertamos en cobros/pagos
    
    
    
    'Todo bien
    DatosOK = True
End Function

Private Sub PonerOpcionesMenu()
PonerOpcionesMenuGeneral Me
End Sub



Private Function SepuedeBorrar() As Boolean
Dim SQL As String
    SepuedeBorrar = False
End Function




Private Sub txtBanco_GotFocus(Index As Integer)
    PonFoco txtBanco(Index)
End Sub

Private Sub txtBanco_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub




Private Sub txtBanco_LostFocus(Index As Integer)
Dim L As Integer

    txtBanco(Index).Text = Trim(txtBanco(Index).Text)
    If txtBanco(Index).Text = "" Then Exit Sub
    
    If Index >= 8 Then
        'ES IBAN
        If Not IBAN_Correcto(Me.txtBanco(Index).Text) Then txtBanco(Index).Text = ""

    Else
    
        Devuelve = 0
        If txtBanco(Index).Text <> "" Then
            If Not IsNumeric(txtBanco(Index).Text) Then
                MsgBox "Cuenta bancaria debe ser numerica", vbExclamation
                Devuelve = 1
            Else
                Select Case Index
                Case 0, 1, 4, 5
                    L = 4
                Case 3, 7
                    L = 10
                Case Else
                    L = 2
                End Select
            End If
        End If
        
        If Devuelve = 1 Then
            txtBanco(Index).Text = ""
            txtBanco(Index).SetFocus
        Else
            txtBanco(Index).Text = Right("0000000000" & txtBanco(Index).Text, L)
            
            
            
            
            
            
            SQL = ""
            CampoDefecto = 4
            If Index < 4 Then CampoDefecto = 0
                
                
            For L = CampoDefecto To CampoDefecto + 4
                SQL = SQL & txtBanco(L).Text
            Next
            
            
            If Len(SQL) = 20 Then
                'OK. Calculamos el IBAN
                If CampoDefecto = 4 Then
                    CampoDefecto = 9
                Else
                    CampoDefecto = 8
                End If
                
                
                If txtBanco(CampoDefecto).Text = "" Then
                    'NO ha puesto IBAN
                    If DevuelveIBAN2("ES", SQL, SQL) Then txtBanco(CampoDefecto).Text = "ES" & SQL
                Else
                    Campos = CStr(Mid(txtBanco(CampoDefecto).Text, 1, 2))
                    If DevuelveIBAN2(CStr(Campos), SQL, SQL) Then
                        If Mid(txtBanco(CampoDefecto).Text, 3) <> SQL Then
                            
                            MsgBox "Codigo IBAN distinto del calculado [" & Campos & SQL & "]", vbExclamation
                            'Text1(49).Text = "ES" & SQL
                        End If
                    End If
                End If
            End If

            
            
            
            
            
            
        End If




    End If

End Sub

Private Sub txtCta_GotFocus(Index As Integer)
    PonFoco txtcta(Index)
End Sub

Private Sub txtCta_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub




Private Sub txtCta_LostFocus(Index As Integer)

    Devuelve = ""
    CampoDefecto = 0
    txtcta(Index).Text = Trim(txtcta(Index).Text)
    If txtcta(Index).Text <> "" Then
        
        If Not IsNumeric(txtcta(Index).Text) Then
            MsgBox "La cuenta debe ser numérica: " & txtcta(Index).Text, vbExclamation
            CampoDefecto = 1
        Else
    
            'DE ULTIMO NIVEL
            CamposBuscaGrid = (txtcta(Index).Text)
            If CuentaCorrectaUltimoNivel(CamposBuscaGrid, Devuelve) Then
                Me.txtcta(Index).Text = CamposBuscaGrid
                'Compruebo que existe en Cuetabanacaria
                CamposBuscaGrid = DevuelveDesdeBD("codmacta", "bancos", "codmacta", txtcta(Index).Text, "T")
                If CamposBuscaGrid = "" Then
                    MsgBox "La cuenta no esta asociada a ninguna cuenta bancaria", vbExclamation
                    Me.txtcta(Index).Text = CamposBuscaGrid
                    Devuelve = ""
                    CampoDefecto = 1
                End If
            Else
                MsgBox Devuelve, vbExclamation
                CampoDefecto = 1
                Devuelve = ""
            End If
        End If
    End If
    Me.txtDescCta(Index).Text = Devuelve
    If CampoDefecto = 1 Then PonleFoco txtcta(Index)
End Sub

Private Sub txtDpto_GotFocus(Index As Integer)
    PonFoco txtdpto(Index)
End Sub

Private Sub txtDpto_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtDpto_LostFocus(Index As Integer)
    txtdpto(Index).Text = Trim(txtdpto(Index).Text)
    Devuelve = ""
    CampoDefecto = 0
    If txtdpto(Index).Text <> "" Then
        If Not IsNumeric(txtdpto(Index).Text) Then
            MsgBox "Código departamento debe ser numérico", vbExclamation
            CampoDefecto = 1
        Else
            Campos = DevDepartamento(txtdpto(Index).Text)
            If Campos <> "" Then
                Devuelve = Campos
            Else
                MsgBox "No existe en el cliente " & Text1(0).Text & "   el departmento : " & txtdpto(Index).Text, vbExclamation
                CampoDefecto = 1
            End If
        End If
    End If
    
    Me.txtDptoDesc(Index).Text = Devuelve
    If Devuelve = "" Then txtdpto(Index).Text = ""
    If CampoDefecto <> 0 Then txtdpto(Index).SetFocus
End Sub

Private Sub txtFecha_GotFocus(Index As Integer)
    PonFoco txtFecha(Index)
End Sub



Private Sub txtfecha_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtfecha_LostFocus(Index As Integer)
    txtFecha(Index).Text = Trim(txtFecha(Index).Text)
    If txtFecha(Index).Text = "" Then Exit Sub
    If Not EsFechaOK(txtFecha(Index)) Then
        MsgBox "Fecha incorrecta: " & txtFecha(Index).Text, vbExclamation
        txtFecha(Index).Text = ""
        txtFecha(Index).SetFocus
    End If
End Sub

Private Sub txtImporte_GotFocus(Index As Integer)
    PonFoco txtImporte(Index)
End Sub

Private Sub txtImporte_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtImporte_LostFocus(Index As Integer)
    
    txtImporte(Index).Text = Trim(txtImporte(Index).Text)
    CampoDefecto = 0
    If txtImporte(Index) <> "" Then
        If Not IsNumeric(txtImporte(Index).Text) Then
            MsgBox "Campo numérico", vbExclamation
            CampoDefecto = 1
        End If
    End If
    If CampoDefecto = 0 Then
        cmdPrimer(0).SetFocus
    Else
        txtImporte(Index).SetFocus
    End If
End Sub

Private Sub txtPago_GotFocus(Index As Integer)
    PonFoco txtPago(Index)
End Sub


Private Sub txtPago_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtPago_LostFocus(Index As Integer)
   
    txtPago(Index).Text = Trim(txtPago(Index).Text)
    Devuelve = ""
    CampoDefecto = 0
    SQL = ""
    If txtPago(Index).Text <> "" Then
        If Not IsNumeric(txtPago(Index).Text) Then
            MsgBox "Forma de pago debe ser numérica", vbExclamation
            CampoDefecto = 1
        Else
            SQL = "tipforpa"
            Campos = DevuelveDesdeBD("nomforpa", "formapago", "codforpa", txtPago(Index).Text, "N", SQL)
            If Campos <> "" Then
                Devuelve = Campos
                If opcion = 0 Then
                    'COBROS. Cuenta banacaria obligada si recibo, es decir
                    ' si forma de pago es 4
                    If SQL <> "4" Then SQL = "" 'EL CUATRO ES FORMA PAGO TRANSFERENCIA
                Else
                    'PAGOS
                    ' Obligado SI transferencia
                    If SQL <> "1" Then SQL = "" 'EL UNO ES FORMA PAGO TRANSFERENCIA
                End If
            Else
                MsgBox "No existe la forma de pago: " & txtPago(Index).Text, vbExclamation
                CampoDefecto = 1
                SQL = ""
            End If
        End If
    End If
    
    'txtPago(Index).Text = ""
    FrameTapaCCC.Tag = SQL
    If Index = 1 Then
        For NumRegElim = 4 To 7
            txtBanco(NumRegElim).Enabled = (SQL <> "")
        Next NumRegElim
    End If
    FrameTapaCCC.Visible = (SQL = "")
    Me.txtPagoDesc(Index).Text = Devuelve
    If Devuelve = "" Then txtPago(Index).Text = ""
    If CampoDefecto <> 0 Then PonleFoco txtPago(Index)
    CamposBuscaGrid = ""
End Sub


Private Function DevDepartamento(ByRef T As String) As String
    Screen.MousePointer = vbHourglass
    DevDepartamento = ""
    CamposBuscaGrid = "Select * from departamentos where codmacta ='" & RecuperaValor(Datos, 5) & "' AND Dpto = " & T
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open CamposBuscaGrid, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then DevDepartamento = DBLet(miRsAux!Descripcion, "T")
    miRsAux.Close
    CamposBuscaGrid = ""
EDev:
    If Err.Number <> 0 Then MuestraError Err.Number, "Obteniendo datos Dpto."
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
End Function


Private Function InsertaPagoCobro(dinerito As Currency, Fecha As Date, Numvenci As Integer, Insertar As Boolean) As Boolean
On Error GoTo EI
    InsertaPagoCobro = False
    If Insertar Then
        Campos = "INSERT INTO tmpfaclin (codusu, codigo, Numfac, Cliente,Fecha, Total) "
                                                                                'pueden ser 3 carcateres
        Campos = Campos & " VALUES (" & vUsu.Codigo & "," & Numvenci & ",'" & Mid(RecuperaValor(Datos, 1) & "   ", 1, 3) & DevNombreSQL(RecuperaValor(Datos, 2))
        Campos = Campos & "','" & DevNombreSQL(RecuperaValor(Datos, 6)) & "','"
        'Los datos del pago
        Campos = Campos & Format(Fecha, FormatoFecha) & "'," & TransformaComasPuntos(CStr(dinerito)) & ")"
        
    Else
        'MODIFICAR
        Campos = "UPDATE tmpfaclin SET fecha = '" & Format(Fecha, FormatoFecha) & "', total ="
        Campos = Campos & TransformaComasPuntos(CStr(dinerito))
        Campos = Campos & " where codusu = " & vUsu.Codigo & " AND codigo = " & Numvenci
    End If
    Conn.Execute Campos
    Campos = ""
    InsertaPagoCobro = True
    Exit Function
EI:
    MuestraError Err.Number, "DatosOK / vencimientos"
End Function



Private Sub GenerarPagosCobros()
    
    
    'COBROS o pagos
    SQL = "Select * from tmpfaclin where codusu = " & vUsu.Codigo & " ORDER By codigo"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    If opcion = 0 Then
        '-------------------------------------
        '
        'COBROS
        '
        '-------------------------------------
        
        'Noviembre 2013.
        'Añadimos el IBAN
        
        'Este trozo es comun para todos los vencimientos
        SQL = "INSERT INTO scobro (numserie, codfaccl, fecfaccl,  codmacta,codforpa,  ctabanc1, iban,codbanco, codsucur, digcontr,"
        SQL = SQL & "cuentaba, ctabanc2, fecultco, impcobro, emitdocum, recedocu, contdocu,"
        SQL = SQL & "text33csb, ultimareclamacion, agente, departamento,codrem, anyorem, siturem, gastos,"
        SQL = SQL & "text41csb,numorden,fecvenci, impvenci) VALUES ('"
        
        ' CLIENTE ->>   Serie|codfac  |anofac|fecfac|CLIENTE|

        Campos = RecuperaValor(Datos, 1) & "'," & RecuperaValor(Datos, 2) & ",'" & Format(RecuperaValor(Datos, 4), FormatoFecha)
        Campos = Campos & "','" & RecuperaValor(Datos, 5) & "',"
        'Forma pago y ctaban y banco
        Campos = Campos & txtPago(0).Text & ",'" & Me.txtcta(0).Text & "',"
        NumRegElim = 0
        If Me.FrameTapaCCC.Tag <> "" Then
            If txtBanco(0).Text <> "" Then NumRegElim = 1
        End If
        If NumRegElim = 0 Then
            Campos = Campos & "NULL,NULL,NULL,NULL,NULL,"
        Else
            'Nov 2013. El IBAN
            Devuelve = "NULL,"
            If Me.txtBanco(8).Text <> "" Then Devuelve = "'" & txtBanco(8).Text & "',"
                
            Campos = Campos & Devuelve & Val(txtBanco(0).Text) & "," & Val(txtBanco(1).Text)
            Campos = Campos & ",'" & txtBanco(2).Text & "','" & txtBanco(3).Text & "',"
        End If
        'ctabanc2,Ultco, impcomro,emitdocum,redocum,contdocu
        Campos = Campos & "NULL,NULL,NULL,0,0,0,'"
        'Text33csb,   Ultirecla, agente
        Campos = Campos & Text1(1).Text & " - " & Text1(2).Text & "',NULL,"
        If Me.txtAgente(0).Text = "" Then
            Campos = Campos & "NULL,"
        Else
            Campos = Campos & "'" & Me.txtAgente(0).Text & "',"
        End If
        
        'Departament
        If Trim(txtdpto(0).Text) = "" Then
            Campos = Campos & "NULL"
        Else
            Campos = Campos & txtdpto(0).Text
        End If
        'codrem, anyorem, siturem, gastos,
        Campos = Campos & ",NULL,NULL,NULL,NULL"
        'Lo metemos dentro de SQL
        SQL = SQL & Campos
        
        
        'Los datos que faltan se generan a partir del RECORSET
        While Not miRsAux.EOF
            'text41csb,numorden,fecvenci, impvenci
            Campos = ",'Vto a fecha: " & Format(miRsAux!Fecha, "dd/mm/yyyy") & "'," & miRsAux!Codigo
            Campos = Campos & ",'" & Format(miRsAux!Fecha, FormatoFecha) & "'," & TransformaComasPuntos(CStr(miRsAux!Total)) & ")"
            Conn.Execute SQL & Campos
            miRsAux.MoveNext
        Wend

    Else
        '-------------------------------------
        '
        'PAGOS
        '
        '-------------------------------------
        
        ' PROVEEDOR     vacio|numregis|anofac|fecfac|PROVEEOD

        
        'Trozo comun
        SQL = "INSERT INTO spagop (ctaprove, numfactu, fecfactu,  codforpa,ctabanc1,"
        SQL = SQL & " imppagad, ctabanc2, emitdocum, fecultpa,contdocu,text1csb,"
        SQL = SQL & "iban,entidad, oficina, CC, cuentaba,"
        SQL = SQL & " text2csb,numorden,fecefect, impefect) VALUES ('"
        '- los datos
        SQL = SQL & RecuperaValor(Datos, 5) & "','" & DevNombreSQL(RecuperaValor(Datos, 2)) & "','" & Format(RecuperaValor(Datos, 4), FormatoFecha) & "',"
        'forpa ,ctabanc1
        SQL = SQL & txtPago(1).Text & ",'" & Me.txtcta(1).Text & "',"
        'impa,ctabanc2,emitdocum,fecultpa,condocu,text1csb,
        SQL = SQL & "NULL,NULL,0,NULL,0,'" & DevNombreSQL(Text1(4).Text) & " - " & Text1(5).Text & "',"
        If Me.FrameTapaCCC.Tag = "" Then
            SQL = SQL & "NULL,NULL,NULL,NULL,NULL,"
        Else
        
            'Nov 2013. El IBAN
            Devuelve = "NULL"
            If Me.txtBanco(9).Text <> "" Then Devuelve = "'" & txtBanco(9).Text & "'"
            SQL = SQL & Devuelve & ","
            
            For NumRegElim = 4 To 7
                SQL = SQL & "'" & txtBanco(NumRegElim).Text & "',"
            Next NumRegElim
        End If
        'Los datos que faltan se generan a partir del RECORSET
        While Not miRsAux.EOF
            'numorden, text2csb,fecefect, impefect
            Campos = "'Vto a fecha: " & Format(miRsAux!Fecha, "dd/mm/yyyy") & "'," & miRsAux!Codigo
            Campos = Campos & ",'" & Format(miRsAux!Fecha, FormatoFecha) & "'," & TransformaComasPuntos(CStr(miRsAux!Total)) & ")"
            Conn.Execute SQL & Campos
            miRsAux.MoveNext
        Wend

    End If
    miRsAux.Close
    Set miRsAux = Nothing
End Sub



Private Function ComprobarExistencia2() As Boolean
Dim IT As ListItem


    If opcion = 0 Then
        SQL = "Select numorden,fecvenci as f1, impvenci as i1 from scobro where numserie = '" & RecuperaValor(Datos, 1)
        SQL = SQL & "' and codfaccl = " & RecuperaValor(Datos, 2)
        SQL = SQL & " and fecfaccl = '" & Format(RecuperaValor(Datos, 4), FormatoFecha) & "'"
        Me.txtLineasCobrosPagos.Text = Text1(1).Text & "  -  " & Text1(2).Text
    Else
        SQL = "Select numorden,fecefect as f1,impefect as i1 from spagop WHERE  ctaprove ='" & RecuperaValor(Datos, 5)
        SQL = SQL & "' AND numfactu = '" & DevNombreSQL(RecuperaValor(Datos, 2))
        SQL = SQL & "' AND fecfactu = '" & Format(RecuperaValor(Datos, 4), FormatoFecha) & "'"
        Me.txtLineasCobrosPagos.Text = Text1(4).Text & "  -  " & Text1(5).Text
        
    End If
    Me.txtLineasCobrosPagos.Text = Me.txtLineasCobrosPagos.Text & "  de " & RecuperaValor(Datos, 6)
    SQL = SQL & " ORDER BY numorden"
    Set miRsAux = New ADODB.Recordset
    CampoDefecto = 0
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        CampoDefecto = 1
        Set IT = ListView1.ListItems.Add
        IT.Text = miRsAux!numorden
        IT.SubItems(1) = Format(miRsAux!F1, "dd/mm/yyyy")
        IT.SubItems(2) = Format(miRsAux!I1, FormatoImporte)
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing

    FrameBloqueo.Visible = CampoDefecto = 1
    ComprobarExistencia2 = CampoDefecto = 0
End Function


Private Function ComprobarExistencia() As Boolean

    'Primero bloqueamos
    ComprobarExistencia = False
    If BloqueoManual(True, "frmVto", Text1(1).Text & Text1(2).Text) Then
        FrameEstaBloq.Visible = False
        ComprobarExistencia = ComprobarExistencia2
    Else
        FrameEstaBloq.Visible = True
    End If
End Function


Private Sub PonerCuentaBanco_Forpa()
    On Error GoTo EPonerCuentaBanco_Forpa
    
    SQL = RecuperaValor(Datos, 5)
    If opcion = 0 Then
        CampoDefecto = 0
    Else
        CampoDefecto = 4
    End If
    SQL = "Select entidad,oficina,CC,cuentaba,forpa,ctabanco,iban from cuentas where codmacta='" & SQL & "'"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        For NumRegElim = 0 To 3
            txtBanco(NumRegElim + CampoDefecto).Text = DBLet(miRsAux.Fields(NumRegElim), "T")
        Next NumRegElim
        
        txtPago(opcion).Text = DBLet(miRsAux!Forpa, "T")
        txtcta(opcion).Text = DBLet(miRsAux!CtaBanco, "T")
        'Nov. 2013
        txtBanco(8 + opcion) = DBLet(miRsAux!IBAN, "T")
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    If txtPago(opcion).Text <> "" Then txtPago_LostFocus CInt(opcion)
    If txtcta(opcion).Text <> "" Then txtCta_LostFocus CInt(opcion)
    Exit Sub
EPonerCuentaBanco_Forpa:
    MuestraError Err.Number, Err.Description
End Sub
