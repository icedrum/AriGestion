VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFacturasCli 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   10890
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   17655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10890
   ScaleWidth      =   17655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   4050
      Index           =   0
      Left            =   270
      TabIndex        =   70
      Top             =   870
      Width           =   17160
      Begin VB.CheckBox chkContab 
         Caption         =   "Contabilizada"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   9720
         TabIndex        =   88
         Tag             =   "Inc|N|N|0||factcli|intconta|||"
         Top             =   600
         Width           =   1785
      End
      Begin VB.TextBox Text1 
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
         Index           =   5
         Left            =   7950
         TabIndex        =   4
         Tag             =   "Forma de pago|N|N|||factcli|codforpa|000||"
         Text            =   "1234567890"
         Top             =   1260
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FCFCE2&
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
         Index           =   4
         Left            =   210
         TabIndex        =   3
         Tag             =   "Cliente|N|N|||factcli|codclien|0000||"
         Text            =   "1234567890"
         Top             =   1260
         Width           =   1350
      End
      Begin VB.TextBox Text1 
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
         Left            =   5160
         MaxLength       =   30
         TabIndex        =   1
         Tag             =   "Fecha|F|N|||factcli|fecfactu|dd/mm/yyyy|N|"
         Top             =   570
         Width           =   1275
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FCFCE2&
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
         Left            =   6645
         TabIndex        =   2
         Tag             =   "Nº factura|N|S|0||factcli|numfactu|0000000|S|"
         Top             =   570
         Width           =   1395
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Left            =   240
         MaxLength       =   3
         TabIndex        =   0
         Tag             =   "Serie|T|N|||factcli|numserie||S|"
         Text            =   "123"
         Top             =   570
         Width           =   510
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Index           =   3
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   76
         Tag             =   "Observaciones|T|S|||factcli|observa|||"
         Top             =   2040
         Width           =   8775
      End
      Begin VB.TextBox Text4 
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
         Index           =   2
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   75
         Text            =   "Text4"
         Top             =   570
         Width           =   4245
      End
      Begin VB.TextBox Text4 
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
         Index           =   4
         Left            =   1650
         Locked          =   -1  'True
         TabIndex        =   74
         Text            =   "Text4"
         Top             =   1260
         Width           =   6135
      End
      Begin VB.TextBox Text4 
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
         Index           =   5
         Left            =   9420
         Locked          =   -1  'True
         TabIndex        =   73
         Text            =   "Text4"
         Top             =   1260
         Width           =   7425
      End
      Begin VB.ComboBox Combo1 
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
         ItemData        =   "frmFacturasCli.frx":0000
         Left            =   9660
         List            =   "frmFacturasCli.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "Tipo retencion|N|N|||factcli|tiporeten|||"
         Top             =   2310
         Width           =   4560
      End
      Begin VB.TextBox Text4 
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
         Index           =   6
         Left            =   11160
         Locked          =   -1  'True
         TabIndex        =   72
         Text            =   "Text4"
         Top             =   3240
         Width           =   4785
      End
      Begin VB.TextBox Text1 
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
         Index           =   6
         Left            =   9690
         TabIndex        =   7
         Tag             =   "Cuenta Retencion|T|S|||factcli|cuereten|||"
         Text            =   "1234567890"
         Top             =   3270
         Width           =   1350
      End
      Begin VB.TextBox Text1 
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
         Index           =   7
         Left            =   14700
         TabIndex        =   6
         Tag             =   "Porcentaje Retencion|N|S|||factcli|retfaccl|##0.00||"
         Text            =   "1234567890"
         Top             =   2280
         Width           =   1230
      End
      Begin VB.TextBox Text1 
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
         Index           =   13
         Left            =   8160
         TabIndex        =   71
         Tag             =   "Exp|F|S|||factcli|fecexped|||"
         Top             =   570
         Width           =   1350
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   6
         Left            =   9120
         Top             =   240
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Serie"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   86
         Top             =   270
         Width           =   555
      End
      Begin VB.Label Label4 
         Caption         =   "Nº Factura"
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
         Left            =   6660
         TabIndex        =   85
         Top             =   270
         Width           =   1140
      End
      Begin VB.Label Label18 
         Caption         =   "Fecha"
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
         Left            =   5190
         TabIndex        =   84
         Top             =   270
         Width           =   930
      End
      Begin VB.Label Label8 
         Caption         =   "Observaciones"
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
         Left            =   240
         TabIndex        =   83
         Top             =   1800
         Width           =   1515
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   1
         Left            =   840
         Top             =   270
         Width           =   240
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   0
         Left            =   6150
         Picture         =   "frmFacturasCli.frx":0004
         Top             =   270
         Width           =   240
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   2
         Left            =   960
         Top             =   990
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Forma de Pago"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   7950
         TabIndex        =   81
         Top             =   990
         Width           =   1545
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   3
         Left            =   9510
         Top             =   990
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Retención"
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
         Left            =   9660
         TabIndex        =   80
         Top             =   2040
         Width           =   1380
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   4
         Left            =   11580
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta Retención"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   9690
         TabIndex        =   79
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "% Retención"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   14700
         TabIndex        =   78
         Top             =   2040
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Expedida"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   8160
         TabIndex        =   77
         Top             =   270
         Width           =   870
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   5
         Left            =   1800
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   210
         TabIndex        =   82
         Top             =   990
         Width           =   675
      End
   End
   Begin VB.Frame Frame4 
      Height          =   2145
      Left            =   9690
      TabIndex        =   60
      Top             =   4920
      Width           =   7725
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FCFCE2&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Index           =   12
         Left            =   5640
         TabIndex        =   24
         Tag             =   "Total Factura|N|S|||factcli|totfaccl|###,###,##0.00||"
         Text            =   "123456789012345"
         Top             =   1590
         Width           =   1935
      End
      Begin VB.TextBox Text1 
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
         Index           =   11
         Left            =   5640
         TabIndex        =   23
         Tag             =   "Importe Retención|N|S|||factcli|trefaccl|###,###,##0.00||"
         Text            =   "123456789012345"
         Top             =   1050
         Width           =   1935
      End
      Begin VB.TextBox Text1 
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
         Index           =   10
         Left            =   1740
         TabIndex        =   22
         Tag             =   "Base Retención|N|S|||factcli|totbasesret|###,###,##0.00||"
         Text            =   "123456789012345"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text1 
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
         Index           =   9
         Left            =   5640
         TabIndex        =   21
         Tag             =   "Importe Iva|N|S|||factcli|totivas|###,###,##0.00||"
         Text            =   "123456789012345"
         Top             =   570
         Width           =   1935
      End
      Begin VB.TextBox Text1 
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
         Index           =   8
         Left            =   1740
         TabIndex        =   20
         Tag             =   "Base Imponible|N|S|||factcli|totbases|###,###,##0.00||"
         Text            =   "123456789012345"
         Top             =   570
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "TOTAL FACTURA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   13
         Left            =   3780
         TabIndex        =   66
         Top             =   1650
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Retención"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   12
         Left            =   3780
         TabIndex        =   65
         Top             =   1110
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Base Retención"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   11
         Left            =   180
         TabIndex        =   64
         Top             =   1140
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "Importe IVA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   10
         Left            =   3780
         TabIndex        =   63
         Top             =   630
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "Base Imponible"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   9
         Left            =   180
         TabIndex        =   62
         Top             =   630
         Width           =   1635
      End
      Begin VB.Label Label7 
         Caption         =   "Totales Factura"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   61
         Top             =   210
         Width           =   1980
      End
   End
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3930
      TabIndex        =   58
      Top             =   90
      Width           =   1815
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   210
         TabIndex        =   59
         Top             =   180
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Datos Fiscales"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Cobros"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Errores NºFactura"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameFiltro 
      Height          =   705
      Left            =   9960
      TabIndex        =   56
      Top             =   90
      Width           =   2445
      Begin VB.ComboBox cboFiltro 
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
         ItemData        =   "frmFacturasCli.frx":008F
         Left            =   120
         List            =   "frmFacturasCli.frx":009C
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   210
         Width           =   2235
      End
   End
   Begin VB.Frame FrameAux2 
      Height          =   2145
      Left            =   270
      TabIndex        =   46
      Top             =   4920
      Width           =   9375
      Begin VB.TextBox txtaux3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   10
         Left            =   8160
         TabIndex        =   69
         Tag             =   "Importe Rec|N|S|||factcli_totales|imporec|###,###,##0.00||"
         Text            =   "ImpRec"
         Top             =   1590
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtaux3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   9
         Left            =   7260
         TabIndex        =   68
         Tag             =   "Importe Iva|N|S|||factcli_totales|impoiva|###,###,##0.00||"
         Text            =   "ImpIva"
         Top             =   1590
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtaux3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   8
         Left            =   6390
         TabIndex        =   67
         Tag             =   "%Ret|N|S|||factcli_totales|porcrec|##0.00||"
         Text            =   "PorRec"
         Top             =   1620
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   7
         Left            =   5550
         TabIndex        =   54
         Tag             =   "%Iva|N|S|||factcli_totales|porciva|##0.00||"
         Text            =   "PorIva"
         Top             =   1620
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   6
         Left            =   4800
         TabIndex        =   53
         Tag             =   "Iva|N|S|||factcli_totales|codigiva|000||"
         Text            =   "Iva"
         Top             =   1620
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   5
         Left            =   4080
         TabIndex        =   52
         Tag             =   "Base Imponible|N|S|||factcli_totales|baseimpo|###,###,##0.00||"
         Text            =   "Base Imponible"
         Top             =   1620
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   4
         Left            =   3330
         TabIndex        =   51
         Tag             =   "Linea|N|N|||factcli_totales|numlinea|||"
         Text            =   "Linea"
         Top             =   1620
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   3
         Left            =   2580
         TabIndex        =   50
         Tag             =   "Año factura|N|N|||factcli_totales|anofactu||S|"
         Text            =   "Año"
         Top             =   1620
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   2
         Left            =   1800
         TabIndex        =   49
         Tag             =   "Fecha|F|N|||factcli_totales|fecfactu|dd/mm/yyyy||"
         Text            =   "Fecha"
         Top             =   1620
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   1
         Left            =   1110
         TabIndex        =   48
         Tag             =   "Nº factura|N|N|0||factcli_totales|numfactu|000000|S|"
         Text            =   "Factura"
         Top             =   1620
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   0
         Left            =   330
         TabIndex        =   47
         Tag             =   "Nº Serie|T|S|||factcli_totales|numserie||S|"
         Text            =   "Serie"
         Top             =   1620
         Visible         =   0   'False
         Width           =   645
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   360
         Top             =   240
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
      Begin MSComctlLib.ListView lw1 
         Height          =   1545
         Left            =   150
         TabIndex        =   55
         Top             =   510
         Width           =   9045
         _ExtentX        =   15954
         _ExtentY        =   2725
         View            =   3
         LabelEdit       =   1
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
         NumItems        =   0
      End
      Begin VB.Label Label3 
         Caption         =   "Desglose Factura"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   57
         Top             =   210
         Width           =   1980
      End
   End
   Begin VB.CheckBox chkVistaPrevia 
      Caption         =   "Vista previa"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   13950
      TabIndex        =   42
      Top             =   270
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   5970
      TabIndex        =   40
      Top             =   90
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   41
         Top             =   210
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Primero"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Anterior"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Siguiente"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Último"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   240
      TabIndex        =   38
      Top             =   90
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   39
         Top             =   180
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Nuevo"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Eliminar"
               Object.Tag             =   "2"
               Object.Width           =   1e-4
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Buscar"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ver Todos"
               Object.Tag             =   "0"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Imprimir"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Salir"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameAux1 
      BorderStyle     =   0  'None
      Height          =   3060
      Left            =   285
      TabIndex        =   34
      Top             =   7125
      Width           =   17190
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   6960
         TabIndex        =   87
         ToolTipText     =   "Buscar cuenta"
         Top             =   2160
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtaux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   13
         Left            =   6480
         MultiLine       =   -1  'True
         TabIndex        =   10
         Tag             =   "A|T|S|||factcli_lineas|ampliaci|||"
         Text            =   "frmFacturasCli.frx":00D3
         Top             =   2160
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   6
         Left            =   7920
         MaxLength       =   15
         TabIndex        =   12
         Tag             =   "Importe|N|N|||factcli_lineas|precio|###,###,##0.00||"
         Text            =   "precio"
         Top             =   2160
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txtaux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   12
         Left            =   9480
         MaxLength       =   15
         TabIndex        =   18
         Tag             =   "Importe|N|N|||factcli_lineas|importe|###,###,##0.00||"
         Text            =   "Importe"
         Top             =   2160
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.CheckBox chkAux2 
         BackColor       =   &H80000005&
         Height          =   255
         Index           =   0
         Left            =   15480
         TabIndex        =   15
         Tag             =   "Aplica Retencion|N|N|0|1|factcli_lineas|aplicret|||"
         Top             =   2160
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox txtaux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   11
         Left            =   14040
         MaxLength       =   15
         TabIndex        =   32
         Tag             =   "Importe Rec|N|S|||factcli_lineas|imporec|###,###,##0.00||"
         Text            =   "ImpRec"
         Top             =   2040
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txtaux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   10
         Left            =   13200
         MaxLength       =   15
         TabIndex        =   31
         Tag             =   "Importe Iva|N|S|||factcli_lineas|impoiva|###,###,##0.00||"
         Text            =   "ImpIva"
         Top             =   2160
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txtaux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   9
         Left            =   12360
         MaxLength       =   15
         TabIndex        =   30
         Tag             =   "% Recargo|N|S|||factcli_lineas|porcrec|##0.00||"
         Text            =   "%rec"
         Top             =   2160
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txtaux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   8
         Left            =   11640
         MaxLength       =   50
         TabIndex        =   17
         Tag             =   "% Iva|N|S|||factcli_lineas|porciva|##0.00||"
         Text            =   "%iva"
         Top             =   2160
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox txtaux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   5
         Left            =   7320
         MaxLength       =   10
         TabIndex        =   11
         Tag             =   "Cantidad|N|N|||factcli_lineas|cantidad|###,###,##0.00||"
         Text            =   "cantidad"
         Top             =   2160
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   7
         Left            =   10680
         MaxLength       =   15
         TabIndex        =   16
         Tag             =   "Codigo Iva|N|N|||factcli_lineas|codigiva|000||"
         Text            =   "Iva"
         Top             =   2160
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Frame FrameToolAux 
         Height          =   555
         Left            =   60
         TabIndex        =   44
         Top             =   0
         Width           =   1605
         Begin MSComctlLib.Toolbar ToolbarAux 
            Height          =   330
            Left            =   180
            TabIndex        =   45
            Top             =   150
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            Style           =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Insertar"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Modificar"
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Eliminar"
               EndProperty
            EndProperty
         End
      End
      Begin VB.TextBox txtaux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   4
         Left            =   4050
         MaxLength       =   15
         TabIndex        =   9
         Tag             =   "Cuenta|T|N|||factcli_lineas|codconce|||"
         Text            =   "clie"
         Top             =   2160
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   3
         Left            =   2880
         TabIndex        =   29
         Tag             =   "Linea|N|N|||factcli_lineas|numlinea||S|"
         Text            =   "linea"
         Top             =   2160
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   2
         Left            =   2220
         TabIndex        =   28
         Tag             =   "Fecha|F|N|||factcli_lineas|fecfactu|dd/mm/yyyy|S|"
         Text            =   "fecha"
         Top             =   2160
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   0
         Left            =   105
         TabIndex        =   8
         Tag             =   "Nº Serie|T|S|||factcli_lineas|numserie||S|"
         Text            =   "Serie"
         Top             =   2145
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   1
         Left            =   840
         MaxLength       =   10
         TabIndex        =   27
         Tag             =   "Nº factura|N|N|0||factcli_lineas|numfactu|000000|S|"
         Text            =   "factura"
         Top             =   2145
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   4800
         TabIndex        =   36
         ToolTipText     =   "Buscar cuenta"
         Top             =   2190
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   4
         Left            =   5040
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   35
         Tag             =   "Cuenta|T|N|||factcli_lineas|nomconce|||"
         Text            =   "Nombre cli"
         Top             =   2190
         Visible         =   0   'False
         Width           =   1365
      End
      Begin MSAdodcLib.Adodc AdoAux 
         Height          =   375
         Index           =   1
         Left            =   3720
         Top             =   480
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
         Caption         =   "AdoAux(1)"
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
      Begin MSDataGridLib.DataGrid DataGridAux 
         Height          =   2400
         Index           =   1
         Left            =   120
         TabIndex        =   37
         Top             =   600
         Width           =   16770
         _ExtentX        =   29580
         _ExtentY        =   4233
         _Version        =   393216
         AllowUpdate     =   0   'False
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   19
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
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
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   285
      TabIndex        =   25
      Top             =   10290
      Width           =   2865
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   26
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
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
      Left            =   16380
      TabIndex        =   14
      Top             =   10350
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
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
      Left            =   15090
      TabIndex        =   13
      Top             =   10350
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   2640
      Top             =   10320
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
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
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   390
      Left            =   17040
      TabIndex        =   43
      Top             =   240
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ayuda"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   8190
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
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
      Left            =   16380
      TabIndex        =   33
      Top             =   10350
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver Todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         HelpContextID   =   2
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   2
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         HelpContextID   =   2
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
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
Attribute VB_Name = "frmFacturasCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


'Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
'Public Event DatoSeleccionado(CadenaSeleccion As String)
Public FACTURA As String  'Con pipes numserie|numfactu|anofactu



Public Datos As String  'Tendra, empipado, numero asiento  y demas

Private Const NO = "No encontrado"

Private Const IdPrograma = ID_FacturasEmitidas


Private WithEvents frmFPag As frmBasico
Attribute frmFPag.VB_VarHelpID = -1
Private WithEvents frmVario As frmBasico
Attribute frmVario.VB_VarHelpID = -1


Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1

Private frmMens As frmMensajes
Attribute frmMens.VB_VarHelpID = -1


Dim AntiguoText1 As String
Private CadenaAmpliacion As String
Private Sql As String


Dim PosicionGrid As Integer

Dim Linliapu As Long
Dim FicheroAEliminar As String



Private Modo As Byte
'*************** MODOS ********************
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'   5.-  Manteniment Llinies

'+-+-Variables comuns a tots els formularis+-+-+

Dim ModoLineas As Byte
'1.- Afegir,  2.- Modificar,  3.- Borrar,  0.-Passar control a Llínies

Dim NumTabMto As Integer 'Indica quin nº de Tab està en modo Mantenimient
Dim TituloLinea As String 'Descripció de la llínia que està en Mantenimient
Dim PrimeraVez As Boolean

Private CadenaConsulta As String 'SQL de la taula principal del formulari
Private Ordenacion As String
Private NombreTabla As String  'Nom de la taula
Private NomTablaLineas As String 'Nom de la Taula de llínies del Mantenimient en que estem

Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

'Private VieneDeBuscar As Boolean
'Per a quan torna 2 poblacions en el mateix codi Postal. Si ve de pulsar prismatic
'de búsqueda posar el valor de població seleccionada i no tornar a recuperar de la Base de Datos

Dim btnPrimero As Byte 'Variable que indica el nº del Botó PrimerRegistro en la Toolbar1
'Dim CadAncho() As Boolean  'array, per a quan cridem al form de llínies
Dim indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos

Dim CadB As String
Dim CadB1 As String
Dim CadB2 As String

Dim PulsadoSalir As Boolean
Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas

Dim B As Boolean

Private BuscaChekc As String

Dim VarieAnt As String


Dim IT As ListItem


Dim cadFiltro As String
Dim I As Long
Dim Ancho As Integer


Private ModificandoLineas As Byte
'0.- A la espera 1.- Insertar   2.- Modificar

'Por si esta en un periodo liquidado, que pueda modificar CONCEPTO , cuentas,
Private ModificaFacturaPeriodoLiquidado As Boolean


Dim IvaCuenta As String
Dim CambiarIva As Boolean

Dim CtaBanco As String
Dim IBAN As String
Dim NomBanco As String

Dim TipForpa As Integer
Dim AntLetraSer As String





Private Sub cboFiltro_Click()
    If PrimeraVez Then Exit Sub
    If Modo = 0 Then Exit Sub
    HacerBusqueda2
End Sub




Private Sub chkContab_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "chkContab(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "chkContab(" & Index & ")|"
    End If
End Sub

Private Sub chkContab_GotFocus(Index As Integer)
    ConseguirFocoChk Modo
End Sub

Private Sub chkContab_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpress KeyAscii
End Sub


Private Sub cmdAceptar_Click()
    Dim Cad As String
    Dim I As Integer
    Dim Limp As Boolean
    Dim B As Boolean

    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BÚSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOK Then
                If InsertarDesdeForm2(Me, 1) Then
                
                    
                    
                
                    'Primera cosa a hacer
                    'ACutalizar contador
                    Cad = DevuelveDesdeBD("numfactu", "contadores", "serfactur", Text1(2).Text, "T")
                    If Val(Cad) = Val(Text1(0).Text) Then
                        'OKKK. hay que incrementar en uno la factura
                        Cad = "UPDATE contadores set numfactu=" & CLng(Cad) + 1
                        Cad = Cad & " WHERE serfactur = " & DBSet(Text1(2).Text, "T")
                        Ejecuta Cad
                    End If
                
                
                    'Segundo. Cobros
                    InsertarEnCobros
                
                    'Situar el adoppal
                    If SituarData1() Then
                        Modo = 2 'para que no inserte de nuevo
                        BotonAnyadirLinea 1, True
                    Else
                        PonerModo 0
                    End If
                End If
            End If
        Case 4  'MODIFICAR
            If DatosOK Then
                '-----------------------------------------
                'Hay que comprobar si ha modificado, o no la clave de la factura
                I = 1
                If Data1.Recordset!numserie = Text1(2).Text Then
                    If Data1.Recordset!NumFactu = CLng(Text1(0).Text) Then
                        If Data1.Recordset!FecFactu = Text1(1).Text Then
                            I = 0
                            'NO HA MODIFICADO NADA
                        End If
                    End If
                End If
            
                'Hacemos MODIFICAR
                Dim RC As Boolean
                If I <> 0 Then
                    MsgBox "No se puede cambiar campos clave  de la factura.", vbExclamation
                    RC = False
                Else
                    RC = ModificarFactura
                End If
                    
                If RC Then
                    '--DesBloqueaRegistroForm Me.Text1(0)
                    TerminaBloquear
                    
                    'LOG
                    vLog.Insertar 5, vUsu, "Factura : " & Text1(2).Text & Text1(0).Text & " " & Text1(1).Text
                    'Creo que no hace falta volver a situar el datagrid
                    'If SituarData1(0) Then
                    PosicionarData
                    
                End If
            End If
        
        Case 5 'LLÍNIES
            
            
            Select Case ModoLineas
                Case 1 'afegir llínia
                    InsertarLinea
                Case 2 'modificar llínies
                    ModificarLinea
                                        
                    '**** parte de contabilizacion de la factura
                    TerminaBloquear
                    
                    
                    PosicionarData
                    
            End Select

    
    End Select
    Screen.MousePointer = vbDefault
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Function ObtenerWhereCP(conWhere As Boolean) As String
Dim Sql As String

    On Error Resume Next
    
    Sql = "numserie= " & DBSet(Text1(2).Text, "T") & " and numfactu = " & DBSet(Text1(0).Text, "N") & " and anofactu = " & DBSet(Text1(14).Text, "N")
    If conWhere Then Sql = " WHERE " & Sql
    ObtenerWhereCP = Sql
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Obteniendo cadena WHERE.", Err.Description
End Function


Private Sub cmdAux_Click(Index As Integer)
        
        If Index = 0 Then
        
        
        
        
        Else
            'Observaciones
            If Modo <> 2 And Modo <> 5 Then Exit Sub
            If txtAux(13).Visible Then
                CadenaConsulta = txtAux(13).Text
            Else
                If Me.AdoAux(1).Recordset.EOF Then Exit Sub
                CadenaConsulta = DBLet(Me.AdoAux(1).Recordset!ampliaci, "T")
             
                
            End If
            
            Set frmZ = New frmZoom
            
            frmZ.pValor = CStr(CadenaConsulta)
            CadenaConsulta = ""
            If txtAux(13).Visible Then
                frmZ.pModo = 3
            Else
                frmZ.pModo = Modo
            End If
            frmZ.Caption = "Ampliacion linea factura"
            frmZ.Show vbModal
            Set frmZ = Nothing
            If txtAux(13).Visible Then
                If CadenaConsulta <> "" Then txtAux(13).Text = CadenaConsulta
            End If
        
        
        
        End If
        
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub DataGridAux_DblClick(Index As Integer)
    cmdAux_Click 1
End Sub

Private Sub Form_Activate()
    
    If PrimeraVez Then
        B = False
        If FACTURA <> "" Then
            B = True
            Modo = 2
            Sql = "Select * from factcli "
            Sql = Sql & " WHERE numserie = " & DBSet(RecuperaValor(FACTURA, 1), "T")
            Sql = Sql & " AND numfactu =" & RecuperaValor(FACTURA, 2)
            Sql = Sql & " AND fecfactu= " & DBSet(RecuperaValor(FACTURA, 3), "F")
            CadenaConsulta = Sql
            PonerCadenaBusqueda
            'BOTON lineas
            
            cboFiltro.ListIndex = 0
            
        Else
            Modo = 0
            CadenaConsulta = "Select * from " & NombreTabla & " WHERE numserie is null"
            Data1.RecordSource = CadenaConsulta
            Data1.Refresh
            
            cboFiltro.ListIndex = vUsu.FiltroFactCli
            
        End If
        
        CargarSqlFiltro
        
        PonerModo CInt(Modo)

        If Modo <> 2 Then
            
            'ESTO LO HE CAMBIADO HOY 9 FEB 2006
            'Antes no estaba el IF
            If FACTURA <> "" Then
                MsgBox "Proceso de sistema. Stop. Frm_Activate"
            End If
        Else

        End If
  
        Toolbar1.Enabled = True
        
        PrimeraVez = False
        
        
    End If
    Screen.MousePointer = vbDefault
    
    
End Sub

Private Sub CargarSqlFiltro()

    Screen.MousePointer = vbHourglass
    
    cadFiltro = ""
    
    Select Case Me.cboFiltro.ListIndex
        Case 0 ' sin filtro
            cadFiltro = "(1=1)"
        
        Case 1 ' ejercicios abiertos
            cadFiltro = "factcli.fecfactu >= " & DBSet("01/01/1900", "F")
        
        Case 2 ' ejercicio actual
            cadFiltro = "factcli.fecfactu between " & DBSet("01/01/1900", "F") & " and " & DBSet("01/01/2200", "F")
        
        Case 3 ' ejercicio siguiente
            cadFiltro = "factcli.fecfactu > " & DBSet("01/01/1900", "F")
    
    End Select
    
    Screen.MousePointer = vbDefault


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Modo > 2 Then Cancel = 1

    Screen.MousePointer = vbDefault
    
    vUsu.ActualizarFiltro "ariconta", IdPrograma, Me.cboFiltro.ListIndex
    
End Sub

Private Sub Form_Load()
Dim I As Integer

    Me.Icon = frmppal.Icon

    LimpiarCampos
    PrimeraVez = True
    PulsadoSalir = False
    CadAncho = False

    
    ' Botonera Principal
    With Me.Toolbar1
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
        .Buttons(5).Image = 1
        .Buttons(6).Image = 2
        .Buttons(8).Image = 16
    End With

    ' Botonera Principal 2
    With Me.Toolbar2
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 47
        .Buttons(2).Image = 44
        .Buttons(3).Image = 42
    End With


    ' desplazamiento
    With Me.ToolbarDes
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 6
        .Buttons(2).Image = 7
        .Buttons(3).Image = 8
        .Buttons(4).Image = 9
    End With
   
    With Me.ToolbarAux
        .HotImageList = frmppal.imgListComun_OM16
        .DisabledImageList = frmppal.imgListComun_BN16
        .ImageList = frmppal.imgListComun16
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
    End With
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
    
    For I = 0 To imgppal.Count - 1
        If I <> 0 And I <> 7 Then imgppal(I).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next I

    CargaFiltros
    
    
    Caption = "Facturas de Cliente"
    
    NumTabMto = 1
    
    LimpiarCampos   'Neteja els camps TextBox
'    ' ******* si n'hi han llínies *******
    DataGridAux(1).ClearFields
    
    '*** canviar el nom de la taula i l'ordenació de la capçalera ***
    NombreTabla = "factcli"
    Ordenacion = " ORDER BY factcli.numserie, factcli.numfactu , factcli.fecfactu"
    '************************************************
    
    'Mirem com està guardat el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = Conn
    '***** canviar el nom de la PK de la capçalera; repasar codEmpre *************
    Data1.RecordSource = "Select * from " & NombreTabla & " where numserie is null"
    Data1.Refresh
       
    
    ModoLineas = 0

       
    CargarColumnas
    
    CargarCombo
    
    
    PonerModoUsuarioGnral 0, "ariconta"
    
    'Maxima longitud cuentas
    txtAux(5).MaxLength = vEmpresa.DigitosUltimoNivel
    PulsadoSalir = False

End Sub

Private Sub CargarColumnas()
Dim Columnas As String
Dim Ancho As String
Dim Alinea As String
Dim Formato As String
Dim Ncol As Integer
Dim C As ColumnHeader

    Columnas = "Linea|Tipo|Descripcion|Base|IVA|Recargo|"
    Ancho = "0|800|2450|1800|1800|1800|"
    'vwColumnRight =1  left=0   center=2
    Alinea = "0|0|0|1|1|1|"
    'Formatos
    Formato = "|||###,###,##0.00|###,###,##0.00|###,###,##0.00|"
    Ncol = 6

    lw1.Tag = "5|" & Ncol & "|"
    
    lw1.ColumnHeaders.Clear
    
    For NumRegElim = 1 To Ncol
         Set C = lw1.ColumnHeaders.Add()
         C.Text = RecuperaValor(Columnas, CInt(NumRegElim))
         C.Width = RecuperaValor(Ancho, CInt(NumRegElim))
         C.Alignment = Val(RecuperaValor(Alinea, CInt(NumRegElim)))
         C.Tag = RecuperaValor(Formato, CInt(NumRegElim))
    Next NumRegElim


End Sub

Private Sub LimpiarCampos()
Dim I As Integer

    On Error Resume Next
    
    Limpiar Me   'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    
   
    Combo1.ListIndex = -1
 

    Me.chkAux2(0).Value = 0 + 6
    
    Me.chkContab(0).Value = 0
    lw1.ListItems.Clear
    
    
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub LimpiarCamposLin(FrameAux As String)
    On Error Resume Next
    
    LimpiarLin Me, FrameAux  'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""

    If Err.Number <> 0 Then Err.Clear
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO s'habiliten, o no, els diversos camps del
'   formulari en funció del modo en que anem a treballar
Private Sub PonerModo(Kmodo As Byte, Optional indFrame As Integer)
Dim I As Integer, NumReg As Byte
Dim B As Boolean

    On Error GoTo EPonerModo
 
    'Actualisa Iconos Insertar,Modificar,Eliminar
    'ActualizarToolbar Modo, Kmodo
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo, ModoLineas
    
    BuscaChekc = ""
       
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)
    If Not Data1.Recordset Is Nothing Then
        DespalzamientoVisible B And (Data1.Recordset.RecordCount > 1)
    End If
    
    Toolbar1.Buttons(8).Enabled = B
    
    B = Modo = 2 Or Modo = 0 Or Modo = 5
    
    For I = 0 To Text1.Count - 1
        Text1(I).Locked = B
        If Modo <> 1 Then
            Text1(I).BackColor = vbWhite
        End If
    Next I
    If Modo = 4 Then
        BloquearTxt Text1(0), True
        BloquearTxt Text1(1), True
        BloquearTxt Text1(2), True
    End If
        
   
    Combo1.Locked = B
   
    
    For I = 0 To imgppal.Count - 1
        imgppal(I).Enabled = Not B
    Next I
    imgppal(6).Enabled = (Text1(8).Text <> "")
    
    ' observaciones
    imgppal(6).Enabled = (Data1.Recordset.RecordCount > 1)
    
    
    B = Modo <> 0 And Modo <> 2
    cmdCancelar.Visible = B
    cmdAceptar.Visible = B
       
    PonerOpcionesMenuGeneral Me
    PonerModoUsuarioGnral Modo, "ariconta"
    
    B = (Modo < 5)
    chkVistaPrevia.Visible = B
    
    Text1(0).Enabled = (Modo = 1 Or Modo = 3)
    
    
    B = (Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) And (NumTabMto = 0))
            
    chkVistaPrevia.Enabled = (Modo <= 2)
    
    PonerLongCampos

    If (Modo < 2) Or (Modo = 3) Then
        CargaGrid 1, False
    End If
    
    B = (Modo = 4) Or (Modo = 2)
    
    DataGridAux(1).Enabled = B
        
    'lineas de factura
    Dim anc As Single
    anc = DataGridAux(1).top
    If DataGridAux(1).Row < 0 Then
        anc = anc + 230
    Else
        anc = anc + DataGridAux(1).RowTop(DataGridAux(1).Row) + 5
    End If
    If Modo = 1 Then
        LLamaLineas 1, Modo, anc
    Else
        LLamaLineas 1, 3, anc
    End If
    
    For I = 0 To txtAux.Count - 1
        txtAux(I).BackColor = vbWhite
    Next I
    
    Frame4.Enabled = (Modo = 1)
    
    
    'txtaux(8).Enabled = (Modo = 1)
    txtAux(9).Enabled = (Modo = 1)
    
    
    ' ponemos en azul clarito
    Text1(0).BackColor = vbMoreLightBlue  ' factura
    Text1(12).BackColor = vbMoreLightBlue ' total factura
    Text1(4).BackColor = vbMoreLightBlue ' codmacta del cliente
    
    
    PonerModoUsuarioGnral Modo, "arigestion"

EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub DespalzamientoVisible(bol As Boolean)
    FrameDesplazamiento.Visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los TEXT1
    PonerLongCamposGnral Me, Modo, 1
End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub

Private Sub Desplazamiento(Index As Integer)
    If Data1.Recordset.EOF Then Exit Sub
    
    Select Case Index
        Case 1
            Data1.Recordset.MoveFirst
        Case 2
            Data1.Recordset.MovePrevious
            If Data1.Recordset.BOF Then Data1.Recordset.MoveFirst
        Case 3
            Data1.Recordset.MoveNext
            If Data1.Recordset.EOF Then Data1.Recordset.MoveLast
        Case 4
            Data1.Recordset.MoveLast
    End Select
    PonerCampos
End Sub

Private Function MontaSQLCarga(Index As Integer, Enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basant-se en la informació proporcionada pel vector de camps
'   crea un SQl per a executar una consulta sobre la base de datos que els
'   torne.
' Si ENLAZA -> Enlaça en el data1
'           -> Si no el carreguem sense enllaçar a cap camp
'--------------------------------------------------------------------
Dim Sql As String
Dim tabla As String
    
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
        Case 0 ' lineas de totales
            tabla = "factcli_totales"
            Sql = "SELECT factcli_totales.numserie, factcli_totales.numfactu, factcli_totales.fecfactu, factcli_totales.anofactu, factcli_totales.numlinea, factcli_totales.baseimpo, factcli_totales.codigiva, factcli_totales.porciva,"
            Sql = Sql & " factcli_totales.porcrec, factcli_totales.impoiva, factcli_totales.imporec "
            Sql = Sql & " FROM " & tabla
            If Enlaza Then
                Sql = Sql & Replace(ObtenerWhereCab(True), "factcli", "factcli_totales")
            Else
                Sql = Sql & " WHERE factcli_totales.numlinea is null"
            End If
            Sql = Sql & " ORDER BY 1,2,3,4,5"
            
       
       
       Case 1 ' lineas de facturas
            tabla = "factcli_lineas"
            Sql = "SELECT factcli_lineas.numserie, factcli_lineas.numfactu, factcli_lineas.fecfactu,  factcli_lineas.numlinea, "
            Sql = Sql & " factcli_lineas.codconce, nomconce, ampliaci,factcli_lineas.cantidad , factcli_lineas.precio,factcli_lineas.importe"
            Sql = Sql & " ,factcli_lineas.codigiva, factcli_lineas.porciva, factcli_lineas.porcrec, factcli_lineas.impoiva,"
            Sql = Sql & " factcli_lineas.imporec, if(aplicret=1,'*','') aplicret "
            Sql = Sql & " From factcli_lineas"

            
            If Enlaza Then
                Sql = Sql & Replace(ObtenerWhereCab(True), "factcli", "factcli_lineas")
            Else
                Sql = Sql & " WHERE factcli_lineas.numlinea is null"
            End If
            Sql = Sql & " ORDER BY 1,2,3,4"
            
            
    End Select
    ' ********************************************************************************
    
    MontaSQLCarga = Sql
End Function









Private Sub frmFact_DatoSeleccionado(CadenaSeleccion As String)
Dim CadB As String
    
    If CadenaSeleccion <> "" Then
        CadB = "numserie = " & DBSet(RecuperaValor(CadenaSeleccion, 1), "T") & " and numfactu = " & DBSet(RecuperaValor(CadenaSeleccion, 2), "N") & " and anofactu = year(" & DBSet(RecuperaValor(CadenaSeleccion, 3), "F") & ") "
        
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub







Private Sub frmFPag_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(5).Text = RecuperaValor(CadenaSeleccion, 1)
        Text4(5).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub



Private Sub frmVario_DatoSeleccionado(CadenaSeleccion As String)
    cadFormula = CadenaSeleccion
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
    If Modo <> 5 Then
        Text1(indice).Text = vCampo
    Else
        CadenaConsulta = vCampo
    End If
End Sub





Private Sub frmF_Selec(vFecha As Date)
    Text1(indice).Text = Format(vFecha, formatoFechaVer)
End Sub





Private Sub imgppal_Click(Index As Integer)
    If (Modo = 2 Or Modo = 5 Or Modo = 0) And (Index <> 6) And (Index <> 8) Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    Select Case Index
    Case 0
        'FECHA FACTURA
        indice = 1
        If Modo = 3 Then
            Set frmF = New frmCal
            frmF.Fecha = Now
            If Text1(1).Text <> "" Then frmF.Fecha = CDate(Text1(1).Text)
            frmF.Show vbModal
            Set frmF = Nothing
            PonFoco Text1(1)
        End If
    Case 1 ' contadores
        If Modo = 3 Or Modo = 1 Then
            cadFormula = ""
            Set frmVario = New frmBasico
            AyudaContadoresBasico frmVario
            Set frmVario = Nothing
            If cadFormula <> "" Then
                Text1(2).Text = RecuperaValor(cadFormula, 1)
                PonFoco Text1(2)
                DoEvents
                PonFoco Text1(1)
                cadFormula = ""
            End If
        End If
    Case 2
        'Cliente
        CadenaDesdeOtroForm = ""
        frmcolClientesBusqueda.Show vbModal
        If CadenaDesdeOtroForm <> "" Then
            Text1(4).Text = CadenaDesdeOtroForm
            Text1_LostFocus 4
            PonFoco Text1(5)
            CadenaDesdeOtroForm = ""
        End If
        
    Case 3 ' forma de pago
        Set frmFPag = New frmBasico
        AyudaFormaPago frmFPag
        Set frmFPag = Nothing
        PonFoco Text1(5)
    
    Case 4
        'Cuenta retencion
        'Set frmCtasRet = New frmColCtas
        'frmCtasRet.DatosADevolverBusqueda = "0|1|2|"
        'frmCtasRet.ConfigurarBalances = 3  'NUEVO
        'frmCtasRet.Show vbModal
        'Set frmCtasRet = Nothing
        PonFoco Text1(6)
        
    Case 5
        'pais
       ' Set frmPais = New frmBasico2
       ' AyudaPais frmPais
       ' Set frmPais = Nothing
        
    Case 8
        ' observaciones
        Screen.MousePointer = vbDefault
        
        indice = 3
        
        Set frmZ = New frmZoom
        frmZ.pValor = Text1(indice).Text
        frmZ.pModo = Modo
        frmZ.Caption = "Observaciones Facturas Cliente"
        frmZ.Show vbModal
        Set frmZ = Nothing
        
   
        
    End Select
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    'BotonEliminar
    HacerToolBar 8
End Sub


Private Sub mnModificar_Click()
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
'    If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    ' ***************************************************************************
    
    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub


Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvançar/Retrocedir els camps en les fleches de desplaçament del teclat.
    KEYdown KeyCode
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    HacerToolBar Button.Index
End Sub

Private Sub BotonBuscar()
Dim I As Integer
' ***** Si la clau primaria de la capçalera no es Text1(0), canviar-ho en <=== *****
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        PonFoco Text1(2) ' <===
        ' *** si n'hi han combos a la capçalera ***
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            PonFoco Text1(kCampo)
        End If
    End If
' ******************************************************************************
End Sub

Private Sub HacerBusqueda()

    CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    
    CadB1 = ObtenerBusqueda2(Me, , 2, "FrameAux1")
    
    HacerBusqueda2
    
End Sub


Private Function MontaSQLPpal() As String
    MontaSQLPpal = "SELECT factcli.*,nomclien FROM factcli,clientes  "
    MontaSQLPpal = MontaSQLPpal & " WHERE factcli.codclien=clientes.codclien "
        
End Function

Private Sub HacerBusqueda2()

    CargarSqlFiltro
    
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia
    ElseIf CadB <> "" Or CadB1 <> "" Or cadFiltro <> "" Then
        
        'CadenaConsulta = "SELECT factcli.*,nomclien FROM factcli,clientes  "
        'CadenaConsulta = CadenaConsulta & " WHERE factcli.codclien=clientes.codclien "
        
        CadenaConsulta = MontaSQLPpal
        
        If CadB <> "" Then CadenaConsulta = CadenaConsulta & " and " & CadB & " "
        If CadB1 <> "" Then CadenaConsulta = CadenaConsulta & " and " & CadB1 & " "
        If cadFiltro <> "" Then CadenaConsulta = CadenaConsulta & " and " & cadFiltro & " "
        
        CadenaConsulta = CadenaConsulta & Ordenacion
        PonerCadenaBusqueda
    Else
        ' *** foco al 1r camp visible de la capçalera que siga clau primaria ***
        PonFoco Text1(0)
        ' **********************************************************************
    End If
    
'    CargaDatosLW

End Sub


Private Sub MandaBusquedaPrevia()
Dim cWhere As String
Dim cWhere1 As String
    
    cWhere = "(numserie, numfactu, fecfactu) in (select factcli.numserie, factcli.numfactu, factcli.fecfactu from "
    cWhere = cWhere & "factcli LEFT JOIN factcli_lineas ON factcli.numserie = factcli_lineas.numserie and factcli.fecfactu = factcli_lineas.fecfactu and factcli.numfactu = factcli_lineas.numfactu "
    cWhere = cWhere & " WHERE (1=1) "
    cWhere1 = ""
    If CadB <> "" Then cWhere1 = cWhere1 & " and " & CadB & " "
    If CadB1 <> "" Then cWhere1 = cWhere1 & " and " & CadB1 & " "
    If cadFiltro <> "" Then cWhere1 = cWhere1 & " and " & cadFiltro & " "
    
    If Trim(cWhere1) <> "and (1=1)" Then
        cWhere = cWhere & cWhere1 & ")"
    Else
        cWhere = ""
    End If


  '  Set frmFact = New frmFacturasCliPrev
  '
  '  frmFact.DatosADevolverBusqueda = "0|1|2|"
  '  frmFact.cWhere = cWhere
  '  frmFact.Show vbModal
  '
  '  Set frmFact = Nothing
    

        
End Sub


Private Sub cmdRegresar_Click()
Dim Cad As String
Dim Aux As String
Dim I As Integer
Dim J As Integer

End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        PonerModo 2
        'Data1.Recordset.MoveLast
        Data1.Recordset.MoveFirst
        PonerCampos
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub

EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub BotonVerTodos()
'Vore tots
    LimpiarCampos 'Neteja els Text1
    CadB = ""
    CadB1 = ""
    
    PonerModo 0
    
    HacerBusqueda2
    
End Sub


Private Sub BotonAnyadir()

    LimpiarCampos 'Huida els TextBox
    
    'Contador de facturas
    
    
    PonerModo 3
    
   
    Combo1.ListIndex = 0

    
    Text1(1).Text = Format(Now, formatoFechaVer)
    Text1(8).Text = "0,00"
    Text1(9).Text = "0,00"
    Text1(2).Text = "FDI"
    
    
    PonFoco Text1(2)
    ' ***********************************************************
    
End Sub


Private Sub BotonModificar()


    
    

    If Not PuedeModificarFactura Then Exit Sub


    PonerModo 4

    
    
    
    DespalzamientoVisible False
    PonFoco Text1(5)
    
    
End Sub


Private Sub BotonEliminar(EliminarDesdeActualizar As Boolean)
    Dim I As Long
    Dim Fec As Date
    Dim SqlLog As String

    'Ciertas comprobaciones
    If Data1.Recordset Is Nothing Then Exit Sub
    If Data1.Recordset.EOF Then Exit Sub
    
    
    
    If Not PuedeModificarFactura Then Exit Sub
    
    
    DataGridAux(1).Enabled = False

    'Comprobamos la fecha pertenece al ejercicio
  '  varFecOk = FechaCorrecta2(CDate(Text1(1).Text))
  '  If varFecOk >= 2 Then
  '      If varFecOk = 2 Then
  '          Sql = varTxtFec
  '      Else
  '          Sql = "La factura pertenece a un ejercicio cerrado."
  '      End If
  '      MsgBox Sql, vbExclamation
  '      Exit Sub
  '  End If

    'Comprobamos si esta liquidado
  '  If Not ComprobarPeriodo2(23) Then Exit Sub
    
   
    
    Sql = Sql & vbCrLf & vbCrLf & "Va usted a eliminar la factura :" & vbCrLf
    Sql = Sql & "Numero : " & Data1.Recordset!NumFactu & vbCrLf
    Sql = Sql & "Fecha  : " & Data1.Recordset!FecFactu & vbCrLf
    Sql = Sql & "Cliente : " & Me.Data1.Recordset!codclien & " - " & Text4(4).Text & vbCrLf
    Sql = Sql & vbCrLf & "          ¿Desea continuar ?" & vbCrLf
    
    If Not EliminarDesdeActualizar Then
        If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
    
    NumRegElim = Data1.Recordset.AbsolutePosition
    Screen.MousePointer = vbHourglass
    'Lo hara en actualizar
    'La borrara desde este mismo form
    Conn.BeginTrans
    
    I = Data1.Recordset!NumFactu
    Fec = Data1.Recordset!FecFactu
    If BorrarFactura Then
        'LOG
        SqlLog = "Factura : " & CStr(DBLet(Data1.Recordset!numserie)) & Format(I, "000000") & " de fecha " & Fec
        SqlLog = SqlLog & vbCrLf & "Cliente : " & Text1(4).Text & " " & Text4(4).Text
        SqlLog = SqlLog & vbCrLf & "Importe : " & Text1(12).Text
        
        vLog.Insertar 6, vUsu, SqlLog
    
       
        Conn.CommitTrans
      
    Else

        Conn.RollbackTrans
    End If
   
    Data1.Refresh
    If Data1.Recordset.EOF Then
        'Solo habia un registro
        LimpiarCampos
        CargaGrid 1, False
        PonerModo 0
        Else
            Data1.Recordset.MoveFirst
            NumRegElim = NumRegElim - 1
            If NumRegElim > 1 Then
                For I = 1 To NumRegElim - 1
                    Data1.Recordset.MoveNext
                Next I
            End If
            PonerCampos
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Function BorrarFactura() As Boolean
    
    On Error GoTo EBorrar
    Sql = " WHERE numserie = '" & Data1.Recordset!numserie & "'"
    Sql = Sql & " AND numfactu = " & Data1.Recordset!NumFactu
    Sql = Sql & " AND fecfactu= " & DBSet(Data1.Recordset!FecFactu, "F")
    'Las lineas
    AntiguoText1 = "DELETE from factcli_totales " & Sql
    Conn.Execute AntiguoText1
    AntiguoText1 = "DELETE from factcli_lineas " & Sql
    Conn.Execute AntiguoText1
    'La factura
    AntiguoText1 = "DELETE from factcli " & Sql
    Conn.Execute AntiguoText1
    
EBorrar:
    If Err.Number = 0 Then
        BorrarFactura = True
    Else
        MuestraError Err.Number, "Eliminar factura"
        BorrarFactura = False
    End If
End Function





Private Sub PonerCampos()
Dim I As Integer
Dim CodPobla As String, desPobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1 'opcio=1: posa el format o els camps de la capçalera
    
    ' *** si n'hi han llínies en datagrids ***
    For I = 1 To DataGridAux.Count ' - 1
        CargaGrid I, True
        'If Not AdoAux(I).Recordset.EOF Then PonerCamposForma2 Me, AdoAux(I), 2, "FrameAux" & I
    Next I
    
    imgppal(6).Enabled = (Text1(8).Text <> "")
    imgppal(6).Visible = (Text1(8).Text <> "")
        
    Text4(2).Text = PonerNombreDeCod(Text1(2), "contadores", "nomregis", "serfactur", "T")
    Text4(4).Text = PonerNombreDeCod(Text1(4), "clientes", "nomclien", "codclien")
    Text4(5).Text = PonerNombreDeCod(Text1(5), "ariconta" & vParam.Numconta & ".formapago", "nomforpa", "codforpa", "N")
    Text4(6).Text = PonerNombreDeCod(Text1(6), "ariconta" & vParam.Numconta & ".cuentas", "nommacta", "codmacta", "T")
    'Text4(25).Text = DevuelveDesdeBDNew(cConta, "departamentos", "descripcion", "codmacta", Text1(4).Text, "T", , "dpto", Text1(25).Text, "N")
    'Text4(26).Text = PonerNombreDeCod(Text1(26), "agentes", "nombre", "codigo", "N")
    

  
    
   
    
    
    CargaDatosLW

    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    
End Sub


Private Sub cmdCancelar_Click()
Dim I As Integer
Dim V

    Select Case Modo
        Case 1, 3 'Búsqueda, Insertar
            'Contador de facturas
            If Modo = 3 Then
                'Intentetamos devolver el contador
                If Text1(0).Text <> "" Then
            '        i = FechaCorrecta2(CDate(Text1(0).Text))
            '        Mc.DevolverContador Mc.TipoContador, i = 0, Mc.Contador
                End If
            End If
            LimpiarCampos
            PonerModo 0


        Case 4  'Modificar
            Modo = 2   'Para que el lostfocus NO haga nada

            PonerCampos
            Modo = 4  'Reestablezco el modo para que vuelva a hahacer ponercampos
            '--DesBloqueaRegistroForm Me.Text1(0)
            TerminaBloquear
            
            PonerModo 2
            'Contador de facturas
                
                
        Case 5 'LLÍNIES
            TerminaBloquear
        
            If ModoLineas = 1 Then 'INSERTAR
                ModoLineas = 0
                DataGridAux(1).AllowAddNew = False
                If Not AdoAux(1).Recordset.EOF Then AdoAux(1).Recordset.MoveFirst
                
'                If AdoAux(1).Recordset.EOF Then
'                    If MsgBox("No se permite una factura sin líneas " & vbCrLf & vbCrLf & "¿ Desea eliminar la factura ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
'                        BotonEliminar True
'                        Exit Sub
'                    Else
'                        ModoLineas = 1
'                        cmdAceptar_Click
'                        Exit Sub
'                    End If
'                End If
                
            End If
            ModoLineas = 0
            LLamaLineas 1, 0, 0
            
            Modo = 2   'Para que el lostfocus NO haga nada


            
            PosicionarData
            PonerCampos
            
    End Select
End Sub


Private Function DatosOK() As Boolean
Dim B As Boolean
Dim Cad As String


    On Error GoTo EDatosOK

    DatosOK = False
    
    'fecha de liquidacion
    Text1(13).Text = Text1(1).Text
    
   
    
    B = CompForm2(Me, 1)
    If Not B Then Exit Function
    
    
    Cad = ""
    
    ' NOV 2007
    ' NUEVA ambitode fecha activa
    ' controles añadidos de la factura de david
    'No puede tener % de retencion sin cuenta de retencion
    I = 6
    If Combo1.ListIndex > 0 Then
        If Text1(6).Text = "" Then
            Cad = "Debe indicar cuenta retencion"
        Else
            If ((Text1(6).Text = "") Xor (Text1(7).Text = "")) And Combo1.ListIndex > 0 Then
                Cad = "- No hay porcentaje de rentención sin cuenta de retención"
                I = 7
            End If
        End If
    Else
        If ((Text1(6).Text <> "") Or (Text1(7).Text <> "")) Then Cad = "- No indique datos rentención"
    End If
    'la forma de pago ha de existir
    If Text4(5).Text = "" And (Modo = 3 Or Modo = 4) Then
        Cad = Cad & vbCrLf & "-No existe la forma de pago"
        I = 5
    End If
    '
    If Text4(4).Text = "" And (Modo = 3 Or Modo = 4) Then
        Cad = Cad & vbCrLf & "-No existe cliente"
        I = 4
    End If
    
    If Cad <> "" Then
        MsgBox "Datos incorrectos:" & vbCrLf & Cad, vbExclamation
        B = False
        PonFoco Text1(I)
    End If


    



    If B And Modo = 3 Then
        'Que no exista YA el numero de factura en ese año
        'select fecfactu from factcli where numserie='1' and numfactu=1 and year(fecfactu)=2016
        Cad = "numfactu=" & Text1(0).Text & " AND year(fecfactu)=" & Year(CDate(Text1(1).Text)) & " AND numserie "
        Cad = DevuelveDesdeBD("fecfactu", "factcli", Cad, Text1(2).Text, "T")
        If Cad <> "" Then
            MsgBox "Ya exsiste la factura " & Text1(0).Text & " con fecha " & Cad, vbExclamation
            B = False
        End If
    End If

    DatosOK = B

EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub PosicionarData()
Dim Cad As String, Indicador As String

    Cad = "(numserie=" & DBSet(Text1(2).Text, "T") & " and numfactu = " & DBSet(Text1(0).Text, "N") & " and fecfactu = " & DBSet(Text1(1).Text, "F") & ") "
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    If SituarDataMULTI(Data1, Cad, Indicador) Then
        If ModoLineas <> 1 Then PonerModo 2
        lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
    End If
    ' ***********************************************************************************
End Sub



Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
    If Index = 2 Then AntLetraSer = Text1(2).Text
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean

Dim RC As String
Dim Correcto As Boolean
Dim Valor As Currency
Dim L As Long
Dim I As Integer
Dim J As Integer

Dim RS As ADODB.Recordset


    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    If (Index = 12 Or Index = 0 Or Index = 4) And Modo = 1 Then
        Text1(Index).BackColor = vbMoreLightBlue ' azul clarito
    End If

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    Select Case Index
        Case 4
            Text4(Index) = ""
        Case 5
            Text4(Index) = ""
        Case 6
            Text4(Index) = ""
    End Select
    
    
    
    If Modo = 1 Then Exit Sub
    If Modo = 5 Then Exit Sub
    
    Select Case Index
        Case 0 'Nro de factura
            PonerFormatoEntero Text1(Index)

        Case 1, 13 '1 - fecha de factura
                   '13- fecha de liquidacion
                   
            Sql = ""
            If Not EsFechaOK(Text1(Index)) Then
                MsgBox "Fecha incorrecta", vbExclamation
                'If Index = 1 Then Text1(14).Text = ""
                PonFoco Text1(Index)
                Exit Sub
            End If
            'ModificandoLineas = FechaCorrecta2(CDate(Text1(Index).Text))
            If ModificandoLineas > 1 Then
                If ModificandoLineas = 2 Then
                    RC = varTxtFec
                Else
                    If ModificandoLineas = 3 Then
                        RC = "ya esta cerrado"
                    Else
                        RC = " todavia no ha sido abierto"
                    End If
                    RC = "La fecha pertenece a un ejercicio que " & RC
                End If
                MsgBox RC, vbExclamation
                Text1(Index).Text = ""
                If Index = 1 Then Text1(14).Text = ""
                PonFoco Text1(Index)
                Exit Sub
            End If
            
            Text1(Index).Text = Format(Text1(Index).Text, formatoFechaVer)
             
            If Index = 1 And Modo <> 1 Then Text1(13).Text = Text1(1).Text
            
           

        Case 2 ' Serie
            
            If Modo = 4 Then
                'Modificando
                If Data1.Recordset!numserie = Text1(Index).Text Then Exit Sub
            End If
            Text1(Index).Text = UCase(Text1(Index).Text)
            If Text1(Index).Text = "" Then
                Text4(Index).Text = ""
                Exit Sub
            End If
            'DevuelveValor("select no..... ontadores where tipore and tiporegi REGEXP '^[0-9]+$' = 0")
            RC = "numfactu"
            Text4(2).Text = DevuelveDesdeBD("nomregis", "contadores", "serfactur", Text1(2).Text, "T", RC)
            If Text4(2).Text = "" Then
                MsgBox "Letra de serie no existe o no es de facturas de cliente. Reintroduzca.", vbExclamation
                Text4(2).Text = ""
                Text1(2).Text = ""
                PonFoco Text1(2)
            Else
                Text1(0).Text = Val(RC) + 1
                PonFoco Text1(4)
            End If
            
        Case 4
                ' cliente
                RC = ""
                If PonerFormatoEntero(Text1(Index)) Then
                    cadMen = "codforpa"
                    RC = DevuelveDesdeBD("nomclien", "clientes", "codclien", Text1(Index).Text, "N", cadMen)
                    If RC = "" Then
                        MsgBox "No existe el cliente. Reintroduzca.", vbExclamation
                        PonFoco Text1(Index)
                    Else
                        If Me.Text1(5).Text = "" Then
                            Text1(5).Text = cadMen
                            Text4(5).Text = PonerNombreDeCod(Text1(5), "ariconta" & vParam.Numconta & ".formapago", "nomforpa", "codforpa", "N")
                        End If
                    End If
                End If
                
                Text4(Index).Text = RC
                
                
        
        Case 5 ' forma de pago
            RC = ""
            If PonerFormatoEntero(Text1(Index)) Then
                RC = PonerNombreDeCod(Text1(Index), "ariconta" & vParam.Numconta & ".formapago", "nomforpa", "codforpa", "N")
                If RC = "" Then
                    MsgBox "No existe la Forma de Pago. Reintroduzca.", vbExclamation
                    PonFoco Text1(Index)
                End If
            End If
            Text4(Index).Text = RC
            
        Case 6
            CuentaCorrectaUltimoNivelTextBox Text1(Index), Text4(6)
            
        Case 7 ' % de retencion
            PonerFormatoDecimal Text1(Index), 4
        
        Case 21 ' codigo de pais
            If Text1(Index).Text <> "" Then
                Text4(Index).Text = PonerNombreDeCod(Text1(Index), "paises", "nompais", "codpais", "T")
                If Text4(Index) = "" Then
                    MsgBox "No existe el País. Reintroduzca.", vbExclamation
                    PonFoco Text1(Index)
                End If
            Else
                Text4(Index).Text = ""
            End If
        
        Case 25 ' departamento
            If Text1(Index).Text <> "" Then
                Text4(Index).Text = DevuelveDesdeBDNew(cConta, "departamentos", "descripcion", "codmacta", Text1(4).Text, "T", , "dpto", Text1(25).Text, "N")
                If Text4(Index) = "" Then
                    MsgBox "No existe el Departamento de este Cliente. Reintroduzca.", vbExclamation
                    PonFoco Text1(Index)
                End If
            Else
                Text4(Index).Text = ""
            End If
            
        Case 26 ' agente
            If Text1(Index).Text <> "" Then
                Text4(Index).Text = PonerNombreDeCod(Text1(Index), "agentes", "nombre", "codigo", "N")
                If Text4(Index) = "" Then
                    MsgBox "No existe el Agente. Reintroduzca.", vbExclamation
                    PonFoco Text1(Index)
                End If
            Else
                Text4(Index).Text = ""
            End If
        
    End Select
End Sub

'++
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 3 Then
        If KeyAscii = teclaBuscar Then
            Select Case Index
                Case 1:  KEYBusqueda KeyAscii, 0 ' fecha de factura
                Case 4:  KEYBusqueda KeyAscii, 2 ' cuenta cliente
                Case 6:  KEYBusqueda KeyAscii, 4 ' cuenta de retencion
                Case 5:  KEYBusqueda KeyAscii, 3 ' forma de pago
                Case 2:  KEYBusqueda KeyAscii, 1 ' serie
                Case 21: KEYBusqueda KeyAscii, 5 ' pais
                Case 25: KEYBusqueda KeyAscii, 9 ' departamento
                Case 26: KEYBusqueda KeyAscii, 10 ' agente
            End Select
         Else
            KEYpress KeyAscii
         End If
    Else
        If Index <> 3 Or (Index = 3 And Text1(Index) = "") Then KEYpress KeyAscii
    End If
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub



Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgppal_Click (indice)
End Sub
'++

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
    
        Case 1 'Datos Fiscales
          
           
        Case 2 'Cartera de Cobros
            If Not Data1.Recordset.EOF Then
                'Set frmMens = New frmMensajes
                
                frmMensajes.Opcion = 27
                frmMensajes.Parametros = Trim(Text1(2).Text) & "|" & Trim(Text1(0).Text) & "|" & Text1(1).Text & "|"
                frmMensajes.Show vbModal
                
                'Set frmMens = Nothing
            End If
    
        Case 3
'            Screen.MousePointer = vbHourglass
'
'            Set frmUtil = New frmUtilidades
'
'            frmUtil.Opcion = 5
'            frmUtil.Show vbModal
'
'            Set frmUtil = Nothing
    End Select

End Sub

'************* LLINIES: ****************************
Private Sub ToolbarAux_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim LINASI As Long
Dim Ampliacion As String
    
    
    If Not PuedeModificarFactura Then Exit Sub
            
    
    If Not BLOQUEADesdeFormulario2(Me, Data1, 1) Then Exit Sub
    
 
    

    
    Select Case Button.Index
        Case 1
                   
            'AÑADIR linea factura
            BotonAnyadirLinea 1, True
        Case 2
            'MODIFICAR linea factura
            BotonModificarLinea 1
        Case 3
            'ELIMINAR linea factura
            BotonEliminarLinea 1
            
            
    End Select


End Sub


Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub

Private Sub BotonEliminarLinea(Index As Integer)
Dim Sql As String
Dim vWhere As String
Dim Eliminar As Boolean
Dim SqlLog As String

    On Error GoTo Error2


    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If AdoAux(Index).Recordset.RecordCount < 1 Then Exit Sub




    ModoLineas = 3 'Posem Modo Eliminar Llínia
    
    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index

    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar(Index) Then Exit Sub
    NumTabMto = Index
    Eliminar = False
   
    vWhere = ObtenerWhereCab(True)
    
    ' ***** independentment de si tenen datagrid o no,
    ' canviar els noms, els formats i el DELETE *****
    Select Case Index
        Case 1 'linea de asiento
            Sql = "¿Seguro que desea eliminar la línea de la factura?"
            Sql = Sql & vbCrLf & "Serie: " & AdoAux(Index).Recordset!numserie & " - " & AdoAux(Index).Recordset!NumFactu & " - " & AdoAux(Index).Recordset!FecFactu & " - " & AdoAux(Index).Recordset!numlinea
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM factcli_lineas "
                Sql = Sql & Replace(vWhere, "factcli", "factcli_lineas") & " and numlinea = " & DBLet(AdoAux(Index).Recordset!numlinea, "N")
                
            End If
        
    End Select

    If Eliminar Then
        NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        Conn.Execute Sql
        
        RecalcularTotales
        
        '**** parte de contabilizacion de la factura
        '--DesBloqueaRegistroForm Me.Text1(0)
        TerminaBloquear
        

        
        'LOG
        
        SqlLog = "Factura : " & Text1(2).Text & Text1(0).Text & " " & Text1(1).Text & " Línea : " & DBLet(Me.AdoAux(1).Recordset!numlinea, "N")
        SqlLog = SqlLog & vbCrLf & "Importe : " & DBLet(Me.AdoAux(1).Recordset!Importe, "N")
        
        vLog.Insertar 8, vUsu, SqlLog
        'Creo que no hace falta volver a situar el datagrid
        
            lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            
            Data1.Refresh
            PonerModo 2
        
        '**** hasta aqui
        
        
        ' *** si n'hi han tabs sense datagrid, posar l'If ***
        If Index <> 3 Then _
            CargaGrid Index, True
        ' ***************************************************
        If Not SituarDataTrasEliminar(AdoAux(Index), NumRegElim, True) Then
'            PonerCampos
            
        End If
        ' *** si n'hi han tabs sense datagrid ***
        If Index = 3 Then CargaFrame 3, True
        ' ***************************************
'        If BLOQUEADesdeFormulario2(Me, data1, 1) Then BotonModificar
        ' *** si n'hi han tabs ***
        SituarTab (NumTabMto)
        ' ************************
    End If
    
    ModoLineas = 0
    PosicionarData
    
    Exit Sub
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando linea", Err.Description
End Sub


Private Sub BotonAnyadirLinea(Index As Integer, Limpia As Boolean)
Dim NumF As String
Dim vWhere As String, vTabla As String
Dim anc As Single
Dim I As Integer

    ModoLineas = 1 'Posem Modo Afegir Llínia

    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If

    NumTabMto = Index
    PonerModo 5, Index

    ' *** posar el nom del les distintes taules de llínies ***
    Select Case Index
        Case 1: vTabla = "factcli_lineas"
    End Select
    ' ********************************************************

    vWhere = ObtenerWhereCab(False)

    Select Case Index
         Case 1   'hlinapu
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            NumF = ""
            NumF = SugerirCodigoSiguienteStr(vTabla, "numlinea", Replace(vWhere, "factcli", "factcli_lineas"))
            ' ***************************************************************

            AnyadirLinea DataGridAux(Index), AdoAux(Index)

            anc = DataGridAux(Index).top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 230 '248
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If

            LLamaLineas Index, ModoLineas, anc

            Select Case Index
                ' *** valor per defecte a l'insertar i formateig de tots els camps ***
                Case 1 'lineas de factura
                    If Limpia Then
                        For I = 0 To txtAux.Count - 1
                            txtAux(I).Text = ""
                        Next I
                    End If
                    txtAux(0).Text = Text1(2).Text 'serie
                    txtAux(1).Text = Text1(0).Text 'numfactu
                    txtAux(2).Text = Text1(1).Text 'fecha
                   
                    
                    txtAux(3).Text = Format(NumF, "0000") 'linea contador
                    
                    
                    If Limpia Then txtAux2(4).Text = ""
                    PonFoco txtAux(4)
                    chkAux2(0).Value = 0
            End Select

    End Select
End Sub

Private Function ExisteEnFactura(Serie As String, NumFactu As String, FecFactu As String, Cuenta As String) As Boolean
Dim Sql As String

    ExisteEnFactura = False
    
    If Serie = "" Or NumFactu = "" Or FecFactu = "" Or Cuenta = "" Then Exit Function

    Sql = "select count(*) from factcli_lineas where numserie = " & DBSet(Serie, "T") & " and numfactu = " & DBSet(NumFactu, "N")
    Sql = Sql & " and fecfactu = " & DBSet(FecFactu, "F") & " and codmacta = " & DBSet(Cuenta, "T")

    ExisteEnFactura = (TotalRegistros(Sql) <> 0)
    
End Function





Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim I As Integer
    Dim J As Integer

    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If AdoAux(Index).Recordset.RecordCount < 1 Then Exit Sub


    ModoLineas = 2 'Modificar llínia

    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If

    NumTabMto = Index
    PonerModo 5, Index
    ' *** bloqueje la clau primaria de la capçalera ***
'    BloquearTxt Text1(0), True
    ' *********************************

    Select Case Index
        Case 0, 1 ' *** pose els index de llínies que tenen datagrid (en o sense tab) ***
            If DataGridAux(Index).Bookmark < DataGridAux(Index).FirstRow Or DataGridAux(Index).Bookmark > (DataGridAux(Index).FirstRow + DataGridAux(Index).VisibleRows - 1) Then
                I = DataGridAux(Index).Bookmark - DataGridAux(Index).FirstRow
                DataGridAux(Index).Scroll 0, I
                DataGridAux(Index).Refresh
            End If

            anc = DataGridAux(Index).top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If

    End Select

    Select Case Index
        ' *** valor per defecte al modificar dels camps del grid ***
        Case 1 'lineas de facturas
            txtAux(0).Text = DataGridAux(Index).Columns(0).Text
            txtAux(1).Text = DataGridAux(Index).Columns(1).Text
            txtAux(2).Text = DataGridAux(Index).Columns(2).Text
            txtAux(3).Text = DataGridAux(Index).Columns(3).Text
            txtAux(4).Text = DataGridAux(Index).Columns(4).Text
            txtAux2(4).Text = DataGridAux(Index).Columns(5).Text 'denominacion
            txtAux(13).Text = DataGridAux(Index).Columns(6).Text 'ampliaci
            txtAux(12).Text = DataGridAux(Index).Columns(14).Text 'centro de coste
            txtAux(5).Text = DataGridAux(Index).Columns(7).Text '
            txtAux(6).Text = DataGridAux(Index).Columns(8).Text '
            txtAux(12).Text = DataGridAux(Index).Columns(9).Text 'centro de coste
            txtAux(7).Text = DataGridAux(Index).Columns(10).Text '
            txtAux(8).Text = DataGridAux(Index).Columns(11).Text '%iva
            txtAux(9).Text = DataGridAux(Index).Columns(12).Text '%retencion
            txtAux(10).Text = DataGridAux(Index).Columns(13).Text 'importe iva
            txtAux(11).Text = DataGridAux(Index).Columns(14).Text 'importe retencion
            
            
            
            If DataGridAux(Index).Columns(15).Text = "*" Then
                chkAux2(0).Value = 1 ' DataGridAux(Index).Columns(14).Text 'aplica retencion
            Else
                chkAux2(0).Value = 0
            End If
            
           
           
            
    End Select

    LLamaLineas Index, ModoLineas, anc
    
    
    PonFoco txtAux(4)
    
    ' ***************************************************************************************
End Sub


Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim B As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    If Index <> 3 Then DeseleccionaGrid DataGridAux(Index)
    ' ***************************************************

    B = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Llínies
    Select Case Index
        Case 1 'lineas de factura
            For jj = 4 To txtAux.Count - 1
                txtAux(jj).Visible = B
                txtAux(jj).top = alto
                If jj >= 7 And jj <= 12 Then BloquearTxt txtAux(jj), True
            Next jj
            
            txtAux2(4).Visible = B
            txtAux2(4).top = alto
            
            chkAux2(0).Visible = B
            chkAux2(0).top = alto
            
            For jj = 0 To 1
                cmdAux(jj).Visible = B
                cmdAux(jj).top = txtAux(4).top
                cmdAux(jj).Height = txtAux(4).Height
            Next jj
            
 
            
    End Select
End Sub



Private Function DatosOkLlin(nomframe As String) As Boolean
Dim Sql As String
Dim B As Boolean


    DatosOkLlin = True
    
    On Error GoTo EDatosOKLlin

    Sql = ""
    DatosOkLlin = False

    B = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not B Then Exit Function
    
    If B And (Modo = 5 And ModoLineas = 1) Then  'insertar
    
    End If
    
    If B And Modo = 5 Then ' tanto si insertamos como si modificamos en lineas
        Sql = ""
        If txtAux2(4).Text = "" Then
            Sql = "Error concepto facturación"
        Else
            Sql = DevuelveDesdeBD("nomconce", "conceptos", "codconce", txtAux(4).Text)
            If Sql <> txtAux2(4).Text Then
                Sql = "Concepto no es el selccionado"
            Else
                Sql = ""
            End If
        End If
        If Sql <> "" Then
            MsgBox Sql, vbExclamation
            B = False
        End If
    End If
    
    DatosOkLlin = B

EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function SepuedeBorrar(ByRef Index As Integer) As Boolean
    SepuedeBorrar = False
    
    SepuedeBorrar = True
End Function


' *********************************************************************************
Private Sub DataGridAux_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
Dim I As Byte

    If ModoLineas <> 1 Then
        Select Case Index
            Case 1 'lineas de facturas
                If DataGridAux(Index).Columns.Count > 2 Then
                End If
        End Select
    End If
End Sub

' ***** si n'hi han varios nivells de tabs *****
Private Sub SituarTab(numTab As Integer)
    On Error Resume Next
    
'    If numTab = 0 Then
'        SSTab1.Tab = 2
'    ElseIf numTab = 1 Then
'        SSTab1.Tab = 1
'    End If
    
    If Err.Number <> 0 Then Err.Clear
End Sub
' **********************************************


Private Sub CargaFrame(Index As Integer, Enlaza As Boolean)
End Sub

' *** si n'hi han tabs sense datagrids ***
Private Sub NetejaFrameAux(nom_frame As String)
Dim Control As Object
    
    For Each Control In Me.Controls
        If (Control.Tag <> "") Then
            If (Control.Container.Name = nom_frame) Then
                If TypeOf Control Is TextBox Then
                    Control.Text = ""
                ElseIf TypeOf Control Is ComboBox Then
                    Control.ListIndex = -1
                End If
            End If
        End If
    Next Control

End Sub
' ****************************************


Private Sub CargaGrid(Index As Integer, Enlaza As Boolean)
Dim B As Boolean
Dim I As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, Enlaza)

    B = DataGridAux(Index).Enabled
    DataGridAux(Index).Enabled = False
    
    AdoAux(Index).ConnectionString = Conn
    AdoAux(Index).RecordSource = MontaSQLCarga(Index, Enlaza)
    AdoAux(Index).CursorType = adOpenDynamic
    AdoAux(Index).LockType = adLockPessimistic
    DataGridAux(Index).ScrollBars = dbgNone
    AdoAux(Index).Refresh
    Set DataGridAux(Index).DataSource = AdoAux(Index)
    
    DataGridAux(Index).AllowRowSizing = False
    DataGridAux(Index).RowHeight = 350
    
    If PrimeraVez Then
        DataGridAux(Index).ClearFields
        DataGridAux(Index).ReBind
        DataGridAux(Index).Refresh
    End If

    For I = 0 To DataGridAux(Index).Columns.Count - 1
        DataGridAux(Index).Columns(I).AllowSizing = False
    Next I
    
    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), tots, PrimeraVez
    
    
    Select Case Index
        
        Case 1 'lineas de factura
            
                ' factcli_lineas.numserie, factcli_lineas.numfactu, factcli_lineas.fecfactu,  factcli_lineas.numlinea,
                'factcli_lineas.codconce, nomconce,ampliaci, factcli_lineas.cantidad , factcli_lineas.precio ,
                'factcli_lineas.codigiva, factcli_lineas.porciva, factcli_lineas.porcrec, factcli_lineas.impoiva,
                'factcli_lineas.imporec,factcli_lineas.importe,


                tots = "N||||0|;N||||0|;N||||0|;N||||0|;S|txtaux(4)|T|Concepto|1005|;S|cmdAux(0)|B|||;S|txtAux2(4)|T|Descripción|3495|"
                tots = tots & ";S|txtaux(13)|T|Ampliacion|2505|;S|cmdAux(1)|B|||;"
                tots = tots & "S|txtaux(5)|T|Uds|805|;S|txtaux(6)|T|Precio|1505|;S|txtaux(12)|T|Total|1155|;"
                tots = tots & "S|txtaux(7)|T|Iva|625|;S|txtaux(8)|T|%Iva|785|;"
                tots = tots & "S|txtaux(9)|T|%Rec|785|;S|txtaux(10)|T|Iva(€)|1154|;"
                tots = tots & "S|txtaux(11)|T|Rec(€)|1155|;S|chkAux2(0)|CB|Ret|400|;"
       
            
            arregla tots, DataGridAux(Index), Me
        
            DataGridAux(Index).Columns(4).Alignment = dbgLeft
            DataGridAux(Index).Columns(5).Alignment = dbgLeft
            DataGridAux(Index).Columns(6).Alignment = dbgLeft
            DataGridAux(Index).Columns(15).Alignment = dbgCenter
            
            B = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))

            If (Enlaza = True) And (Not AdoAux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
            
            Else
                For I = 0 To 3
                    txtAux(I).Text = ""
                Next I
                
            End If
    End Select
    DataGridAux(Index).ScrollBars = dbgAutomatic
      
    ' **** si n'hi han llínies en grids i camps fora d'estos ****
    If Not AdoAux(Index).Recordset.EOF Then
        DataGridAux_RowColChange Index, 1, 1
    Else
        LimpiarCamposFrame Index
    End If
    ' **********************************************************
      
    'Obtenemos las sumas
'    ObtenerSumas
    If Enlaza Then CargaDatosLW
    
    PonerModoUsuarioGnral Modo, "arigestion"

      
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
End Sub


Private Sub InsertarLinea()
'Inserta registre en les taules de Llínies
Dim nomframe As String
Dim B As Boolean
Dim Limp As Boolean
Dim Cad As String



    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 1: nomframe = "FrameAux1"
    End Select
    ' ***************************************************************
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        Conn.BeginTrans
        
        B = True
        If CambiarIva Then B = ActualizarIva
    
        If B And InsertarDesdeForm2(Me, 2, nomframe) Then
        
            B = RecalcularTotales
            
            If B Then RecalcularTotalesFactura
        
            If B Then
                Conn.CommitTrans
            Else
                Conn.RollbackTrans
            End If
            
            B = BLOQUEADesdeFormulario2(Me, Data1, 1)
            
            Select Case NumTabMto
                Case 0, 1 ' *** els index de les llinies en grid (en o sense tab) ***
                    CargaGrid NumTabMto, True
                    
                    DataGridAux(1).AllowAddNew = False
                    
                    If Not AdoAux(1).Recordset.EOF Then PosicionGrid = DataGridAux(1).FirstRow
                    CargaGrid 1, True
                    Limp = True

                    txtAux(11).Text = ""
                    If Limp Then
                        txtAux2(5).Text = ""
                        txtAux2(12).Text = ""
                        For I = 0 To 11
                            txtAux(I).Text = ""
                        Next I
                    End If
                    ModoLineas = 0
                    If B Then
                            BotonAnyadirLinea NumTabMto, True
                    End If
            End Select
           
        Else
           Conn.RollbackTrans
        End If
    End If
End Sub

Private Function ActualizarIva() As Boolean
Dim Sql As String

    On Error GoTo eActualizarIva
    
    ActualizarIva = False
    
    Sql = "update cuentas set codigiva = " & DBSet(txtAux(7).Text, "N") & " where codmacta = " & DBSet(txtAux(5).Text, "T")
    Conn.Execute Sql
    
    ActualizarIva = True
    Exit Function
    
eActualizarIva:
    MuestraError Err.Number, "Actualizar Iva", Err.Description
End Function


Private Sub ModificarLinea()
'Modifica registre en les taules de Llínies
Dim nomframe As String
Dim V As Integer
Dim Cad As String
    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 1: nomframe = "FrameAux1" 'apuntes
    End Select
    ' **************************************************************

    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        Conn.BeginTrans
        
        B = True
        If CambiarIva Then B = ActualizarIva
        
        If B And ModificaDesdeFormulario2(Me, 2, nomframe) Then
        
            B = RecalcularTotales
            
            'LOG
            vLog.Insertar 7, vUsu, "Factura : " & Text1(2).Text & Text1(0).Text & " " & Text1(1).Text & " Linea : " & txtAux(4).Text
        
            If B Then
                Conn.CommitTrans
            Else
                Conn.RollbackTrans
            End If
            
            ' *** si cal que fer alguna cosa abas d'insertar ***
            If NumTabMto = 0 Then
            End If
            ' ******************************************************
            ModoLineas = 0

            If NumTabMto <> 3 Then
                V = AdoAux(NumTabMto).Recordset.Fields(3) 'el 2 es el nº de llinia
                CargaGrid NumTabMto, True
            End If

            ' *** si n'hi han tabs ***
            SituarTab (NumTabMto)

            ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
            If NumTabMto <> 3 Then
                DataGridAux(NumTabMto).SetFocus
                AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(3).Name & " =" & V)
            End If
            ' ***********************************************************

            LLamaLineas NumTabMto, 0
            
        Else
            Conn.RollbackTrans
        End If
    End If
        
End Sub




Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & "factcli.numserie=" & DBSet(Text1(2).Text, "T") & " and factcli.numfactu=" & DBSet(Text1(0).Text, "N") & " and fecfactu = " & DBSet(Text1(1).Text, "F")
    ' *******************************************************
    
    ObtenerWhereCab = vWhere
End Function

' *** neteja els camps dels tabs de grid que
'estan fora d'este, i els camps de descripció ***
Private Sub LimpiarCamposFrame(Index As Integer)
End Sub
' ***********************************************


Private Sub PonerModoUsuarioGnral(Modo As Byte, aplicacion As String)
Dim RS As ADODB.Recordset
Dim Cad As String
    
    On Error Resume Next

    Cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(aplicacion, "T")
    Cad = Cad & " and codigo = " & DBSet(IdPrograma, "N") & " and codusu = " & DBSet(vUsu.id, "N")
    
    Set RS = New ADODB.Recordset
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        Toolbar1.Buttons(1).Enabled = DBLet(RS!creareliminar, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(2).Enabled = DBLet(RS!Modificar, "N") And (Modo = 2)
        Toolbar1.Buttons(3).Enabled = DBLet(RS!creareliminar, "N") And (Modo = 2)
        
        Toolbar1.Buttons(5).Enabled = DBLet(RS!Ver, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(6).Enabled = DBLet(RS!Ver, "N") And (Modo = 0 Or Modo = 2)
        
        Toolbar1.Buttons(8).Enabled = DBLet(RS!Imprimir, "N") And Modo <= 2
        
        Me.Toolbar2.Buttons(1).Enabled = False
        Me.Toolbar2.Buttons(3).Enabled = False
        
        ToolbarAux.Buttons(1).Enabled = DBLet(RS!creareliminar, "N") And (Modo = 2)
        ToolbarAux.Buttons(2).Enabled = DBLet(RS!Modificar, "N") And (Modo = 2 And Me.Data1.Recordset.RecordCount > 0)
        ToolbarAux.Buttons(3).Enabled = DBLet(RS!creareliminar, "N") And (Modo = 2 And Me.Data1.Recordset.RecordCount > 0)
        
        vUsu.LeerFiltros "arigestion", IdPrograma
        
    End If
    
    RS.Close
    Set RS = Nothing
    
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    AntiguoText1 = txtAux(Index).Text
    ConseguirFoco txtAux(Index), Modo
End Sub


Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

'++
Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 4:  KEYImage KeyAscii, 0 ' cta base
            'Case 7:  KEYImage KeyAscii, 1 ' iva
            'Case 12:  KEYImage KeyAscii, 2 ' Centro Coste
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYImage(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
  '  cmdAux_Click (Indice)
End Sub
'++


Private Sub txtAux_LostFocus(Index As Integer)
    Dim RC As String
    Dim Importe As Currency
    Dim IVA As String
    
        If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
        
        Select Case Index
        Case 4
            RC = ""
            IVA = ""
            If PonerFormatoEntero(txtAux(Index)) Then
                IVA = "codigiva"
                RC = DevuelveDesdeBD("nomconce", "conceptos", "codconce", txtAux(Index).Text, "N", IVA)
                If RC = "" Then
                    IVA = ""
                    MsgBox "No se encuentra el concepto: " & txtAux(Index).Text, vbExclamation
                    txtAux(Index).Text = ""
                    PonFoco txtAux(Index)
                    
                End If
                
            End If
            txtAux2(Index).Text = RC
            If IVA <> "" Then
                'Pondremos los datos del iva
                txtAux(7).Text = IVA
                IVA = "porcerec"
                RC = DevuelveDesdeBD("porceiva", "ariconta" & vParam.Numconta & ".tiposiva", "codigiva", txtAux(7).Text, , IVA)
                If RC = "" Then
                    MsgBox "IVA no encontrado en contabilidad: " & txtAux(7).Text, vbExclamation
                    IVA = ""
                Else
                    txtAux(8).Text = Format(RC, FormatoDec10d2)
                    txtAux(9).Text = Format(IVA, FormatoDec10d2)
                End If
            End If
            If IVA = "" Then
            
            
            End If
            
        Case 5, 6
            PonerFormatoDecimal txtAux(Index), 1
    
        Case 6

        End Select

        If Index = 5 Or Index = 6 Or Index = 7 Then CalcularIVA


End Sub




Private Sub HacerToolBar(Boton As Integer)

    
    Select Case Boton
        Case 1
            BotonAnyadir
        Case 2
            BotonModificar
        Case 3
            BotonEliminar False
        Case 5
            BotonBuscar
        Case 6
            BotonVerTodos
        Case 8
            'Imprimir factura
            
            ImprimirFactura
           



    End Select
End Sub


Private Function ModificarFactura() As Boolean
Dim B1 As Boolean


    On Error GoTo EModificar
         
        ModificarFactura = False
     
                    
        Conn.BeginTrans
        'Comun
        
        B = RecalcularTotalesFactura
        
        If B Then B = ModificaDesdeFormulario2(Me, 1)
        
        
        If B Then InsertarEnCobros
        
  
EModificar:
        If Err.Number <> 0 Or Not B Then
            MuestraError Err.Number
            Conn.RollbackTrans
            ModificarFactura = False
            B1 = False
        Else
            Conn.CommitTrans
            ModificarFactura = True
        End If
        
End Function


'##### Nuevo para el ambito de fechas
Private Function AmbitoDeFecha(DesbloqueAsiento As Boolean) As Boolean
        AmbitoDeFecha = False
       ' varFecOk = FechaCorrecta2(CDate(Text1(1).Text))
        If varFecOk > 1 Then
            If varFecOk = 2 Then
                MsgBox varTxtFec, vbExclamation
            Else
                MsgBox "El asiento pertenece a un ejercicio cerrado.", vbExclamation
            End If
        Else
            AmbitoDeFecha = True
        End If
    
'        If DesbloqueAsiento Then DesBloqAsien
End Function


Private Sub LanzaPantalla(Index As Integer)
Dim miI As Integer
        '----------------------------------------------------
        '----------------------------------------------------
        '
        ' Dependiendo de index lanzaremos una opcion uotra
        '
        '----------------------------------------------------
        
        'De momento solo para el 5. Cliente
        miI = -1
        Select Case Index
        Case 0
            txtAux(0).Text = ""
            miI = 3
        Case 3
            txtAux(3).Text = ""
            miI = 0
        Case 4
            txtAux(4).Text = ""
            miI = 1
            
        Case 8
            txtAux(8).Text = ""
            miI = 2
        End Select
       ' If miI >= 0 Then cmdAux_Click miI
End Sub

Private Function SituarData1() As Boolean
    Dim Sql As String
    
    On Error GoTo ESituarData1
    
    
    'Si es insertar, lo que hace es simplemente volver a poner el el recordset
    'este unico registro
    'If Insertar Then
        Sql = MontaSQLPpal
        Sql = Sql & " AND numserie =" & DBSet(Text1(2).Text, "T")
        Sql = Sql & " AND fecfactu=" & DBSet(Text1(1).Text, "F") & " AND numfactu = " & Text1(0).Text
        Data1.RecordSource = Sql
    
    'End If
    
    Data1.Refresh
    With Data1.Recordset
        If .EOF Then Exit Function
        .MoveLast
        .MoveFirst
        While Not Data1.Recordset.EOF
            If CStr(.Fields!numserie) = Text1(2).Text Then
                If CStr(.Fields!NumFactu) = Val(Text1(0).Text) Then
                    If Format(CStr(.Fields!FecFactu), formatoFechaVer) = Text1(1).Text Then
                        SituarData1 = True
                        Exit Function
                    End If
                End If
            End If
            .MoveNext
        Wend
    End With
ESituarData1:
        If Err.Number <> 0 Then Err.Clear
        Limpiar Me
        PonerModo 0
        lblIndicador.Caption = ""
        SituarData1 = False
End Function


'********************************************************
'
' FUNCIONES CORRESPONDIENTES A LA INSERCION DE DOCUMENTOS
'
'********************************************************


Private Sub CargaDatosLW()
Dim C As String
Dim bs As Byte
    bs = Screen.MousePointer
    C = Me.lblIndicador.Caption
    lblIndicador.Caption = "Leyendo "
    lblIndicador.Refresh
    CargaDatosLW2
    Me.lblIndicador.Caption = C
    Screen.MousePointer = bs
End Sub

Private Sub CargaDatosLW2()
Dim Cad As String
Dim RS As ADODB.Recordset
Dim IT As ListItem
Dim ElIcono As Integer
Dim GroupBy As String
Dim Orden As String
Dim C As String


    On Error GoTo ECargaDatosLW
    
    
    Select Case CByte(RecuperaValor(lw1.Tag, 1))
    Case 5 ' imagenes
        Cad = "select h.numlinea,  h.codigiva, tt.nombriva,  h.baseimpo, h.impoiva, h.imporec from factcli_totales h "
        Cad = Cad & " inner join ariconta" & vParam.Numconta & ".tiposiva tt on h.codigiva = tt.codigiva  WHERE "
        Cad = Cad & " numserie=" & DBSet(Data1.Recordset!numserie, "T")
        Cad = Cad & " and numfactu=" & Data1.Recordset!NumFactu
        Cad = Cad & " and fecfactu=" & DBSet(Data1.Recordset!FecFactu, "F")
        GroupBy = ""
        BuscaChekc = "numlinea"
        
    End Select
    
    
    'BuscaChekc="" si es la opcion de precios especiales
    Cad = Cad & " ORDER BY 1"
    
    lw1.ListItems.Clear
    
    Set RS = New ADODB.Recordset
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    While Not RS.EOF
        Set IT = lw1.ListItems.Add

        IT.Text = RS!numlinea
        IT.SubItems(1) = Format(RS!codigiva, "000")
        IT.SubItems(2) = RS!nombriva
        IT.SubItems(3) = Format(RS!Baseimpo, "###,###,##0.00")
        IT.SubItems(4) = Format(RS!Impoiva, "###,###,##0.00")
        If DBLet(RS!ImpoRec) <> 0 Then
            IT.SubItems(5) = Format(RS!ImpoRec, "###,###,##0.00")
        Else
            IT.SubItems(5) = " "
        End If
        
        Set IT = Nothing

        RS.MoveNext
    Wend
    Set RS = Nothing
    
    
    
    Exit Sub
ECargaDatosLW:
    MuestraError Err.Number
    Set RS = Nothing
    
End Sub


Private Sub AnyadirAlListview(vpaz As String, DesdeBD As Boolean)
Dim J As Integer
Dim Aux As String
Dim IT As ListItem
Dim Contador As Integer
    If Dir(vpaz, vbArchive) = "" Then
'        MsgBox "No existe el archivo: " & vpaz, vbExclamation
    Else
        Set IT = lw1.ListItems.Add()

        IT.Text = Me.Adodc1.Recordset!Orden '"Nuevo " & Contador
        
        IT.SubItems(1) = Me.Adodc1.Recordset.Fields(5)  'Abs(DesdeBD)   'DesdeBD 0:NO  numero: el codigo en la BD
        IT.SubItems(2) = vpaz
        IT.SubItems(3) = Me.Adodc1.Recordset.Fields(0)
        
        Set IT = Nothing
    End If
End Sub



Private Sub EliminarImagen()
Dim Sql As String
Dim Mens As String
    
    On Error GoTo eEliminarImagen

    Mens = "Va a proceder a eliminar de la lista correspondiente al asiento. " & vbCrLf & vbCrLf & "¿ Desea continuar ?" & vbCrLf & vbCrLf
    
    If MsgBox(Mens, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        Sql = "delete from hcabapu_fichdocs where numasien = " & DBSet(Text1(0).Text, "N") & " and fechaent = " & DBSet(Text1(1).Text, "F") & " and numdiari = " & DBSet(Text1(2).Text, "N") & " and codigo = " & Me.lw1.SelectedItem.SubItems(3)
        Conn.Execute Sql
        FicheroAEliminar = lw1.SelectedItem.SubItems(2)
        CargaDatosLW
        
    End If
    Exit Sub

eEliminarImagen:
    MuestraError Err.Number, "Eliminar imágen", Err.Description
End Sub


Private Sub CargaFiltros()
Dim Aux As String
    

    cboFiltro.Clear
    
    cboFiltro.AddItem "Sin Filtro "
    cboFiltro.ItemData(cboFiltro.NewIndex) = 0
    cboFiltro.AddItem "Ejercicios Abiertos "
    cboFiltro.ItemData(cboFiltro.NewIndex) = 1
    cboFiltro.AddItem "Ejercicio Actual "
    cboFiltro.ItemData(cboFiltro.NewIndex) = 2
    cboFiltro.AddItem "Ejercicio Siguiente "
    cboFiltro.ItemData(cboFiltro.NewIndex) = 3

End Sub
    


Private Sub CargarCombo()
Dim RS As ADODB.Recordset
Dim Sql As String
Dim J As Long
    
    Combo1.Clear
    'Tipo de retencion
    Set RS = New ADODB.Recordset
    Sql = "SELECT * FROM usuarios.wtiporeten ORDER BY codigo"
    RS.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        Combo1.AddItem RS!Descripcion
        Combo1.ItemData(Combo1.NewIndex) = RS!Codigo
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing




End Sub

Private Function ComprobarPeriodo2(indice As Integer) As Boolean
Dim Cerrado As Boolean

'
'
'    'Primero pondremos la fecha a año periodo
'    i = Year(CDate(Text1(Indice).Text))
'    If vParam.periodos = 0 Then
'        'Trimestral
'        Ancho = ((Month(CDate(Text1(Indice).Text)) - 1) \ 3) + 1
'        Else
'        Ancho = Month(CDate((Text1(Indice).Text)))
'    End If
'    Cerrado = False
'    If i < vParam.anofactu Then
'        Cerrado = True
'    Else
'        If i = vParam.anofactu Then
'            'El mismo año. Comprobamos los periodos
'            If vParam.perfactu >= Ancho Then _
'                Cerrado = True
'        End If
'    End If
'    ComprobarPeriodo2 = True
'    ModificaFacturaPeriodoLiquidado = False
'    If Cerrado Then
'        ModificaFacturaPeriodoLiquidado = True
'        Sql = "La fecha "
'        If Indice = 0 Then
'            Sql = Sql & "factura"
'        Else
'            Sql = Sql & "liquidacion"
'        End If
'        Sql = Sql & " corresponde a un periodo ya liquidado. " & vbCrLf
'
'        If vUsu.Nivel = 0 Then
'
'            Sql = Sql & vbCrLf & " ¿Desea continuar igualmente ?"
'
'            If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then ComprobarPeriodo2 = False
'
'        Else
'            MsgBox Sql, vbExclamation
'
'            ComprobarPeriodo2 = False
'
'        End If
'
'        '[Monica]12/09/2016: no tocar cartera
'        ModificarCobros = False
'
'    End If
End Function


Private Sub CargarDatosCuenta(Cuenta As String)
Dim RS As ADODB.Recordset
Dim Sql As String

    On Error GoTo eTraerDatosCuenta
    
    Sql = "select * from cuentas where codmacta = " & DBSet(Cuenta, "T")
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Text1(5).Text = ""
    Text4(5).Text = ""
    
    For I = 15 To 21
        Text1(I).Text = ""
    Next I
    
    If Not RS.EOF Then
        Text1(5).Text = DBLet(RS!Forpa, "N")
        Text4(5).Text = PonerNombreDeCod(Text1(5), "formapago", "nomforpa", "codforpa", "N")
        
        Text1(15).Text = DBLet(RS!Nommacta, "T")
        Text1(16).Text = DBLet(RS!dirdatos, "T")
        Text1(17).Text = DBLet(RS!codposta, "T")
        Text1(18).Text = DBLet(RS!desPobla, "T")
        Text1(19).Text = DBLet(RS!desProvi, "T")
        Text1(20).Text = DBLet(RS!nifdatos, "T")
        Text1(21).Text = DBLet(RS!codpais, "T")
        Text4(21).Text = PonerNombreDeCod(Text1(21), "paises", "nompais", "codpais", "T")
    End If
    Exit Sub
    
eTraerDatosCuenta:
    MuestraError Err.Number, "Cargar Datos de Cuenta", Err.Description

End Sub


Private Function AnyadeCadenaFiltro() As String
Dim Aux As String

    Aux = ""
    If vUsu.FiltroFactCli <> 0 Then
        '-------------------------------- INICIO
        I = Year(Now)
        If vUsu.FiltroFactCli < 3 Then
            'INicio = actual
            Aux = " anofactu >= " & I
            Else
            Aux = " anofactu >=" & I + 1
        End If
        I = Year(Now) + 1
        If vUsu.FiltroFactCli = 2 Then
            Aux = Aux & " AND anofactu <= " & I
        Else
            Aux = Aux & " AND anofactu <= " & I + 1
        End If
        
    End If  'filtro=0
    AnyadeCadenaFiltro = Aux
End Function



Private Sub CalcularIVA()
Dim J As Integer
Dim Base As Currency
Dim Aux As Currency
Dim impor As Currency
    
    impor = ImporteFormateado(txtAux(5).Text)
    Base = ImporteFormateado(txtAux(6).Text)
    
    impor = Round(Base * impor, 2)
    txtAux(12).Text = Format(impor, FormatoImporte)
    'EL iva
    Aux = ImporteFormateado(txtAux(8).Text) / 100
    
    If Aux = 0 Then
        Base = 0
        If txtAux(10).Text = "" Then
            txtAux(10).Text = ""
        Else
            txtAux(10).Text = "0,00"
        End If
    Else
        Base = Round((Aux * impor), 2)
        txtAux(10).Text = Format(Base, FormatoImporte)
    End If
    
    'Recargo
    Aux = ImporteFormateado(txtAux(9).Text) / 100
    If Aux = 0 Then
        txtAux(11).Text = ""
    Else
        Base = Base + Round(Aux * impor, 2)
        txtAux(11).Text = Format(Round((Aux * impor), 2), FormatoImporte)
    End If
    
    
End Sub

Private Function RecalcularTotales() As Boolean
Dim Sql As String
Dim SqlInsert As String
Dim SqlValues As String
Dim I As Long
Dim RS As ADODB.Recordset

Dim Baseimpo As Currency
Dim Basereten As Currency
Dim Impoiva As Currency
Dim ImpoRec As Currency
Dim Imporeten As Currency
Dim TotalFactura As Currency

    On Error GoTo eRecalcularTotales

    RecalcularTotales = False

    Sql = Replace(ObtenerWhereCab(True), "factcli", "factcli_totales")
    Sql = "delete from factcli_totales  " & Sql
    Conn.Execute Sql
    
    SqlInsert = "insert into factcli_totales (numserie,numfactu,fecfactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec) values "
    
    Sql = "select codigiva, porciva, porcrec, sum(importe) baseimpo, sum(coalesce(impoiva,0)) imporiva, sum(coalesce(imporec,0)) imporrec from factcli_lineas "
    Sql = Sql & Replace(ObtenerWhereCab(True), "factcli", "factcli_lineas")
    Sql = Sql & " group by 1,2,3"
    Sql = Sql & " order by 1,2,3"
    
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 1
    
    SqlValues = ""
    
    Baseimpo = 0
    Basereten = 0
    Impoiva = 0
    ImpoRec = 0
    Imporeten = 0
    
    TotalFactura = 0
    
    While Not RS.EOF
        Sql = "(" & DBSet(Text1(2).Text, "T") & "," & DBSet(Text1(0).Text, "N") & "," & DBSet(Text1(1).Text, "F") & ","
        Sql = Sql & DBSet(I, "N") & "," & DBSet(RS!Baseimpo, "N") & "," & DBSet(RS!codigiva, "N") & "," & DBSet(RS!porciva, "N") & "," & DBSet(RS!porcrec, "N") & ","
        Sql = Sql & DBSet(RS!ImporIVA, "N") & "," & DBSet(RS!imporrec, "N") & "),"
        
        SqlValues = SqlValues & Sql
        
        Baseimpo = Baseimpo + DBLet(RS!Baseimpo, "N")
        Impoiva = Impoiva + DBLet(RS!ImporIVA, "N")
        ImpoRec = ImpoRec + DBLet(RS!imporrec, "N")
        
        I = I + 1
        RS.MoveNext
    Wend
    Set RS = Nothing
    
    If SqlValues <> "" Then
        SqlValues = Mid(SqlValues, 1, Len(SqlValues) - 1)
        Conn.Execute SqlInsert & SqlValues
    End If
    
    
    RecalcularTotales = RecalcularTotalesFactura
    Exit Function
    
eRecalcularTotales:
    MuestraError Err.Number, "Recalcular Totales", Err.Description
End Function


Private Function RecalcularTotalesFactura() As Boolean
Dim Sql As String
Dim SqlInsert As String
Dim SqlValues As String
Dim I As Long
Dim RS As ADODB.Recordset

Dim Baseimpo As Currency
Dim Basereten As Currency
Dim Impoiva As Currency
Dim ImpoRec As Currency
Dim Imporeten As Currency
Dim TotalFactura As Currency
Dim PorcRet As Currency

Dim TipoRetencion As Integer

    On Error GoTo eRecalcularTotalesFactura

    RecalcularTotalesFactura = False

    TipoRetencion = DevuelveValor("select tipo from usuarios.wtiporeten where codigo = " & DBSet(Combo1.ListIndex, "N"))
    
    Baseimpo = 0
    Basereten = 0
    Impoiva = 0
    Imporeten = 0
    ImpoRec = 0
    TotalFactura = 0
    
    Sql = "select aplicret, sum(importe) baseimpo, sum(coalesce(impoiva,0)) imporiva, sum(coalesce(imporec,0)) imporrec from factcli_lineas "
    Sql = Sql & " where numserie = " & DBSet(Text1(2).Text, "T") & " and numfactu = " & DBSet(Text1(0).Text, "N") & " and fecfactu = " & DBSet(Text1(1).Text, "F")
    Sql = Sql & " group by 1 order by 1"
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        Baseimpo = Baseimpo + DBLet(RS!Baseimpo, "N")
        Impoiva = Impoiva + DBLet(RS!ImporIVA, "N")
        ImpoRec = ImpoRec + DBLet(RS!imporrec, "N")
        
        If RS!aplicret = 1 Then
            Basereten = Basereten + DBLet(RS!Baseimpo, "N")
            
            If TipoRetencion = 1 Then
                Basereten = Basereten + DBLet(RS!ImporIVA, "N")
            End If
        End If
        
        RS.MoveNext
    Wend
    Set RS = Nothing
    
    PorcRet = ImporteFormateado(Text1(7).Text)
    
    If PorcRet = 0 Then Basereten = 0
    
    If PorcRet = 0 Then
        Imporeten = 0
    Else
        Imporeten = Round((PorcRet * Basereten / 100), 2)
    End If
    
    TotalFactura = Baseimpo + Impoiva + ImpoRec - Imporeten
    
    Text1(8).Text = Format(Baseimpo, FormatoImporte)
    Text1(10).Text = Format(Basereten, FormatoImporte)
    Text1(9).Text = Format(Impoiva, FormatoImporte)
    Text1(11).Text = Format(Imporeten, FormatoImporte)
    Text1(12).Text = Format(TotalFactura, FormatoImporte)
    
    If PorcRet = 0 Then
        Text1(10).Text = ""
        Text1(11).Text = ""
    End If
    
    Sql = "update factcli set "
    Sql = Sql & " totbases = " & DBSet(Baseimpo, "N")
    Sql = Sql & ", totivas = " & DBSet(Impoiva, "N")
    Sql = Sql & ", totrecargo = " & DBSet(ImpoRec, "N")
    Sql = Sql & ", totfaccl = " & DBSet(TotalFactura, "N")
    Sql = Sql & ", totbasesret = " & DBSet(Basereten, "N", "S")
    Sql = Sql & ", trefaccl = " & DBSet(Imporeten, "N", "S")
    Sql = Sql & " where numserie= " & DBSet(Text1(2).Text, "T") & " and numfactu= " & DBSet(Text1(0).Text, "N") & " and fecfactu = " & DBSet(Text1(1).Text, "F")
    
    Conn.Execute Sql
    
    
    RecalcularTotalesFactura = True
    Exit Function
    
eRecalcularTotalesFactura:
    MuestraError Err.Number, "Recalcular Totales Factura", Err.Description
End Function


Private Function IntegrarFactura() As Boolean
Dim SqlLog As String

    IntegrarFactura = False
    
    SqlLog = "Factura : " & Text1(2).Text & " " & Text1(0).Text & " de fecha " & Text1(1).Text
    SqlLog = SqlLog & vbCrLf & "Línea   : " & DBLet(Me.AdoAux(1).Recordset!numlinea, "N")
    SqlLog = SqlLog & vbCrLf & "Cuenta  : " & DBLet(Me.AdoAux(1).Recordset!Codmacta, "T") & " " & DBLet(Me.AdoAux(1).Recordset!Nommacta, "T")
    SqlLog = SqlLog & vbCrLf & "Importe : " & DBLet(Me.AdoAux(1).Recordset!Baseimpo, "N")
    
End Function




Private Function TieneRegistros() As Boolean
    On Error Resume Next
    TieneRegistros = False
    If Data1.Recordset.RecordCount > 0 Then TieneRegistros = True
End Function

Private Sub ImprimirFactura()

    cadFormula = "2"
    If Modo = 2 Then
        If Not Me.Data1.Recordset.EOF Then cadFormula = Data1.Recordset!numserie & "|" & Data1.Recordset!NumFactu & "|" & Data1.Recordset!FecFactu & "|"
    End If
    frmFacturasList.DatosFactura = cadFormula
    frmFacturasList.Show vbModal
   
End Sub



Private Sub InsertarEnCobros()
Dim RR As ADODB.Recordset
Dim EsUNaCuota As Boolean
    
    If Modo = 4 Then
        'Borramos
        cadFormula = ObtenerWhereCab(True)
        cadFormula = Replace(cadFormula, "factcli.", "factcli_vtos.")
        Conn.Execute "DELETE from factcli_vtos " & cadFormula
        
        'En la conta
        
        cadFormula = ObtenerWhereCab(True)
        cadFormula = Replace(cadFormula, "factcli.", "cobros.")
        Conn.Execute "DELETE from ariconta" & vParam.Numconta & ".cobros " & cadFormula
        
    End If
    Set RR = New ADODB.Recordset

    'Habra que  hacer una clase, de momento, por velocidad, no tengo otra opcion
    cadFormula = ObtenerWhereCab(False)
    cadFormula = " from factcli,clientes where factcli.codclien=clientes.codclien AND " & cadFormula
    cadFormula = "licencia,PobClien ,codposta ,ProClien ,NIFClien ,codpais ,IBAN ,totfaccl " & cadFormula
    cadFormula = "SELECT factcli.codclien, factcli.codforpa ,numserie,NumFactu ,FecFactu ,NomClien ,DomClien," & cadFormula
 
    RR.Open cadFormula, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    'Tesoreria
    cadFormula = vParam.BancoPropioFacturacionContabilidad()
    EsUNaCuota = False
    If Text1(2).Text = "ASO" Or Text1(2).Text = "CUO" Then EsUNaCuota = True
   
    
    
    InsertarEnTesoreria EsUNaCuota, RR, cadFormula, "", Msg
        
End Sub


Private Function PuedeModificarFactura() As Boolean
    
    PuedeModificarFactura = False
    
    If Val(Data1.Recordset!intconta) = 1 Then
        MsgBox "Factura contabilizada", vbExclamation
        Exit Function
    End If
    
    
    'Si el periodo...
    
    
    
    
    'Llegado aqui, ya puede
    PuedeModificarFactura = True
End Function
