VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmExpediente 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   10620
   ClientLeft      =   -15
   ClientTop       =   -30
   ClientWidth     =   14640
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExpediente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10620
   ScaleWidth      =   14640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lw1 
      Height          =   1335
      Left            =   7200
      TabIndex        =   63
      Top             =   4920
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   2355
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   177
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Base"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "% IVA"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "IVA €"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Rec "
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Total"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame FrameAux2 
      BorderStyle     =   0  'None
      Height          =   2340
      Left            =   6120
      TabIndex        =   51
      Top             =   2040
      Width           =   8295
      Begin VB.TextBox txtaux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   350
         Index           =   16
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   19
         Tag             =   "Importe|N|N|||expedientes_acuenta|importe|#,##0.00||"
         Text            =   "importe"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.TextBox txtaux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   350
         Index           =   14
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   17
         Tag             =   "Fecha|F|N|||expedientes_acuenta|fechaent|dd/mm/yyyy||"
         Text            =   "fec"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.TextBox txtAux2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   350
         Index           =   0
         Left            =   2760
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   60
         Text            =   "ampconce"
         Top             =   1200
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   1
         Left            =   2640
         TabIndex        =   59
         ToolTipText     =   "Buscar cuenta"
         Top             =   1440
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtaux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   350
         Index           =   11
         Left            =   480
         TabIndex        =   57
         Tag             =   "Exp|N|S|||expedientes_acuenta|numexped|0000|S|"
         Text            =   "ex"
         Top             =   1440
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Frame FrameToolAux 
         Height          =   555
         Index           =   2
         Left            =   120
         TabIndex        =   55
         Top             =   0
         Width           =   1605
         Begin MSComctlLib.Toolbar ToolbarAux 
            Height          =   330
            Index           =   2
            Left            =   180
            TabIndex        =   56
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
         Height          =   350
         Index           =   15
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   18
         Tag             =   "Forma pago|N|N|||expedientes_acuenta|codforpa|000||"
         Text            =   "forpa"
         Top             =   1200
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   350
         Index           =   13
         Left            =   960
         TabIndex        =   16
         Tag             =   "Linea|N|N|||expedientes_acuenta|numlinea|000|S|"
         Text            =   "linea"
         Top             =   1200
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   350
         Index           =   12
         Left            =   360
         TabIndex        =   54
         Tag             =   "Año|N|S|||expedientes_acuenta|anoexped||S|"
         Text            =   "año"
         Top             =   1560
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   350
         Index           =   9
         Left            =   120
         TabIndex        =   53
         Tag             =   "TipoReg|N|S|||expedientes_acuenta|tiporegi||S|"
         Text            =   "tipoR"
         Top             =   1080
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   350
         Index           =   10
         Left            =   480
         MaxLength       =   10
         TabIndex        =   52
         Tag             =   "Serie|T|S|||expedientes_acuenta|numserie||S|"
         Text            =   "seri"
         Top             =   1200
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSAdodcLib.Adodc AdoAux 
         Height          =   375
         Index           =   2
         Left            =   360
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
         Height          =   1680
         Index           =   2
         Left            =   120
         TabIndex        =   58
         Top             =   600
         Width           =   8010
         _ExtentX        =   14129
         _ExtentY        =   2963
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Pagos a cuenta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   270
         Index           =   7
         Left            =   1920
         TabIndex        =   62
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   3105
      Index           =   6
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Tag             =   "To|T|S|||Expedientes|observac|||"
      Text            =   "frmExpediente.frx":000C
      Top             =   3120
      Width           =   5805
   End
   Begin VB.Frame FrameAux1 
      BorderStyle     =   0  'None
      Height          =   3420
      Left            =   120
      TabIndex        =   45
      Top             =   6480
      Width           =   14295
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   2
         Left            =   6480
         TabIndex        =   65
         ToolTipText     =   "Buscar cuenta"
         Top             =   2160
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   0
         Left            =   4800
         TabIndex        =   48
         ToolTipText     =   "Buscar cuenta"
         Top             =   2190
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtaux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   350
         Index           =   1
         Left            =   840
         MaxLength       =   10
         TabIndex        =   7
         Tag             =   "Serie|T|S|||expedientes_lineas|numserie||S|"
         Text            =   "seri"
         Top             =   2145
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtaux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   350
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Tag             =   "TipoReg|N|S|||expedientes_lineas|tiporegi||S|"
         Text            =   "tipoR"
         Top             =   2160
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   350
         Index           =   2
         Left            =   2220
         TabIndex        =   8
         Tag             =   "Exp|N|S|||expedientes_lineas|numexped|0000|S|"
         Text            =   "exped"
         Top             =   2160
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   350
         Index           =   4
         Left            =   3330
         TabIndex        =   10
         Tag             =   "Linea|N|N|||expedientes_lineas|numlinea|000|S|"
         Text            =   "linea"
         Top             =   2160
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   350
         Index           =   5
         Left            =   4050
         MaxLength       =   15
         TabIndex        =   11
         Tag             =   "Concepto|N|N|||expedientes_lineas|codconce|000||"
         Text            =   "concep"
         Top             =   2160
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Frame FrameToolAux 
         Height          =   555
         Index           =   1
         Left            =   60
         TabIndex        =   46
         Top             =   0
         Width           =   1605
         Begin MSComctlLib.Toolbar ToolbarAux 
            Height          =   330
            Index           =   1
            Left            =   180
            TabIndex        =   47
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
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   350
         Index           =   7
         Left            =   9120
         MaxLength       =   15
         TabIndex        =   14
         Tag             =   "Importe|N|N|||expedientes_lineas|importe|#,##0.00||"
         Text            =   "importe"
         Top             =   2160
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.TextBox txtaux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   350
         Index           =   6
         Left            =   8370
         TabIndex        =   13
         Tag             =   "Ampli|T|S|||expedientes_lineas|ampliaci|||"
         Text            =   "Ampliaci"
         Top             =   2160
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   350
         Index           =   8
         Left            =   5160
         MaxLength       =   50
         TabIndex        =   12
         Tag             =   "Conceot|T|N|||expedientes_lineas|nomconce|||"
         Text            =   "nomconce"
         Top             =   2160
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.TextBox txtaux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   350
         Index           =   3
         Left            =   2910
         TabIndex        =   9
         Tag             =   "Año|N|S|||expedientes_lineas|anoexped||S|"
         Text            =   "año"
         Top             =   2160
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.CheckBox chkAux 
         BackColor       =   &H80000005&
         Height          =   255
         Index           =   0
         Left            =   11160
         TabIndex        =   15
         Tag             =   "Pagado|N|N|0|1|expedientes_lineas|pagado|||"
         Top             =   2160
         Visible         =   0   'False
         Width           =   285
      End
      Begin MSAdodcLib.Adodc AdoAux 
         Height          =   375
         Index           =   1
         Left            =   360
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
      Begin VB.TextBox txtAux2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   350
         Index           =   5
         Left            =   10080
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   49
         Text            =   "ampconce"
         Top             =   2160
         Visible         =   0   'False
         Width           =   885
      End
      Begin MSDataGridLib.DataGrid DataGridAux 
         Height          =   2640
         Index           =   1
         Left            =   120
         TabIndex        =   50
         Top             =   720
         Width           =   14250
         _ExtentX        =   25135
         _ExtentY        =   4657
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lineas de expediente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   270
         Index           =   6
         Left            =   1800
         TabIndex        =   61
         Top             =   240
         Width           =   2355
      End
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   3
      Left            =   3600
      MaxLength       =   10
      TabIndex        =   1
      Tag             =   "Num albar|N|N|||Expedientes|numexped|00000|S|"
      Text            =   "Text1"
      Top             =   1320
      Width           =   1875
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   360
      Index           =   1
      Left            =   1320
      MaxLength       =   30
      TabIndex        =   36
      Text            =   "commor"
      Top             =   2190
      Width           =   4515
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   9240
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Tag             =   "Situacion|N|N|||Expedientes|codsitua|||"
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   360
      Index           =   0
      Left            =   720
      MaxLength       =   30
      TabIndex        =   33
      Text            =   "commor"
      Top             =   1320
      Width           =   2595
   End
   Begin VB.Frame FrameDesplazamiento 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   3810
      TabIndex        =   30
      Top             =   180
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   31
         Top             =   180
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   120
      TabIndex        =   28
      Top             =   180
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   29
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
   Begin VB.CheckBox chkVistaPrevia 
      Caption         =   "Vista previa"
      Height          =   300
      Left            =   6480
      TabIndex        =   22
      Top             =   480
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   5
      Left            =   120
      MaxLength       =   30
      TabIndex        =   3
      Tag             =   "Cod. cliente|N|N|||Expedientes|codclien|00000|N|"
      Text            =   "Text1"
      Top             =   2190
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   2
      Left            =   120
      TabIndex        =   0
      Tag             =   "Tipo registro|T|N|||Expedientes|tiporegi||S|"
      Text            =   "0000000000"
      Top             =   1320
      Width           =   465
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      TabIndex        =   24
      Top             =   9960
      Width           =   4215
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   210
         Width           =   3795
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   13320
      TabIndex        =   21
      Top             =   10065
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   12120
      TabIndex        =   20
      Top             =   10080
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   4200
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
   Begin VB.CommandButton cmdRegresar 
      Cancel          =   -1  'True
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   13320
      TabIndex        =   23
      Top             =   10080
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   390
      Left            =   14160
      TabIndex        =   32
      Top             =   120
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
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   4
      Left            =   5760
      TabIndex        =   2
      Tag             =   "Fec. expediente|FH|N|||Expedientes|fecexped|dd/mm/yyyy hh:nn:ss||"
      Text            =   "Text1"
      Top             =   1320
      Width           =   3195
   End
   Begin VB.Frame FrameDatosClavePrimariFalta 
      Caption         =   "Frame2"
      Enabled         =   0   'False
      Height          =   1575
      Left            =   480
      TabIndex        =   39
      Top             =   3720
      Width           =   4815
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   8
         Left            =   480
         TabIndex        =   44
         Tag             =   "fechaaccion|FH|N|||Expedientes|fecha|dd/mm/yyyy hh:nn:ss||"
         Text            =   "0000000000"
         Top             =   960
         Width           =   2385
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   7
         Left            =   2280
         TabIndex        =   43
         Tag             =   "usuario|T|N|||Expedientes|usuario|||"
         Text            =   "usuario"
         Top             =   480
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   1
         Left            =   360
         TabIndex        =   41
         Tag             =   "Tipo registro|T|N|||Expedientes|numserie||S|"
         Text            =   "tipo"
         Top             =   480
         Width           =   465
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   0
         Left            =   1200
         TabIndex        =   40
         Tag             =   "Tipo registro|T|N|||Expedientes|anoexped||S|"
         Text            =   "anoexp"
         Top             =   480
         Width           =   825
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Totales"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   6120
      TabIndex        =   64
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Número"
      Height          =   225
      Index           =   4
      Left            =   3600
      TabIndex        =   42
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Situación"
      Height          =   240
      Index           =   0
      Left            =   9240
      TabIndex        =   38
      Top             =   1080
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nombre"
      Height          =   240
      Index           =   1
      Left            =   2040
      TabIndex        =   37
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Observaciones"
      Height          =   240
      Index           =   11
      Left            =   120
      TabIndex        =   35
      Top             =   2880
      Width           =   1980
   End
   Begin VB.Image imgppal 
      Height          =   240
      Index           =   2
      Left            =   6480
      Picture         =   "frmExpediente.frx":0010
      Top             =   1080
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      Height          =   225
      Index           =   2
      Left            =   5760
      TabIndex        =   34
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Image imgCC 
      Height          =   480
      Left            =   1440
      Top             =   1920
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cod. cliente"
      Height          =   240
      Index           =   5
      Left            =   120
      TabIndex        =   27
      Top             =   1965
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tipo expediente"
      Height          =   240
      Index           =   3
      Left            =   120
      TabIndex        =   26
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
         Shortcut        =   ^F
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
Attribute VB_Name = "frmExpediente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public numExpediente As String   'tiporegi|numer|anño|
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)



Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private Const IdPrograma = ID_Expedientes
Private WithEvents frmCC As frmBasico
Attribute frmCC.VB_VarHelpID = -1
Private WithEvents frmco As frmConceptos
Attribute frmco.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom
Attribute frmZ.VB_VarHelpID = -1
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'  Variables comunes a todos los formularios
Private Modo As Byte
Private CadenaConsulta As String
Private CadB As String

Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean
Private DevfrmCCtas As String

Private BuscaChekc As String
Dim NumTabMto As Integer 'Indica quin nº de Tab està en modo Mantenimient
Dim ModoLineas As Byte
Dim PrimeraVez As Boolean 'para los grids


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub



Private Sub chkAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
    Dim Cad As String
    Dim I As Integer
    
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case Modo
    Case 3
        If DatosOK Then
            '-----------------------------------------
            'Hacemos insertar
            If InsertarDesdeForm2(Me, 1) Then
                Cad = "UPDATE contadores set numalbar=numalbar + 1 where tiporegi='0'"
                Ejecuta Cad
                SituarData1
                PonerModo 2
                lblIndicador.Caption = ""
            End If
        End If
    Case 4
            'Modificar
            If DatosOK Then
                '-----------------------------------------
                'Hacemos insertar
                If ModificaDesdeFormulario2(Me, 1) Then
                    TerminaBloquear
                    lblIndicador.Caption = ""
                    If SituarData1 Then
                        PonerModo 2
                        
                    Else
                        LimpiarCampos
                        PonerModo 0
                    End If
                End If
            End If
    Case 5
            Select Case ModoLineas
                Case 1 'afegir llínia
                    InsertarLinea
                Case 2 'modificar llínies
                    ModificarLinea
                                        
                    '**** parte de contabilizacion de la factura
                    TerminaBloquear
            
                    
                    PosicionarData
             End Select

    Case 1
        HacerBusqueda
    End Select
        
Error1:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MsgBox Err.Number & " - " & Err.Description, vbExclamation
End Sub

Private Sub cmdAux_Click(Index As Integer)
    CadB = ""
    Select Case Index
    Case 0
        Set frmco = New frmConceptos
        frmco.DatosADevolverBusqueda = "0|1|3|"
        frmco.Show vbModal
        Set frmco = Nothing
        If CadB <> "" Then
            txtaux(5).Text = Format(RecuperaValor(CadB, 1), "000")
            txtaux(8).Text = RecuperaValor(CadB, 2)
            CadB = RecuperaValor(CadB, 3)
            If CadB = "" Then CadB = "0"
            txtaux(7).Text = Format(CadB, FormatoImporte)
            PonerFoco txtaux(6)
        End If

    Case 1
        'IVA
        
        Set frmCC = New frmBasico '
        AyudaFormaPago frmCC
        Set frmCC = Nothing
        If CadB <> "" Then
            
            txtaux(15).Text = Format(RecuperaValor(CadB, 1), "000")
            txtAux2(0).Text = RecuperaValor(CadB, 2)
            PonerFoco txtaux(16)
        End If
            
            
    Case 2
            'Observaciones
            If Modo <> 2 And Modo <> 5 Then Exit Sub
            If txtaux(6).Visible Then
                CadB = txtaux(6).Text
            Else
                If Me.AdoAux(1).Recordset.EOF Then Exit Sub
                CadB = DBLet(Me.AdoAux(1).Recordset!ampliaci, "T")
            End If
            Set frmZ = New frmZoom
            frmZ.pValor = CadB
            If txtaux(6).Visible Then
                frmZ.pModo = 3
            Else
                frmZ.pModo = Modo
            End If
            frmZ.Caption = "Ampliacion linea expediente"
            frmZ.Show vbModal
            Set frmZ = Nothing
            If txtaux(6).Visible Then
                If CadB <> "" Then txtaux(6).Text = CadB
            End If
    End Select
    CadB = ""
End Sub

Private Sub cmdCancelar_Click()
Dim InserLin As Boolean

Select Case Modo
Case 1, 3
    LimpiarCampos
    PonerModo 0
Case 4
    'Modificar
    lblIndicador.Caption = ""
    TerminaBloquear
    PonerModo 2
    PonerCampos
    
    
Case 5
                TerminaBloquear
            InserLin = False
            If ModoLineas = 1 Then 'INSERTAR
                ModoLineas = 0
                DataGridAux(NumTabMto).AllowAddNew = False
                If Not AdoAux(NumTabMto).Recordset.EOF Then AdoAux(NumTabMto).Recordset.MoveFirst
                
                InserLin = True
                
            End If
            ModoLineas = 0
            LLamaLineas NumTabMto, 0, 0
                        
            If InserLin Then
                PosicionarData
                PonerCampos
            Else
                PonerModo 2
            End If
End Select

End Sub


' Cuando modificamos el data1 se mueve de lugar, luego volvemos
' ponerlo en el sitio
' Para ello con find y un SQL lo hacemos
' Buscamos por el codigo, que estara en un text u  otro
' Normalmente el text(0)
Private Function SituarData1() As Boolean
    Dim Sql As String
    On Error GoTo ESituarData1
            'Actualizamos el recordset
            If numExpediente <> "" Then
                Sql = " numexped = " & Text1(3).Text & " AND anoexped =" & Text1(0).Text
                Data1.RecordSource = "select * from " & NombreTabla & " WHERE " & Sql
            End If
            Data1.Refresh
            '#### A mano.
            'El sql para que se situe en el registro en especial es el siguiente
            Sql = " numexped = " & Text1(3).Text & " AND anoexped =" & Text1(0).Text
            
            SituarDataMULTI Data1, Sql, Me.lblIndicador, False
            If Data1.Recordset.EOF Then GoTo ESituarData1
            SituarData1 = True
        Exit Function
ESituarData1:
        If Err.Number <> 0 Then Err.Clear
        Limpiar Me
        PonerModo 0
        lblIndicador.Caption = ""
        SituarData1 = False
End Function

Private Sub BotonAnyadir()
    LimpiarCampos
    'Añadiremos el boton de aceptar y demas objetos para insertar
    cmdAceptar.Caption = "Aceptar"
    PonerModo 3
    'Escondemos el navegador y ponemos insertando
    DespalzamientoVisible False
    SugerirCodigoSiguiente
    Combo1.ListIndex = 0
    
    '###A mano
    Text1(4).SetFocus
End Sub

Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        '### A mano
        '################################################
        'Si pasamos el control aqui lo ponemos en amarill
        PonFoco Text1(4)
        Else
            HacerBusqueda
            If Data1.Recordset.EOF Then
                 '### A mano
                Text1(kCampo).Text = ""
                PonFoco Text1(kCampo)
            End If
    End If
End Sub

Private Sub BotonVerTodos()
    'Ver todos
    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub

Private Sub Desplazamiento(Index As Integer)
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
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub

Private Sub BotonModificar()
    '---------
    'MODIFICAR
    '----------
    'Añadiremos el boton de aceptar y demas objetos para insertar
   ' cmdAceptar.Caption = "Modificar"
    PonerModo 4
    'Escondemos el navegador y ponemos insertando
    'Como el campo 1 es clave primaria, NO se puede modificar
    '### A mano
    BloquearTxt Text1(3), True, True
    
    DespalzamientoVisible False
    PonFoco Text1(8)
End Sub

Private Sub BotonEliminar()

'
    Dim Cad As String
    Dim I As Integer

    If Modo <> 2 Then Exit Sub

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    'Comprobamos si se puede eliminar
    I = 0
    If Not SePuedeEliminar Then I = 1
     
    Set miRsAux = Nothing
    If I = 1 Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '### a mano
    
    
    
    Cad = "Seguro que desea eliminar de el expediente:"
    Cad = Cad & vbCrLf & "ID: " & Text1(3).Text
    Cad = Cad & vbCrLf & "Cliente: " & Data1.Recordset!codclien & " " & Me.Text2(1).Text
    
    
    I = MsgBox(Cad, vbQuestion + vbYesNo)
    'Borramos
    If I = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        Data1.Recordset.Delete
        Data1.Refresh
        If Data1.Recordset.EOF Then
            'Solo habia un registro
            LimpiarCampos
            PonerModo 0
            If Me.numExpediente <> "" Then Unload Me
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
    End If
Error2:
        Screen.MousePointer = vbDefault
        If Err.Number > 0 Then MsgBox Err.Number & " - " & Err.Description
End Sub




Private Sub cmdRegresar_Click()

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If



    
    

    RaiseEvent DatoSeleccionado(CStr(Text1(4).Text & "|" & Text2(4).Text & "|"))
    Unload Me
    Screen.MousePointer = vbDefault
End Sub



Private Sub DataGridAux_DblClick(Index As Integer)
    If Index = 1 Then cmdAux_Click 2
   
End Sub

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

Private Sub Form_Activate()
    If Me.Tag = 1 Then
        Me.Tag = 0
        Data1.ConnectionString = Conn
        If numExpediente <> "" Then
            CadB = RecuperaValor(numExpediente, 1)
            If CadB <> "-1" Then
                CadB = "tiporegi = " & CadB & " AND numexped =" & RecuperaValor(numExpediente, 2)
                CadB = CadB & " AND anoexped =" & RecuperaValor(numExpediente, 3)
                CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
                 
                PonerCadenaBusqueda
            Else
                'Como no esta establecido
                Data1.RecordSource = "Select * from " & NombreTabla & " WHERE numexped = -1"
                Data1.Refresh
            
                BotonAnyadir
                Text1(5).Text = RecuperaValor(numExpediente, 2)
                Text2(1).Text = RecuperaValor(numExpediente, 3)
            End If
        Else
            'ASignamos un SQL al DATA1
            Data1.RecordSource = "Select * from " & NombreTabla & " WHERE numexped = -1"
            Data1.Refresh
            If DatosADevolverBusqueda = "" Then
                PonerModo 0
            Else
                PonerModo 1
            End If
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


'++
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then Unload Me
End Sub
'++


Private Sub Form_Load()
Dim I As Integer

    Me.Tag = 1
    Me.Icon = frmppal.Icon


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
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With

    For I = 1 To 2
        With Me.ToolbarAux(I)
            .HotImageList = frmppal.imgListComun_OM16
            .DisabledImageList = frmppal.imgListComun_BN16
            .ImageList = frmppal.imgListComun16
            .Buttons(1).Image = 3
            .Buttons(2).Image = 4
            .Buttons(3).Image = 5
        End With
    Next


    Me.imgCC.Picture = frmppal.imgIcoForms.ListImages(1).Picture

    
    DespalzamientoVisible False


    LimpiarCampos

    
    '## A mano
    NombreTabla = "expedientes"
    Ordenacion = " ORDER BY anoexped,numexped"
    CargaDatosFijos
    
        
    PonerOpcionesMenu
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    PrimeraVez = True

End Sub

Private Sub CargaDatosFijos()
    Me.Combo1.Clear
    
    Set miRsAux = New ADODB.Recordset
    BuscaChekc = "select   codsitua , nomsitua from tipositexped order by codsitua"
    miRsAux.Open BuscaChekc, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    While Not miRsAux.EOF
        Combo1.AddItem miRsAux!nomsitua
        Combo1.ItemData(Combo1.NewIndex) = I
        I = I + 1
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing

    BuscaChekc = ""
End Sub

Private Sub LimpiarCampos()
    Limpiar Me   'Metodo general
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.Combo1.ListIndex = -1


    lw1.ListItems.Clear

  
    
End Sub




Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub




Private Sub frmCC_DatoSeleccionado(CadenaSeleccion As String)
    'Centro de coste
    CadB = CadenaSeleccion
End Sub



Private Sub frmco_DatoSeleccionado(CadenaSeleccion As String)
    CadB = CadenaSeleccion
End Sub

Private Sub frmF_Selec(vFecha As Date)
    CadB = Format(vFecha, formatoFechaVer)
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
    CadB = vCampo
End Sub

Private Sub imgCC_Click()
    'Lanzaremos el vista previa
    CadenaDesdeOtroForm = ""
    frmcolClientesBusqueda.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
        Text1(5).Text = CadenaDesdeOtroForm
        Text1_LostFocus 5
    End If
End Sub

Private Sub imgCuentas_Click(Index As Integer)
 Screen.MousePointer = vbHourglass

End Sub


Private Sub imgppal_Click(Index As Integer)
    If Modo = 2 Or Modo = 0 Then Exit Sub
    Set frmF = New frmCal
    frmF.Fecha = Now
    CadB = ""
    If Me.Text1(Index + 2).Text <> "" Then frmF.Fecha = Text1(Index + 2).Text
    frmF.Show vbModal
    If CadB <> "" Then Text1(Index + 2).Text = CadB & " " & Format(Now, "hh:mm:ss")
    
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




'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 4: KEYCta KeyAscii, 4
            Case 5: KEYCta KeyAscii, 5
            Case 10: KEYCta KeyAscii, 10
            Case 12: KEYCta KeyAscii, 12
            Case 13: KEYCta KeyAscii, 13
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYCta(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgCuentas_Click (indice)
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
' Cunado el campo de texto pierde el enfoque
' Es especifico de cada formulario y en el podremos controlar
' lo que queramos, desde formatear un campo si asi lo deseamos
' hasta pedir que nos devuelva los datos de la empresa
'----------------------------------------------------------------
'----------------------------------------------------------------
Private Sub Text1_LostFocus(Index As Integer)
    Dim Valor As Currency
    Dim Sql As String
    Dim mTag As CTag
    Dim I As Integer
    Dim Sql2 As String
    
    
    
    If Modo <> 2 Then
        If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    Else
        Exit Sub
    End If
  
    'Si queremos hacer algo ..
    Select Case Index
           
        Case 5
            If Text1(Index).Text = "" Then
                Text2(1).Text = ""
                Exit Sub
            End If
            If Not PonerFormatoEntero(Text1(Index)) Then
                DevfrmCCtas = ""
            Else
                DevfrmCCtas = DevuelveDesdeBD("nomclien", "clientes", "codclien", Text1(Index).Text)
                If DevfrmCCtas = "" Then
                    MsgBox "No existe cliente: " & Text1(Index).Text, vbExclamation
                    Text1(Index).Text = ""
                    Exit Sub
                
                    
                End If
            End If
            Text2(1).Text = DevfrmCCtas
            
        Case 4
            PonerFormatoFechaHora Text1(Index)
        '....
    End Select
    '---
End Sub

Private Sub HacerBusqueda()
Dim Cad As String
Dim CadB As String

CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)

 

If chkVistaPrevia = 1 Then
    MandaBusquedaPrevia CadB
    Else
        'Se muestran en el mismo form
        If CadB <> "" Then
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
            PonerCadenaBusqueda
        End If
End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)

'    Set frmBan = New frmBasico2
'
'    AyudaBanco frmBan, , CadB
'
'    Set frmBan = Nothing

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

Private Sub PonerCampos()
Dim N  As Byte
    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1
    Text2(0).Text = "EXPEDIENTES"
   
    Text2(1).Text = DevuelveDesdeBD("nomclien", "clientes", "codclien", Text1(5).Text)
    
    
    For N = 1 To DataGridAux.Count ' - 1
 
            CargaGrid CInt(N), True
            If Not AdoAux(N).Recordset.EOF Then _
                PonerCamposForma2 Me, AdoAux(N), 2, "FrameAux" & N

    Next N
    
    lblIndicador.Caption = "Totales"
    lblIndicador.Refresh
    Totales
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount

End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
Private Sub PonerModo(Kmodo As Integer, Optional indFrame As Integer)
    Dim I As Integer
    Dim B As Boolean
    Dim Obj
    
    BuscaChekc = ""
    
    Modo = Kmodo
    
    B = (Modo = 2)
    If B Then
        PonerContRegIndicador
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    B = (Modo = 0 Or Modo = 2)
    
    'chkVistaPrevia.Visible = (Modo = 1)
    
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)
    DespalzamientoVisible B And Me.Data1.Recordset.RecordCount > 1
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.Visible = B
    Else
        cmdRegresar.Visible = False
    End If

    
    
    
    'Modo insertar o modificar
    B = (Kmodo >= 3) '-->Luego not b sera kmodo<3
    cmdAceptar.Visible = B Or Modo = 1
    cmdCancelar.Visible = B Or Modo = 1
    mnOpciones.Enabled = Not B
    If cmdCancelar.Visible Then
        cmdCancelar.Cancel = True
        Else
        cmdCancelar.Cancel = False
    End If
    Toolbar1.Buttons(6).Enabled = Not B And vUsu.Nivel < 2
    Toolbar1.Buttons(1).Enabled = Not B
    Toolbar1.Buttons(2).Enabled = Not B
    
    If Kmodo = 0 Then lblIndicador.Caption = ""
    
     Dim anc As Single
    anc = DataGridAux(1).top
    If DataGridAux(1).Row < 0 Then
        anc = anc + 230
    Else
        anc = anc + DataGridAux(1).RowTop(DataGridAux(1).Row) + 5
    End If
    For I = 1 To 2
        If Modo = 1 Then
            LLamaLineas I, Modo, anc
        Else
            LLamaLineas I, 3, anc
        End If
    Next
    For I = 0 To txtaux.Count - 1
        txtaux(I).BackColor = vbWhite
    Next I
    
    
    
    
    '### A mano
    'Aqui añadiremos controles para datos especificos. Esto es, si hay imagenes en el form
    ' o cualquier objeto que dependiendo en el modo en el que esteos se visualizaran o no
    ' Bloqueamos los campos de texto y demas controles en funcion
    ' del modo en el que estamos.
    ' Es decir, si estamos en modo busqueda, insercion o modificacion estaran enables
    ' si no  disable. la variable b nos devuelve esas opciones
    BloquearTxt Me.Text1(2), Modo <> 1
    BloquearTxt Me.Text1(3), Modo <> 1, True
    B = Modo = 2 Or Modo = 0
    BloquearTxt Me.Text1(4), B
    BloquearTxt Me.Text1(5), B
    BloquearTxt Me.Text1(6), B
    Combo1.Locked = B
    

    Me.imgCC.Enabled = Not B
    
    
    If (Modo < 2) Or (Modo = 3) Then
        CargaGrid 1, False
        CargaGrid 2, False
    End If
    B = (Modo = 4) Or (Modo = 2)
    DataGridAux(1).Enabled = B
    
    
    PonerModoUsuarioGnral Modo, "arigestion"

    
End Sub

Private Sub PonerContRegIndicador()
'si estamos en modo ver registros muestra el numero de registro en el que estamos
'situados del total de registros mostrados: 1 de 24
Dim cadReg As String

    If (Modo = 2 Or Modo = 0) Then
        cadReg = PonerContRegistros(Me.Data1)
        If CadB = "" Then
            lblIndicador.Caption = cadReg
        Else
            lblIndicador.Caption = "BUSQUEDA: " & cadReg
        End If
    End If
End Sub


Private Function DatosOK() As Boolean
Dim B As Boolean

    
    DatosOK = False
    B = CompForm2(Me, 1)
    If Not B Then Exit Function

  
    If Modo = 3 And B Then
        Text1(0).Text = Year(CDate(Text1(4).Text))
        Text1(3).Text = Val(DevuelveDesdeBD("numalbar", "contadores", "tiporegi", "0")) + 1
        
    End If
    DatosOK = B
End Function

'Esto es para que cuando pincha en siguiente le sugerimos
'Se puede comentar todo y asi no hace nada ni da error
'El SQL es propio de cada tabla
Private Sub SugerirCodigoSiguiente()
    
    
    Text1(0).Text = "0"  'año exp
    Text1(1).Text = "AEX"
    Text1(2).Text = "0"
    Text2(0).Text = "EXPEDIENTES"
    Text1(3).Text = "000000"
            Text1(7).Text = vUsu.Login
        Text1(8).Text = Now
    
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            If Modo = 2 Or Modo = 0 Then BotonAnyadir
        Case 2
            If Modo = 2 Then
                If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
            End If
                
        Case 3
            If Modo = 2 Then BotonEliminar
        Case 5
            If Modo = 2 Or Modo = 0 Then BotonBuscar
        Case 6
            If Modo = 2 Or Modo = 0 Then BotonVerTodos
        
        Case 8
            ImprimirExpe
        Case Else
    
    End Select
End Sub


Private Sub DespalzamientoVisible(bol As Boolean)
    FrameDesplazamiento.Visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub






Private Sub PonerFoco(ByRef Text As TextBox)
    On Error Resume Next
    Text.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub PonerOpcionesMenu()
PonerOpcionesMenuGeneral Me
End Sub



Private Function SePuedeEliminar() As Boolean
Dim B As Boolean
Dim Cad As String

    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset
    
    SePuedeEliminar = False
    
    If Me.AdoAux(1).Recordset.RecordCount > 0 Then
        MsgBox "Tiene lineas de expediente asociadas", vbExclamation
        Exit Function
    End If
    If Me.AdoAux(2).Recordset.RecordCount > 0 Then
        MsgBox "Tiene pagos a cuenta ", vbExclamation
        Exit Function
    End If
    
    
'    cad = "Select count(*) from factcli where codclien = " & Data1.Recordset.Fields(0)
'    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    NumRegElim = 0
'    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
'    miRsAux.Close
'
'    If NumRegElim > 0 Then
'        MsgBox "Tiene facturas realizadas", vbExclamation
'        Exit Function
'    End If
'
    
    
    SePuedeEliminar = True
    Screen.MousePointer = vbDefault
End Function

Private Sub ToolbarAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 1
            'AÑADIR linea factura
            BotonAnyadirLinea Index, True
        Case 2
            'MODIFICAR linea factura
            BotonModificarLinea Index
        Case 3
            'ELIMINAR linea factura
            BotonEliminarLinea Index
            

    End Select


End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

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
        Toolbar1.Buttons(2).Enabled = DBLet(RS!Modificar, "N") And (Modo = 2 And Me.Data1.Recordset.RecordCount > 0)
        Toolbar1.Buttons(3).Enabled = DBLet(RS!creareliminar, "N") And (Modo = 2 And Me.Data1.Recordset.RecordCount > 0)
        
        Toolbar1.Buttons(5).Enabled = DBLet(RS!Ver, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(6).Enabled = DBLet(RS!Ver, "N") And (Modo = 0 Or Modo = 2)
        
        Toolbar1.Buttons(8).Enabled = DBLet(RS!Imprimir, "N") And (Modo = 0 Or Modo = 2)
        
        
        For I = 1 To 2
            ToolbarAux(I).Buttons(1).Enabled = DBLet(RS!creareliminar, "N") And (Modo = 2)
            ToolbarAux(I).Buttons(2).Enabled = DBLet(RS!Modificar, "N") And (Modo = 2 And Me.Data1.Recordset.RecordCount > 0)
            ToolbarAux(I).Buttons(3).Enabled = DBLet(RS!creareliminar, "N") And (Modo = 2 And Me.Data1.Recordset.RecordCount > 0)
        Next
        
    Else
        
        Toolbar1.Buttons(1).Enabled = vUsu.Nivel < 2 And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(2).Enabled = vUsu.Nivel < 2 And (Modo = 2 And Me.Data1.Recordset.RecordCount > 0)
        Toolbar1.Buttons(3).Enabled = vUsu.Nivel < 2 And (Modo = 2 And Me.Data1.Recordset.RecordCount > 0)
        
        Toolbar1.Buttons(5).Enabled = vUsu.Nivel < 2 And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(6).Enabled = vUsu.Nivel < 2 And (Modo = 0 Or Modo = 2)
        
        Toolbar1.Buttons(8).Enabled = vUsu.Nivel < 2 And (Modo = 0 Or Modo = 2)
        
        
        For I = 1 To 2
            ToolbarAux(I).Buttons(1).Enabled = vUsu.Nivel < 2 And (Modo = 2)
            ToolbarAux(I).Buttons(2).Enabled = vUsu.Nivel < 2 And (Modo = 2 And Me.Data1.Recordset.RecordCount > 0)
            ToolbarAux(I).Buttons(3).Enabled = vUsu.Nivel < 2 And (Modo = 2 And Me.Data1.Recordset.RecordCount > 0)
        Next
    End If
    
    RS.Close
    Set RS = Nothing
    
End Sub



Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
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
            For jj = 4 To 8
                txtaux(jj).Visible = B
                txtaux(jj).top = alto
            Next jj
            
            txtAux2(5).Visible = B
            txtAux2(5).top = alto
           
            
            chkAux(0).Visible = B
            chkAux(0).top = alto
            
                For I = 0 To 2
                    If I <> 1 Then
                        cmdAux(I).Visible = B
                        cmdAux(I).top = txtaux(5).top
                        cmdAux(I).Height = txtaux(5).Height
                    End If
                Next
           
        Case 2
            For jj = 13 To 16
                txtaux(jj).Visible = B
                txtaux(jj).top = alto
            Next jj
            
            txtAux2(0).Visible = B
            txtAux2(0).top = alto
           
        
            cmdAux(1).Visible = B
            cmdAux(1).top = txtaux(13).top
            cmdAux(1).Height = txtaux(13).Height
    
    End Select
End Sub


Private Sub ModificarLinea()
'Modifica registre en les taules de Llínies
Dim nomframe As String
Dim V As Integer
Dim Cad As String
    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 1: nomframe = "FrameAux1" 'apuntes
        Case 2: nomframe = "FrameAux2" 'apuntes
    End Select
    ' **************************************************************

    If DatosOkLlin(nomframe) Then
        If ModificaDesdeFormulario2(Me, 2, nomframe) Then
        
        
            ModoLineas = 0

           
            V = AdoAux(NumTabMto).Recordset.Fields(4) 'el 4 es el nº de llinia
            CargaGrid NumTabMto, True
            
            
            ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
            
                DataGridAux(NumTabMto).SetFocus
                AdoAux(NumTabMto).Recordset.Find "numlinea =" & V
         
            ' ***********************************************************

            LLamaLineas NumTabMto, 0
        End If
    End If
        
End Sub

Private Function DatosOkLlin(nomframe As String) As Boolean
Dim RS As ADODB.Recordset
Dim Sql As String
Dim B As Boolean
Dim cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte


    DatosOkLlin = True
    
    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False

    B = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not B Then Exit Function
    
    If B And (Modo = 5 And ModoLineas = 1) Then  'insertar
    
    End If
    
    If B And Modo = 5 Then ' tanto si insertamos como si modificamos en lineas
     
       
        
    End If
    
    DatosOkLlin = B

EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function



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
    PrimeraVez = False
    
    Select Case Index
        
        Case 1 'lineas de expediente
            
                'tiporegi,numserie,numexped,anoexped, numlinea,codconce,nomconce,ampliaci,importe,nomsitua,pagado"
                tots = "N||||0|;N||||0|;N||||0|;N||||0|;S|txtaux(4)|T|Lin|605|;S|txtaux(5)|T|Concepto|1055|;S|cmdAux(0)|B|||;S|txtaux(8)|T|Descripcion|3695|;"
                tots = tots & "S|cmdAux(2)|B|||;S|txtaux(6)|T|Ampliacion|4905|;S|txtaux(7)|T|Importe|1725|;"
                tots = tots & "S|txtAux2(5)|T|Situacion|1300|;N||||0|;S|chkAux(0)|CB|Pag|450|;"
  
                
                arregla tots, DataGridAux(Index), Me
                
                DataGridAux(Index).Columns(4).Alignment = dbgLeft
                DataGridAux(Index).Columns(5).Alignment = dbgLeft
                DataGridAux(Index).Columns(6).Alignment = dbgLeft
                 
                B = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
                
                If (Enlaza = True) And (Not AdoAux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
                
                Else
                    For I = 0 To 4
                        txtaux(I).Text = ""
                    Next I
                    txtAux2(5).Text = ""
                   
                End If
        Case 2
                'tiporegi,numserie,numexped,anoexped, numlinea,fecha,codforpa,nomforpa,importe"
                tots = "N||||0|;N||||0|;N||||0|;N||||0|;S|txtaux(13)|T|Lin|585|;S|txtaux(14)|T|Fecha|1255|;S|txtaux(15)|T|Cod.|855|;S|cmdAux(1)|B|||;"
                tots = tots & "S|txtAux2(0)|T|Desc. pago|3305|;S|txtaux(16)|T|Importe|1525|;"
             
                
                arregla tots, DataGridAux(Index), Me
                
'                DataGridAux(Index).Columns(4).Alignment = dbgLeft
'                DataGridAux(Index).Columns(5).Alignment = dbgLeft
'                DataGridAux(Index).Columns(16).Alignment = dbgLeft
'
                B = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
                
                If (Enlaza = True) And (Not AdoAux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
                
                Else
                    For I = 13 To 16
                        txtaux(I).Text = ""
                    Next I
                    txtAux2(0).Text = ""
                   
                End If
    End Select
    DataGridAux(Index).ScrollBars = dbgAutomatic
      
    ' **** si n'hi han llínies en grids i camps fora d'estos ****
    If Not AdoAux(Index).Recordset.EOF Then
        DataGridAux_RowColChange Index, 1, 1
    Else
        
    End If
    ' **********************************************************
      
 
    PonerModoUsuarioGnral Modo, "arigestion"

      
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
End Sub


Private Function MontaSQLCarga(Index As Integer, Enlaza As Boolean) As String
Dim Sql As String
    
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
        Case 1 ' lineas de totales
            Sql = "select tiporegi,numserie,numexped,anoexped, numlinea,codconce,nomconce,ampliaci,importe,nomsitua,pagado,if(pagado=1,'Si','') Chkpagado"
            Sql = Sql & " from expedientes_lineas,tipositexped where expedientes_lineas.codsitua=tipositexped.codsitua"
            
        Case 2
            Sql = "select tiporegi,numserie,numexped,anoexped, numlinea,fechaent,expedientes_acuenta.codforpa,nomforpa,importe"
            Sql = Sql & " from expedientes_acuenta, ariconta1.formapago where expedientes_acuenta.codforpa=formapago.codforpa"
            
    End Select
    If Enlaza Then
        Sql = Sql & " AND tiporegi =" & Data1.Recordset!tiporegi & " AND  numexped  ="
        Sql = Sql & Data1.Recordset!numexped & " AND  anoexped =" & Data1.Recordset!anoexped

    Else
        Sql = Sql & " AND numlinea = -1"
    End If

    MontaSQLCarga = Sql
End Function


Private Sub Totales()
Dim Sql As String
Dim impor As Currency
Dim impR As Currency
    lw1.ListItems.Clear
    Sql = "select conceptos.codigiva, sum(importe) base,porceiva,porcerec"
    Sql = Sql & " from expedientes_lineas ,conceptos,ariconta1.tiposiva iva where expedientes_lineas.codconce= conceptos.codconce"
    Sql = Sql & " AND iva.codigiva=conceptos.codigiva"
    Sql = Sql & " AND tiporegi =" & Data1.Recordset!tiporegi & " AND  numexped  ="
    Sql = Sql & Data1.Recordset!numexped & " AND  anoexped =" & Data1.Recordset!anoexped
    Sql = Sql & " group by 1"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 1
    While Not miRsAux.EOF
        lw1.ListItems.Add , , " "
        lw1.ListItems(I).SubItems(1) = Format(miRsAux!Base, FormatoImporte)
        impor = miRsAux!Base
        impor = Round((miRsAux!Base * miRsAux!porceiva) / 100, 2)
        lw1.ListItems(I).SubItems(2) = Format(miRsAux!porceiva, FormatoImporte)
        lw1.ListItems(I).SubItems(3) = Format(impor, FormatoImporte)
        impR = Round((miRsAux!Base * miRsAux!porcerec) / 100, 2)
        If impR = 0 Then
            Sql = " "
        Else
            Sql = Format(impR, FormatoImporte)
        End If
        lw1.ListItems(I).SubItems(4) = Sql
        lw1.ListItems(I).SubItems(5) = Format(miRsAux!Base + impor + impR, FormatoImporte)
        I = I + 1
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
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
        Case 1: vTabla = "expedientes_lineas"
        Case 2: vTabla = "expedientes_acuenta"
    End Select
    ' ********************************************************

    vWhere = ObtenerWhereCab(False)

    
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
      
            NumF = SugerirCodigoSiguienteStr(vTabla, "numlinea", Replace(vWhere, "expedientes", "expedientes_lineas"))
            ' ***************************************************************

            AnyadirLinea DataGridAux(Index), AdoAux(Index)

            anc = DataGridAux(Index).top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 230 '248
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If

            LLamaLineas Index, ModoLineas, anc

                If Limpia Then
                    For I = 0 To txtaux.Count - 1
                        txtaux(I).Text = ""
                    Next I
                End If
                
                
                
    Select Case Index
         Case 1
                txtaux(0).Text = Text1(2).Text '
                txtaux(1).Text = Text1(1).Text '
                txtaux(2).Text = Text1(3).Text '
                txtaux(3).Text = Text1(0).Text '
                
                txtaux(4).Text = Format(NumF, "0000") 'linea contador
                
                
                If Limpia Then txtAux2(5).Text = ""
                  
                ' antes si hay retencion se marca como que hay que aplicarle retencion
                chkAux(0).Value = 0
            
                ' traemos la cuenta de contrapartida habitual
                PonFoco txtaux(5)

        
    Case 2
            txtaux(9).Text = Text1(2).Text '
            txtaux(10).Text = Text1(1).Text '
            txtaux(11).Text = Text1(3).Text '
            txtaux(12).Text = Text1(0).Text '
            
            txtaux(13).Text = Format(NumF, "0000") 'linea contador
            
            
            If Limpia Then txtAux2(0).Text = ""
             
            PonFoco txtaux(14)

    End Select
End Sub

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " tiporegi =" & Data1.Recordset!tiporegi & " AND  numexped  ="
    vWhere = vWhere & Data1.Recordset!numexped & " AND  anoexped =" & Data1.Recordset!anoexped

    ObtenerWhereCab = vWhere
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


    Select Case Index
        ' *** valor per defecte al modificar dels camps del grid ***
        Case 1 'lineas de facturas
            txtaux(0).Text = DataGridAux(Index).Columns(0).Text
            txtaux(1).Text = DataGridAux(Index).Columns(1).Text
            txtaux(2).Text = DataGridAux(Index).Columns(2).Text
            txtaux(3).Text = DataGridAux(Index).Columns(3).Text
            txtaux(4).Text = DataGridAux(Index).Columns(4).Text
            
            txtaux(5).Text = DataGridAux(Index).Columns(5).Text '
            txtaux(6).Text = DataGridAux(Index).Columns(7).Text '
            txtaux(7).Text = DataGridAux(Index).Columns(8).Text '
            txtaux(8).Text = DataGridAux(Index).Columns(6).Text '
        
            txtAux2(5).Text = DataGridAux(Index).Columns(9).Text '
            If DataGridAux(Index).Columns(10).Text = 1 Then
                chkAux(0).Value = 1 '
            Else
                chkAux(0).Value = 0
            End If
            
        Case 2
            txtaux(9).Text = DataGridAux(Index).Columns(0).Text
            txtaux(10).Text = DataGridAux(Index).Columns(1).Text
            txtaux(11).Text = DataGridAux(Index).Columns(2).Text
            txtaux(12).Text = DataGridAux(Index).Columns(3).Text
            txtaux(13).Text = DataGridAux(Index).Columns(4).Text
            
            txtaux(14).Text = DataGridAux(Index).Columns(5).Text '
            txtaux(15).Text = DataGridAux(Index).Columns(6).Text '
            txtaux(16).Text = DataGridAux(Index).Columns(8).Text '
        
            txtAux2(0).Text = DataGridAux(Index).Columns(7).Text
    End Select

    LLamaLineas Index, ModoLineas, anc
    
    If Index = 1 Then
        PonFoco txtaux(4)
    Else
        PonFoco txtaux(14)
    End If
    ' ***************************************************************************************
End Sub



Private Sub BotonEliminarLinea(Index As Integer)
Dim Sql As String
Dim vWhere As String
Dim Eliminar As Boolean
Dim SqlLog As String

    On Error GoTo Error2
    
    If Modo < 2 Then Exit Sub
    If AdoAux(Index).Recordset.EOF Then Exit Sub
    
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
            Sql = "¿Seguro que desea eliminar la línea del expediente?" & vbCrLf
            Sql = Sql & vbCrLf & "Linea: " & AdoAux(Index).Recordset!numlinea & " - " & AdoAux(Index).Recordset!nomconce & " - " & DBLet(AdoAux(Index).Recordset!Importe, "N")
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM expedientes_lineas "
                Sql = Sql & Replace(vWhere, "expedientes", "expedientes_lineas") & " and numlinea = " & DBLet(AdoAux(Index).Recordset!numlinea, "N")
                
            End If
      Case 2 'linea de asiento
            Sql = "¿Seguro que desea eliminar pago a cuenta?" & vbCrLf
            Sql = Sql & vbCrLf & "Linea: " & AdoAux(Index).Recordset!numlinea & " - " & AdoAux(Index).Recordset!nomforpa & " - " & DBLet(AdoAux(Index).Recordset!Importe, "N")
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM expedientes_acuenta "
                Sql = Sql & Replace(vWhere, "expedientes", "expedientes_acuenta") & " and numlinea = " & DBLet(AdoAux(Index).Recordset!numlinea, "N")
                
            End If
    End Select

    If Eliminar Then
        NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        Conn.Execute Sql
        
        
        
        '**** parte de contabilizacion de la factura
        '--DesBloqueaRegistroForm Me.Text1(0)
        TerminaBloquear
        
        
        
        'LOG
        'vLog.Insertar 8, vUsu, SqlLog
        
        'Creo que no hace falta volver a situar el datagrid
        If True Then
            lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            PonerModo 2
       
        End If
        '**** hasta aqui
        
        
        ' *** si n'hi han tabs sense datagrid, posar l'If ***
        If Index <> 3 Then _
            CargaGrid Index, True
        ' ***************************************************
        If Not SituarDataTrasEliminar(AdoAux(Index), NumRegElim, True) Then
'            PonerCampos
            
        End If
      
      
        
    End If
    
    ModoLineas = 0
    PosicionarData
    
    Exit Sub
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando linea", Err.Description
End Sub


Private Function SepuedeBorrar(Index As Integer) As Boolean
    'Ha pulsado en un tool bar u en otro
    
    'En ppio siempre se borra
    SepuedeBorrar = True
End Function


Private Sub PosicionarData()
Dim Cad As String, Indicador As String

    Cad = ObtenerWhereCab(False)
    
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


Private Sub txtaux_GotFocus(Index As Integer)
    ConseguirFoco txtaux(Index), 3
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim C As String

    txtaux(Index).Text = Trim(txtaux(Index).Text)
    
    Select Case Index
    Case 4
        PonerFormatoEntero txtaux(Index)
        
    Case 5
        C = ""
        If txtaux(Index).Text <> "" Then
            If PonerFormatoEntero(txtaux(Index)) Then
                CadB = "preciocon"
                C = DevuelveDesdeBD("nomconce", "conceptos", "codconce", txtaux(Index).Text, "N", CadB)
                If C = "" Then
                    MsgBox "No existe el concepto " & txtaux(Index).Text, vbExclamation
                Else
                    txtaux(7).Text = Format(CadB, FormatoImporte)
                End If
                CadB = ""
            End If
        End If
        txtaux(8).Text = C
        
    Case 14
        If Not PonerFormatoFecha(txtaux(Index)) Then txtaux(Index).Text = ""
        
    Case 7, 16
        If Not PonerFormatoDecimal(txtaux(Index), 1) Then txtaux(Index).Text = ""
        
    Case 15
        C = ""
        If txtaux(Index).Text <> "" Then
            If PonerFormatoEntero(txtaux(Index)) Then
                C = DevuelveDesdeBD("nomforpa", "ariconta1.formapago", "codforpa", txtaux(Index).Text, "N")
                If C = "" Then MsgBox "No existe la forma de pago " & txtaux(Index).Text, vbExclamation
                
            End If
        End If
        txtAux2(0).Text = C

        
        
    End Select
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
        Case 2: nomframe = "FrameAux2"
    End Select
    ' ***************************************************************
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        
        B = True

        If B And InsertarDesdeForm2(Me, 2, nomframe) Then
        
                DataGridAux(NumTabMto).AllowAddNew = False
                    
                CargaGrid NumTabMto, True
                Limp = True

                    
                ModoLineas = 0
                If B Then BotonAnyadirLinea NumTabMto, True
      
        End If
    End If
End Sub




Private Sub ImprimirExpe()

    If Modo <> 2 Then Exit Sub
    If Data1.Recordset.EOF Then Exit Sub

    cadParam = ""
    numParam = 0
    
    'indRPT = "0106-00"
    
    
    cadFormula = "{expedientes.anoexped}=" & Data1.Recordset!anoexped & " AND {expedientes.numexped}=" & Data1.Recordset!numexped
    cadFormula = cadFormula & " AND {expedientes.tiporegi}='" & Data1.Recordset!tiporegi & "'"
    
    'FALTA
    'If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = "rExpediente.rpt"
    
    ImprimeGeneral

End Sub
