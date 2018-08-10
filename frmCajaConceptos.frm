VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCajaConceptos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Conceptos CAJA"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   14400
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCajaConceptos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   14400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3840
      TabIndex        =   16
      Top             =   30
      Visible         =   0   'False
      Width           =   1575
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   120
         TabIndex        =   17
         Top             =   150
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Tasas"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "sd"
            EndProperty
         EndProperty
      End
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
      TabIndex        =   15
      ToolTipText     =   "Buscar cuenta"
      Top             =   5640
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Left            =   5520
      TabIndex        =   14
      Text            =   "Dat"
      Top             =   5760
      Width           =   2955
   End
   Begin VB.TextBox txtAux 
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
      Left            =   3960
      TabIndex        =   2
      Tag             =   "Cuenta|T|N|||caja_conceptos|codmacta||N|"
      Text            =   "Dat"
      Top             =   5640
      Width           =   800
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   60
      TabIndex        =   10
      Top             =   30
      Width           =   3585
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3750
         TabIndex        =   12
         Top             =   270
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   11
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
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
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
      ItemData        =   "frmCajaConceptos.frx":000C
      Left            =   2280
      List            =   "frmCajaConceptos.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Tag             =   "Tipo concepto|N|N|||caja_conceptos|tipoconcepto|||"
      Top             =   5640
      Width           =   615
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
      Left            =   11880
      TabIndex        =   4
      Top             =   8070
      Visible         =   0   'False
      Width           =   1035
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
      Left            =   13200
      TabIndex        =   5
      Top             =   8070
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFE1FF&
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
      ForeColor       =   &H00000000&
      Height          =   350
      Index           =   1
      Left            =   900
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "Denominación|T|N|||caja_conceptos|nomconceC|||"
      Text            =   "Dato2"
      Top             =   5640
      Width           =   1425
   End
   Begin VB.TextBox txtAux 
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
      Left            =   60
      TabIndex        =   0
      Tag             =   "Código concepto|N|N|0||caja_conceptos|codconceC|000|S|"
      Text            =   "Dat"
      Top             =   5640
      Width           =   800
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmCajaConceptos.frx":0010
      Height          =   6855
      Left            =   60
      TabIndex        =   6
      Top             =   900
      Width           =   14190
      _ExtentX        =   25030
      _ExtentY        =   12091
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   23
      RowDividerStyle =   6
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
      Left            =   13200
      TabIndex        =   9
      Top             =   8070
      Visible         =   0   'False
      Width           =   1035
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
      TabIndex        =   7
      Top             =   7920
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
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   120
         Width           =   2550
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   375
      Left            =   12000
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
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   390
      Left            =   13560
      TabIndex        =   13
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
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
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
Attribute VB_Name = "frmCajaConceptos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)
Public vWhere As String

Private Const IdPrograma = ID_CajaConceptos

Private WithEvents frmCta As frmBasico
Attribute frmCta.VB_VarHelpID = -1
Private CadenaConsulta As String
Private CadB As String

Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas
Dim Modo As Byte

Dim SQL As String

'----------------------------------------------
'----------------------------------------------
'   Deshabilitamos todos los botones menos
'   el de salir
'   Ademas mostramos aceptar y cancelar
'   Modo 0->  Normal
'   Modo 1 -> Lineas BUSCAR
'   Modo 2 -> Recorrer registros
'   Modo 3 -> Lineas  INSERTAR
'   Modo 4 -> Lineas MODIFICAR
'----------------------------------------------
'----------------------------------------------

Private Sub PonerModo(vModo)
Dim B As Boolean
Dim I As Integer

    Modo = vModo

    B = (Modo = 2)
    If B Then
        PonerContRegIndicador
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    B = (Modo = 0 Or Modo = 2)
        
    For I = 0 To txtaux.Count - 1
        txtaux(I).Visible = Not B
    Next
   
    txtAux2(2).Visible = Not B
    
    Combo1.Visible = Not B
    
    cmdAux(0).Visible = Not B And vUsu.Nivel < 1
    
    
    
    For I = 0 To txtaux.Count - 1
        txtaux(I).BackColor = vbWhite
    Next I
    Combo1.BackColor = vbWhite
    'Prueba
    
    
    cmdAceptar.Visible = Not B
    cmdCancelar.Visible = Not B
    DataGrid1.Enabled = B
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.Visible = B
    End If
    txtaux(0).Enabled = (Modo <> 4)
    
    PonerModoUsuarioGnral Modo, "arigestion"
    
End Sub

Private Sub PonerContRegIndicador()
'si estamos en modo ver registros muestra el numero de registro en el que estamos
'situados del total de registros mostrados: 1 de 24
Dim cadReg As String

    If (Modo = 2 Or Modo = 0) Then
        cadReg = PonerContRegistros(Me.Adodc1)
        If CadB = "" Then
            lblIndicador.Caption = cadReg
        Else
            lblIndicador.Caption = "BUSQUEDA: " & cadReg
        End If
    End If
End Sub

Private Sub BotonAnyadir()
    Dim NumF As String
    Dim anc As Single
    
    'Obtenemos la siguiente numero de factura
    NumF = SugerirCodigoSiguiente
    'Situamos el grid al final
    DataGrid1.AllowAddNew = True
    If Adodc1.Recordset.RecordCount > 0 Then
        DataGrid1.HoldFields
        Adodc1.Recordset.MoveLast
        DataGrid1.Row = DataGrid1.Row + 1
    End If
    
    If DataGrid1.Row < 0 Then
        anc = DataGrid1.top + 260
    Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.top
    End If
    
    txtaux(0).Text = NumF
    For I = 1 To 2
        txtaux(I).Text = ""
    Next
    txtAux2(2).Text = ""

    Combo1.ListIndex = 0
    LLamaLineas anc, 3
    
    
    'Ponemos el foco
    PonFoco txtaux(0)

End Sub



Private Sub BotonVerTodos()
    CargaGrid ""
End Sub

Private Sub BotonBuscar()
    CargaGrid "codconce = -1"
    'Buscar
    Limpiar Me
    Combo1.ListIndex = -1
    
    LLamaLineas DataGrid1.top + 250, 1
    PonFoco txtaux(0)
End Sub

Private Sub BotonModificar()
    '---------
    'MODIFICAR
    '----------
    Dim Cad As String
    Dim anc As Single
    Dim I As Integer
    If Adodc1.Recordset.EOF Then Exit Sub
    If Adodc1.Recordset.RecordCount < 1 Then Exit Sub



    Screen.MousePointer = vbHourglass
    Me.lblIndicador.Caption = "MODIFICAR"
    DeseleccionaGrid
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 320
        Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.top
    End If

    'Llamamos al form
    
    txtaux(0).Text = DataGrid1.Columns(0).Text
    txtaux(1).Text = DataGrid1.Columns(1).Text
    I = Adodc1.Recordset!tipoconcepto
    SituarCombo Combo1, I
   
    txtaux(2).Text = DataGrid1.Columns(2).Text
    txtAux2(2).Text = DataGrid1.Columns(3).Text
    
    
    LLamaLineas anc, 4
   
   'Como es modificar
   PonFoco txtaux(1)
   
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
PonerModo xModo
'Fijamos el ancho
For I = 0 To txtaux.Count - 1
    txtaux(I).top = alto
Next

txtAux2(2).top = alto
Combo1.top = alto - 15
cmdAux(0).top = alto - 15
End Sub




Private Sub BotonEliminar()
Dim SQL As String
    On Error GoTo Error2
    'Ciertas comprobaciones
    If Adodc1.Recordset.EOF Then Exit Sub
   
    
    If Not SepuedeBorrar Then Exit Sub
    '### a mano
    SQL = "Seguro que desea eliminar el concepto de caja:"
    SQL = SQL & vbCrLf & "Código: " & Adodc1.Recordset.Fields(0)
    SQL = SQL & vbCrLf & "Denominación: " & Adodc1.Recordset.Fields(1)
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        SQL = "Delete from caja_conceptos where codconcec=" & Adodc1.Recordset!CodConceC
        Conn.Execute SQL
        CargaGrid ""
        Adodc1.Recordset.Cancel
    End If
    Exit Sub
Error2:
        Screen.MousePointer = vbDefault
        MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub


Private Sub adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  If adReason = adRsnMove And adStatus = adStatusOK Then PonLblIndicador Me.lblIndicador, Adodc1
End Sub

Private Function CadWhere() As String
Dim H As Integer
Dim C As String

    C = Me.Adodc1.RecordSource
    H = InStr(1, C, " WHERE ")
    If H > 0 Then
        C = Mid(C, H + 6)
    Else
        C = ""
    End If
    C = Replace(C, "1=1  AND ", "")

    H = InStr(1, C, " ORDER BY ")
    If H > 0 Then
        C = Mid(C, 1, H)
    Else
        C = ""
    End If
    
    
    CadWhere = C
End Function

Private Sub cmdAceptar_Click()
Dim I As Integer
Dim CadB As String

    Select Case Modo
    Case 1
        'HacerBusqueda
        CadB = ObtenerBusqueda(Me)
        If CadB <> "" Then
            PonerModo 0
            CargaGrid CadB
        End If
    Case 3
        If DatosOK Then
            '-----------------------------------------
            'Hacemos insertar
            If InsertarDesdeForm(Me) Then
                'MsgBox "Registro insertado.", vbInformation
                CargaGrid
                BotonAnyadir
            End If
        End If
    Case 4
            'Modificar
            If DatosOK Then
                '-----------------------------------------
                'Hacemos insertar
                If ModificaDesdeFormulario(Me) Then
                
                    If Adodc1.Recordset.Fields(0) = vParam.CajaConceptoTarjetaCredito Then
                        SQL = "Antiguos" & vbCrLf
                        For I = 1 To 2
                            SQL = SQL & DBLet(Adodc1.Recordset.Fields(I), "T") & " // "
                        Next I
                        SQL = SQL & vbCrLf & "Nuevos" & vbCrLf
                        For I = 1 To 2
                            SQL = SQL & txtaux(I).Text & " // "
                        Next I
                        SQL = "Modificar concepto caja tarjeta credito: " & vParam.CajaConceptoTarjetaCredito & vbCrLf & SQL
                        vLog.Insertar 13, vUsu, SQL
                        
                        
                    End If
                    I = Adodc1.Recordset.Fields(0)
                    PonerModo 0
                    CargaGrid CadWhere
                    Adodc1.Recordset.Find (Adodc1.Recordset.Fields(0).Name & " =" & I)
                End If
            End If
    End Select


End Sub

Private Sub cmdAux_Click(Index As Integer)
  Set frmCta = New frmBasico
    SQL = ""
    AyudaCtasContabilidad frmCta
    If SQL <> "" Then
        
        Me.txtaux(2).Text = RecuperaValor(SQL, 1)
        Me.txtAux2(2).Text = RecuperaValor(SQL, 2)
    End If
    Set frmCta = Nothing
End Sub

Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 1
            CargaGrid
        Case 3
            DataGrid1.AllowAddNew = False
            'CargaGrid
            If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveFirst
            
    End Select
    PonerModo 0
    lblIndicador.Caption = ""
    DataGrid1.SetFocus
End Sub

Private Sub cmdRegresar_Click()
Dim Cad As String

    If Adodc1.Recordset.EOF Then
        MsgBox "Ningún registro a devolver.", vbExclamation
        Exit Sub
    End If
    
    
  
    Cad = Adodc1.Recordset.Fields(0) & "|"
    Cad = Cad & Adodc1.Recordset.Fields(1) & "|"
    Cad = Cad & Adodc1.Recordset.Fields(3) & "|"

    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub

Private Sub cmdRegresar_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Combo1_GotFocus()
    If Modo = 1 Then
        Combo1.BackColor = vbLightBlue
    End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Combo1_LostFocus()
    Combo1.BackColor = vbWhite
End Sub



Private Sub DataGrid1_DblClick()
    If cmdRegresar.Visible Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

'++
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then Unload Me
End Sub
'++
Private Sub Form_Load()

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

'
'    With Me.Toolbar2
'        .HotImageList = frmppal.imgListComun_OM
'        .DisabledImageList = frmppal.imgListComun_BN
'        .ImageList = frmppal.ImgListComun
'        .Buttons(2).Image = 31
'    End With
'
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With


    '## A mano
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    CargaCombo
    
    cmdRegresar.Visible = (DatosADevolverBusqueda <> "")
    

    
    PonerModo 0
    CadAncho = False
    PonerOpcionesMenu  'En funcion del usuario
    'Cadena consulta
   
    CadenaConsulta = "Select codconcec,nomconcec,caja_conceptos.codmacta,nommacta "
    CadenaConsulta = CadenaConsulta & ",if(tipoconcepto=0,""General"",if(tipoconcepto>=0,""Genérico"",if(tipoconcepto>100,"""","""")))"
    CadenaConsulta = CadenaConsulta & " as QueTipo,tipoconcepto  FROM "
    CadenaConsulta = CadenaConsulta & " caja_conceptos  LEFT JOIN ariconta" & vParam.Numconta & ".cuentas LaCta on lacta.codmacta=caja_conceptos.codmacta WHERE 1=1"
    
    If vWhere <> "" Then CadenaConsulta = CadenaConsulta & " and " & vWhere
    CargaGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub



Private Sub frmIv_DatoSeleccionado(CadenaSeleccion As String)
    CadB = CadenaSeleccion
End Sub

Private Sub frmCta_DatoSeleccionado(CadenaSeleccion As String)
    SQL = CadenaSeleccion
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
    Dim SQL As String
    Dim Rs As ADODB.Recordset
    
    SQL = "Select Max(codconcec) from caja_conceptos "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, , , adCmdText
    SQL = "1"
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            SQL = CStr(Rs.Fields(0) + 1)
        End If
    End If
    Rs.Close
    SugerirCodigoSiguiente = SQL
End Function

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
                BotonAnyadir
        Case 2
                BotonModificar
        Case 3
                BotonEliminar
        Case 5
                BotonBuscar
        Case 6
                BotonVerTodos
        Case 8
                'frmConceptosList.Show vbModal
                If Modo < 3 Then frmConceptosList.Show vbModal
        Case Else
    End Select
End Sub



Private Sub CargaGrid(Optional SQL As String)
    Dim J As Integer
    Dim TotalAncho As Integer
    Dim I As Integer
    
    Adodc1.ConnectionString = Conn
    If SQL <> "" Then
        SQL = CadenaConsulta & " AND " & SQL
        Else
        SQL = CadenaConsulta
    End If
    SQL = SQL & " ORDER BY codconcec"
    Adodc1.RecordSource = SQL
    Adodc1.CursorType = adOpenDynamic
    Adodc1.LockType = adLockOptimistic
    Adodc1.Refresh
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 350 '290
    
    ' codconcec,nomconcec,caja_conceptos.codmacta,nommacta "
    ',if(tipoconcepto=0,""General"",if(tipoconcepto>=0,""Genérico"",if(tipoconcepto>100,"""","""")))  QueTipo,tipoconcepto  FROM "
    
    
    
    'Nombre producto
    I = 0
        DataGrid1.Columns(I).Caption = "Cod."
        DataGrid1.Columns(I).Width = 600
        DataGrid1.Columns(I).NumberFormat = "000"
    I = 1
        DataGrid1.Columns(I).Caption = "Descripcion"
        DataGrid1.Columns(I).Width = 5500
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
                
    I = 2
        DataGrid1.Columns(I).Caption = "Cuenta"
        DataGrid1.Columns(I).Width = 1750
        
    I = 3
        DataGrid1.Columns(I).Caption = "Nom. cuenta"
        DataGrid1.Columns(I).Width = 4000
    I = 4
        DataGrid1.Columns(I).Caption = "Tipo"
        DataGrid1.Columns(I).Width = 1200
        
    DataGrid1.Columns(5).Visible = False
    
        'Fiajamos el cadancho
    If Not CadAncho Then
        'La primera vez fijamos el ancho y alto de  los txtaux
        txtaux(0).Left = DataGrid1.Left + 340
        txtaux(0).Width = DataGrid1.Columns(0).Width - 60
        txtaux(1).Left = txtaux(0).Left + txtaux(0).Width + 45
        txtaux(1).Width = DataGrid1.Columns(1).Width - 60
        txtaux(2).Left = txtaux(1).Left + txtaux(1).Width + 45
        txtaux(2).Width = DataGrid1.Columns(2).Width - 60
        txtAux2(2).Left = txtaux(2).Left + txtaux(2).Width + 45
        txtAux2(2).Width = DataGrid1.Columns(3).Width - 60
        Combo1.Left = txtAux2(2).Left + txtAux2(2).Width + 45
        Combo1.Width = DataGrid1.Columns(4).Width
        
        

      
        
        cmdAux(0).Left = txtAux2(2).Left - 180
        cmdAux(0).Height = txtAux2(2).Height
        CadAncho = True
    End If
    'Habilitamos modificar y eliminar
    If vUsu.Nivel < 2 Then
        Toolbar1.Buttons(7).Enabled = Not Adodc1.Recordset.EOF
        Toolbar1.Buttons(8).Enabled = Not Adodc1.Recordset.EOF
    End If
    DataGrid1.AllowRowSizing = False
End Sub



Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    'Solo hay una opcion
    'Button.Index
    If Modo = 1 Or Modo > 2 Then Exit Sub
    If Adodc1.Recordset.EOF Then Exit Sub
    
    
    If Val(Adodc1.Recordset!tipoconcepto) <> 0 Then
        MsgBox "Este tipo de cuota no lleva control de movimientos", vbExclamation
        Exit Sub
    End If
    
    frmGestionTasasMov.Concepto = Adodc1.Recordset!codconce
    frmGestionTasasMov.Show vbModal
    
    'Refrescamos y situamos
    
    I = Adodc1.Recordset.Fields(0)
    PonerModo 0
    CargaGrid CadWhere
    Adodc1.Recordset.Find (Adodc1.Recordset.Fields(0).Name & " =" & I)
End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    ConseguirFoco txtaux(Index), Modo
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim C As String

    If Not PerderFocoGnral(txtaux(Index), Modo) Then Exit Sub

    txtaux(Index).Text = Trim(txtaux(Index).Text)
    If txtaux(Index).Text = "" Then Exit Sub
    If Modo = 1 Then Exit Sub 'Busquedas
    Select Case Index
    Case 0
        If Not IsNumeric(txtaux(0).Text) Then
            MsgBox "Código concepto tiene que ser numérico", vbExclamation
            Exit Sub
        End If
        txtaux(0).Text = Format(txtaux(0).Text, "000")
    Case 1
    
            
    Case 2
            
    
        CuentaCorrectaUltimoNivelTextBox txtaux(Index), txtAux2(Index)
        
    End Select
End Sub


Private Function DatosOK() As Boolean
Dim Datos As String
Dim B As Boolean
txtaux(1).Text = UCase(txtaux(1).Text)
B = CompForm(Me)
If Not B Then Exit Function

If Modo = 3 Then
    'Estamos insertando
     Datos = DevuelveDesdeBD("nomconceC", "caja_conceptos", "codconcec", txtaux(0).Text, "N")
     If Datos <> "" Then
        MsgBox "Ya existe el concepto : " & txtaux(0).Text, vbExclamation
        B = False
    End If
End If
DatosOK = B
End Function

Private Sub CargaCombo()
    Combo1.Clear

    Combo1.AddItem "General"
    Combo1.ItemData(Combo1.NewIndex) = 0
    

End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub



Private Function SepuedeBorrar() As Boolean
Dim SQL As String
    SepuedeBorrar = False
    
    If Adodc1.Recordset!CodConceC = 0 Then
        MsgBox "No se puede eliminar el concepto 0", vbExclamation
        Exit Function
    End If
    
    If Adodc1.Recordset!CodConceC = vParam.CajaConceptoTarjetaCredito Then
        MsgBox "No se puede eliminar el concepto de tarjeta de credito ", vbExclamation
        Exit Function
    End If
    
    
    
'    Msg = "expedientes_lineas|clientes_fiscal|clientes_laboral|clientes_cuotas|factcli_lineas|"
'    MsgErr = ""
'    For I = 1 To 5
'        SQL = DevuelveDesdeBD("count(*)", RecuperaValor(Msg, CInt(I)), "codconce", CStr(Me.Adodc1.Recordset!codconce))
'        If Val(SQL) > 0 Then
'            MsgErr = MsgErr & "- " & RecuperaValor("Expedientes|Fiscal|Laboral|Cuotas|Facturas|", CInt(I)) & " (" & SQL & ")" & vbCrLf
'        End If
'    Next
'
    Msg = ""
    If MsgErr <> "" Then
        MsgBox "Concepto en: " & vbCrLf & vbCrLf & MsgErr, vbExclamation
    Else
        SepuedeBorrar = True
    End If
End Function


Private Sub DeseleccionaGrid()
    On Error GoTo EDeseleccionaGrid
        
    While DataGrid1.SelBookmarks.Count > 0
        DataGrid1.SelBookmarks.Remove 0
    Wend
    Exit Sub
EDeseleccionaGrid:
        Err.Clear
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModoUsuarioGnral(Modo As Byte, aplicacion As String)
Dim Rs As ADODB.Recordset
Dim Cad As String
    
    On Error Resume Next

    Cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(aplicacion, "T")
    Cad = Cad & " and codigo = " & DBSet(IdPrograma, "N") & " and codusu = " & DBSet(vUsu.id, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Toolbar1.Buttons(1).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(2).Enabled = DBLet(Rs!Modificar, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(3).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 0 Or Modo = 2)
        
        Toolbar1.Buttons(5).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(6).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2)
        
        Toolbar1.Buttons(8).Enabled = DBLet(Rs!Imprimir, "N") And (Modo = 0 Or Modo = 2)
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub



