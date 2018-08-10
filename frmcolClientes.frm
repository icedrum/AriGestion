VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#17.2#0"; "Codejock.Controls.v17.2.0.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#17.2#0"; "Codejock.ReportControl.v17.2.0.ocx"
Begin VB.Form frmcolClientes 
   Caption         =   "Clientes"
   ClientHeight    =   9555
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14580
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9555
   ScaleWidth      =   14580
   StartUpPosition =   3  'Windows Default
   Begin XtremeReportControl.ReportControl wndReportControl 
      Height          =   6735
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   14295
      _Version        =   1114114
      _ExtentX        =   25215
      _ExtentY        =   11880
      _StockProps     =   64
      MultipleSelection=   0   'False
      FreezeColumnsAbs=   0   'False
      MultiSelectionMode=   -1  'True
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   14055
      Begin XtremeSuiteControls.FlatEdit txtSearchBar 
         Height          =   420
         Index           =   1
         Left            =   1320
         TabIndex        =   1
         Tag             =   "nomclien"
         Top             =   150
         Width           =   3300
         _Version        =   1114114
         _ExtentX        =   5821
         _ExtentY        =   741
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Appearance      =   12
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSearchBar 
         Height          =   420
         Index           =   2
         Left            =   4800
         TabIndex        =   2
         Tag             =   "nifclien"
         Top             =   150
         Width           =   1980
         _Version        =   1114114
         _ExtentX        =   3492
         _ExtentY        =   741
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Appearance      =   12
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton cmdSearch 
         Height          =   300
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "F1 - Lanzar busqueda"
         Top             =   180
         Width           =   375
         _Version        =   1114114
         _ExtentX        =   661
         _ExtentY        =   529
         _StockProps     =   79
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   6
         BuddyControl    =   "txtSearchBar"
      End
      Begin XtremeSuiteControls.FlatEdit txtSearchBar 
         Height          =   420
         Index           =   3
         Left            =   7080
         TabIndex        =   3
         Tag             =   "matricula"
         Top             =   150
         Width           =   1980
         _Version        =   1114114
         _ExtentX        =   3492
         _ExtentY        =   741
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Appearance      =   12
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSearchBar 
         Height          =   420
         Index           =   4
         Left            =   10440
         TabIndex        =   4
         Tag             =   "licencia"
         Top             =   150
         Width           =   1980
         _Version        =   1114114
         _ExtentX        =   3492
         _ExtentY        =   741
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Appearance      =   12
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSearchBar 
         Height          =   420
         Index           =   0
         Left            =   480
         TabIndex        =   0
         Tag             =   "codclien"
         Top             =   150
         Width           =   540
         _Version        =   1114114
         _ExtentX        =   952
         _ExtentY        =   741
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Appearance      =   12
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSearchBar 
         Height          =   420
         Index           =   5
         Left            =   12720
         TabIndex        =   13
         Tag             =   "telefono"
         Top             =   150
         Width           =   1980
         _Version        =   1114114
         _ExtentX        =   3492
         _ExtentY        =   741
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Appearance      =   12
         UseVisualStyle  =   0   'False
      End
   End
   Begin VB.Frame FrameAux 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   14175
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
         TabIndex        =   9
         Top             =   120
         Width           =   3585
         Begin MSComctlLib.Toolbar Toolbar1 
            Height          =   330
            Left            =   240
            TabIndex        =   10
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
                  Object.ToolTipText     =   "Impresión avanzada"
               EndProperty
               BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
                  Style           =   3
               EndProperty
            EndProperty
         End
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   3720
         TabIndex        =   12
         Top             =   240
         Width           =   3375
      End
   End
   Begin MSComctlLib.StatusBar statusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   9180
      Width           =   14580
      _ExtentX        =   25718
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmcolClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim PrimeraVez As Boolean



'Dim iconArray(0 To 9) As Long
'Dim RowExpanded(0 To 49) As Boolean
'Dim RowVisible(0 To 49) As Boolean
Dim MaxRowIndex As Long
Dim fntBold As StdFont
Dim fntStrike As StdFont


Dim Clientes As String






Private Sub VerTodos()
    For I = 0 To 5
        txtSearchBar(I).Text = ""
    Next
    CargaDatos "", False

End Sub




Private Sub PonerModoUsuarioGnral(Modo As Byte, aplicacion As String)
Dim Rs As ADODB.Recordset
Dim Cad As String
    
    On Error Resume Next

    Cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(aplicacion, "T")
    Cad = Cad & " and codigo = " & ID_Clientes & " and codusu = " & DBSet(vUsu.id, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Toolbar1.Buttons(1).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(2).Enabled = DBLet(Rs!Modificar, "N") And (Modo = 2)
        Toolbar1.Buttons(3).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 2)
        
        Toolbar1.Buttons(5).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(6).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2)
        
        Toolbar1.Buttons(8).Enabled = DBLet(Rs!Imprimir, "N") And (Modo = 0 Or Modo = 2)
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub





Private Sub cmdRegresar_Click()
        
End Sub

Private Sub cmdSearch_Click()
Dim Cad1 As String
Dim cad2 As String

    cad2 = ""
    J = 0
    For I = 0 To 5
        Me.txtSearchBar(I).Text = Trim(Me.txtSearchBar(I).Text)
        If Me.txtSearchBar(I).Text <> "" Then
            If SeparaCampoBusqueda(IIf(I = 0, "N", "T"), txtSearchBar(I).Tag, txtSearchBar(I).Text, Cad1) = 0 Then
                If J > 0 Then cad2 = cad2 & " AND  "
                J = J + 1
                cad2 = cad2 & Cad1
            End If
        
        End If
    Next
    
    
        
    If J = 0 Then
        txtSearchBar(0).SetFocus
        Exit Sub
    End If
        
    
    
    
    
    CargaDatos cad2, False
    
    
    On Error Resume Next
     If wndReportControl.SelectedRows.Count > 0 Then
        wndReportControl.SetFocus
    Else
         For I = 0 To 5
            If txtSearchBar(I).Text <> "" Then
                txtSearchBar(I).SetFocus
                Exit For
            End If
        Next
    End If
    
    Err.Clear
End Sub



Private Sub Form_Activate()
    Dim Record As ReportRecord

    If PrimeraVez Then
        PrimeraVez = False
        PonerModoUsuarioGnral 2, "arigestion"
        Me.Refresh
        DoEvents
        CargaDatos "", False
        wndReportControl.Populate
         HacerPrimeravez
    End If
End Sub


Private Sub HacerPrimeravez()
    On Error Resume Next
    txtSearchBar(4).SetFocus
    Err.Clear
End Sub


Private Sub Form_Load()
    PrimeraVez = True
    Me.Icon = frmppal.Icon
    
    
    wndReportControl.Icons = ReportControlGlobalSettings.Icons
    wndReportControl.PaintManager.NoItemsText = "Ningún registro "
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
        .Buttons(5).Visible = False
        .Buttons(8).Image = 16
        .Buttons(9).Image = 32
        .Buttons(9).Enabled = True
    End With
    
  
    
    EstablecerFuente
    
    
    CreateReportControl
    Me.Frame1.BorderStyle = 0
    
    'Buscar
    cmdSearch.BuddyControl = "Buscar"
    Set cmdSearch.Icon = frmppal.CommandBars.Icons.GetImage(ID_SEARCH_ICON, 16)

End Sub


Private Sub EstablecerFuente()

    On Error GoTo eEstablecerFuente
    'The following illustrate how to change the different fonts used in the ReportControl
    Dim TextFont As StdFont
    Set TextFont = Me.Font
    TextFont.SIZE = 11
    Set wndReportControl.PaintManager.TextFont = TextFont
    Set wndReportControl.PaintManager.CaptionFont = TextFont
    Set wndReportControl.PaintManager.PreviewTextFont = TextFont
    
    'This font will be used in the BeforeDrawRow when automatic formatting is selected
    'This simply applies Strikethrough to the currently set text font
    'Set fntStrike = wndReportControl.PaintManager.TextFont
    'fntStrike.Strikethrough = True
    
    'This font will be used in the BeforeDrawRow when automatic formatting is selected
    'This simply applies Bold to the currently set text font
    'Set fntBold = wndReportControl.PaintManager.TextFont
    'fntBold.Bold = True


    Exit Sub
eEstablecerFuente:
    MuestraError Err.Number, Err.Description

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    FrameAux.Move 0, 0, ScaleWidth, FrameAux.Height
    Frame1.Move 0, FrameAux.Height + 60, ScaleWidth, Frame1.Height
    wndReportControl.Move 0, Frame1.Height + Frame1.top, ScaleWidth, ScaleHeight - statusBar.Height - Frame1.Height - Frame1.top
    SituaBusquedas
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub SituaBusquedas()

    On Error Resume Next
    J = 1100
    For I = 0 To 5
        Me.txtSearchBar(I).Left = J + (6 * I)
        k = (Me.wndReportControl.Columns(I + 4).Width * 15) - 30
        txtSearchBar(I).Width = k - 60
        J = J + k
        
        'Me.txtSearchBar(I).Text = I
    Next
    Err.Clear
End Sub





























Public Sub CreateReportControl()
    'Start adding columns
    Dim Column As ReportColumn
    wndReportControl.Columns.DeleteAll
    'Adds a new ReportColumn to the ReportControl's collection of columns, growing the collection by 1.
    Set Column = wndReportControl.Columns.Add(COLUMN_IMPORTANCE, "Socio", 18, False)
    Column.Icon = COLUMN_IMPORTANCE_ICON
    Set Column = wndReportControl.Columns.Add(COLUMN_ICON, "Cuotas", 18, False)
    Column.Icon = COLUMN_MAIL_ICON
    Set Column = wndReportControl.Columns.Add(COLUMN_ATTACHMENT, "Laboral", 18, False)
    Column.Icon = COLUMN_ATTACHMENT_ICON
    Set Column = wndReportControl.Columns.Add(3, "Fiscal", 18, False)
    Column.Icon = COLUMN_ATTACHMENT_ICON
    
    
    Set Column = wndReportControl.Columns.Add(4, "ID", 30, True)
    Column.Alignment = xtpAlignmentRight
    
    Set Column = wndReportControl.Columns.Add(5, "Nombre", 200, True)
    Set Column = wndReportControl.Columns.Add(6, "DNI", 60, True)
    Set Column = wndReportControl.Columns.Add(7, "Matricula", 55, True)
    Set Column = wndReportControl.Columns.Add(8, "Licencia", 55, True)
    Set Column = wndReportControl.Columns.Add(9, "Telefono", 55, True)
    

    wndReportControl.PaintManager.MaxPreviewLines = 1
    wndReportControl.PaintManager.HorizontalGridStyle = xtpGridNoLines
                  
    
    'This font will be used in the BeforeDrawRow when automatic formatting is selected
    'This simply applies Strikethrough to the currently set text font
    Set fntStrike = wndReportControl.PaintManager.TextFont
    fntStrike.Strikethrough = True
    
    'This font will be used in the BeforeDrawRow when automatic formatting is selected
    'This simply applies Bold to the currently set text font
    Set fntBold = wndReportControl.PaintManager.TextFont
    fntBold.Bold = True
    
    'Any time you add or delete rows(by removing the attached record), you must call the
    'Populate method so the ReportControl will display the changes.
    'If rows are added, the rows will remain hidden until Populate is called.
    'If rows are deleted, the rows will remain visible until Populate is called.
    wndReportControl.Populate
    
    wndReportControl.SetCustomDraw xtpCustomBeforeDrawRow
End Sub






Private Sub Label(Visible As Boolean)
    If Visible Then
        Label1.Caption = "Leyendo registros BD"
    Else
        Label1.Caption = ""
    End If
    Label1.Refresh
End Sub



'Cuando modifiquemos o insertemos, pondremos el SQL entero
Public Sub CargaDatos(ByVal Sql As String, EsTodoSQL As Boolean)
Dim Aux  As String
Dim Inicial As Integer
Dim N As Integer
Dim V As Boolean
Dim T1 As Single



    On Error GoTo eCargaDatos

    Screen.MousePointer = vbHourglass
    statusBar.Panels(1).Text = "Leyendo BD"
    V = True
    Label V
    wndReportControl.ShowItemsInGroups = False
    wndReportControl.Records.DeleteAll
    wndReportControl.Populate
    
    Set miRsAux = New ADODB.Recordset
    
    If EsTodoSQL Then
        Stop
    Else
        If Sql <> "" Then Sql = " WHERE " & Sql
            
        Sql = " FROM clientes" & Sql
        Sql = "SELECT codclien,nomclien,nifclien,matricula,licencia,essocio,situclien,telefono " & Sql
        
        Sql = Sql & " ORDER BY codclien"
    End If
    
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Inicial = 0
    Clientes = ""
    T1 = Timer
    While Not miRsAux.EOF
        AddRecord2
        Clientes = Clientes & ", " & miRsAux!CodClien
        N = N + 1
        If N > 40 Then
            Sql = "1"
            
            If Timer - T1 > 0.75 Then
                V = Not V
                Label V
                T1 = Timer
            End If
            wndReportControl.Populate
            
            Label1.Caption = "Carga iconos"
            Label1.Refresh
        
            
            PonerIconosRs Inicial, Me.wndReportControl.Rows.Count - 1
                    
            
            Sql = "2"
            'Haremos ahora el poplate
            wndReportControl.Populate
            
            'Movemos variables
            Sql = "3  " & Inicial + N
            Inicial = Inicial + N - 1
            Clientes = ""
            N = 0
            
            Sql = "4  " & Inicial
            
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
        
    Sql = "N>0"
    If N > 0 Then
        wndReportControl.Populate
        
        
        Sql = "Iconos last"
        PonerIconosRs Inicial, Me.wndReportControl.Rows.Count - 1
    
        'Haremos ahora el poplate
        wndReportControl.Populate
        
        'Movemos variables
        Inicial = Inicial + N - 1
        Clientes = ""
        N = 0
    End If
    
    
    Sql = "Tool butt"
    'Me.Toolbar1.Buttons(2).Enabled = wndReportControl.Records.Count > 0
    'Me.Toolbar1.Buttons(3).Enabled = wndReportControl.Records.Count > 0
    
    
    
eCargaDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description, Sql
    
    
    
    statusBar.Panels(1).Text = ""
    Label1.Caption = ""
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerIconosRs(Inicial As Integer, Final As Integer)
Dim RN As ADODB.Recordset
Dim Cad As String
Dim C As Integer
Dim I As Integer

    On Error GoTo ePonerIconosRs

    Clientes = Mid(Clientes, 2)
    Set RN = New ADODB.Recordset
    For I = 1 To 3
        Cad = IIf(I = 1, "clientes_cuotas", IIf(I = 2, "clientes_laboral", "clientes_fiscal"))
        Cad = "Select distinct(codclien) from " & Cad & " WHERE codclien IN (" & Clientes & ")"
        RN.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If RN.EOF Then
            If Inicial = Final Then wndReportControl.Rows(Inicial).Record(I).Icon = -1
        Else
            While Not RN.EOF
                For C = Inicial To Final
                    If Val(wndReportControl.Rows(C).Record(4).Value) = RN.Fields(0) Then
                        wndReportControl.Rows(C).Record(I).Icon = COLUMN_ATTACHMENT_ICON
                        wndReportControl.Rows(C).Record(I).SortPriority = 1
                        Exit For
                    End If
                Next
                RN.MoveNext
            Wend
        End If
        RN.Close
    Next
    
ePonerIconosRs:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set RN = Nothing
End Sub


'socio, pendiente , nombre, matricula,licencia
'Leera los datos de mirsaux
Private Sub AddRecord2()

Dim Record As ReportRecord
Dim Socio As Boolean
Dim OtroIcono As Boolean
Dim NoActivo As Boolean

    On Error GoTo eAddRecord2

    'Adds a new Record to the ReportControl's collection of records, this record will
    'automatically be attached to a row and displayed with the Populate method
    Set Record = wndReportControl.Records.Add()
    
    Dim Item As ReportRecordItem
    'Socio
    Set Item = Record.AddItem("")
    Socio = miRsAux!esSocio
    NoActivo = DBLet(miRsAux!situclien, "N")
    
    Item.SortPriority = IIf(Socio, 1, 0)

    
    If NoActivo Then
        Item.Icon = 7
    Else
        Item.Icon = IIf(Socio, RECORD_UNREAD_MAIL_ICON, -1)
        
    End If
    'Cuota
    Set Item = Record.AddItem("")
    OtroIcono = False
    Item.SortPriority = 0
    Item.Icon = -1
    
    'Laboral
    Set Item = Record.AddItem("")
    OtroIcono = False
    Item.SortPriority = 0
    Item.Icon = -1
    'Fiscal
    Set Item = Record.AddItem("")
    OtroIcono = False
    Item.SortPriority = 0
    Item.Icon = -1
    
    
    
    ' '  codclien,nomclien,nifclien,matricula,licencia,essocio "
    Set Item = Record.AddItem(CStr(miRsAux!CodClien))
    Item.Value = CLng(miRsAux!CodClien)
    Set Item = Record.AddItem(DBLet(miRsAux!NomClien, "T"))
    If NoActivo Then Item.ForeColor = vbRed
    
    Record.AddItem CStr(miRsAux!NIFClien)
    Record.AddItem CStr(DBLet(miRsAux!Matricula, "T"))
    
    Set Item = Record.AddItem(DBLet(miRsAux!licencia, "T"))
    If Val(DBLet(miRsAux!licencia, "N")) > 0 Then Item.Value = CLng(DBLet(miRsAux!licencia, "N"))
    
    
    Record.AddItem DBLet(miRsAux!telefono, "T") & " "
    
    
    'Adds the PreviewText to the Record.  PreviewText is the text displayed for the ReportRecord while in PreviewMode
    Record.PreviewText = "ID: " & miRsAux!CodClien
    
    
    
    
eAddRecord2:
    
End Sub









'Subroutine that randomly generates a date.  If you group by a column with a date, the records will
'be grouped by how recent the date is in respect to the current date
Public Sub GetDate(ByVal Item As ReportRecordItem, Optional Week = -1, Optional Day = -1, Optional Month = -1, Optional Year = -1, _
                        Optional Hour = -1, Optional Minute = -1)
    Dim WeekDay As String
    Dim TimeOfDay As String
    
    'Initialize random number generator
    Randomize
        
    'Random number fomula
    'Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
    
    'If no week day was provided, randomly select a week day
    If (Week = -1) Then
        Week = Int((7 * Rnd) + 1)
    End If
    
    'Determine the week text
    Select Case Week
        Case 1:
            WeekDay = "Sun"
        Case 2:
            WeekDay = "Mon"
        Case 3:
            WeekDay = "Tue"
        Case 4:
            WeekDay = "Wed"
        Case 5:
            WeekDay = "Thu"
        Case 6:
            WeekDay = "Fri"
        Case 7:
            WeekDay = "Sat"
    End Select
    
    'If no month was provided, randomly select a month
    If (Month = -1) Then
        Month = Int((DatePart("m", Now) - 1 + 1) * Rnd + 1)
    End If
     
    'If no day was provided, randomly select a day
    If (Day = -1) Then
        Day = Int((31 - 1 + 1) * Rnd + 1)
    End If
    
    'If no year was provided, randomly select a year
    If (Year = -1) Then
        Year = Int((2004 - 2003 + 1) * Rnd + 2003)
    End If
    
    'If no hour was provided, randomly select a hour
    If (Hour = -1) Then
        Hour = Int((12 - 1 + 1) * Rnd + 1)
    End If
    
    'If no minute was provided, randomly select a minute
    If (Minute = -1) Then
        Minute = Int((60 - 10 + 1) * Rnd + 10)
    End If
     
    'Randomly select AM or PM
    If (Int(2 * Rnd + 1) = 1) Then
        TimeOfDay = "PM"
    Else
        TimeOfDay = "AM"
    End If
       
    'This block of code determines the GroupPriority, GroupCaption, and SortPriority of the Item
    'based on the date or generated provided.  If the date is the current date, then this Item will
    'have High group and sort priority, and will be given the "Date: Today" groupcaption.
    If (Month = DatePart("m", Now)) And (Day = DatePart("d", Now)) And (Year = DatePart("yyyy", Now)) Then
        Item.GroupCaption = "Date: Today"
        Item.GroupPriority = 0
        Item.SortPriority = 0
    ElseIf (Month = DatePart("m", Now)) And (Year = DatePart("yyyy", Now)) Then
        Item.GroupCaption = "Date: This Month"
        Item.GroupPriority = 1
        Item.SortPriority = 1
    ElseIf (Year = DatePart("yyyy", Now)) Then
        Item.GroupCaption = "Date: This Year"
        Item.GroupPriority = 2
        Item.SortPriority = 2
    Else
        Item.GroupCaption = "Date: Older"
        Item.GroupPriority = 3
        Item.SortPriority = 3
    End If
    
    'Assign the DateTime string to the value of the ReportRecordItem
    Item.Value = WeekDay & " " & Month & "/" & Day & "/" & Year & " " & Hour & ":" & Minute & " " & TimeOfDay
End Sub


Private Sub Form_Unload(Cancel As Integer)
    '
  
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)



    Select Case Button.Index
    Case 1
        'nuevo
        AbrirCliente Nothing
    Case 2
        'modificar
        If Me.wndReportControl.SelectedRows(0) Is Nothing Then Exit Sub
        AbrirCliente Me.wndReportControl.SelectedRows(0)
    Case 3
        'eliminar
        
        If Me.wndReportControl.SelectedRows(0) Is Nothing Then Exit Sub
        NumRegElim = CLng(wndReportControl.SelectedRows(0).Record(4).Caption)
        If PuedeBorrarCliente(NumRegElim, True) Then
        
            Msg = " Codigo: " & wndReportControl.SelectedRows(0).Record(4).Caption & vbCrLf & "Nombre: "
            Msg = Msg & wndReportControl.SelectedRows(0).Record(5).Caption & vbCrLf
            Msg = "Va a eliminar el cliente:" & vbCrLf & Msg & vbCrLf & "¿Continuar?"
            If MsgBox(Msg, vbQuestion + vbYesNoCancel) = vbYes Then
                 If EliminarCliente(NumRegElim) Then
                    wndReportControl.RemoveRowEx wndReportControl.SelectedRows(0)
                    wndReportControl.Populate
                End If
            End If
            Msg = ""
         End If
    Case 6
        'ver todos
        For I = 0 To 4
            txtSearchBar(I).Text = ""
        Next
        CargaDatos "", False
    Case 100
        'Aqui no entra. Es como imprmiriamos el wndrptcontrol
            wndReportControl.PrintPreviewOptions.Title = "Clientes"
            wndReportControl.PrintOptions.Header.Font.Bold = False
            wndReportControl.PrintOptions.Header.Font.SIZE = 24
            wndReportControl.PrintOptions.Header.Font.Name = "Verdana"
            wndReportControl.PrintOptions.Header.TextLeft = "Gremial"
            wndReportControl.PrintOptions.Footer.Font.SIZE = 8
            wndReportControl.PrintOptions.Footer.FormatString = "ReportControl Demo" & vbCrLf & _
                            "&bDemo data not filtered" & _
                            "&bPage &p of &P"
            
            wndReportControl.PrintPreview True
    
    
    Case 8
        frmClientesList.Show vbModal
    End Select
End Sub

Private Sub txtSearchBar_GotFocus(Index As Integer)
  '  ConseguirFoco txtSearchBar, 3
    txtSearchBar(Index).SelStart = 0
    txtSearchBar(Index).SelLength = Len(txtSearchBar(Index).Text)
End Sub

Private Sub txtSearchBar_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        cmdSearch_Click
    ElseIf KeyCode = vbKeyF2 Then
        VerTodos
    Else
        If Shift = 4 Then
            If KeyCode = vbKeyA Then
                cmdSearch_Click
            Else
                If KeyCode = vbKeyV Then VerTodos
            End If
        End If
    End If

End Sub

Private Sub txtSearchBar_KeyPress(Index As Integer, KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me

End Sub




Public Sub SetColor(id As Integer)

    Set wndReportControl.Icons = ReportControlGlobalSettings.Icons
    
    wndReportControl.ToolTipContext.Style = frmppal.CommandBars.ToolTipContext.Style
    Dim HexColor As Long
    If id = ID_OPTIONS_STYLEBLACK2010 Then
        'HexColor = &H393839
        HexColor = &H949294
    ElseIf id = ID_OPTIONS_STYLESILVER2010 Then
        'HexColor = &H73716B
        HexColor = &HBDB2AD
    Else
        HexColor = &HBD9E84
    End If
    
   ' FrameBorder.BackColor = HexColor
   ' FrameReportTop.BackColor = frmShortcutBar.wndShortcutBar.PaintManager.PaneBackgroundColor
End Sub





Private Sub wndReportControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        
        If wndReportControl.SelectedRows.Count > 0 Then
            AbrirCliente wndReportControl.SelectedRows(0)
        End If
    End If
End Sub

Private Sub wndReportControl_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row Is Nothing Then Exit Sub
  
        AbrirCliente Row

    
End Sub


Private Sub AbrirCliente(ByVal Row As XtremeReportControl.IReportRow)
Dim Leer As Boolean
Dim Sql As String
    
    Screen.MousePointer = vbHourglass
    If Row Is Nothing Then
        CadenaDesdeOtroForm = ""
        frmCliente.IdCliente = -1
        frmCliente.Show vbModal
   
    Else
        frmCliente.IdCliente = Val(Row.Record(4).Caption)
        frmCliente.Show vbModal
    End If
    
    Leer = True
    If Row Is Nothing Then
        If CadenaDesdeOtroForm = "" Then Leer = False
    Else
        CadenaDesdeOtroForm = Row.Record(4).Caption
    End If

    If Not Leer Then Exit Sub
    
    Set miRsAux = New ADODB.Recordset
    Screen.MousePointer = vbHourglass
    Sql = "SELECT codclien,nomclien,nifclien,matricula,licencia,essocio "
    Sql = Sql & " FROM clientes  where codclien=" & CadenaDesdeOtroForm
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        
        If Row Is Nothing Then
            AddRecord2
             Clientes = ", " & CadenaDesdeOtroForm
            PonerIconosRs Me.wndReportControl.Rows.Count - 1, Me.wndReportControl.Rows.Count - 1
            wndReportControl.Rows(Me.wndReportControl.Rows.Count - 1).EnsureVisible
            wndReportControl.Rows(Me.wndReportControl.Rows.Count - 1).Selected = True
        Else
            Row.Record(5).Caption = DBLet(miRsAux!NomClien, "T")
            Row.Record(6).Caption = CStr(miRsAux!NIFClien)
            Row.Record(7).Caption = CStr(DBLet(miRsAux!Matricula, "T"))
            Row.Record(8).Caption = CStr(DBLet(miRsAux!licencia, "T"))
            Clientes = ", " & Row.Record(4).Caption
            PonerIconosRs Row.Index, Row.Index
        End If
        wndReportControl.Populate
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
End Sub




