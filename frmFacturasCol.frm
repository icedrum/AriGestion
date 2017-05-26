VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#17.2#0"; "Codejock.Controls.v17.2.0.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#17.2#0"; "Codejock.ReportControl.v17.2.0.ocx"
Begin VB.Form frmFacturasCol 
   Caption         =   "Facturas"
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
   StartUpPosition =   2  'CenterScreen
   Begin XtremeReportControl.ReportControl wndReportControl2 
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
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "Limpiar"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   9120
      Width           =   975
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "Regresar"
      Height          =   375
      Left            =   13080
      TabIndex        =   8
      Top             =   9120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   14055
      Begin XtremeSuiteControls.FlatEdit txtSearchBar 
         Height          =   420
         Index           =   1
         Left            =   1320
         TabIndex        =   1
         Tag             =   "numfactu"
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
         Tag             =   "fecfactu"
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
         Left            =   0
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
         Tag             =   "codclien"
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
         Tag             =   "nomclien"
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
         Tag             =   "codtipom"
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
         Left            =   12960
         TabIndex        =   11
         Tag             =   "totfaccl"
         Top             =   150
         Width           =   1500
         _Version        =   1114114
         _ExtentX        =   2646
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
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   1680
      TabIndex        =   10
      Top             =   9120
      Width           =   3375
   End
End
Attribute VB_Name = "frmFacturasCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public vSQL As String  'Siempre mandaran WHERE. O "" o lo que hay puesto
Public Event DatoSeleccionado(CadenaSeleccion As String)

Dim PrimeraVez As Boolean




Dim iconArray(0 To 9) As Long
Dim RowExpanded(0 To 49) As Boolean
Dim RowVisible(0 To 49) As Boolean
Dim MaxRowIndex As Long
Dim fntBold As StdFont
Dim fntStrike As StdFont









Private Sub VerTodos()
    For I = 0 To 4
        txtSearchBar(I).Text = ""
    Next
    CargaDatos "", False

End Sub










Private Sub cmdRegresar_Click()
Dim Devuelve As String
   If wndReportControl2.SelectedRows.Count > 0 Then
        Devuelve = wndReportControl2.SelectedRows(0).Record(1).Caption & "|" & wndReportControl2.SelectedRows(0).Record(2).Caption & "|" & wndReportControl2.SelectedRows(0).Record(3).Caption & "|"
       RaiseEvent DatoSeleccionado(Devuelve)
       Unload Me
   End If

End Sub

Private Sub cmdSearch_Click()
Dim Cad1 As String
Dim cad2 As String
Dim Tipo As String

    cad2 = ""
    J = 0
    For I = 0 To 5
        Me.txtSearchBar(I).Text = Trim(Me.txtSearchBar(I).Text)
        If Me.txtSearchBar(I).Text <> "" Then
            
            Tipo = "N"
            If I = 0 Or I = 4 Then Tipo = "T"
            If I = 2 Then Tipo = "F"
            If SeparaCampoBusqueda(Tipo, txtSearchBar(I).Tag, txtSearchBar(I).Text, Cad1) = 0 Then
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
     If wndReportControl2.SelectedRows.Count > 0 Then
        wndReportControl2.SetFocus
    Else
         For I = 0 To 4
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
        Me.Refresh
        DoEvents
        CargaDatos "", False
        wndReportControl2.Populate
         HacerPrimeravez
    End If
End Sub


Private Sub HacerPrimeravez()
    On Error Resume Next
    txtSearchBar(1).SetFocus
    Err.Clear
End Sub


Private Sub Form_Load()
    PrimeraVez = True
    Me.Icon = frmppal.Icon
    
    
    wndReportControl2.Icons = ReportControlGlobalSettings.Icons
    wndReportControl2.PaintManager.NoItemsText = "Ningún registro "
     
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
    Set wndReportControl2.PaintManager.TextFont = TextFont
    Set wndReportControl2.PaintManager.CaptionFont = TextFont
    Set wndReportControl2.PaintManager.PreviewTextFont = TextFont
    
    Exit Sub
eEstablecerFuente:
    MuestraError Err.Number, Err.Description

End Sub

Private Sub Form_Resize()
Dim Fondo
    On Error Resume Next
    
    Fondo = Me.cmdRegresar.Height + 240
    Frame1.Move 0, 0, ScaleWidth, Frame1.Height
    wndReportControl2.Move 0, Frame1.Height + Frame1.top, ScaleWidth, ScaleHeight - Fondo - Frame1.Height - Frame1.top
    SituaBusquedas
    
    cmdRegresar.Move ScaleWidth - cmdRegresar.Width - 600, ScaleHeight - Fondo + 120
    cmdLimpiar.Move 240, ScaleHeight - Fondo + 120
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub SituaBusquedas()

    On Error Resume Next
    J = 400
    For I = 0 To 5
        If I = 0 Then
            Me.txtSearchBar(I).Left = J
            Me.txtSearchBar(I).Width = (Me.wndReportControl2.Columns(I + 1).Width * 15) - 220
        Else
            Me.txtSearchBar(I).Left = txtSearchBar(I - 1).Left + txtSearchBar(I - 1).Width + 30
            Me.txtSearchBar(I).Width = (Me.wndReportControl2.Columns(I + 1).Width * 15) - 30
        End If
        
        Debug.Print txtSearchBar(I).Width
        'Me.txtSearchBar(I).Text = I
    Next
    Err.Clear
End Sub





























Public Sub CreateReportControl()
    'Start adding columns
    Dim Column As ReportColumn
    wndReportControl2.Columns.DeleteAll
    'Adds a new ReportColumn to the ReportControl's collection of columns, growing the collection by 1.
    Set Column = wndReportControl2.Columns.Add(0, "T", 18, False)
    
    
    Set Column = wndReportControl2.Columns.Add(1, "Tipo", 8, True)
    Column.Alignment = xtpAlignmentLeft
    
    Set Column = wndReportControl2.Columns.Add(2, "Numero", 10, True)
    Column.Alignment = xtpAlignmentRight
    Set Column = wndReportControl2.Columns.Add(3, "Fecha", 15, True)
    Column.Alignment = xtpAlignmentRight
    Set Column = wndReportControl2.Columns.Add(4, "Cod.", 15, True)
    Column.Alignment = xtpAlignmentRight
    Set Column = wndReportControl2.Columns.Add(5, "Nombre", 55, True)
    Set Column = wndReportControl2.Columns.Add(6, "Total", 15, True)
    Column.Alignment = xtpAlignmentRight

    wndReportControl2.PaintManager.MaxPreviewLines = 1
    wndReportControl2.PaintManager.HorizontalGridStyle = xtpGridNoLines
                  
    
    'This font will be used in the BeforeDrawRow when automatic formatting is selected
    'This simply applies Strikethrough to the currently set text font
    Set fntStrike = wndReportControl2.PaintManager.TextFont
    fntStrike.Strikethrough = True
    
    'This font will be used in the BeforeDrawRow when automatic formatting is selected
    'This simply applies Bold to the currently set text font
    Set fntBold = wndReportControl2.PaintManager.TextFont
    fntBold.Bold = True
    
    'Any time you add or delete rows(by removing the attached record), you must call the
    'Populate method so the ReportControl will display the changes.
    'If rows are added, the rows will remain hidden until Populate is called.
    'If rows are deleted, the rows will remain visible until Populate is called.
    wndReportControl2.Populate
    
    wndReportControl2.SetCustomDraw xtpCustomBeforeDrawRow
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




    On Error GoTo eCargaDatos

    Screen.MousePointer = vbHourglass
    
    V = True
    Label V
    wndReportControl2.ShowItemsInGroups = False
    wndReportControl2.Records.DeleteAll
    wndReportControl2.Populate
    
    Set miRsAux = New ADODB.Recordset
    
    If EsTodoSQL Then
        Stop
    Else
        
        If Sql <> "" Then Sql = " AND " & Sql
        If vSQL <> "" Then Sql = Sql & " AND " & vSQL
            
        Sql = " FROM factcli,clientes where factcli.codclien=clientes.codclien" & Sql
        Sql = "SELECT numserie,numfactu,fecfactu,factcli.codclien,nomclien,totfaccl  " & Sql
        
        Sql = Sql & " ORDER BY numserie,numfactu,fecfactu "
    End If
    
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Inicial = 0
    

    While Not miRsAux.EOF
        AddRecord2
    
        N = N + 1
        If N > 40 Then
            wndReportControl2.Populate
            N = 0
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
        
    
    wndReportControl2.Populate
    
    
eCargaDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description, Sql
    
    
    
    
    Label1.Caption = ""
    Screen.MousePointer = vbDefault
End Sub




'socio, pendiente , nombre, matricula,licencia
'Leera los datos de mirsaux
Private Sub AddRecord2()
Dim Record As ReportRecord
Dim OtroIcono As Boolean
Dim Impor As Currency


    On Error GoTo eAddRecord2

    'Adds a new Record to the ReportControl's collection of records, this record will
    'automatically be attached to a row and displayed with the Populate method
    Set Record = wndReportControl2.Records.Add()
    
    Dim Item As ReportRecordItem
    'Socio
    Set Item = Record.AddItem("")
    
    Item.SortPriority = IIf(True, 1, 0)

    
    
    '  Item.Icon = IIf(Socio, RECORD_UNREAD_MAIL_ICON, -1)
    
 
    
    
    Set Item = Record.AddItem(CStr(miRsAux!numSerie))
    
    Set Item = Record.AddItem(Format(miRsAux!NumFactu, "00000"))
    Set Item = Record.AddItem(Format(miRsAux!Fecfactu, "yyy-mm-dd"))
    Item.Caption = Format(miRsAux!Fecfactu, "dd/mm/yyyy")
    
    Set Item = Record.AddItem(CStr(miRsAux!CodClien))
    Item.Caption = Format(miRsAux!CodClien, "0000")
    
    Record.AddItem CStr(DBLet(miRsAux!NomClien, "T"))
    
    Impor = DBLet(miRsAux!totfaccl, "N")
    Set Item = Record.AddItem(Impor)
    Item.Caption = Format(Impor, FormatoImporte)
    
    
   
    
    
    
    
eAddRecord2:
    
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

    Set wndReportControl2.Icons = ReportControlGlobalSettings.Icons
    
    wndReportControl2.ToolTipContext.Style = frmppal.CommandBars.ToolTipContext.Style
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





Private Sub wndReportControl2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdRegresar_Click
    End If
End Sub

Private Sub wndReportControl2_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    cmdRegresar_Click
End Sub


