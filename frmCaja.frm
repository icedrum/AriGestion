VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#17.2#0"; "Codejock.ReportControl.v17.2.0.ocx"
Begin VB.Form frmCaja 
   Caption         =   "CAJA"
   ClientHeight    =   9120
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16080
   LinkTopic       =   "Form1"
   ScaleHeight     =   9120
   ScaleWidth      =   16080
   StartUpPosition =   1  'CenterOwner
   Begin XtremeReportControl.ReportControl wndReportControl 
      Height          =   9615
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   16215
      _Version        =   1114114
      _ExtentX        =   28601
      _ExtentY        =   16960
      _StockProps     =   64
      FreezeColumnsAbs=   0   'False
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   14895
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   9000
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   240
         Width           =   2775
      End
      Begin VB.ComboBox Combo1 
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
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   270
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   11880
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   240
         Width           =   1575
      End
      Begin VB.Frame FrameBotonGnral2 
         Height          =   705
         Left            =   3120
         TabIndex        =   4
         Top             =   120
         Width           =   1575
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   330
            Left            =   120
            TabIndex        =   5
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
                  Object.ToolTipText     =   "Cierre caja"
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Errores NºFactura"
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame FrameBotonGnral 
         Height          =   705
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   2865
         Begin MSComctlLib.Toolbar Toolbar1 
            Height          =   330
            Left            =   120
            TabIndex        =   3
            Top             =   180
            Width           =   2655
            _ExtentX        =   4683
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
                  Object.Visible         =   0   'False
                  Style           =   3
               EndProperty
               BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Imprimir"
               EndProperty
               BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
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
      Begin VB.Label Label2 
         Caption         =   "Cierre"
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
         Left            =   7680
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   7560
         X2              =   7560
         Y1              =   240
         Y2              =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Cierre"
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
         Left            =   8400
         TabIndex        =   9
         Top             =   345
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim PrimVez As Boolean
Dim Importe As Currency



Private Sub Combo1_Click()
    If PrimVez Then Exit Sub
    ObtenerCierreUltimo
    MostrarDatos
End Sub

Private Sub Form_Activate()
    If PrimVez Then
        PrimVez = False
        MostrarDatos
    End If
End Sub

Private Sub Form_Load()
Dim TextFont

    Me.Icon = frmppal.Icon
    PrimVez = True
    
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
        
        .Buttons(8).Image = 16
        .Buttons(9).Image = 32
        .Buttons(9).Enabled = True
        
        'Ocultamos
        '.Buttons(9).Visible = False
        .Buttons(5).Visible = False
        .Buttons(6).Visible = False
        
    End With
        
    
 ' Botonera Principal 2
    With Me.Toolbar2
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 27
        .Buttons(2).Image = 25
        
        '.Buttons(3).Image = 42
        .Buttons(3).Visible = False
    End With
    
    
    
    CargaCombo
    wndReportControl.AllowColumnReorder = False
    CreateReportControlCaja
   '
   '
   ' '
   ' Dim TextFont As StdFont
    Set TextFont = Label1.Font
    TextFont.SIZE = 11
    Set wndReportControl.PaintManager.TextFont = TextFont
    Label1.Caption = ""
    
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Frame1.Width = Me.Width - 240
    wndReportControl.Move 60, Me.Frame1.Height + 120, Me.Width - 320, Me.Height - Me.Frame1.Height - 480
    
    Me.Text1.Move Frame1.Width - Text1.Width - 640
    Me.Text2.Move Text1.Left - Text2.Width - 120
    Label2.Move Text2.Left - Label2.Width - 120
    Err.Clear
End Sub





Public Sub CreateReportControlCaja()
    'gestadministrativa  id usuario fechacreacion llevados importe fechafinalizacion
    Dim Column As ReportColumn
    
    wndReportControl.Columns.DeleteAll
    wndReportControl.PaintManager.MaxPreviewLines = 1
    wndReportControl.PaintManager.HorizontalGridStyle = xtpGridNoLines
    'Adds a new ReportColumn to the ReportControl's collection of columns, growing the collection by 1.
    Set Column = wndReportControl.Columns.Add(COLUMN_IMPORTANCE, "Tipo", 18, False)
    'The value assigned to the icon property corresponds to the index of an icon in the collection of wndReportControl.Icons
    'I.e. The icon at index=1 in the collection will be displayed in the column header.  The index of the icon depends on the
    'order it is added to the collection.  (Icons are added after the records near the bottom of the Form_Load)
    Column.Icon = COLUMN_IMPORTANCE_ICON
    
    Set Column = wndReportControl.Columns.Add(2, "Fecha", 90, True)
    Set Column = wndReportControl.Columns.Add(3, "Origen", 60, True)
    'Febrero 17
    'Cuenta contable
    Set Column = wndReportControl.Columns.Add(4, "Destino", 110, True)
    
    Set Column = wndReportControl.Columns.Add(5, "Ampliacion", 150, True)
    Set Column = wndReportControl.Columns.Add(6, "Importe", 55, True)
    Column.Alignment = xtpAlignmentRight
    Set Column = wndReportControl.Columns.Add(7, "Sal", 15, True)
    Column.ToolTip = "Salida"
    Column.Alignment = xtpAlignmentRight
End Sub


Private Sub MostrarDatos()

    Label1.Caption = "Leyendo BD"
    Label1.Refresh
    Screen.MousePointer = vbHourglass
    
    
    populateInbox
    
    
    Label1.Caption = ""
    Screen.MousePointer = vbDefault
End Sub



Public Sub populateInbox()
Dim C As String
Dim F As Date


    wndReportControl.Records.DeleteAll
  
    C = "Select * from caja where usuario = " & DBSet(Combo1.Text, "T")
    'Si ha habido un cierre de caja
    If Text2.Text <> "" Then C = C & " AND feccaja >" & DBSet(Text2.Text, "FH")
    C = C & " order by feccaja asc"
    Importe = Me.Text1.Tag
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        
        AddRecordCaja False
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    wndReportControl.Populate
    
    
    'Añado un item de total
    If Me.wndReportControl.Rows.Count > 0 Then
        AddRecordCaja True
        wndReportControl.Populate
    End If
    
End Sub

Private Sub AddRecordCaja(EsElDeTotal As Boolean)
Dim Origen As Byte
Dim Aux As String
Dim ClavePRimaria As String
Dim EsFactura As Boolean
  
    Dim Record As ReportRecord
    'Adds a new Record to the ReportControl's collection of records, this record will
    'automatically be attached to a row and displayed with the Populate method
    Set Record = wndReportControl.Records.Add()
    
    Dim Item As ReportRecordItem
   
    
    If EsElDeTotal Then
        Record.Tag = "T"
        Record.AddItem ("")
        Record.AddItem ("")
        Record.AddItem ("")
        Record.AddItem ("")
        Record.AddItem ("")
        Set Item = Record.AddItem("TOTAL")
        Item.Bold = True
        
        Set Item = Record.AddItem(Format(Importe, FormatoImporte))
        If Importe < 0 Then
            Item.BackColor = vbRed
        Else
            Item.BackColor = vbBlue
        End If
        Item.ForeColor = vbWhite
        Exit Sub
    End If
    
    
    
    
    
    
    
    
    Aux = ""
    ClavePRimaria = ""
    EsFactura = False
    ''0=Cobro 1=Pago 2=Cierre',
    If miRsAux!tipomovi = 2 Then
        'Apunte de cierre
        Origen = 1
        Aux = "CIERRE"
    ElseIf miRsAux!tipomovi = 1 Then
        'PAGO
        Origen = 7
        
        
        'Gestion administrativa
        If Not IsNull(miRsAux!TipoRegi) Then
            If miRsAux!TipoRegi = 0 And DBLet(miRsAux!numSerie, "T") <> "" Then
                'Falria revisar ya que he aadido el numlinea mas adelante
                Aux = "EXP " & Format(miRsAux!Numdocum, "00000") & "/" & Right(CStr(miRsAux!anoexped), 2)
                'ClavePRimaria = "numexped =" & miRsAux!Numdocum & " AND tiporegi =" & miRsAux!TipoRegi
                'ClavePRimaria = ClavePRimaria & " AND anoexped=" & miRsAux!anoexped
                EsFactura = False
            End If
        End If
        
        
        
        
        
        
    Else
        'Cobro
        If IsNull(miRsAux!TipoRegi) Then
            'Es un cobro "manual"
            Origen = 6
        Else
            'Puede pagar una factura o un expediente
            Origen = 4
            
            'El expediente puede estar factura, y el cobro haberse realizado sobre la factura
            EsFactura = True
            If miRsAux!TipoRegi = 0 Then
                If miRsAux!numSerie = "1" Then
                    'Falria revisar ya que he aadido el numlinea mas adelante
                    Aux = "EXP " & Format(miRsAux!Numdocum, "00000") & "/" & Right(CStr(miRsAux!anoexped), 2)
                    ClavePRimaria = "numexped =" & miRsAux!Numdocum & " AND tiporegi =" & miRsAux!TipoRegi
                    ClavePRimaria = ClavePRimaria & " AND anoexped=" & miRsAux!anoexped
                    EsFactura = False
                End If
            End If
            
            
            If EsFactura Then
                Aux = miRsAux!numSerie & Format(miRsAux!Numdocum, "00000")
                ClavePRimaria = "XXXXXX"
            End If
            
        End If
    End If
   
   
   
    'Adds a new ReportRecordItem to the Record, this can be thought of as adding a cell to a row
    
    Set Item = Record.AddItem("")
    If Origen = 1 Then
        'Assigns an icon to the item
        Item.Icon = Origen
        Item.ToolTip = "Cierre"
    ElseIf Origen = 7 Then
        Item.Icon = Origen
        Item.ToolTip = "Pago"
        
    ElseIf Origen = 4 Then
        Item.Icon = Origen
        Item.ToolTip = "Factura/expediente"
    Else
        '6 manual
        Item.Icon = Origen
        Item.ToolTip = "Manual"
    End If


      

    Set Item = Record.AddItem("")
    
    Set Item = Record.AddItem(Format(miRsAux!feccaja, "dd/mm/yyyy hh:nn:ss"))
    Item.Caption = Format(miRsAux!feccaja, "dd/mm/yyyy hh:nn")
    
    'Origen del pago
    Set Item = Record.AddItem(Aux)
    Item.Tag = ClavePRimaria
     
     
    'DEstino
    Aux = DBLet(miRsAux!codmacta, "T")
    Msg = " "
    If Aux <> "" Then
        Msg = DevuelveDesdeBD("nommacta", "ariconta" & vParam.Numconta & ".cuentas", "codmacta", Aux, "T")
        If Msg = "" Then Msg = "NO encontado"
    End If
    Set Item = Record.AddItem(Msg)
    Item.Tag = Aux
     
     
    Record.AddItem DBLet(miRsAux!Ampliacion, "T")
   
    
    Set Item = Record.AddItem("")
    'Specifys the format that the price will be displayed
    'Item.Format = " %s"
    Item.Format = "%.2f"
    Item.Value = CCur(miRsAux!Importe)
    
       
    ' ''0=Cobro 1=Pago 2=Cierre',
    If miRsAux!tipomovi = 1 Then
        Importe = Importe - miRsAux!Importe
        Item.Caption = Format(-Item.Value, FormatoImporte)
        Item.ForeColor = vbRed
    Else
        Importe = Importe + miRsAux!Importe
        Item.Caption = Format(Item.Value, FormatoImporte)
    End If
    
    Msg = ""
    If miRsAux!tipomovi > 0 Then Msg = "*"
    Set Item = Record.AddItem(Msg)
        
    'Adds the PreviewText to the Record.  PreviewText is the text displayed for the ReportRecord while in PreviewMode
    'Record.PreviewText = miRsAux!NomClien
    
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    HacerToolBar Button.Index
End Sub
Private Sub HacerToolBar(indice As Integer)
Dim B As Boolean
Dim GroupRow
    B = False
    If indice <> 1 And indice < 8 Then
        'Si esta "cerrado" ya no puedo hacer nada
        
        If Me.wndReportControl.Records.Count = 0 Then Exit Sub
        If wndReportControl.SelectedRows.Count = 0 Then Exit Sub
        
        'Es un agrupado
        If wndReportControl.SelectedRows(0).Record.Tag = "T" Then Exit Sub
        Msg = Trim(wndReportControl.SelectedRows(0).Record(3).Caption)
        If Msg <> "" Then
            MsgBox "No se puede editar el movimiento de caja.Esta vinculado", vbExclamation
            Exit Sub
        End If
        
    End If
    
    Select Case indice
    Case 1
        CadenaDesdeOtroForm = ""
        frmMensajes.Parametros = Combo1.Text & "||"
        frmMensajes.Opcion = 4
        frmMensajes.Show vbModal
        If CadenaDesdeOtroForm <> "" Then B = True
           
    Case 2
        
        CadenaDesdeOtroForm = ""
        Msg = wndReportControl.SelectedRows(0).Record(2).Value & "|" & wndReportControl.SelectedRows(0).Record(5).Caption
        Msg = Msg & "|" & Abs(wndReportControl.SelectedRows(0).Record(6).Caption) & "|"
        If wndReportControl.SelectedRows(0).Record(7).Caption = "*" Then Msg = Msg & "1"
        Msg = Msg & "|"
        'Codmacta|nommacta
        Msg = Msg & wndReportControl.SelectedRows(0).Record(4).Tag & "|" & Trim(wndReportControl.SelectedRows(0).Record(4).Value) & "|"
        frmMensajes.Parametros = Combo1.Text & "|" & Msg
        frmMensajes.Opcion = 4
        frmMensajes.Show vbModal
        If CadenaDesdeOtroForm <> "" Then B = True
    Case 3
        
        'Eliminaremos el movimiento
        If wndReportControl.SelectedRows(0).Record(7).Caption = "*" Then
            Msg = "Salida"
        Else
            Msg = "Entrada"
        End If
        Msg = "Tipo: " & Msg & "   €" & Abs(wndReportControl.SelectedRows(0).Record(6).Caption) & vbCrLf
        Msg = "Caja: " & Combo1.Text & "  Fecha: " & wndReportControl.SelectedRows(0).Record(2).Value & vbCrLf & Msg
        Msg = Msg & "Ampliacion: " & wndReportControl.SelectedRows(0).Record(5).Caption & vbCrLf
        Msg = Msg & "Cuenta: " & wndReportControl.SelectedRows(0).Record(4).Tag & "  " & Trim(wndReportControl.SelectedRows(0).Record(4).Value) & vbCrLf
        
        If MsgBox("Va a eliminar el dato de caja: " & vbCrLf & vbCrLf & Msg & vbCrLf & "¿Continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
                
        MsgErr = DBSet(wndReportControl.SelectedRows(0).Record(2).Value, "FH")
        MsgErr = "DELETE FROM caja WHERE usuario = " & DBSet(Combo1.Text, "T") & " AND feccaja=" & MsgErr
        If Ejecuta(MsgErr) Then
            B = True
            vLog.Insertar 7, vUsu, Msg
        End If
        
    Case 8
        
        ImprimirProceso
    Case 9
        If Combo1.ListIndex < 0 Then Exit Sub
        frmMensajes.Opcion = 10
        frmMensajes.Parametros = Me.Combo1.Text
        frmMensajes.Show vbModal
    End Select
    If B Then MostrarDatos
End Sub

Private Sub ImprimirProceso()


    InicializarVbles True
    
    '{caja.feccaja} >= DateTime (2016, 12, 31, 12, 20, 10)
    
    cadNomRPT = "rCaja.rpt"
    
    
    If Me.Text2.Text = "" Then
        'No hay ciere
        Text2.Text = ""
        Msg = ""
        Importe = 0
        cadFormula = "1 = 1"
    Else

        If Not PonerDesdeHasta("caja.feccaja", "FH", Text2, Nothing, Nothing, Nothing, Msg) Then Exit Sub
        Importe = ImporteFormateado(Text1.Text)
        Msg = Text2.Text
    End If
    cadParam = cadParam & "|ImporteIncial=" & TransformaComasPuntos(CStr(Importe)) & "|"
    cadParam = cadParam & "|UltimoCierre= """ & Msg & """|"
    numParam = numParam + 2
    cadFormula = cadFormula & " AND {caja.usuario}= """ & Combo1.Text & """"
   
   


    ImprimeGeneral
    
    
    
End Sub



Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)


    'Es un agrupado
    If Me.wndReportControl.Records.Count = 0 Then Exit Sub
    
   
    I = wndReportControl.Rows.Count - 1
            'Es un agrupado
    If wndReportControl.Records(I).Tag <> "T" Then
        MsgBox "No hay sumatorio caja", vbExclamation
        Exit Sub
    End If
    
    'Veré si estan los parametros
    If vParam.CtaGastosCaja = "" Or vParam.CtaIngresosCaja = "" Then
        MsgBox "Falta configurar datos caja", vbExclamation
        Exit Sub
    End If
    
    'FALTA####
    'HABLAR CON MANOLO
    'Cad usuario tendra SU caja , con lo cual , el parametros vParam.CtaCaja sera
    'la raiz para las cajas, o no
    
    
    
    cadFormula = Combo1.Text & "|" & wndReportControl.Records(I).Item(6).Caption & "|"
    
    CadenaDesdeOtroForm = ""
    frmMensajes.Parametros = cadFormula
    frmMensajes.Opcion = 5
    frmMensajes.Show vbModal
    
    
    
    If CadenaDesdeOtroForm <> "" Then
        Screen.MousePointer = vbHourglass
        Conn.BeginTrans
        If HacerProcesCierreCaja() Then
            Conn.CommitTrans
            espera 0.5
            
            'Volvemos a cargar
            ObtenerCierreUltimo
            MostrarDatos
            DoEvents
            espera 0.5
            ImprimirCierreCaja



        Else
            Conn.RollbackTrans
        End If
        Screen.MousePointer = vbDefault
    End If
End Sub



Private Sub CargaCombo()


    Combo1.Clear
    'Conceptos
    Set miRsAux = New ADODB.Recordset
        
    Msg = "Select * from usuarios.usuarios where nivelariges <> -1 "
    If vUsu.Nivel <> 0 Then Msg = Msg & " AND login= " & DBSet(vUsu.Login, "T")
    Msg = Msg & "  order by login"
    miRsAux.Open Msg, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    
    While Not miRsAux.EOF
        Combo1.AddItem miRsAux!Login
        Combo1.ItemData(Combo1.NewIndex) = miRsAux!codusu
        If miRsAux!Login = vUsu.Login Then Combo1.ListIndex = Combo1.NewIndex
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    If Combo1.Text <> "" Then ObtenerCierreUltimo
End Sub




Private Sub ObtenerCierreUltimo()
Dim Cad As String
Dim F As Date
    
    Cad = "feccaja"
    Msg = " usuario=" & DBSet(Combo1.Text, "T") & " AND 1 "
    Msg = DevuelveDesdeBD("importe", "caja_param", Msg, "1 ORDER BY  feccaja DESC", "N", Cad)
    If Msg = "" Then Msg = "0": Cad = ""
    Text1.Text = Format(Msg, FormatoImporte)
    Text1.Tag = CCur(Msg)
    'Le sumo un segundo a la fecha para el listado , ya que no puedo poner >  y pone>=
    If Cad <> "" Then
        F = CDate(Cad)
        F = DateAdd("s", 1, F)
        Cad = F
    End If
    Text2.Text = Cad
End Sub

Private Sub wndReportControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub wndReportControl_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    HacerToolBar 2
End Sub



Private Function HacerProcesCierreCaja() As Boolean
Dim ColApu As Collection
Dim Aux As String
Dim aux2 As String
Dim CtaCaja As String
Dim FechaCierre As Date
Dim ImporteParaCaja As Currency
Dim Serie As String
Dim FechaFactura As String



On Error GoTo eHacerProcesoCierreCaja
    HacerProcesCierreCaja = False
    
    
    CtaCaja = vParam.CtaCaja
    FechaCierre = CDate(RecuperaValor(CadenaDesdeOtroForm, 1))
    
    'Si queda importe meteremos una linea con el importe retirado para hacer cuadrar en caja
    'Ejmplo.
    '
    '   Saldo anterior  100
    '   Cobro            40
    '   Gasto            25
    '               -------
    '                   115
    '
    '   Dejamos caja     60  --> Retiramos 55 para ingresar
        
        
    'Iremos uno a uno cuadrando todos los movimientos
    Set ColApu = New Collection
    ImporteParaCaja = 0
    
        '' Llevara
        '       codmacta | docum | codconce | ampliaci | imported|importeH |
    For I = 0 To Me.wndReportControl.Rows.Count - 2 'El ultimo no se procesa
        ImporteParaCaja = ImporteParaCaja + ImporteFormateado(wndReportControl.Rows(I).Record(6).Caption)

        
        'Origen = 7  Pago
        'Origen = 4 Then Factura/expediente
        '6 manual Item.Icon = Origen
        If wndReportControl.Rows(I).Record(0).Icon = 7 Then
            'PAGO
             Aux = vParam.CtaGastosCaja & "|C:" & Combo1.Text & "|2|" & wndReportControl.Rows(I).Record(4).Caption & "|"
             Aux = Aux & Replace(wndReportControl.Rows(I).Record(6).Caption, "-", "") & "|" & "|" & CtaCaja & "|"
             
        ElseIf wndReportControl.Rows(I).Record(0).Icon = 4 Then
            'Factura expediente
            'Buscare el cliente del expediente
            
            Aux = wndReportControl.Rows(I).Record(3).Caption
            If Mid(Aux, 1, 3) = "EXP" Then
                'EXP 00006/17
                Aux = wndReportControl.Rows(I).Record(3).Tag
                Aux = Aux & " AND 1"
                Aux = DevuelveDesdeBD("codclien", "expedientes", Aux, "1")
                aux2 = DevuelveCuentaContableCliente(False, Aux)
                If aux2 = "" Then Err.Raise 513, "Error obteniendo cuenta contable: " & wndReportControl.Rows(I).Record(3).Caption
                '       codmacta | docum | codconce | ampliaci | imported|importeH |
                Aux = Replace(wndReportControl.Rows(I).Record(3).Caption, "XP ", "")
                Aux = aux2 & "|" & Aux & "|1|EXP pago a cuenta " & Mid(wndReportControl.Rows(I).Record(2).Caption, 1, 10) & " "
                Aux = Aux & "||" & wndReportControl.Rows(I).Record(6).Caption & "|" & CtaCaja & "|"
            Else
                '       codmacta | docum | codconce | ampliaci | imported|importeH |
                
                
                'Es una factura
                Serie = Mid(wndReportControl.Rows(I).Record(3).Caption, 1, 3)
                Aux = "numserie= '" & Serie & "' and numfactu= " & Mid(wndReportControl.Rows(I).Record(3).Caption, 4)
                Aux = Aux & " AND 1"
                FechaFactura = "fecfactu"
                Aux = DevuelveDesdeBD("codclien", "factcli", Aux, "1", "N", FechaFactura)
                
                aux2 = DevuelveCuentaContableCliente(Serie = "CUO", Val(Aux))
                If aux2 = "" Then Err.Raise 513, "Error obteniendo cuenta contable: " & wndReportControl.Rows(I).Record(3).Caption
                
                Aux = wndReportControl.Rows(I).Record(3).Caption
                Aux = aux2 & "|" & Aux & "|1|Factura " & wndReportControl.Rows(I).Record(3).Caption & "|"
                If ImporteFormateado(wndReportControl.Rows(I).Record(6).Caption) > 0 Then
                    aux2 = "|" & wndReportControl.Rows(I).Record(6).Caption
                Else
                    aux2 = Replace(wndReportControl.Rows(I).Record(6).Caption, "-", "") & "|"
                End If
                Aux = Aux & aux2 & "|" & CtaCaja & "|"
                
                Aux = Aux & Serie & "|" & Mid(wndReportControl.Rows(I).Record(3).Caption, 4) & "|" & FechaFactura & "|"
            End If
            
            
            
            
        Else
            'MANUAL
           
            '       codmacta | docum | codconce | ampliaci | imported|importeH |
            If wndReportControl.Rows(I).Record(4).Tag <> "" Then
                aux2 = wndReportControl.Rows(I).Record(4).Tag
            Else
                aux2 = vParam.CtaIngresosCaja
            End If
            Aux = aux2 & "|C:" & Combo1.Text & "|1|"
            aux2 = wndReportControl.Rows(I).Record(5).Caption
            aux2 = Mid(aux2 & " " & wndReportControl.Rows(I).Record(4).Caption, 1, 50)
            Aux = Aux & aux2 & "|"
            Aux = Aux & "|" & wndReportControl.Rows(I).Record(6).Caption & "|" & CtaCaja & "|"
        End If
        ColApu.Add Aux
        
    Next
    'Cuadre de caja
    aux2 = "Cierre caja : " & Format(FechaCierre, "dd/mm/yyyy hh:nn")
    Aux = CtaCaja & "|C:" & Combo1.Text & "|1|" & aux2 & "|"
    Aux = Aux & Format(ImporteParaCaja, FormatoImporte) & "|" & "|" & "|"
    ColApu.Add Aux
        
    'Si llevamos a banco, metermeos dos lineas mas
    ImporteParaCaja = RecuperaValor(CadenaDesdeOtroForm, 2)
    If ImporteParaCaja <> 0 Then
    
        aux2 = "Sale caja " & Combo1.Text & " a banco "
        Aux = CtaCaja & "|C:" & Combo1.Text & "|1|" & aux2 & "|"
        Aux = Aux & "|" & Format(ImporteParaCaja, FormatoImporte) & "|" & vParam.CtaBanco & "|"
        ColApu.Add Aux
            
        
        aux2 = "Sale caja " & Combo1.Text & " a banco "
        Aux = vParam.CtaBanco & "|C:" & Combo1.Text & "|1|" & aux2 & "|"
        Aux = Aux & Format(ImporteParaCaja, FormatoImporte) & "|" & "|" & CtaCaja & "|"
        ColApu.Add Aux
    End If
        
    'NO LLEVAMOS A BANCO   NO LLEVAMPS
'''''
'''''                        'Los apuntes da caja--banco
'''''                        'A banco
'''''                        Aux = RecuperaValor(CadenaDesdeOtroForm, 2)
'''''                        Importe = ImporteFormateado(Aux)
'''''                        Aux = vParam.CtaBanco & "|C:" & Combo1.Text & "|1|" & aux2 & "|"
'''''                        Aux = Aux & Format(Importe, FormatoImporte) & "||" & CtaCaja & "|"
'''''                        ColApu.Add Aux
'''''
'''''                        'Desde caja
'''''                        Aux = CtaCaja & "|C:" & Combo1.Text & "|1|" & aux2 & "|"
'''''                        Aux = Aux & "|" & Format(Importe, FormatoImporte) & "|" & vParam.CtaBanco & "|"
'''''                        ColApu.Add Aux
'''''
   
    'Ahora ya , con las lineas de apuntes, mandariamos a crear el apunte con las col de apuntes
    Aux = "Cierre caja    Usuario: " & Combo1.Text & "   Fecha cierre: " & Format(FechaCierre, "dd/mm/yyyy hh:nn") & vbCrLf
    Aux = Aux & "Lineas : " & Me.wndReportControl.Rows.Count - 1 & vbCrLf
    Aux = Aux & "Queda caja : " & RecuperaValor(CadenaDesdeOtroForm, 3) & "    "
    Aux = Aux & "Llevado banco : " & RecuperaValor(CadenaDesdeOtroForm, 2) & vbCrLf
    
    If CrearApunteDesdeColeccion(FechaCierre, Aux, ColApu) Then
        'AHora haremos dos cosas mas.
        'Añadir el dinero que sale de caja para llevar a banco
        aux2 = "Lineas: " & wndReportControl.Rows.Count - 1 'ampliacion
        Aux = RecuperaValor(CadenaDesdeOtroForm, 2)
        Importe = ImporteFormateado(Aux)
        Aux = "INSERT INTO caja (usuario,feccaja,tipomovi,importe,ampliacion) VALUES ("
        Aux = Aux & DBSet(Combo1.Text, "T") & "," & DBSet(FechaCierre, "FH") & ",2,"
        Aux = Aux & DBSet(Importe, "N") & "," & DBSet(aux2, "T") & ")"
        Conn.Execute Aux
        
        'Poner la caja el dinero que queda
        Aux = RecuperaValor(CadenaDesdeOtroForm, 3)
        Importe = ImporteFormateado(Aux)
        Aux = "REPLACE INTO caja_param(usuario,feccaja,importe) VALUES ("
        Aux = Aux & DBSet(Combo1.Text, "T") & "," & DBSet(FechaCierre, "FH") & "," & DBSet(Importe, "N") & ")"
        Conn.Execute Aux
        
        HacerProcesCierreCaja = True
    End If
  
eHacerProcesoCierreCaja:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
    Set ColApu = Nothing
End Function




Private Sub ImprimirCierreCaja()
Dim Aux As String
Dim F As Date



    InicializarVbles True
    '{caja.feccaja} >= DateTime (2016, 12, 31, 12, 20, 10)
    cadNomRPT = "rCajaCierre.rpt"
    
    
    F = RecuperaValor(CadenaDesdeOtroForm, 1)
    
    'rpt
    cadFormula = Year(F) & "," & Month(F) & "," & Day(F) & "," & Hour(F) & "," & Minute(F) & "," & Second(F)
    cadFormula = "{caja.feccaja} <=  DateTime (" & cadFormula & ")"
    'select
    Msg = "feccaja < " & DBSet(F, "FH")
    Msg = Msg & " AND usuario=" & DBSet(Combo1.Text, "T") & " AND 1 "
    Aux = "feccaja"
    Msg = DevuelveDesdeBD("importe", "caja_param", Msg, "1 ORDER BY  feccaja DESC", "N", Aux)
    If Msg = "" Then Msg = "0": Aux = ""
    Importe = CCur(Msg)
    'Le sumo un segundo a la fecha para el listado , ya que no puedo poner >  y pone>=
    If Aux <> "" Then
        F = CDate(Aux)
        Aux = Year(F) & "," & Month(F) & "," & Day(F) & "," & Hour(F) & "," & Minute(F) & "," & Second(F)
        cadFormula = cadFormula & " AND {caja.feccaja} >  DateTime (" & Aux & " )"
    End If
    
    cadParam = cadParam & "|ImporteIncial=" & TransformaComasPuntos(CStr(Importe)) & "|"
    cadParam = cadParam & "|UltimoCierre= """ & RecuperaValor(CadenaDesdeOtroForm, 1) & """|"
    numParam = numParam + 2
    cadFormula = cadFormula & " AND {caja.usuario}= """ & Combo1.Text & """"
    ImprimeGeneral
    
End Sub
