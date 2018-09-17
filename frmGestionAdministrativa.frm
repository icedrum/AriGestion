VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#17.2#0"; "Codejock.ReportControl.v17.2.0.ocx"
Begin VB.Form frmGestionAdministrativa 
   Caption         =   "Gestion tasas administrativas"
   ClientHeight    =   9120
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16080
   LinkTopic       =   "Form1"
   ScaleHeight     =   9120
   ScaleWidth      =   16080
   StartUpPosition =   2  'CenterScreen
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
      Width           =   13815
      Begin VB.Frame FrameBotonGnral2 
         Height          =   705
         Left            =   3120
         TabIndex        =   6
         Top             =   120
         Width           =   2055
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   330
            Left            =   120
            TabIndex        =   7
            Top             =   150
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Style           =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Cerrar proceso"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Contabilizar"
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Entrega documentacion"
               EndProperty
            EndProperty
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Desglosado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6240
         TabIndex        =   5
         Top             =   360
         Width           =   2415
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
      Begin VB.Label Label1 
         Caption         =   "Desglosar"
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
         Left            =   12360
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmGestionAdministrativa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PrimVez As Boolean
Dim QueID As Long
Dim Situacion As Byte

Private Sub Check1_Click()
    If PrimVez Then Exit Sub
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
        .Buttons(9).Visible = False
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
        
        .Buttons(3).Image = 42
        .Buttons(3).Visible = True
        If LCase(vUsu.Login) <> "root" And LCase(vUsu.Login) <> "antonio" Then .Buttons(3).Visible = False
    End With
    
    
    
    Me.Check1.Value = 1

    CreateReportControl
   '
   '
   ' '
   ' Dim TextFont As StdFont
    Set TextFont = Label1.Font
    TextFont.SIZE = 10
    Set wndReportControl.PaintManager.TextFont = TextFont
    Label1.Caption = ""
    
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Frame1.Width = Me.Width - 240
    wndReportControl.Move 60, Me.Frame1.Height + 120, Me.Width - 320, Me.Height - Me.Frame1.Height - 120
    
    Err.Clear
End Sub


Private Sub CreateReportControl()

    If Me.Check1.Value = 1 Then
        CreateReportControlDesglosado
    Else
        CreateReportControlSindesglose
    End If
    Check1.Tag = Val(Check1.Value)
End Sub


Public Sub CreateReportControlSindesglose()
    'gestadministrativa  id usuario fechacreacion llevados importe fechafinalizacion
    Dim Column As ReportColumn
    
    wndReportControl.Columns.DeleteAll
    
    'Adds a new ReportColumn to the ReportControl's collection of columns, growing the collection by 1.
    Set Column = wndReportControl.Columns.Add(COLUMN_IMPORTANCE, "Situacion", 18, False)
    'The value assigned to the icon property corresponds to the index of an icon in the collection of wndReportControl.Icons
    'I.e. The icon at index=1 in the collection will be displayed in the column header.  The index of the icon depends on the
    'order it is added to the collection.  (Icons are added after the records near the bottom of the Form_Load)
    Column.Icon = COLUMN_IMPORTANCE_ICON
    
    Set Column = wndReportControl.Columns.Add(2, "ID", 30, True)
    Set Column = wndReportControl.Columns.Add(3, "Creada", 120, True)
    Set Column = wndReportControl.Columns.Add(4, "Usuario", 120, True)
    Set Column = wndReportControl.Columns.Add(5, "Por Banco", 120, True)
    Set Column = wndReportControl.Columns.Add(6, "Cerrada", 140, True)
    Set Column = wndReportControl.Columns.Add(7, "Lineas", 25, True)
    Set Column = wndReportControl.Columns.Add(8, "Importe", 125, True)
    Column.Alignment = xtpAlignmentRight
End Sub


Public Sub CreateReportControlDesglosado()
    'Start adding columns
    Dim Column As ReportColumn
    
    wndReportControl.Columns.DeleteAll
    
    'Adds a new ReportColumn to the ReportControl's collection of columns, growing the collection by 1.
    Set Column = wndReportControl.Columns.Add(COLUMN_IMPORTANCE, "Anterior", 18, False)
    'The value assigned to the icon property corresponds to the index of an icon in the collection of wndReportControl.Icons
    'I.e. The icon at index=1 in the collection will be displayed in the column header.  The index of the icon depends on the
    'order it is added to the collection.  (Icons are added after the records near the bottom of the Form_Load)
    Column.Icon = COLUMN_IMPORTANCE_ICON
    Set Column = wndReportControl.Columns.Add(1, "ID", 15, True)
    Set Column = wndReportControl.Columns.Add(2, "Entr.", 15, True)
    Column.ToolTip = "Documentacion entregada"
    Set Column = wndReportControl.Columns.Add(3, "Licencia", 30, True)
    Set Column = wndReportControl.Columns.Add(4, "Nombre", 140, True)
    Set Column = wndReportControl.Columns.Add(5, "Conce.", 25, True)
    Set Column = wndReportControl.Columns.Add(6, "Descripción", 115, True)
    Set Column = wndReportControl.Columns.Add(7, "Expediente", 35, True)
    Set Column = wndReportControl.Columns.Add(8, "F.Exp", 35, True)
    Set Column = wndReportControl.Columns.Add(9, "Importe", 40, True)
    Column.Alignment = xtpAlignmentRight


    
    

    Set Column = wndReportControl.Columns.Add(9, "claveLineasExpediente", 0, True)
    Column.Visible = False
    
    
    wndReportControl.PaintManager.MaxPreviewLines = 1
    wndReportControl.PaintManager.HorizontalGridStyle = xtpGridNoLines
                  
    
    
    
    'wndReportControl.PaintManager.VerticalGridStyle = xtpGridSolid
    
    'This code below will add a column to the GroupOrder collection of columns.
    'This will cause the columns in the ReportControl to be grouped by column "COLUMN_FROM",
    'which has an index of zero (0) in the GroupOrder collection. Columns are first grouped
    'in the order that they are added to the GroupOrder collection.
    wndReportControl.GroupsOrder.Add wndReportControl.Columns(1)

    'This will cause the column at index 0 of the GroupOrder collection to be displayed
    'in ascending order.
    wndReportControl.GroupsOrder(0).SortAscending = False
            
  
    
    'Any time you add or delete rows(by removing the attached record), you must call the
    'Populate method so the ReportControl will display the changes.
    'If rows are added, the rows will remain hidden until Populate is called.
    'If rows are deleted, the rows will remain visible until Populate is called.
    wndReportControl.Populate
    
    wndReportControl.SetCustomDraw xtpCustomBeforeDrawRow
End Sub


Private Sub MostrarDatos()
Dim Importe As Currency
Dim Total As Currency
Dim GroupRow As XtremeReportControl.ReportGroupRow
Dim DocumeTacionEntragada As Boolean
    

    Label1.Caption = "Leyendo BD"
    Label1.Refresh
    Screen.MousePointer = vbHourglass
    
   
    
    
    Screen.MousePointer = vbHourglass
    If Val(Check1.Tag) <> Check1.Value Then CreateReportControl
    
    QueID = -1
    
    populateInbox
    wndReportControl.Populate
    
    
    
    If Me.Check1.Value = 1 Then

      Importe = 0
      For I = Me.wndReportControl.Rows.Count - 1 To 0 Step -1
          If wndReportControl.Rows(I).GroupRow Then
              'Es la del grupo
              'Debug.Print ""
              
              
              Set GroupRow = wndReportControl.Rows(I)
              Msg = Mid(GroupRow.GroupCaption, InStr(1, GroupRow.GroupCaption, ":") + 1)
              Msg = DevuelveDesdeBD("If(pagoporbanco=1,'Banco','Caja')", "gestadministrativa", "id", Msg)
              Msg = "    " & Msg & "    " & Format(Importe, FormatoImporte) & "€" & IIf(DocumeTacionEntragada, "   ENTREGADA", "")
              GroupRow.GroupCaption = GroupRow.GroupCaption & Msg
              Importe = 0
              DocumeTacionEntragada = True
          Else
              'Debug.Print wndReportControl.Rows(I).Record.Item(7).Value
              Importe = Importe + wndReportControl.Rows(I).Record.Item(9).Value
              If wndReportControl.Rows(I).Record.Item(2).Value = "" Then DocumeTacionEntragada = False
          End If
      Next I
      
      
      
    End If
    Label1.Caption = ""
    Screen.MousePointer = vbDefault
End Sub



Public Sub populateInbox()
Dim C As String
Dim F As Date


    wndReportControl.Records.DeleteAll
    wndReportControl.Populate
    If Me.Check1.Value = 1 Then
        C = "select l.numexped,l.anoexped,fecexped,licencia,nomclien ,pagado ,nomconce,e.codclien,l.importe"
        C = C & " ,l.codconce ,l.tiporegi,l.numlinea , l.codsitua,l.docuentregada"
        C = C & " from expedientes e,expedientes_lineas l,clientes"
        C = C & " Where pagado>0 and e.tiporegi = L.tiporegi And e.numexped = L.numexped And e.anoexped = L.anoexped And e.CodClien = Clientes.CodClien"
        C = C & " ORDER BY pagado desc"
    Else
        C = "Select * from gestadministrativa order by fechacreacion desc"
    End If
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        If Me.Check1.Value Then
            AddRecordDes
        Else
            AddRecordSin
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    Set miRsAux = Nothing
    
    
    
End Sub

Private Sub AddRecordDes()

Dim Aux As String


  
    Dim Record As ReportRecord
    'Adds a new Record to the ReportControl's collection of records, this record will
    'automatically be attached to a row and displayed with the Populate method
    Set Record = wndReportControl.Records.Add()
    
    Dim Item As ReportRecordItem
    
    
    
    If QueID <> miRsAux!pagado Then
        Situacion = 1  'abierta
        Aux = DevuelveDesdeBD("concat(if(fechafinalizacion  is null,'0','1'),'|',contabilizada,'|')", "gestadministrativa", "ID", miRsAux!pagado)
        If Aux <> "" Then
            If RecuperaValor(Aux, 1) = "1" Then
                If RecuperaValor(Aux, 2) = 1 Then
                    'Contabilizada
                    Situacion = 4 'contab
                Else
                    Situacion = 6 'en proceso
                End If
            End If
        End If
        QueID = miRsAux!pagado
    End If
    'Adds a new ReportRecordItem to the Record, this can be thought of as adding a cell to a row
    
    Set Item = Record.AddItem("")
    If Situacion = 1 Then
        'Assigns an icon to the item
        Item.Icon = Situacion
        'Assigns a GroupCaption to the item, this is displayed int he grouprow when grouped by the column
        'this item belong to.
        Item.GroupCaption = "Abierta"
        'Sets the group priority of the item when grouped, the lower the number the higher the priority,
        'Highest priority is displayed first
        Item.GroupPriority = IMPORTANCE_LOW
        'Sets the sort priority of the item when the colulmn is sorted, the lower the number the higher the priority,
        'Highest priority is sorted displayed first, then by value
        Item.SortPriority = IMPORTANCE_LOW
    ElseIf Situacion = 6 Then
        Item.Icon = Situacion   ' RECORD_IMPORTANCE_LOW_ICON
        Item.GroupCaption = "En proceso"
        Item.GroupPriority = IMPORTANCE_NORMAL
        Item.SortPriority = IMPORTANCE_NORMAL
        
    Else
        Item.Icon = Situacion
        Item.GroupCaption = "Contabilizada"
        
        Item.GroupPriority = IMPORTANCE_HIGH
        Item.SortPriority = IMPORTANCE_HIGH
    End If
    Item.ToolTip = Item.GroupCaption




'    If (Anterior = IMPORTANCE_NORMAL) Then
'        Item.GroupCaption = "Importance: Normal"
'        Item.GroupPriority = IMPORTANCE_NORMAL
'        Item.SortPriority = IMPORTANCE_NORMAL
'    End If
      

    Set Item = Record.AddItem("")
    Item.Value = Format(miRsAux!pagado, "000")
    Item.Caption = ""
    
    Set Item = Record.AddItem("")
    Item.Value = IIf(miRsAux!docuEntregada = 1, "Si", "")
    Item.Caption = IIf(miRsAux!docuEntregada = 1, "Si", "-")
    Item.ToolTip = IIf(miRsAux!docuEntregada = 1, "Entregada", "Pendiente de entregar")
    'If miRsAux!docuentregada = 1 Then Stop
    
    
    Record.AddItem (CStr(miRsAux!licencia))
    Record.AddItem CStr(miRsAux!NomClien) & " (" & miRsAux!CodClien & ")"
    Record.AddItem CStr(miRsAux!codconce)
    Record.AddItem CStr(miRsAux!nomconce)
    

   
    Set Item = Record.AddItem(Year(miRsAux!fecexped) & Format(miRsAux!numexped, "0000000"))
    Item.Caption = Format(miRsAux!numexped, "00000") & "/" & Year(miRsAux!fecexped) - 2000
    
    Set Item = Record.AddItem(Format(miRsAux!fecexped, FormatoFecha))
    Item.Caption = Format(miRsAux!fecexped, "dd/mm/yyyy")
    
    Set Item = Record.AddItem("")
    'Specifys the format that the price will be displayed
    'Item.Format = " %s"
    Item.Format = "%.2f"
    Item.Value = CCur(miRsAux!Importe)
    Item.Caption = Format(Item.Value, FormatoImporte)
    

    
    
    
    
    'Adds the PreviewText to the Record.  PreviewText is the text displayed for the ReportRecord while in PreviewMode
    'Record.PreviewText = miRsAux!NomClien
    
End Sub


Private Sub AddRecordSin()
Dim Situacion As Byte
Dim Aux As String
    
  
    Dim Record As ReportRecord
    Set Record = wndReportControl.Records.Add()
    
    Dim Item As ReportRecordItem
    
    
    
    'Adds a new ReportRecordItem to the Record, this can be thought of as adding a cell to a row
    Set Item = Record.AddItem("")
    Situacion = 1
    If Not IsNull(miRsAux!fechafinalizacion) Then
        If miRsAux!contabilizada = 0 Then
            Situacion = 6
        Else
            Situacion = 4
        End If
    End If
    
    If Situacion = 1 Then
        'Assigns an icon to the item
        Item.Icon = Situacion
        Item.ToolTip = "Abierta"
        
    ElseIf Situacion = 6 Then
        Item.Icon = Situacion   ' RECORD_IMPORTANCE_LOW_ICON
        Item.ToolTip = "En proceso"
       
    Else
        Item.Icon = Situacion
        Item.ToolTip = "Contabilizada"
    End If




      

      
    Record.AddItem ("")
    Set Item = Record.AddItem(Format(miRsAux!id, "0000"))
    Item.Value = Val(miRsAux!id)
    
    Set Item = Record.AddItem("")
    Item.Caption = Format(miRsAux!fechacreacion, "dd/mm/yyyy hh:nn")
    Item.Value = Format(miRsAux!fechacreacion, "yyyymmddhhnnss")
    
    
    Record.AddItem CStr(miRsAux!Usuario)
    
    Record.AddItem CStr(IIf(DBLet(miRsAux!pagoporbanco, "N") = 1, "SI", " "))
    
    
    
    Set Item = Record.AddItem("")
    
    If Not IsNull(miRsAux!fechafinalizacion) Then
        Item.Caption = Format(miRsAux!fechacreacion, "dd/mm/yyyy hh:nn")
        Item.Value = Format(miRsAux!fechacreacion, "yyyymmddhhnnss")
    End If
    Record.AddItem CStr(miRsAux!llevados)
    Set Item = Record.AddItem(Format(miRsAux!Importe, FormatoImporte))
    Item.Value = miRsAux!Importe
    

    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim B As Boolean
Dim id As Long
Dim GroupRow
    B = False
    If Button.Index <> 1 Then
        'Si esta "cerrado" ya no puedo hacer nada
        
        If Me.wndReportControl.Records.Count = 0 Then Exit Sub
        If wndReportControl.SelectedRows.Count = 0 Then Exit Sub
        
        'Es un agrupado
        If Me.Check1.Value = 1 Then
            If wndReportControl.SelectedRows(0).GroupRow Then
                Set GroupRow = wndReportControl.Rows(wndReportControl.SelectedRows(0).Index)
                Msg = GroupRow.GroupCaption
                Msg = Trim(Mid(Msg, 4))
                Msg = Trim(Mid(Msg, 1, InStr(1, Msg, " ")))
                id = Msg
            Else
                id = wndReportControl.SelectedRows(0).Record(1).Caption
            End If
        Else
            id = wndReportControl.SelectedRows(0).Record(2).Caption
        End If
        
        If Button.Index <> 8 Then
            'Button.Index = 2 Or Button.Index = 3
            Msg = DevuelveDesdeBD("fechafinalizacion", "gestadministrativa", "id", CStr(id))
            If Msg <> "" Then
                MsgBox "No se puede eliminar/modificar. Ya esta cerrado el proceso", vbExclamation
                Exit Sub
            End If
        End If
    End If
    
    Select Case Button.Index
    Case 1
        CadenaDesdeOtroForm = ""
        frmMensajes.Parametros = ""
        frmMensajes.Opcion = 3
        frmMensajes.Show vbModal
        If CadenaDesdeOtroForm <> "" Then
            B = True
            ImprimirProceso CLng(CadenaDesdeOtroForm)
        End If
    Case 2
        
        CadenaDesdeOtroForm = ""
        frmMensajes.Parametros = CStr(id)
        frmMensajes.Opcion = 3
        frmMensajes.Show vbModal
        If CadenaDesdeOtroForm <> "" Then B = True
    Case 3
        'If EliminarProceso_(id) Then B = True
         
         If EliminarProcoso(id) Then B = True
    Case 8
        
        ImprimirProceso id
    End Select
    If B Then MostrarDatos
End Sub

Private Sub ImprimirProceso(idd As Long)
    InicializarVbles True
    cadNomRPT = "rGestionAdm.rpt"
    cadFormula = "{expedientes_lineas.pagado}=" & idd
    Msg = DevuelveDesdeBD("concat(usuario,'|',fechacreacion,'|',coalesce(fechafinalizacion,''),'|',contabilizada,'|')", "gestadministrativa", "id", CStr(idd))
    cadselect = "ID " & Format(idd, "0000") & "   Usuario: " & RecuperaValor(Msg, 1) & "   Fecha: "
    cadselect = cadselect & Format(RecuperaValor(Msg, 2), "dd/mm/yyyy")
    MsgErr = RecuperaValor(Msg, 3)
    If MsgErr <> "" Then
        cadselect = cadselect & "     CERRADO: " & MsgErr
        If RecuperaValor(Msg, 4) = "1" Then cadselect = cadselect & " (Contab.)"
    End If
    cadParam = cadParam & "pdh1=""" & cadselect & """|"
    MsgErr = ""

    ImprimeGeneral
End Sub


Private Function EliminarProcoso(idd As Long) As Boolean
    EliminarProcoso = False
    
    Msg = DevuelveDesdeBD("concat(usuario,'|',fechacreacion,'|',coalesce(fechafinalizacion,''),'|')", "gestadministrativa", "id", CStr(idd))
    cadselect = vbCrLf & "ID " & Format(idd, "0000") & vbCrLf & "Usuario: " & RecuperaValor(Msg, 1) & vbCrLf & "Fecha: "
    cadselect = cadselect & RecuperaValor(Msg, 2) & vbCrLf
    
    
    Msg = "Va a eliminar el proceso. " & cadselect & vbCrLf & "¿Continuar?"
    If MsgBox(Msg, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
        
    Msg = "UPDATE expedientes_lineas set  codsitua=0, pagado =0 where "
    Msg = Msg & " pagado =" & idd
    Conn.Execute Msg

    
    Msg = "DELETE FROM  gestadministrativa  WHERE id =" & idd
    Conn.Execute Msg
    EliminarProcoso = True
End Function

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim id As String
Dim F As Date
Dim GroupRow
Dim B As Boolean
Dim PorBanco As Byte
Dim PuedoEntregarDocumentos As Boolean

    'Es un agrupado
    If Me.wndReportControl.Records.Count = 0 Then Exit Sub
    If wndReportControl.SelectedRows.Count = 0 Then Exit Sub
    If Me.Check1.Value = 1 Then
        If wndReportControl.SelectedRows(0).GroupRow Then
            Set GroupRow = wndReportControl.Rows(wndReportControl.SelectedRows(0).Index)
            Msg = GroupRow.GroupCaption
            Msg = Trim(Mid(Msg, 4))
            Msg = Trim(Mid(Msg, 1, InStr(1, Msg, " ")))
            id = Msg
        Else
            id = wndReportControl.SelectedRows(0).Record(1).Caption
        End If
    Else
        id = wndReportControl.SelectedRows(0).Record(2).Caption
    End If


    


    Set miRsAux = New ADODB.Recordset
    Msg = "Select fechafinalizacion,contabilizada,pagoPorBanco,quebanco from gestadministrativa where id =" & id
    miRsAux.Open Msg, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Msg = ""
    PuedoEntregarDocumentos = False
    If Val(miRsAux!contabilizada) = 1 Then
        Msg = "Ya esta contabilizada"
        If Button.Index = 3 Then PuedoEntregarDocumentos = True
        
    Else
        If IsNull(miRsAux!fechafinalizacion) Then
            'No esta cerrada, todavia
            If Button.Index = 2 Then Msg = "Falta cerrar proceso"
        Else
            If Button.Index = 1 Then Msg = "Ya esta cerrado el proceso"
            If Button.Index = 3 Then PuedoEntregarDocumentos = True
        End If
    End If
    If miRsAux!pagoporbanco = 0 Then
        PorBanco = 101
    Else
        PorBanco = miRsAux!QueBanco
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    
    If Button.Index = 3 Then
        If Not PuedoEntregarDocumentos Then
            MsgBox "Debe estar cerrada la gestion administrativa", vbExclamation
            Exit Sub
        End If
        If LCase(vUsu.Login) <> "root" And LCase(vUsu.Login) <> "antonio" Then
            MsgBox "NO puede realizar la entrega de documentos desde aqui." & vbCrLf & vbCrLf & "Vaya al mantenimiento de clientes/socios", vbExclamation
            Exit Sub
        End If
        'Entrega masiva documentos
        
        frmGestionAdmCierre.QueTasa_ = CLng(id)
        frmGestionAdmCierre.EntregaDocumentos2 = 1
        frmGestionAdmCierre.Show vbModal
        MostrarDatos
        Exit Sub
    End If
    
    If Msg <> "" Then
        MsgBox Msg, vbExclamation
        Exit Sub
    End If
    
    CadenaDesdeOtroForm = ""
    If Button.Index = 1 Then
        frmGestionAdmCierre.QueTasa_ = CLng(id)
        frmGestionAdmCierre.EntregaDocumentos2 = 0
        frmGestionAdmCierre.Show vbModal
        If CadenaDesdeOtroForm = "" Then Exit Sub
    Else
        'Lanzamos el frm para pedir fecha
        ' Sin cancelar es que no quiere continuar
        If PorBanco < 100 Then
            CadenaDesdeOtroForm = PorBanco
        Else
            CadenaDesdeOtroForm = "NO"
        End If
        frmMensajes.Opcion = 9
        frmMensajes.Show vbModal
        If CadenaDesdeOtroForm = "" Then Exit Sub
        F = CDate(RecuperaValor(CadenaDesdeOtroForm, 1))
        CadenaDesdeOtroForm = RecuperaValor(CadenaDesdeOtroForm, 2)
        If PorBanco = 1 Then
            If CadenaDesdeOtroForm = "2" Then
                CadenaDesdeOtroForm = vParam.CtaBanco2
            Else
                CadenaDesdeOtroForm = vParam.CtaBanco
            End If
            
                
                
        Else
            CadenaDesdeOtroForm = ""
        End If
        
    End If
    
    
    
    If Button.Index <> 1 Then
        'Se cierra en el fomulario
        'B = CerrarProceso(id)
        Conn.BeginTrans
        B = HacerProcesContabilizacionGestion(CLng(id), F, CadenaDesdeOtroForm)
    
        If B Then
            Conn.CommitTrans
        Else
            Conn.RollbackTrans
        End If
    Else
        B = True
    End If
    
    If B Then MostrarDatos
    
End Sub



Private Function CerrarProceso(id As String) As Boolean
    On Error GoTo eCerrarProceso
    Msg = "UPDATE expedientes_lineas set  codsitua=2 where "
    Msg = Msg & " pagado =" & id
    Conn.Execute Msg

    
    
    Msg = "UPDATE gestadministrativa  set fechafinalizacion=now()  where id= " & id
    Conn.Execute Msg
    
    CerrarProceso = True
    Exit Function
eCerrarProceso:
    MuestraError Err.Number, Err.Description
    
End Function





Private Function HacerProcesContabilizacionGestion(id As Long, Fecha As Date, QueBanco As String) As Boolean
Dim ColApu As Collection
Dim Aux As String
Dim aux2 As String
Dim CtaCajaBanco As String
Dim ImporteParaCajaBanco As Currency
Dim EsACaja As Boolean
Dim usuario2 As String
Dim Creacion As Date
Dim SegundoAux As String
On Error GoTo eHacerProcesContabilizacionGestion
    HacerProcesContabilizacionGestion = False
    
    
       
    
        
    'Iremos uno a uno cuadrando todos los movimientos
    Set ColApu = New Collection
    Set miRsAux = New ADODB.Recordset
    
    Aux = "Select usuario,llevados,importe,fechafinalizacion,pagoPorBanco,fechacreacion from gestadministrativa where id =" & id
    miRsAux.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'no puede ser eof
    EsACaja = miRsAux!pagoporbanco = 0
    usuario2 = miRsAux!Usuario
    Creacion = miRsAux!fechacreacion
    miRsAux.Close
    
    
    'Si es a caja, no dejare meter con fecha anterior a cierre
    If EsACaja Then
        Aux = "Select max(feccaja) from caja_param WHERE usuario =" & DBSet(usuario2, "T")
        miRsAux.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Aux = ""
        If Not miRsAux.EOF Then
            If Not IsNull(miRsAux.Fields(0)) Then
                If miRsAux.Fields(0) >= Fecha Then Err.Raise 513, , "Fecha cierre de caja(" & usuario2 & ") mayor fecha contabilizacion"
            End If
        End If
        miRsAux.Close
        
    End If
    ImporteParaCajaBanco = 0
        '' Llevara
        '       codmacta | docum | codconce | ampliaci | imported|importeH |
    
    Aux = "select pagado, l.numexped,l.anoexped,l.nomconce,l.numserie,e.codclien,l.importe ,l.codconce ,l.tiporegi,l.numlinea , l.codsitua,codmacta,codclien "
    Aux = Aux & " from expedientes e,expedientes_lineas l,conceptos Where"
    Aux = Aux & " pagado=" & id & " and e.tiporegi = L.tiporegi And e.numexped = L.numexped And e.anoexped = L.anoexped And"
    Aux = Aux & " conceptos.codconce=l.codconce"

    
    miRsAux.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    aux2 = ""
    While Not miRsAux.EOF
        I = I + 1
        
        If IsNull(miRsAux!codmacta) Then Err.Raise 513, , "Cuenta nula para concepto: " & miRsAux!nomconce
        
        If EsACaja Then
            
        
            'Meteremos en la caja del usuario. Guardamos los SQLs
            'usuario,feccaja,tipomovi,tiporegi,numserie,numdocum,anoexped,importe,ampliacion,codmacta,numlinea
            Aux = "(" & DBSet(usuario2, "T") & ",'" & Format(Fecha, "yyyy-mm-dd") & " " & Format(Now, "hh:mm:") & Format(I, "00") & "',1," 'pago:1
            Aux = Aux & DBSet(miRsAux!TipoRegi, "T") & "," & DBSet(miRsAux!numSerie, "T") & "," & DBSet(miRsAux!numexped, "T") & "," & DBSet(miRsAux!anoexped, "N")
            Aux = Aux & "," & DBSet(miRsAux!Importe, "N") & "," & DBSet(miRsAux!nomconce, "T") & ",'" & miRsAux!codmacta & "'," & miRsAux!numlinea & ")"
            
            
        Else
            'Va a banco,. directamente a contabilidad
            'ImporteParaCajaBanco = ImporteParaCajaBanco + miRsAux!Importe
            
            
            '   col: 'codmacta | docum | codconce | ampliaci | imported|importeH |ctacontrpar|
        
            'aux2 = DevuelveCuentaContableCliente(False, CStr(miRsAux!CodClien))
            'Nov 2017
            aux2 = miRsAux!codmacta
            
            
            Aux = miRsAux!numSerie & Format(miRsAux!numexped, "0000") & "/" & miRsAux!anoexped - 2000
            
            '   col: 'codmacta | docum | codconce |
            SegundoAux = QueBanco & "|" & Aux & "|1|"
            Aux = aux2 & "|" & Aux & "|1|"
            
            
            aux2 = DevuelveDesdeBD("nomclien", "clientes", "codclien", miRsAux!CodClien)
            
            'Para el contra apunte
            SegundoAux = SegundoAux & aux2 & "||" & miRsAux!Importe & "|" & miRsAux!codmacta & "|"
            
            'Acabo de montar el apunte de la cuota
            '   ampliaci
            Aux = Aux & aux2 & "|"
            ' imported|importeH |ctacontrpar|
            Aux = Aux & miRsAux!Importe & "|" & "|" & QueBanco & "|"
    
            
            
    
        End If
        ColApu.Add Aux
        
        'Contra apunte banco
        If Not EsACaja Then ColApu.Add SegundoAux   'el contrapunte banco
    
        
        
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If ColApu.Count = 0 Then Err.Raise 513, , "No se han encontrado lineas"
    'Cuadre de /banco
    
    
    
    If EsACaja Then
        'Directamente metemos el pago desde la caja
        Stop  'no deberia
        Aux = ""
        For I = 1 To ColApu.Count
            Aux = Aux & ", " & ColApu(I)
        Next
        Aux = Mid(Aux, 2)
        Aux = "INSERT INTO caja (usuario,feccaja,tipomovi,tiporegi,numserie,numdocum,anoexped,importe,ampliacion,codmacta,numlinea) VALUES " & Aux
        Conn.Execute Aux
        
    Else
    
        'Apunte linea-banco, linea -banco
        'aux2 = "Gestion administrativa " & Format(id, "0000") & "  F" & Format(Creacion, "dd/mm/yyyy hh:nn")
        'Aux = "GA: " & id & "-" & Format(ColApu.Count, "00")
        'Aux = QueBanco & "|" & Aux & "|1|" & aux2 & "||"
        '' imported|importeH |ctacontrpar|
        'Aux = Aux & Format(ImporteParaCajaBanco, FormatoImporte) & "|" & "|"
        'ColApu.Add Aux
        
        Aux = "Gestion administrativa " & Format(id, "0000") & "  Fecha " & Format(Creacion, "dd/mm/yyyy hh:nn") & vbCrLf
        Aux = Aux & "Usuario gestion:" & usuario2 & "   Tasas: " & Format(ColApu.Count, "00") & vbCrLf
        Aux = Aux & "Usuario actual :" & vUsu.Login
        If Not CrearApunteDesdeColeccion(Fecha, Aux, ColApu) Then Err.Raise 513, , "Crear apunte en contabilidad"
        
        
    End If
    
    Aux = "UPDATE expedientes_lineas SET codsitua = 3 WHERE pagado=" & id
    Conn.Execute Aux
    
    Aux = "UPDATE gestadministrativa SET contabilizada = 1 "
    If QueBanco <> "" Then
        If QueBanco = vParam.CtaBanco Then Aux = Aux & " , quebanco=1"
        If QueBanco = vParam.CtaBanco2 Then Aux = Aux & " , quebanco=2"
    End If
    
    Aux = Aux & " Where id = " & id
    Conn.Execute Aux
    
    
    HacerProcesContabilizacionGestion = True
eHacerProcesContabilizacionGestion:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
    Set ColApu = Nothing
End Function

