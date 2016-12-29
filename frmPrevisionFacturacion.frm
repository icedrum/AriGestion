VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#17.2#0"; "Codejock.ReportControl.v17.2.0.ocx"
Begin VB.Form frmPrevisionFacturacion 
   Caption         =   "Facturas periodicas"
   ClientHeight    =   11040
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16425
   LinkTopic       =   "Form1"
   ScaleHeight     =   11040
   ScaleWidth      =   16425
   StartUpPosition =   2  'CenterScreen
   Begin XtremeReportControl.ReportControl wndReportControl 
      Height          =   9615
      Left            =   120
      TabIndex        =   8
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
      TabIndex        =   7
      Top             =   0
      Width           =   16215
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
         ForeColor       =   &H00000000&
         Height          =   360
         Index           =   1
         Left            =   9960
         MaxLength       =   30
         TabIndex        =   4
         Tag             =   "Fecha|F|N|||factcli|fecfactu|dd/mm/yyyy|N|"
         Top             =   450
         Width           =   1395
      End
      Begin VB.CommandButton cmdFacturar 
         Caption         =   "Facturar"
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
         Left            =   15120
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   12960
         MaxLength       =   10
         TabIndex        =   13
         Tag             =   "Num albar|N|N|||Expedientes|numexped|00000|S|"
         Top             =   450
         Width           =   1395
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         ItemData        =   "frmPrevisionFacturacion.frx":0000
         Left            =   7080
         List            =   "frmPrevisionFacturacion.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   480
         Width           =   2655
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         ItemData        =   "frmPrevisionFacturacion.frx":0004
         Left            =   240
         List            =   "frmPrevisionFacturacion.frx":0014
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   480
         Width           =   2175
      End
      Begin VB.CommandButton cmdVerdatos 
         Caption         =   "Previsión"
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
         Left            =   11520
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         ItemData        =   "frmPrevisionFacturacion.frx":004D
         Left            =   5040
         List            =   "frmPrevisionFacturacion.frx":005E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         ItemData        =   "frmPrevisionFacturacion.frx":0089
         Left            =   2640
         List            =   "frmPrevisionFacturacion.frx":0096
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label1 
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
         Height          =   240
         Left            =   9960
         TabIndex        =   15
         Top             =   240
         Width           =   600
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   0
         Left            =   10800
         Picture         =   "frmPrevisionFacturacion.frx":00B2
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   12960
         TabIndex        =   14
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Filtro"
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
         Index           =   2
         Left            =   7080
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.Line Line1 
         X1              =   2520
         X2              =   2520
         Y1              =   240
         Y2              =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ver agrupado por"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   1665
      End
      Begin VB.Label Label3 
         Caption         =   "Periodo facturacion"
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
         Index           =   0
         Left            =   5040
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo facturación"
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
         Index           =   4
         Left            =   2640
         TabIndex        =   9
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmPrevisionFacturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Dim CargandoDatos As Boolean
Dim AmpliacionesFacturacion As String
Dim ctacontabanco As String   'cuenta bancaria en contabilidad 572...

'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'
' Cargar datos
'
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
Public Sub CreateReportControl()
    'Start adding columns
    Dim Column As ReportColumn
    
    wndReportControl.Columns.DeleteAll
    
    'Adds a new ReportColumn to the ReportControl's collection of columns, growing the collection by 1.
    Set Column = wndReportControl.Columns.Add(COLUMN_IMPORTANCE, "Anterior", 18, False)
    'The value assigned to the icon property corresponds to the index of an icon in the collection of wndReportControl.Icons
    'I.e. The icon at index=1 in the collection will be displayed in the column header.  The index of the icon depends on the
    'order it is added to the collection.  (Icons are added after the records near the bottom of the Form_Load)
    Column.Icon = COLUMN_IMPORTANCE_ICON
    
    Set Column = wndReportControl.Columns.Add(2, "Cod.Cli", 30, True)
    Set Column = wndReportControl.Columns.Add(3, "Nombre", 140, True)
    Set Column = wndReportControl.Columns.Add(4, "Conce.", 25, True)
    Set Column = wndReportControl.Columns.Add(5, "Descripción", 125, True)
    Set Column = wndReportControl.Columns.Add(6, "Licencia", 40, True)
    Set Column = wndReportControl.Columns.Add(7, "Fecha", 35, True)
    Set Column = wndReportControl.Columns.Add(8, "Importe", 40, True)
    Column.Alignment = xtpAlignmentRight
    Set Column = wndReportControl.Columns.Add(9, "CodigoIVA", 0, True)
    Column.Visible = False
    Set Column = wndReportControl.Columns.Add(10, "tipoconcepto", 0, True)
    Column.Visible = False
    
    
    wndReportControl.PaintManager.MaxPreviewLines = 1
    wndReportControl.PaintManager.HorizontalGridStyle = xtpGridNoLines
                  
    
    
    
    'wndReportControl.PaintManager.VerticalGridStyle = xtpGridSolid
    
    'This code below will add a column to the GroupOrder collection of columns.
    'This will cause the columns in the ReportControl to be grouped by column "COLUMN_FROM",
    'which has an index of zero (0) in the GroupOrder collection. Columns are first grouped
    'in the order that they are added to the GroupOrder collection.
    wndReportControl.GroupsOrder.Add wndReportControl.Columns(Combo1(2).ItemData(Combo1(2).ListIndex))

    'This will cause the column at index 0 of the GroupOrder collection to be displayed
    'in ascending order.
    wndReportControl.GroupsOrder(0).SortAscending = True
            
  
    
    'Any time you add or delete rows(by removing the attached record), you must call the
    'Populate method so the ReportControl will display the changes.
    'If rows are added, the rows will remain hidden until Populate is called.
    'If rows are deleted, the rows will remain visible until Populate is called.
    wndReportControl.Populate
    
    wndReportControl.SetCustomDraw xtpCustomBeforeDrawRow
End Sub

Public Sub populateInbox()
Dim C As String
Dim F As Date


    wndReportControl.Records.DeleteAll
    CargandoDatos = True
    
    If Combo1(0).ListIndex = 1 Then
        C = "clientes_laboral"
    ElseIf Combo1(0).ListIndex = 2 Then
        C = "clientes_fiscal"
    Else
        C = "clientes_cuotas"
    End If
    
        
    C = "select t.codclien,nomclien,t.codconce,nomconce,fecultfac,importe,codigiva,tipoconcepto,licencia FROM  " & C
    C = C & " as t,clientes,conceptos where t.codconce=conceptos.codconce and"
    C = C & " t.codclien=clientes.codclien and  periodicidad = " & Combo1(1).ListIndex '0 Men, 1 Trim  2 Semes 3 anual
    I = Combo1(1).ItemData(Combo1(1).ListIndex)
    F = CDate(Text1(1).Text) 'fercha facturacion
    F = DateAdd("m", -I, F)
    C = C & " and (fecultfac is null or fecultfac <=" & DBSet(F, "F") & ")"
    If Combo1(3).ListIndex > 0 Then
        'Lleva filtro
        C = C & " AND t.codconce= " & Combo1(3).ItemData(Combo1(3).ListIndex)
    End If
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        AddRecord
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    Set miRsAux = Nothing
    
    CargandoDatos = False
    
End Sub

Private Sub AddRecord()
Dim Anterior As Boolean
Dim Aux As String
    
  
    Dim Record As ReportRecord
    'Adds a new Record to the ReportControl's collection of records, this record will
    'automatically be attached to a row and displayed with the Populate method
    Set Record = wndReportControl.Records.Add()
    
    Dim Item As ReportRecordItem
    
    Anterior = Not IsNull(miRsAux!fecultfac)
    
    'Adds a new ReportRecordItem to the Record, this can be thought of as adding a cell to a row
    Set Item = Record.AddItem("")
    If Not Anterior Then
        'Assigns an icon to the item
        Item.Icon = RECORD_IMPORTANCE_HIGH_ICON
        'Assigns a GroupCaption to the item, this is displayed int he grouprow when grouped by the column
        'this item belong to.
        Item.GroupCaption = "Facturados"
        'Sets the group priority of the item when grouped, the lower the number the higher the priority,
        'Highest priority is displayed first
        Item.GroupPriority = IMPORTANCE_HIGH
        'Sets the sort priority of the item when the colulmn is sorted, the lower the number the higher the priority,
        'Highest priority is sorted displayed first, then by value
        Item.SortPriority = IMPORTANCE_HIGH
    Else
        Item.Icon = RECORD_IMPORTANCE_LOW_ICON
        Item.GroupCaption = "Nuevos"
        Item.GroupPriority = IMPORTANCE_LOW
        Item.SortPriority = IMPORTANCE_LOW
    End If





'    If (Anterior = IMPORTANCE_NORMAL) Then
'        Item.GroupCaption = "Importance: Normal"
'        Item.GroupPriority = IMPORTANCE_NORMAL
'        Item.SortPriority = IMPORTANCE_NORMAL
'    End If
      
      

      
    Record.AddItem ("")
    Set Item = Record.AddItem(CStr(miRsAux!codclien))
    Item.Value = Val(miRsAux!codclien)
    
    Record.AddItem CStr(miRsAux!NomClien)
    Record.AddItem CStr(miRsAux!codconce)
    Record.AddItem CStr(miRsAux!nomconce)
    Record.AddItem DBLet(miRsAux!licencia, "T")
    
    Set Item = Record.AddItem("")

    If IsNull(miRsAux!fecultfac) Then
        Item.Caption = "-"
    Else
        GetDate Item
    End If
    
    
    Set Item = Record.AddItem("")
    'Specifys the format that the price will be displayed
    'Item.Format = " %s"
    Item.Format = "%.2f"
    Item.Value = CCur(miRsAux!Importe)
    Item.Caption = Format(Item.Value, FormatoImporte)
    'Assigns the properties based on the value of Price
    Select Case miRsAux!Importe
        Case Is <= 5:
                Item.GroupCaption = "Importe < 5"
                Item.GroupPriority = 1
        Case Is <= 20:
                Item.GroupCaption = "Importe 5-20"
                Item.GroupPriority = 0
        Case Is > 20:
                Item.GroupCaption = "Importe > 20"
                Item.GroupPriority = 3
    End Select
    
    Record.AddItem CStr(miRsAux!codigiva)
    Record.AddItem CStr(miRsAux!tipoconcepto)
    
    
    'Adds the PreviewText to the Record.  PreviewText is the text displayed for the ReportRecord while in PreviewMode
    Record.PreviewText = miRsAux!NomClien
    
End Sub

Private Sub MostrarDatos()
Dim Importe As Currency
Dim Total As Currency
Dim GroupRow As XtremeReportControl.ReportGroupRow



    If Me.Combo1(0).ListIndex < 0 Then Exit Sub
    If Me.Combo1(1).ListIndex < 0 Then Exit Sub
    If Text1(1).Text = "" Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    CargandoDatos = True
    
    
    Screen.MousePointer = vbHourglass
    If Combo1(2).Tag <> Combo1(2).ListIndex Then
        Combo1(2).Tag = Combo1(2).ListIndex
        CreateReportControl
    End If
    populateInbox
    wndReportControl.Populate
    
    
    Total = 0
    Importe = 0
    For I = Me.wndReportControl.Rows.Count - 1 To 0 Step -1
        If wndReportControl.Rows(I).GroupRow Then
            'Es la del grupo
            'Debug.Print ""
            Set GroupRow = wndReportControl.Rows(I)
            GroupRow.GroupCaption = GroupRow.GroupCaption & "    " & Format(Importe, FormatoImporte) & "€"
            Importe = 0
        Else
            'Debug.Print wndReportControl.Rows(I).Record.Item(7).Value
            Importe = Importe + wndReportControl.Rows(I).Record.Item(8).Value
            Total = Total + wndReportControl.Rows(I).Record.Item(8).Value
        End If
    Next I
    
    If Total = 0 Then
        txtTotal.Text = ""
    Else
        txtTotal.Text = Format(Total, FormatoImporte)
    End If
    
    
    CargandoDatos = False
    Screen.MousePointer = vbDefault
End Sub







Private Sub cmdVerdatos_Click()
    MostrarDatos
End Sub

Private Sub CargaFiltro()
Dim C As String
            
        Combo1(3).Clear
        Combo1(3).AddItem "Sin filtro"
        If Combo1(0).ListIndex >= 0 Then
            If Combo1(0).ListIndex = 0 Then
                C = "1,2"
            ElseIf Combo1(0).ListIndex = 1 Then
                C = 3
            Else
                C = 4
            End If
            Set miRsAux = New ADODB.Recordset
            C = "Select codconce,nomconce from conceptos where tipoconcepto IN (" & C & ")"
            miRsAux.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                Combo1(3).AddItem miRsAux!nomconce
                Combo1(3).ItemData(Combo1(3).NewIndex) = miRsAux!codconce
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            Set miRsAux = Nothing
        End If
        Combo1(3).ListIndex = 0
End Sub

Private Sub Combo1_Click(Index As Integer)
    If Index = 0 Then
        CargaFiltro
    End If
End Sub

Private Sub cmdFacturar_Click()
Dim F2 As Date
    'Algunas comprobaciones
    If Text1(1).Text = "" Then Exit Sub
    If Me.wndReportControl.Records.Count = 0 Then Exit Sub
        
    'Que la fecha de factura este dentro de los ejercicios contables
    
    
    'Si tiene fecha ult factura, comprobar algo ms
       
    
    
    
    'Ultima comprobacion. Cuenta banco en contabilidad
    ctacontabanco = vParam.BancoPropioFacturacionContabilidad()
    If ctacontabanco = "" Then
        MsgBox "No existe cuenta banco en contabilidad", vbExclamation
        Exit Sub
    End If
    
    'Todo ha ido bien
    'Facturamos a inicio facturacion
    '0 Men, 1 Trim  2 Semes 3 anual
    If Combo1(1).ListIndex = 0 Then
        AmpliacionesFacturacion = Format(Text1(1).Text, "mmmm")
        J = 1
        
    ElseIf Combo1(1).ListIndex = 1 Then
        I = ((Month(Text1(1).Text) - 1) \ 3) + 1
        If (I Mod 2) = 0 Then
            AmpliacionesFacturacion = "o"
        Else
            AmpliacionesFacturacion = "er"
        End If
        AmpliacionesFacturacion = I & AmpliacionesFacturacion & "  trimestre"
        J = 3
    ElseIf Combo1(1).ListIndex = 0 Then
        I = Month(Text1(1).Text)
        If I < 7 Then
            AmpliacionesFacturacion = "Primer "
        Else
            AmpliacionesFacturacion = "Segundo "
        End If
        AmpliacionesFacturacion = AmpliacionesFacturacion & " semestre"
        J = 6
    Else
        AmpliacionesFacturacion = "Cuota anual "
        J = 12
    End If
    AmpliacionesFacturacion = AmpliacionesFacturacion & " " & Year(Text1(1).Text)
    
    Msg = "Facturacion: " & AmpliacionesFacturacion & vbCrLf & vbCrLf
    Msg = Msg & "Fecha: " & Text1(1).Text & vbCrLf
    Msg = Msg & "Nºfacturas: " & Me.wndReportControl.Rows.Count & vbCrLf
    Msg = Msg & "Total: " & txtTotal.Text & vbCrLf & vbCrLf & "¿Continuar?"
    
    
    
    'Metemos la observacione que sera Fecha periodo desde-hasta
    AmpliacionesFacturacion = AmpliacionesFacturacion & "|"
    F2 = CDate(Text1(1).Text)
    F2 = DateAdd("m", J, F2)
    F2 = DateAdd("d", -1, F2)
    AmpliacionesFacturacion = AmpliacionesFacturacion & Text1(1).Text & " al " & Format(F2, "dd/mm/yyyy") & "|"
    
    
    
    
    
    
    
    If MsgBox(Msg, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    If BloqueoManual(Me.Name, "1") Then
        Screen.MousePointer = vbHourglass
        
        
        If HacerFacturacion Then
            Unload Me
        Else
            MostrarDatos
        End If
        
        
        Screen.MousePointer = vbDefault
        DesBloqueoManual Me.Name
    End If
        
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub Form_Load()
    Me.Icon = frmppal.Icon
    wndReportControl.Icons = ReportControlGlobalSettings.Icons
    Combo1(2).ListIndex = 2
    Combo1(2).Tag = Combo1(2).ListIndex
    CreateReportControl
    
    
    '
    Dim TextFont As StdFont
    Set TextFont = Label3(0).Font
    TextFont.SIZE = 10
    Set wndReportControl.PaintManager.TextFont = TextFont
    Combo1_Click 0
    
    Me.Text1(1).Text = Format(Now, "dd/mm/yyyy")

End Sub

Private Sub Form_Resize()
On Error Resume Next
    Frame1.Width = Me.Width - 240
    wndReportControl.Move 60, Me.Frame1.Height + 120, Me.Width - 320, Me.Height - Me.Frame1.Height - 120
    
    Err.Clear
End Sub


Public Sub GetDate(ByVal Item As ReportRecordItem)
    
    'Assign the DateTime string to the value of the ReportRecordItem
 ', ,
    Item.Value = Format(DatePart("d", miRsAux!fecultfac), "00") & "/" & Format(DatePart("m", miRsAux!fecultfac), "00") & "/" & DatePart("yyyy", miRsAux!fecultfac)  '& " " & Hour & ":" & Minute & " " & TimeOfDay
End Sub


Private Sub frmF_Selec(vFecha As Date)
    Text1(1).Text = Format(vFecha, formatoFechaVer)
End Sub

Private Sub imgppal_Click(Index As Integer)
    
    
    Set frmF = New frmCal
    frmF.Fecha = Now
    If Me.Text1(1).Text <> "" Then frmF.Fecha = Text1(1).Text
    frmF.Show vbModal
    Set frmF = Nothing
    

End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), 3
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    'Solo hay index=1
    If Text1(Index).Text <> "" Then
        If Not EsFechaOK(Text1(Index)) Then
            MsgBox "Fecha incorrecta", vbExclamation
            Text1(Index).Text = ""
            PonFoco Text1(Index)
            
        End If
    End If


End Sub


'--------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------
'
'   FACTURACION
'
'--------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------
Private Function HacerFacturacion() As Boolean
Dim CuotaAnterior As Integer
Dim DatosIVA As String   'codigiva|porceiva|
Dim AlgunFallo As Boolean
   
    


    AlgunFallo = False
    Set miRsAux = New ADODB.Recordset
    CuotaAnterior = -1
    'Vamos p'alla
    For I = 1 To Me.wndReportControl.Rows.Count - 1
        
        'Para cada cuota crearemos su factura
        Conn.BeginTrans
        If GenerarFacturaItem(CuotaAnterior, DatosIVA) Then
            Conn.CommitTrans
        Else
            AlgunFallo = True
            Conn.RollbackTrans
        End If
        
        
    
    Next
    Set miRsAux = Nothing
    HacerFacturacion = Not AlgunFallo
End Function


'El I ya lo lleva(la variable "i")
'Dim DatosIVA As String   'codigiva|porceiva|
Private Function GenerarFacturaItem(IVA_Anterior As Integer, DatosIVA As String) As Boolean
Dim rsContador As ADODB.Recordset
Dim Cad As String
Dim TipoContador As Integer
Dim Aux As Currency
Dim ImporIVA As Currency
Dim ImporRecar As Currency
Dim IVA As Currency
Dim IvaRecar As Currency
Dim Bases As Currency
Dim tabla As String
Dim EsUNaCuota As Boolean

    On Error GoTo eGenerarFacturaItem
    GenerarFacturaItem = False
    
    
        
        
    'Veamos la datos del iva de la cuota
    Cad = wndReportControl.Rows(I).Record(9).Caption
    If IVA_Anterior <> CInt(Cad) Then
        'Ha cambiad loa cuota. Leemos los datos del IVA
        Cad = "select * from ariconta" & vParam.Numconta & ".tiposiva WHERE codigiva = " & Cad
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        DatosIVA = miRsAux!porceiva & "|" & miRsAux!porcerec & "|"
        IVA_Anterior = miRsAux!codigiva
        miRsAux.Close
    End If
        
    IVA = CCur(RecuperaValor(DatosIVA, 1))
    IvaRecar = CCur(RecuperaValor(DatosIVA, 2))
        
        
        
    'Tabla
    If Combo1(0).ListIndex = 1 Then
        tabla = "clientes_laboral"
    ElseIf Combo1(0).ListIndex = 2 Then
        tabla = "clientes_fiscal"
    Else
        tabla = "clientes_cuotas"
    End If
        
    If wndReportControl.Rows(I).Record(10).Caption = "3" Then
        TipoContador = 4   'registro para laboral fiscal
    ElseIf wndReportControl.Rows(I).Record(10).Caption = "4" Then
        TipoContador = 4   'registro para laboral fiscal
    ElseIf wndReportControl.Rows(I).Record(10).Caption = "1" Then
        TipoContador = 2   'registro para cuotas asoc
    Else
        TipoContador = 3   'registro para cuotas normales
    End If
    
    
    
   
    
    
    
    Cad = "select tlinea.codclien,codforpa,iban,tlinea.codconce from " & tabla
    Cad = Cad & " as tlinea , clientes ,conceptos where tlinea.codclien=" & Me.wndReportControl.Rows(I).Record(2).Caption
    Cad = Cad & " and  tlinea.codconce=" & Me.wndReportControl.Rows(I).Record(4).Caption
    Cad = Cad & "  and tlinea.codclien=clientes.codclien and tlinea.codconce = conceptos.codconce"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    
    
    
    'Cogemos el contador
    Cad = "SELECT * FROM contadores WHERE tiporegi=" & TipoContador & " FOR update"
    Set rsContador = New ADODB.Recordset
    rsContador.Open Cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText  'no puede ser eof
        
        
 
    'Cebecera factura
    Cad = "insert into `factcli` (`numserie`,`numfactu`,`fecfactu`,`codclien`,`codforpa`,`numexped`,`fecexped`,"
    Cad = Cad & "`observa`,`totbases`,`totbasesret`,`totivas`,`totrecargo`,`totfaccl`,`retfaccl`,`trefaccl`,"
    Cad = Cad & "`cuereten`,`tiporeten`,`intconta`,`usuario`,`fecha`) values ("
    Cad = Cad & DBSet(rsContador!serfactur, "T") & "," & rsContador!NumFactu + 1 & "," & DBSet(Me.Text1(1).Text, "F") & ","
    Cad = Cad & wndReportControl.Rows(I).Record(2).Caption & "," & miRsAux!codforpa & ",NULL,NULL,"
    
    'Amoliacion
    Cad = Cad & DBSet(RecuperaValor(AmpliacionesFacturacion, 1), "T") & ","
    
    '`totbases`,`totbasesret`,
    Bases = CCur(wndReportControl.Rows(I).Record(8).Caption)
    Cad = Cad & DBSet(Bases, "N") & ",NULL,"
    '`totivas`,`totrecargo`,`totfaccl`,`retfaccl`,`trefaccl`
    ImporIVA = Round((Bases * IVA) / 100, 2)
    ImporRecar = Round((Bases * IvaRecar) / 100, 2)
    Cad = Cad & DBSet(ImporIVA, "N") & "," & DBSet(ImporIVA, "N") & "," & DBSet(ImporIVA + ImporIVA + Bases, "N") & ",NULL,NULL,"
    'cuereten`,`tiporeten`,`intconta`,`usuario`,`fecha`) values ("
    Cad = Cad & "NULL,0,0," & DBSet(vUsu.Login, "T") & "," & DBSet(Now, "FH") & ")"
    Conn.Execute Cad
    
    
    
    
    Cad = "insert into `factcli_lineas` (numserie,numfactu,fecfactu,numlinea,codconce,nomconce,ampliaci,cantidad,precio,"
    Cad = Cad & "importe,codigiva,porciva,porcrec,impoiva,imporec,aplicret) values ("
    Cad = Cad & DBSet(rsContador!serfactur, "T") & "," & rsContador!NumFactu + 1 & "," & DBSet(Me.Text1(1).Text, "F") & ",1,"
    Cad = Cad & wndReportControl.Rows(I).Record(4).Caption & "," & DBSet(wndReportControl.Rows(I).Record(5).Caption, "T") & ","
    'Amplicacion, cantidad,precio,importe  AmpliacionesFacturacion
    Cad = Cad & DBSet(RecuperaValor(AmpliacionesFacturacion, 2), "T")
    Cad = Cad & ",1," & DBSet(Bases, "T") & "," & DBSet(Bases, "T") & ","
    'codigiva,porciva,porcrec,impoiva,imporec,aplicret
    Cad = Cad & IVA_Anterior & "," & DBSet(IVA, "N") & "," & DBSet(IvaRecar, "N") & "," & DBSet(ImporIVA, "N") & "," & DBSet(ImporRecar, "N") & "," & "0)"
    Conn.Execute Cad
    
    
    'Sumatorios de IVA
    Cad = "INSERT INTO factcli_totales(numserie,numfactu,fecfactu,numlinea,baseimpo,codigiva,porciva,porcrec,"
    Cad = Cad & "impoiva,imporec) VALUES ("
    Cad = Cad & DBSet(rsContador!serfactur, "T") & "," & rsContador!NumFactu + 1 & "," & DBSet(Me.Text1(1).Text, "F") & ",1,"
    Cad = Cad & DBSet(Bases, "N") & "," & IVA_Anterior & ","
    Cad = Cad & DBSet(IVA, "N") & "," & DBSet(IvaRecar, "N") & "," & DBSet(ImporIVA, "N") & "," & DBSet(ImporRecar, "N") & ")"
    Conn.Execute Cad
    
    Cad = "UPDATE contadores SET numfactu = " & rsContador!NumFactu + 1 & " WHERE tiporegi = " & rsContador!tiporegi
    Conn.Execute Cad
    
    'Grabamos la ultima fecha de factura en las clientes_tabla
    
    Cad = "UPDATE " & tabla & " SET fecultfac =" & DBSet(Text1(1).Text, "F")
    Cad = Cad & " WHERE codclien=" & Me.wndReportControl.Rows(I).Record(2).Caption
    Cad = Cad & " and codconce=" & Me.wndReportControl.Rows(I).Record(4).Caption
    Conn.Execute Cad
    
    
    miRsAux.Close
    
    'Los cobros asociados
    Cad = "SELECT codclien, codforpa ," & rsContador!NumFactu + 1 & " NumFactu ,"
    Cad = Cad & DBSet(Text1(1).Text, "F") & " FecFactu , " & DBSet(rsContador!serfactur, "T") & " as numserie,"
    Cad = Cad & "NomClien ,DomClien,licencia,PobClien ,codposta ,ProClien ,NIFClien ,codpais ,IBAN, "
    Cad = Cad & DBSet(ImporIVA + ImporIVA + Bases, "N") & " as totfaccl"
    Cad = Cad & " from clientes where codclien=" & Me.wndReportControl.Rows(I).Record(2).Caption
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    EsUNaCuota = False
    If rsContador!serfactur = "ASO" Or rsContador!serfactur = "CUO" Then EsUNaCuota = True
    
    
    If Not InsertarEnTesoreria(EsUNaCuota, miRsAux, ctacontabanco, "", Msg) Then Err.Raise 513, Msg, Msg
    
    
    
    
    
    GenerarFacturaItem = True
    rsContador.Close
    miRsAux.Close
    Exit Function
eGenerarFacturaItem:
    MuestraError Err.Number, "Facturando: ", Err.Description
    Set rsContador = Nothing
    Set miRsAux = Nothing
    Set miRsAux = New ADODB.Recordset
End Function



