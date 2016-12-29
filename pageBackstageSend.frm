VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#17.2#0"; "Codejock.CommandBars.v17.2.0.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#17.2#0"; "Codejock.ReportControl.v17.2.0.ocx"
Begin VB.Form pageBackstageSend 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13155
   FontTransparent =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   479
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   877
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeReportControl.ReportControl wndReportControl 
      Height          =   5295
      Left            =   10320
      TabIndex        =   2
      Top             =   -360
      Visible         =   0   'False
      Width           =   12255
      _Version        =   1114114
      _ExtentX        =   21616
      _ExtentY        =   9340
      _StockProps     =   64
      FreezeColumnsAbs=   0   'False
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5895
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   10398
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   317
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   317
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Resumido"
         Object.Width           =   317
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Inicio"
         Object.Width           =   317
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Fin"
         Object.Width           =   317
      EndProperty
   End
   Begin XtremeCommandBars.BackstageSeparator lblBackstageSeparator1 
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   4695
      _Version        =   1114114
      _ExtentX        =   8281
      _ExtentY        =   450
      _StockProps     =   2
      MarkupText      =   ""
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cambiar empresa"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005B5B5B&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "pageBackstageSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ItemCheck As Boolean

Private Sub Seleccionado()

End Sub

Private Sub Form_Load()
    BuscaEmpresas
    ItemCheck = False
End Sub






Private Sub BuscaEmpresas()
Dim Prohibidas As String
Dim Rs As ADODB.Recordset
Dim cad As String
Dim SQL As String
Dim ItmX As ListItem

ListView1.ListItems.Clear

Dim I As Integer

For I = 1 To 5
    ListView1.ColumnHeaders(I).Width = CInt(RecuperaValor("60|420|200|100|100|", I))
Next

'Cargamos las prohibidas
Prohibidas = DevuelveProhibidas

'Cargamos las empresas
Set Rs = New ADODB.Recordset

'[Monica]11/04/2014: solo debe de salir las ariconta
Rs.Open "Select * from usuarios.empresasarigestion where arigestion like 'arigestion%' ORDER BY Codempre", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText

While Not Rs.EOF
    cad = "|" & Rs!codempre & "|"
    If InStr(1, Prohibidas, cad) = 0 Then
        cad = Rs!nomempre
        Set ItmX = ListView1.ListItems.Add()
        ItmX.Text = Rs!codempre
        
        
        ItmX.SubItems(1) = Rs!nomempre
        ItmX.SubItems(2) = Rs!nomresum
       ' cad = "fechafin"
       ' Sql = DevuelveDesdeBD("fechaini", "arigestion" & Rs!codempre & ".parametros", "1", "1", "N", cad)
        SQL = " "
        cad = " "
        ItmX.SubItems(3) = SQL
        ItmX.SubItems(4) = cad
        
            
        cad = Rs!AriGestion & "|" & Rs!nomresum '& "|" & Rs!Usuario & "|" & Rs!Pass & "|"
        
        If Rs!codempre = vEmpresa.codempre Then
            ItmX.Bold = True
            Set ListView1.SelectedItem = ItmX
        End If
            
       ' ItmX.Tag = Cad
       ' ItmX.ToolTipText = Rs!CONTA
        
        
       
    End If
    Rs.MoveNext
Wend
Rs.Close
End Sub


Private Function DevuelveProhibidas() As String
Dim Rs As ADODB.Recordset
Dim cad As String
Dim I As Integer
    On Error GoTo EDevuelveProhibidas
    DevuelveProhibidas = ""
    Set Rs = New ADODB.Recordset
    I = vUsu.Codigo Mod 1000
    Rs.Open "Select * from usuarios.usuarioempresaAriGestion WHERE codusu =" & I, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    cad = ""
    While Not Rs.EOF
        cad = cad & Rs.Fields(1) & "|"
        Rs.MoveNext
    Wend
    If cad <> "" Then cad = "|" & cad
    Rs.Close
    DevuelveProhibidas = cad
EDevuelveProhibidas:
    Err.Clear
    Set Rs = Nothing
End Function






Public Sub CreateReportControl()
    'Start adding columns
    Dim Column As ReportColumn
    
    'Adds a new ReportColumn to the ReportControl's collection of columns, growing the collection by 1.
    Set Column = wndReportControl.Columns.Add(COLUMN_IMPORTANCE, "Importance", 18, False)
    'The value assigned to the icon property corresponds to the index of an icon in the collection of wndReportControl.Icons
    'I.e. The icon at index=1 in the collection will be displayed in the column header.  The index of the icon depends on the
    'order it is added to the collection.  (Icons are added after the records near the bottom of the Form_Load)
    Column.Icon = COLUMN_IMPORTANCE_ICON
    Set Column = wndReportControl.Columns.Add(COLUMN_ICON, "Message Class", 18, False)
    Column.Icon = COLUMN_MAIL_ICON
    Set Column = wndReportControl.Columns.Add(COLUMN_ATTACHMENT, "Attachment", 18, False)
    Column.Icon = COLUMN_ATTACHMENT_ICON
    Set Column = wndReportControl.Columns.Add(COLUMN_FROM, "From", 130, True)
    Set Column = wndReportControl.Columns.Add(COLUMN_SUBJECT, "Subject", 180, True)
    Set Column = wndReportControl.Columns.Add(COLUMN_RECEIVED, "Received", 90, True)
    Set Column = wndReportControl.Columns.Add(COLUMN_SIZE, "Size", 55, False)
    Set Column = wndReportControl.Columns.Add(COLUMN_CHECK, "Checked", 18, False)
    Column.Icon = COLUMN_CHECK_ICON
    Set Column = wndReportControl.Columns.Add(COLUMN_PRICE, "Price", 80, True)
    Column.Visible = False
    Set Column = wndReportControl.Columns.Add(COLUMN_SENT, "Sent", 150, True)
    Column.Visible = False
    Set Column = wndReportControl.Columns.Add(COLUMN_CREATED, "Created", 80, True)
    Column.Visible = False
    Set Column = wndReportControl.Columns.Add(COLUMN_CONVERSATION, "Conversation", 80, True)
    Column.Visible = False
    Set Column = wndReportControl.Columns.Add(COLUMN_CONTACTS, "Contacts", 30, True)
    Column.Visible = False
    Set Column = wndReportControl.Columns.Add(COLUMN_MESSAGE, "Message", 80, True)
    Column.Visible = False
    Set Column = wndReportControl.Columns.Add(COLUMN_CC, "CC", 80, True)
    Column.Visible = False
    Set Column = wndReportControl.Columns.Add(COLUMN_CATEGORIES, "Categories", 30, True)
    Column.Visible = False
    Set Column = wndReportControl.Columns.Add(COLUMN_AUTOFORWARD, "Auto Forward", 30, True)
    Column.Visible = False
    Set Column = wndReportControl.Columns.Add(COLUMN_DO_NOT_AUTOARCH, "Do not autoarchive", 30, True)
    Column.Visible = False
    Set Column = wndReportControl.Columns.Add(COLUMN_DUE_BY, "Due by", 30, True)
    Column.Visible = False
    
    wndReportControl.PaintManager.MaxPreviewLines = 1
    wndReportControl.PaintManager.HorizontalGridStyle = xtpGridNoLines
                  
    'Add Records to the ReportControl
    populateInbox
    
    'wndReportControl.PaintManager.VerticalGridStyle = xtpGridSolid
    
    'This code below will add a column to the GroupOrder collection of columns.
    'This will cause the columns in the ReportControl to be grouped by column "COLUMN_FROM",
    'which has an index of zero (0) in the GroupOrder collection. Columns are first grouped
    'in the order that they are added to the GroupOrder collection.
    wndReportControl.GroupsOrder.Add wndReportControl.Columns(COLUMN_FROM)

    'This will cause the column at index 0 of the GroupOrder collection to be displayed
    'in ascending order.
    wndReportControl.GroupsOrder(0).SortAscending = True
            
    'The following illustrate how to change the different fonts used in the ReportControl
    'Dim TextFont As StdFont
    'Set TextFont = Me.Font
    'TextFont.Size = 15
    'Set wndReportControl.PaintManager.TextFont = TextFont
    'Set wndReportControl.PaintManager.CaptionFont = TextFont
    'Set wndReportControl.PaintManager.PreviewTextFont = TextFont
    
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


Public Sub populateInbox()
    'Variables used to provide data to the AddRecord subroutine
    Dim Sent As Boolean
    Dim Created As Boolean
    Dim Received As Boolean
    Dim Message As String
    Dim From As String
    Dim Subject As String
    Dim EmptyString As String

    EmptyString = ""
    'Used to specify if the current (Today's) date is used in date related items
    'If False, a date is randomly generated
    Sent = False
    Created = False
    Received = False
    
    'Start Adding Records with the AddRecord helper subroutine
    'Look at the AddRecord subroutine for a more detailed explanation
    Message = "Simple Preview Text"
    Subject = "Undeliverable Mail"
    AddRecord IMPORTANCE_HIGH, False, False, "postmaster@mail.codejock.com", Subject, Sent, 7, True, 5, _
                Received, Created, Subject, EmptyString, Message, EmptyString, EmptyString, EmptyString, _
                EmptyString, EmptyString, Message
                
    Message = "Breaks words. Lines are automatically broken between words if a word would extend past the edge of the rectangle specified by the lpRect parameter. A carriage return-linefeed sequence also breaks the line."
    Subject = "Hi Mary Jane"
    AddRecord IMPORTANCE_NORMAL, False, False, "Peter Parker", "RE: " + Subject, Sent, 14, False, 4.3, _
                Received, Created, Subject, EmptyString, Message, EmptyString, EmptyString, EmptyString, _
                EmptyString, EmptyString, Message
                
    Subject = ""
    Message = "If you have several conditions to be tested together, and you know that one is more likely to pass or fail than the others, you can use a feature called 'short circuit evaluation' to speed the execution of your script. When JScript evaluates a logical expression, it only evaluates as many sub-expressions as required to get a result."
    AddRecord IMPORTANCE_NORMAL, True, False, "James Howlett", "RE:" + Subject, Sent, 24, False, 56, _
        Received, Created, Subject, EmptyString, Message, EmptyString, EmptyString, EmptyString, _
        EmptyString, EmptyString, Message


End Sub
Private Sub AddRecord(Importance As Integer, Checked As Boolean, Attachment As Boolean, _
                        From As String, Subject As String, Sent As Boolean, SIZE As Integer, _
                        Read As Boolean, Price As Double, Received As Boolean, Created As Boolean, _
                        Conversation As String, Contact As String, Message As String, CC As String, _
                        Categories As String, Autoforward As String, DoNotAutoarch As String, _
                        DueBy As String, Preview As String)
  
    Dim Record As ReportRecord
    'Adds a new Record to the ReportControl's collection of records, this record will
    'automatically be attached to a row and displayed with the Populate method
    Set Record = wndReportControl.Records.Add()
    
    Dim Item As ReportRecordItem
    
    'Adds a new ReportRecordItem to the Record, this can be thought of as adding a cell to a row
    Set Item = Record.AddItem("")
    If (Importance = IMPORTANCE_HIGH) Then
        'Assigns an icon to the item
        Item.Icon = RECORD_IMPORTANCE_HIGH_ICON
        'Assigns a GroupCaption to the item, this is displayed int he grouprow when grouped by the column
        'this item belong to.
        Item.GroupCaption = "Importance: Hight"
        'Sets the group priority of the item when grouped, the lower the number the higher the priority,
        'Highest priority is displayed first
        Item.GroupPriority = IMPORTANCE_HIGH
        'Sets the sort priority of the item when the colulmn is sorted, the lower the number the higher the priority,
        'Highest priority is sorted displayed first, then by value
        Item.SortPriority = IMPORTANCE_HIGH
    End If
    If (Importance = IMPORTANCE_LOW) Then
        Item.Icon = RECORD_IMPORTANCE_LOW_ICON
        Item.GroupCaption = "Importance: Low"
        Item.GroupPriority = IMPORTANCE_LOW
        Item.SortPriority = IMPORTANCE_LOW
    End If
    If (Importance = IMPORTANCE_NORMAL) Then
        Item.GroupCaption = "Importance: Normal"
        Item.GroupPriority = IMPORTANCE_NORMAL
        Item.SortPriority = IMPORTANCE_NORMAL
    End If
      
    Set Item = Record.AddItem("")
    Item.GroupCaption = IIf(Read, "Message Status: Read", "Message Status: Unread")
    Item.GroupPriority = IIf(Read, 0, 1)
    Item.SortPriority = IIf(Read, 0, 1)
    Item.Icon = IIf(Read, RECORD_READ_MAIL_ICON, RECORD_UNREAD_MAIL_ICON)
       
    Set Item = Record.AddItem("")
    Item.Checked = IIf(Attachment, ATTACHMENTS_TRUE, ATTACHMENTS_FALSE)
    Item.GroupCaption = IIf(Attachment, "Attachments: With Attachments", "Attachments: No Attachments")
    Item.GroupPriority = IIf(Attachment, 0, 1)
    Item.SortPriority = IIf(Attachment, 0, 1)
    If (Attachment) Then
        Item.Icon = IIf(Attachment, COLUMN_ATTACHMENT_NORMAL_ICON, COLUMN_ATTACHMENT_ICON)
    End If
    
    Record.AddItem From
    Record.AddItem Subject
       
    Set Item = Record.AddItem("")
    If (Received) Then
      '  GetDate Item, DatePart("w", Now), DatePart("d", Now), DatePart("m", Now), DatePart("yyyy", Now)
    Else
      '  GetDate Item
    End If
        
    Record.AddItem SIZE
    
    Set Item = Record.AddItem("")
    'Adds a checkbox to the item
    Item.HasCheckbox = True
    'Determine if the chekbox will have a checkmark based on the value of Checked
    Item.Checked = IIf(Checked, CHECKED_TRUE, CHECKED_FALSE)
    'Sets the group caption
    Item.GroupCaption = IIf(Checked, "Check State: Checked", "Check State: Unchecked")
    Item.GroupPriority = IIf(Checked, 0, 1)
    Item.SortPriority = IIf(Checked, 0, 1)
           
    Set Item = Record.AddItem(Price)
    'Specifys the format that the price will be displayed
    Item.Format = "$ %s"
    'Assigns the properties based on the value of Price
    Select Case Price
        Case Is <= 5:
                Item.GroupCaption = "Record Price < 5"
                Item.GroupPriority = 1
        Case Is <= 20:
                Item.GroupCaption = "Record Price 5-20"
                Item.GroupPriority = 0
        Case Is > 20:
                Item.GroupCaption = "Record Price > 20"
                Item.GroupPriority = 3
    End Select
    
    Set Item = Record.AddItem("")
    If (Sent) Then
       ' GetDate Item, DatePart("w", Now), DatePart("d", Now), DatePart("m", Now), DatePart("yyyy", Now)
    Else
       ' GetDate Item
    End If
    
    Set Item = Record.AddItem("")
    If (Created) Then
      '  GetDate Item, DatePart("w", Now), DatePart("d", Now), DatePart("m", Now), DatePart("yyyy", Now)
    Else
       ' GetDate Item
    End If
    
    Record.AddItem Conversation
    Record.AddItem Contact
    Record.AddItem Message
    Record.AddItem CC
    Record.AddItem Categories
    Record.AddItem Autoforward
    Record.AddItem DoNotAutoarch
    Record.AddItem Contact
    Record.AddItem DueBy
    
    'Adds the PreviewText to the Record.  PreviewText is the text displayed for the ReportRecord while in PreviewMode
    Record.PreviewText = Preview
    
End Sub


Private Sub Form_LostFocus()
    ItemCheck = False
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.ListView1.Width = Me.ScaleWidth - ListView1.Left - 240
End Sub

Private Sub ListView1_DblClick()
    If Not ItemCheck Then Exit Sub
    If ListView1.SelectedItem Is Nothing Then Exit Sub
   frmppal.CambiarEmpresa CInt(ListView1.SelectedItem.Text)
    
    
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    ItemCheck = True
    
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ItemCheck = True
End Sub
