VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#17.2#0"; "Codejock.ReportControl.v17.2.0.ocx"
Begin VB.Form frmGestionTasasMov 
   Caption         =   "TASAS. Gestion compra-venta"
   ClientHeight    =   9120
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16080
   LinkTopic       =   "Form1"
   ScaleHeight     =   9120
   ScaleWidth      =   16080
   StartUpPosition =   3  'Windows Default
   Begin XtremeReportControl.ReportControl wndReportControl 
      Height          =   9015
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   16215
      _Version        =   1114114
      _ExtentX        =   28601
      _ExtentY        =   15901
      _StockProps     =   64
      FreezeColumnsAbs=   0   'False
   End
   Begin VB.Frame FrameConce 
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   13815
      Begin VB.TextBox txtTasas 
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
         Height          =   480
         Index           =   2
         Left            =   9240
         Locked          =   -1  'True
         TabIndex        =   10
         Tag             =   "imgConcepto"
         Top             =   120
         Width           =   1425
      End
      Begin VB.TextBox txtTasas 
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
         Height          =   480
         Index           =   1
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   9
         Tag             =   "imgConcepto"
         Top             =   120
         Width           =   4785
      End
      Begin VB.TextBox txtTasas 
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
         Height          =   480
         Index           =   0
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   8
         Tag             =   "imgConcepto"
         Top             =   120
         Width           =   1185
      End
      Begin VB.Label Label2 
         Caption         =   "Concepto"
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
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label2 
         Caption         =   "Stock"
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
         Index           =   0
         Left            =   8280
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
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
         TabIndex        =   5
         Top             =   120
         Width           =   1575
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   330
            Left            =   120
            TabIndex        =   6
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
                  Object.ToolTipText     =   "Quitar stock"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Object.ToolTipText     =   "Contabilizar"
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
         Left            =   13320
         TabIndex        =   4
         Top             =   480
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmGestionTasasMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Concepto As Integer   'De que concepto vamos a ver movimientos de tas...

Dim PrimVez As Boolean
Dim Cantidad As Integer


Private Sub Form_Activate()
    If PrimVez Then
        PrimVez = False
        
        MostrarDatos True
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
        .Buttons(2).Enabled = False
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
        
        '.Buttons(3).Image = 42
        .Buttons(3).Visible = False
    End With
    
    
    
    FrameConce.BorderStyle = 0

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
    FrameConce.Move 60, Frame1.top + Frame1.Height + 90, Me.Width - 320, FrameConce.Height
    wndReportControl.Move 60, FrameConce.top + Me.FrameConce.Height + 120, Me.Width - 320, Me.Height - Me.FrameConce.Height - FrameConce.top - 120
    
    Err.Clear
End Sub



Private Sub CreateReportControl()
    'gestadministrativa  id usuario fechacreacion llevados importe fechafinalizacion
    Dim Column As ReportColumn
    
    wndReportControl.Columns.DeleteAll
    
    'Adds a new ReportColumn to the ReportControl's collection of columns, growing the collection by 1.
    Set Column = wndReportControl.Columns.Add(COLUMN_IMPORTANCE, "Tipo", 18, False)
    Column.Icon = COLUMN_IMPORTANCE_ICON
    '    C = "Select id,codconce,tipomovi,usuario,fechamov,numserie,numdocum,anoexped,ampliacion,numlinea,cantidad"
    Set Column = wndReportControl.Columns.Add(1, "Fecha", 35, True)
    Set Column = wndReportControl.Columns.Add(2, "Usuario", 30, True)
    Set Column = wndReportControl.Columns.Add(3, "Documento", 40, True)
    Set Column = wndReportControl.Columns.Add(4, "Ampliacion", 90, True)
    Set Column = wndReportControl.Columns.Add(5, "Cantidad", 15, True)
    Column.Alignment = xtpAlignmentRight
End Sub




Private Sub MostrarDatos(PRimera As Boolean)

    Label1.Caption = "Leyendo BD"
    Label1.Refresh
    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset
    
    If PRimera Then
        miRsAux.Open "Select * from conceptos where codconce=" & Concepto, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        'No puede ser eof
        Me.txtTasas(0).Text = Format(miRsAux!codconce, "000")
        Me.txtTasas(1).Text = miRsAux!nomconce
        Me.txtTasas(2).Text = Format(DBLet(miRsAux!stock, "N"), "#,##0")
        Me.txtTasas(2).Tag = DBLet(miRsAux!stock, "N")
        miRsAux.Close
    End If
    populateInbox
    wndReportControl.Populate
    
    
    Set miRsAux = Nothing
    
    Label1.Caption = ""
    Screen.MousePointer = vbDefault
End Sub



Public Sub populateInbox()
Dim C As String
Dim F As Date


    wndReportControl.Records.DeleteAll
    C = "Select id,codconce,tipomovi,usuario,fechamov,numserie,numdocum,anoexped,ampliacion,numlinea,cantidad"
    C = C & " from tasas where codconce = " & Concepto & " order by fechamov"
    miRsAux.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        AddRecordSin
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
End Sub


Private Sub AddRecordSin()
Dim Situacion As Byte
Dim Aux As String
    
  
    Dim Record As ReportRecord
    Set Record = wndReportControl.Records.Add()
    
    Dim Item As ReportRecordItem
    
    ' id,codconce,tipomovi,usuario,fechamov,numserie,numdocum,anoexped,ampliacion,numlinea,cantidad"
    
    'Adds a new ReportRecordItem to the Record, this can be thought of as adding a cell to a row
    Set Item = Record.AddItem("")
    
    Situacion = miRsAux!tipomovi
    If Situacion = 1 Then
        'Assigns an icon to the item
        Item.Icon = Situacion
        Item.ToolTip = "Compra"
           
    Else
        Item.Icon = 4
        Item.ToolTip = "Venta"
    End If




      


    
    Set Item = Record.AddItem("")
    Item.Caption = Format(miRsAux!fechamov, "dd/mm/yyyy hh:nn")
    Item.Value = Format(miRsAux!fechamov, "yyyymmddhhnnss")
    
    
    Record.AddItem CStr(miRsAux!Usuario)
    
    Set Item = Record.AddItem("")
    If miRsAux!tipomovi = 0 Then
        'VENTA
        Stop
        Item.Caption = miRsAux!numserie & Format(miRsAux!numdocum, "000000") & "/" & miRsAux!anoexped & " -" & miRsAux!numlinea
        
        'Ampliacion
        Aux = miRsAux!Ampliacion
    Else
        Item.Caption = "Compra"
        'Ampliacion
        Aux = " "
    End If
    'Ampliacion
    Set Item = Record.AddItem(Aux)
    
    I = miRsAux!Cantidad
    If miRsAux!tipomov = 0 Then I = -I
    Set Item = Record.AddItem(Format(I, "#,##0"))
    Item.Value = I
    

    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim C As String

    If Button.Index = 1 Then
        'NUEVO
        CadenaDesdeOtroForm = ""
        frmMensajes.Opcion = 7
        frmMensajes.Parametros = Me.txtTasas(0).Text & " " & Me.txtTasas(1).Text & "|"
        frmMensajes.Show vbModal
        If CadenaDesdeOtroForm <> "" Then
               
               
                'tasas(id,codconce,tipomovi,usuario,fechamov,numdocum,cantidad)
               C = "INSERT INTO tasas(codconce,tipomovi,usuario,fechamov,ampliacion,cantidad)"
               C = C & "  VALUES (" & Concepto & ",1," & DBSet(vUsu.Login, "T")
               C = C & "," & DBSet(RecuperaValor(CadenaDesdeOtroForm, 1), "FH") & ",'COMPRA'"
               C = C & "," & DBSet(RecuperaValor(CadenaDesdeOtroForm, 2), "N") & ")"
               Conn.Execute C
               
               txtTasas(2).Tag = Val(txtTasas(2).Tag) + Val(RecuperaValor(CadenaDesdeOtroForm, 2))
               C = "UPDATE conceptos set stock= " & DBSet(txtTasas(2).Tag, "N")
               C = C & " WHERE codconce=" & Concepto
               Conn.Execute C
                               
               Me.txtTasas(2).Text = Format(txtTasas(2).Tag, "#,##0")
               MostrarDatos False
               
        End If
    Else
        If Me.wndReportControl.Records.Count = 0 Then Exit Sub
        
        If Button.Index = 2 Then
            
        End If
    End If
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Msg = DevuelveDesdeBD("stock", "conceptos", "codconce", CStr(Concepto))
    If Msg = "" Then Exit Sub
    
    Msg = "Concepto: " & Me.txtTasas(0).Text & " " & Me.txtTasas(1).Text & vbCrLf
    Msg = Msg & "Cantidad: " & Me.txtTasas(2).Text
    
    If MsgBox("Va a dejar de llevar control de movimientos:" & vbCrLf & Msg & vbCrLf & "¿Continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    Conn.Execute "UPDATE conceptos set stock= NULL WHERE codconce=" & Concepto
    vLog.Insertar 9, vUsu, Msg
    Unload Me
End Sub
