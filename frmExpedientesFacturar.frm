VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExpedientesFacturar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturar expedientes"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   13905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView ListView1 
      Height          =   5655
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   9975
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Serie"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Numero"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fecha"
         Object.Width           =   2188
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Licencia"
         Object.Width           =   1905
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Nombre"
         Object.Width           =   6950
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Forma de pago"
         Object.Width           =   4657
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Importe"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "TipoRegi"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Anoexped"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13575
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5520
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   6
         Top             =   360
         Width           =   1515
      End
      Begin VB.TextBox txtFecha 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1920
         MaxLength       =   30
         TabIndex        =   4
         Top             =   360
         Width           =   1515
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
         Left            =   10920
         TabIndex        =   3
         Top             =   360
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
         Left            =   12240
         TabIndex        =   2
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label lblInf 
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
         Left            =   7560
         TabIndex        =   8
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Base imponible"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   4080
         TabIndex        =   7
         ToolTipText     =   "Fecha alta asociado"
         Top             =   420
         Width           =   1305
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   0
         Left            =   1560
         Picture         =   "frmExpedientesFacturar.frx":0000
         ToolTipText     =   "Fecha alta asociado"
         Top             =   420
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha factura"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   5
         ToolTipText     =   "Fecha alta asociado"
         Top             =   420
         Width           =   1155
      End
   End
End
Attribute VB_Name = "frmExpedientesFacturar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Dim PrimeraVez As Boolean
Dim Sql As String

Private Sub cmdAceptar_Click()
    Sql = ""
    For I = 1 To Me.ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked Then Sql = Sql & "X"
    Next
    
    
    If Sql = "" Then
        MsgBox "Seleccione algun expediente para facturar", vbExclamation
        Exit Sub
    End If
    I = Len(Sql)
    
    
    'Fecha
    Sql = ""
    If Me.txtFecha(0).Text = "" Then
        MsgBox "Indique fecha factura", vbExclamation
        Exit Sub
    Else
        If Not FechaFacturaOK(CDate(txtFecha(0).Text)) Then Exit Sub
    End If
    
    Sql = "Va a facturar " & I & " expedientes.  ¿Continuar?"
    If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    
    Screen.MousePointer = vbHourglass
    If BloqueoManual("FACT_EXP", "1") Then
        
        For I = 1 To ListView1.ListItems.Count
            
            If Me.ListView1.ListItems(I).Checked Then
                Conn.BeginTrans
                lblInf.Caption = ListView1.ListItems(I).Text & ListView1.ListItems(I).SubItems(1)
                lblInf.Refresh
                If FacturarExpediente(ListView1.ListItems(I).SubItems(7), ListView1.ListItems(I).SubItems(1), ListView1.ListItems(I).SubItems(8), CDate(txtFecha(0).Text)) Then
                    Conn.CommitTrans
                Else
                    Conn.RollbackTrans
                End If
                
            End If
        Next I
        CargaDatos
        DesBloqueoManual "FACT_EXP"
    End If
    Screen.MousePointer = vbDefault
    lblInf.Caption = ""
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        CargaDatos
        
    End If
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Load()
    PrimeraVez = True
    Me.Icon = frmppal.Icon
    
    
  
    
    
    
    
    
    Me.Frame1.BorderStyle = 0
    txtImporte.Text = "0,00"
    txtImporte.Tag = 0
    

End Sub



Private Sub Form_Resize()
    On Error Resume Next
    
    Frame1.Move 0, 0 + 60, ScaleWidth, Frame1.Height
    ListView1.Move 0, Frame1.Height + Frame1.top, ScaleWidth, ScaleHeight - Frame1.Height - Frame1.top
    
    If Err.Number <> 0 Then Err.Clear
End Sub













'Cuando modifiquemos o insertemos, pondremos el SQL entero
Public Sub CargaDatos()
Dim Sql As String
Dim IT

    
    Set miRsAux = New ADODB.Recordset
    
    Sql = "Select expedientes.numserie,expedientes.numexped,expedientes.anoexped,fecexped,licencia,nomclien,expedientes.tiporegi,expedientes.anoexped,"
    Sql = Sql & " nomforpa, sum(if(gestionadm=1,if(expedientes_lineas.codsitua>1,0,1),0)) nosepuede,sum(importe) eltotal"
    Sql = Sql & " FROM expedientes,expedientes_lineas,clientes,conceptos,ariconta" & vParam.Numconta & ".formapago WHERE"
    Sql = Sql & " expedientes.TipoRegi = expedientes_lineas.TipoRegi And expedientes.numexped = expedientes_lineas.numexped"
    Sql = Sql & " AND  expedientes.anoexped = expedientes_lineas.anoexped  and expedientes.codclien =clientes.codclien"
    Sql = Sql & " and conceptos.codconce=expedientes_lineas.codconce and formapago.codforpa=clientes.codforpa"
    Sql = Sql & " and expedientes.codsitua<2 group by 1,2,3 ORDER BY 1,3,2"


        
    ListView1.ListItems.Clear
    
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = ListView1.ListItems.Add
        IT.Text = miRsAux!numSerie
        IT.SubItems(1) = Format(miRsAux!numexped, "00000")
        IT.SubItems(2) = Format(miRsAux!fecexped, "dd/mm/yy")
        IT.SubItems(3) = DBLet(miRsAux!licencia, "T")
        IT.SubItems(4) = DBLet(miRsAux!NomClien, "T")
        IT.SubItems(5) = DBLet(miRsAux!nomforpa, "T")
        IT.SubItems(6) = Format(DBLet(miRsAux!eltotal, "N"), FormatoImporte)
        IT.SubItems(7) = miRsAux!TipoRegi
        IT.SubItems(8) = miRsAux!anoexped
        If miRsAux.Fields!nosepuede = 0 Then
            IT.Tag = 0
        Else
            IT.Tag = 1
            IT.ForeColor = vbRed
        End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
        
    Set miRsAux = Nothing
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Sql = vFecha
End Sub

Private Sub imgppal_Click(Index As Integer)
    Set frmC = New frmCal
    frmC.Fecha = Now
    Sql = ""
    If Me.txtFecha(Index).Text <> "" Then frmC.Fecha = txtFecha(Index).Text
    frmC.Show vbModal
    If Sql <> "" Then
        txtFecha(Index).Text = Format(Sql, "dd/mm/yyyy")
    End If
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim Im As Currency
    If Item.Tag = 1 Then
        Item.Checked = False
    Else
        I = 1
        If Not Item.Checked Then I = -1
        Im = ImporteFormateado(Item.SubItems(6))
        txtImporte.Tag = txtImporte.Tag + (Im * I)
        txtImporte.Text = Format(txtImporte.Tag, FormatoImporte)
    End If
End Sub

Private Sub txtFecha_GotFocus(Index As Integer)
    ConseguirFoco txtFecha(Index), 4
End Sub

Private Sub txtFecha_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub txtfecha_LostFocus(Index As Integer)
Dim B As Boolean
    If txtFecha(Index).Text = "" Then Exit Sub
    B = True
    If Index = 4 Or Index = 5 Or Index = 6 Then
        If Not EsFechaHoraOK(txtFecha(Index)) Then B = False
    Else
        If Not EsFechaOK(txtFecha(Index)) Then B = False
    End If
    If Not B Then
        txtFecha(Index).Text = ""
        MsgBox "Fecha incorrecta", vbExclamation
        PonFoco txtFecha(Index)
        Exit Sub
   
    End If

End Sub

