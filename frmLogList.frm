VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogList 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11745
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   11745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameConceptoDer 
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6315
      Left            =   7080
      TabIndex        =   14
      Top             =   0
      Width           =   4635
      Begin VB.OptionButton optLog 
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
         Index           =   0
         Left            =   510
         TabIndex        =   28
         Top             =   5790
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton optLog 
         Caption         =   "Trabajador"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1830
         TabIndex        =   27
         Top             =   5820
         Width           =   1755
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2145
         Index           =   0
         Left            =   360
         TabIndex        =   23
         Top             =   780
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   3784
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   6174
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2325
         Index           =   1
         Left            =   390
         TabIndex        =   24
         Top             =   3390
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   4101
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "1800"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Trabajador"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   390
         TabIndex        =   26
         Top             =   3060
         Width           =   1320
      End
      Begin VB.Label Label3 
         Caption         =   "Acción"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   390
         TabIndex        =   25
         Top             =   480
         Width           =   960
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   4020
         Picture         =   "frmLogList.frx":0000
         ToolTipText     =   "Puntear al haber"
         Top             =   510
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   5
         Left            =   3990
         Picture         =   "frmLogList.frx":014A
         ToolTipText     =   "Puntear al haber"
         Top             =   3060
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   2
         Left            =   3690
         Picture         =   "frmLogList.frx":0294
         ToolTipText     =   "Quitar al haber"
         Top             =   3060
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   3690
         Picture         =   "frmLogList.frx":03DE
         ToolTipText     =   "Quitar al haber"
         Top             =   510
         Width           =   240
      End
   End
   Begin VB.Frame FrameConcepto 
      Caption         =   "Selección"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   6915
      Begin VB.TextBox txtFecha 
         Alignment       =   1  'Right Justify
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
         Index           =   1
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "imgConcepto"
         Top             =   1170
         Width           =   1305
      End
      Begin VB.TextBox txtFecha 
         Alignment       =   1  'Right Justify
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
         Index           =   0
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "imgConcepto"
         Top             =   750
         Width           =   1305
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   930
         Picture         =   "frmLogList.frx":0528
         Top             =   1170
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   930
         Picture         =   "frmLogList.frx":05B3
         Top             =   780
         Width           =   240
      End
      Begin VB.Label lblAsiento 
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
         Index           =   0
         Left            =   2550
         TabIndex        =   22
         Top             =   990
         Width           =   4095
      End
      Begin VB.Label lblAsiento 
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
         Index           =   1
         Left            =   2550
         TabIndex        =   21
         Top             =   1440
         Width           =   4095
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   20
         Top             =   1170
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   19
         Top             =   810
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   18
         Top             =   450
         Width           =   960
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
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
      Left            =   10290
      TabIndex        =   4
      Top             =   6450
      Width           =   1215
   End
   Begin VB.CommandButton cmdAccion 
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
      Index           =   1
      Left            =   8730
      TabIndex        =   2
      Top             =   6450
      Width           =   1455
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "&Imprimir"
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
      Left            =   120
      TabIndex        =   3
      Top             =   6390
      Width           =   1335
   End
   Begin VB.Frame FrameTipoSalida 
      Caption         =   "Tipo de salida"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   5
      Top             =   3660
      Width           =   6915
      Begin VB.CommandButton PushButtonImpr 
         Caption         =   "Propiedades"
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
         Left            =   5190
         TabIndex        =   17
         Top             =   720
         Width           =   1515
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   6450
         TabIndex        =   16
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   6450
         TabIndex        =   15
         Top             =   1200
         Width           =   255
      End
      Begin VB.TextBox txtTipoSalida 
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
         Index           =   2
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1680
         Width           =   4665
      End
      Begin VB.TextBox txtTipoSalida 
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
         Index           =   1
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1200
         Width           =   4665
      End
      Begin VB.TextBox txtTipoSalida 
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
         Index           =   0
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   720
         Width           =   3345
      End
      Begin VB.OptionButton optTipoSal 
         Caption         =   "eMail"
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
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Width           =   975
      End
      Begin VB.OptionButton optTipoSal 
         Caption         =   "PDF"
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
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   1680
         Width           =   975
      End
      Begin VB.OptionButton optTipoSal 
         Caption         =   "Archivo csv"
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
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   1515
      End
      Begin VB.OptionButton optTipoSal 
         Caption         =   "Impresora"
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
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmLogList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ***********************************************************************************************************
' ***********************************************************************************************************
' ***********************************************************************************************************
'
'  3 espacios
'       -Los desde hasta,
'       -las opciones / ordenacion
'       -el tipo salida
'
' ***********************************************************************************************************
' ***********************************************************************************************************
' ***********************************************************************************************************

Public Legalizacion As String

Public Cuenta As String
Public Descripcion As String
Public FecDesde As String
Public FecHasta As String


Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1


Private Sql As String
Dim Cad As String
Dim RC As String
Dim I As Integer
Dim IndCodigo As Integer
Dim PrimeraVez As String

Dim V340()   'Llevara un str

Public Sub InicializarVbles(AñadireElDeEmpresa As Boolean)
    cadFormula = ""
    cadselect = ""
    cadParam = "|"
    numParam = 0
    cadNomRPT = ""
    conSubRPT = False
    cadPDFrpt = ""
    ExportarPDF = False
    vMostrarTree = False
    
    If AñadireElDeEmpresa Then
        cadParam = cadParam & "pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1
    End If
    
End Sub







Private Sub cmdAccion_Click(Index As Integer)
    
    If Not DatosOK Then Exit Sub
    
    
    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    
    InicializarVbles True
    
    
    ReDim V340(5)
    Cad = ""
    I = 0
    CadenaDesdeOtroForm = ""
    For NumRegElim = 1 To Me.ListView1(0).ListItems.Count
        If Me.ListView1(0).ListItems(NumRegElim).Checked Then
            I = I + 1
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & ", " & Me.ListView1(0).ListItems(NumRegElim).Text
            V340(3) = V340(3) & "," & Mid(ListView1(0).ListItems(NumRegElim).Key, 2)
        End If
    Next NumRegElim
    If I = 0 Then Cad = " - Accion"
    If I = Me.ListView1(0).ListItems.Count Then
        'VAN TODOS. No pongo ninguno
        CadenaDesdeOtroForm = ""
        V340(3) = ""
    Else
        'QUito la primeracom
        CadenaDesdeOtroForm = "Acciones: " & Trim(Mid(CadenaDesdeOtroForm, 2))
        V340(3) = Mid(V340(3), 2)
    End If
    V340(0) = CadenaDesdeOtroForm
    
    I = 0
    For NumRegElim = 1 To Me.ListView1(1).ListItems.Count
        If Me.ListView1(1).ListItems(NumRegElim).Checked Then
            I = I + 1
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & ", " & Me.ListView1(1).ListItems(NumRegElim).Text
            V340(4) = V340(4) & ",'" & DevNombreSQL(ListView1(1).ListItems(NumRegElim).Text) & "'"
        End If
    Next NumRegElim
    If I = 0 Then Cad = Cad & vbCrLf & " - Trabajador"
    If I = Me.ListView1(1).ListItems.Count Then
        'VAN TODOS. No pongo ninguno
        CadenaDesdeOtroForm = ""
        V340(4) = ""
    Else
        'QUito la primeracom
        CadenaDesdeOtroForm = "Trabajadores: " & Trim(Mid(CadenaDesdeOtroForm, 2))
        V340(4) = Mid(V340(4), 2)
    End If
    V340(1) = CadenaDesdeOtroForm
    If Cad <> "" Then
        Cad = "Debe seleccionar: " & vbCrLf & Cad
        MsgBox Cad, vbExclamation
        Exit Sub
    End If
    
    
    Cad = ""
    If txtFecha(0).Text <> "" Then Cad = "desde " & txtFecha(0).Text
    If txtFecha(1).Text <> "" Then Cad = Cad & "      hasta " & txtFecha(1).Text
    If Cad <> "" Then Cad = "Fechas: " & Trim(Cad)
    V340(5) = Cad
        
    
    
    
    
    
    If Not MontaSQL Then Exit Sub
    
    If Not HayRegParaInforme("tmppendientes", "codusu = " & vUsu.Codigo) Then Exit Sub
    
    If optTipoSal(1).Value Then
        'EXPORTAR A CSV
        AccionesCSV
    
    Else
        'Tanto a pdf,imprimiir, preevisualizar como email van COntral Crystal
    
        If optTipoSal(2).Value Or optTipoSal(3).Value Then
            ExportarPDF = True 'generaremos el pdf
        Else
            ExportarPDF = False
        End If
        SoloImprimir = False
        If Index = 0 Then SoloImprimir = True 'ha pulsado impirmir
        
        AccionesCrystal
    End If
    
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub Form_Load()
    PrimeraVez = True

    Me.Icon = frmppal.Icon
        
    'Otras opciones
    Me.Caption = "Listado de Log"

    
    PrimeraVez = True
     
    
    'txtFecha(0).Text = vParam.fechaini
    'txtFecha(1).Text = vParam.fechafin
     
    CargaListLog
    
    
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
End Sub



Private Sub frmF_Selec(vFecha As Date)
    txtFecha(IndCodigo).Text = Format(vFecha, formatoFechaVer)
End Sub


Private Sub imgCheck_Click(Index As Integer)
    NumRegElim = 0
    If Index > 1 Then NumRegElim = 1
    
    Cad = "0"
    If (Index Mod 2) = 1 Then Cad = "1"
    For I = 1 To Me.ListView1(NumRegElim).ListItems.Count
        ListView1(NumRegElim).ListItems(I).Checked = Cad = "1"
    Next
End Sub

Private Sub imgFec_Click(Index As Integer)
    
    Screen.MousePointer = vbHourglass
    
    Select Case Index
    Case 0, 1, 2
        IndCodigo = Index
    
        'FECHA
        Set frmF = New frmCal
        frmF.Fecha = Now
        If txtFecha(Index).Text <> "" Then frmF.Fecha = CDate(txtFecha(Index).Text)
        frmF.Show vbModal
        Set frmF = Nothing
        PonFoco txtFecha(Index)
        
    End Select
    
    Screen.MousePointer = vbDefault

End Sub



Private Sub optTipoSal_Click(Index As Integer)
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), Index
End Sub


Private Sub PushButton2_Click(Index As Integer)
    'FILTROS
'    If Index = 0 Then
'        frmppal.cd1.Filter = "*.csv|*.csv"
'
'    Else
'        frmppal.cd1.Filter = "*.pdf|*.pdf"
'    End If
'    frmppal.cd1.InitDir = App.Path & "\Exportar" 'PathSalida
'    frmppal.cd1.FilterIndex = 1
'    frmppal.cd1.ShowSave
'    If frmppal.cd1.FileTitle <> "" Then
'        If Dir(frmppal.cd1.FileName, vbArchive) <> "" Then
'            If MsgBox("El archivo ya existe. Reemplazar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
'        End If
'        txtTipoSalida(Index + 1).Text = frmppal.cd1.FileName
'    End If
End Sub

Private Sub PushButtonImpr_Click()
  '  frmppal.cd1.ShowPrinter
    PonerDatosPorDefectoImpresion Me, True
End Sub



Private Sub LanzaFormAyuda(Nombre As String, indice As Integer)
    Select Case Nombre
    Case "imgFecha"
        imgFec_Click indice
    End Select
    
End Sub



Private Sub AccionesCSV()
Dim Sql2 As String

    'Monto el SQL
    If Me.optLog(0) Then
        Sql = "Select  `tmppendientes`.`nomforpa` Fecha, `tmppendientes`.`nombre` Trabajador, `tmppendientes`.`Situacion` Accion, `tmppendientes`.`observa` Detalle"
        Sql = Sql & "  FROM  `tmppendientes` `tmppendientes`"
        Sql = Sql & " where codusu = " & vUsu.Codigo
        Sql = Sql & " order by 1,2,3,4"
    Else
        Sql = "Select  `tmppendientes`.`nombre` Trabajador, `tmppendientes`.`nomforpa` Fecha,`tmppendientes`.`Situacion` Accion,  `tmppendientes`.`observa` Detalle"
        Sql = Sql & "  FROM  `tmppendientes` `tmppendientes`"
        Sql = Sql & " where codusu = " & vUsu.Codigo
        Sql = Sql & " order by `tmppendientes`.`nombre`, `tmppendientes`.`nomforpa` "
    End If
    'LLamos a la funcion
    GeneraFicheroCSV Sql, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal()
Dim indRPT As String
Dim nomDocu As String
    
    vMostrarTree = False
    conSubRPT = False
        
    
    
    'Borrador
    Cad = ""
    Cad = V340(0)
    If V340(5) <> "" Then
        If Cad <> "" Then Cad = Cad & "     "
        Cad = Cad & V340(5)
    End If
    Cad = "pdh1= """ + Cad + """|"
    
    Cad = Cad & "pdh2= """ + V340(1) + """|"
    Cad = Cad & "Emp= """ & vEmpresa.nomempre & """|"
    
    
    If optLog(0).Value Then
        indRPT = "1412-00" '"rListLogCon.rpt"
    Else
        indRPT = "1412-01" '"rListLogConTr.rpt"
    End If

    
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu


    cadFormula = "{tmppendientes.codusu}=" & vUsu.Codigo

    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then CopiarFicheroASalida False, txtTipoSalida(2).Text, (Legalizacion <> "")
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 2
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
End Sub


Private Function MontaSQL() As Boolean
Dim Sql As String
Dim Sql2 As String
Dim RC As String
Dim RC2 As String

            
    MontaSQL = ListadoAcciones
    
           
End Function

Private Function ListadoAcciones() As Boolean
    
    On Error GoTo eListadoAcciones
    
    ListadoAcciones = False
    
    Conn.Execute "Delete from tmppendientes WHERE codusu = " & vUsu.Codigo
    
    '               secuenc  fecha      accion   trab      desc
    'z347carta(codusu,nif,otralineadir,razosoci,dirdatos,parrafo1)
    
    Cad = "select " & vUsu.Codigo & ",'A',right(concat(""0000000"",@rownum:=@rownum+1),7), slog.fecha,"
    Cad = Cad & "date_format(slog.fecha,'%d/%m/%Y %H:%i:%s'),titulo,usuario,descripcion"
    Cad = Cad & " from slog,tmppresu1,(SELECT @rownum:=0) r "
    Cad = Cad & " where tmppresu1.codusu=" & vUsu.Codigo & " and slog.accion=tmppresu1.codigo"
    If V340(3) <> "" Then Cad = Cad & " AND slog.accion IN (" & V340(3) & ")"
    If V340(4) <> "" Then Cad = Cad & " AND usuario IN (" & V340(4) & ")"
    If txtFecha(0).Text <> "" Then Cad = Cad & " AND  slog.fecha >= '" & Format(txtFecha(0).Text, FormatoFecha) & " 00:00:00'"
    If txtFecha(1).Text <> "" Then Cad = Cad & " AND  slog.fecha <= '" & Format(txtFecha(1).Text, FormatoFecha) & " 23:59:59'"
    
    Cad = "INSERT INTO tmppendientes(codusu,serie_cta,factura,fecha,nomforpa,Situacion,nombre,observa) " & Cad
    Conn.Execute Cad
    
    ListadoAcciones = True
    
    Exit Function
eListadoAcciones:
    MuestraError Err.Number, Err.Description
End Function





Private Sub txtfecha_LostFocus(Index As Integer)
    txtFecha(Index).Text = Trim(txtFecha(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub


    PonerFormatoFecha txtFecha(Index)
End Sub

Private Sub txtFecha_GotFocus(Index As Integer)
    ConseguirFoco txtFecha(Index), 3
End Sub

Private Sub txtFecha_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0
        
        LanzaFormAyuda "imgFecha", Index
    End If
End Sub

Private Function DatosOK() As Boolean
    
    DatosOK = False
    

    DatosOK = True

End Function


Private Sub txtNIF_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtTipoSalida_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub



Private Sub CargaListLog()
Dim IT As ListItem
   

    Set miRsAux = New ADODB.Recordset
    Cad = "Select * from tmppresu1 where tmppresu1.codusu=" & vUsu.Codigo & " ORDER BY codigo"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = ListView1(0).ListItems.Add(, CStr("K" & miRsAux!Codigo), miRsAux!Titulo)
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    Cad = "Select usuario from slog group by 1 order by 1"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = ListView1(1).ListItems.Add(, miRsAux!Usuario, miRsAux!Usuario)
            
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
End Sub



