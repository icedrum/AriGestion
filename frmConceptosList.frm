VERSION 5.00
Begin VB.Form frmConceptosList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Conceptos"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10695
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   2175
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   10395
      Begin VB.OptionButton Option1 
         Caption         =   "Tesoreria"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   5280
         TabIndex        =   23
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox txtConcepto 
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
         Top             =   1200
         Width           =   1065
      End
      Begin VB.TextBox txtDescConcepto 
         BackColor       =   &H80000018&
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   21
         Tag             =   "imgConcepto"
         Top             =   720
         Width           =   4425
      End
      Begin VB.TextBox txtDescConcepto 
         BackColor       =   &H80000018&
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   20
         Tag             =   "imgConcepto"
         Top             =   1200
         Width           =   4425
      End
      Begin VB.TextBox txtConcepto 
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
         Top             =   720
         Width           =   1065
      End
      Begin VB.Image imgCon 
         Height          =   360
         Index           =   1
         Left            =   840
         Top             =   1200
         Width           =   360
      End
      Begin VB.Image imgCon 
         Height          =   360
         Index           =   0
         Left            =   840
         Top             =   720
         Width           =   360
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
         Index           =   0
         Left            =   240
         TabIndex        =   22
         Top             =   1200
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
         Index           =   4
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   615
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
         TabIndex        =   18
         Top             =   1440
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
         Index           =   0
         Left            =   2550
         TabIndex        =   17
         Top             =   990
         Width           =   4095
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
      Left            =   9120
      TabIndex        =   4
      Top             =   5160
      Width           =   1335
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
      Left            =   7560
      TabIndex        =   2
      Top             =   5160
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
      Top             =   5160
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
      Top             =   2280
      Width           =   10395
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
         Left            =   7080
         TabIndex        =   15
         Top             =   600
         Width           =   1515
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   9240
         TabIndex        =   14
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   9240
         TabIndex        =   13
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
         Width           =   7305
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
         Width           =   7305
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
         Top             =   600
         Width           =   5025
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
         Top             =   1720
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
         Top             =   600
         Value           =   -1  'True
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmConceptosList"
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

Private WithEvents frmCo As frmConceptos
Attribute frmCo.VB_VarHelpID = -1
Dim PrimeraVez As String






Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then Unload Me
End Sub
Private Sub Form_Load()
    PrimeraVez = True

    Me.Icon = frmppal.Icon
        
    'Otras opciones
    Me.Caption = "Listado de clientes"

    
    PrimeraVez = True
     
    Me.imgCon(0).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Me.imgCon(1).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub frmCo_DatoSeleccionado(CadenaSeleccion As String)
    CadenaDesdeOtroForm = CadenaSeleccion
End Sub

Private Sub imgCon_Click(Index As Integer)
        Set frmCo = New frmConceptos
        CadenaDesdeOtroForm = ""
        frmCo.DatosADevolverBusqueda = "0|1|3|"
        frmCo.Show vbModal
        Set frmCo = Nothing
        If CadenaDesdeOtroForm <> "" Then
            txtConcepto(Index).Text = Format(RecuperaValor(CadenaDesdeOtroForm, 1), "000")
            txtDescConcepto(Index).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
            CadenaDesdeOtroForm = "2"
        End If

End Sub



Private Sub optTipoSal_Click(Index As Integer)
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), Index
End Sub


Private Sub PushButton2_Click(Index As Integer)
    'FILTROS
    'If Index = 0 Then
    '    frmppal.cd1.Filter = "*.csv|*.csv"
    '
    'Else
    '    frmppal.cd1.Filter = "*.pdf|*.pdf"
    'End If
    'frmppal.cd1.InitDir = App.Path & "\Exportar" 'PathSalida
    'frmppal.cd1.FilterIndex = 1
    'frmppal.cd1.ShowSave
    'If frmppal.cd1.FileTitle <> "" Then
    '    If Dir(frmppal.cd1.FileName, vbArchive) <> "" Then
    '        If MsgBox("El archivo ya existe. Reemplazar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    '    End If
    '    txtTipoSalida(Index + 1).Text = frmppal.cd1.FileName
    'End If
End Sub

Private Sub PushButtonImpr_Click()
    'frmppal.cd1.ShowPrinter
    PonerDatosPorDefectoImpresion Me, True
End Sub



Private Sub LanzaFormAyuda(Nombre As String, indice As Integer)
    Select Case Nombre
    Case "imgFecha"
        'imgFec_Click Indice
    End Select
    
End Sub




Private Sub cmdAccion_Click(Index As Integer)
 Dim C As String
   
    
    
    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    
    InicializarVbles True
    
    If Not PonerDesdeHasta("conceptos.codconce", "N", Me.txtConcepto(0), Me.txtDescConcepto(0), Me.txtConcepto(1), Me.txtDescConcepto(1), "pDH1=""Concepto ") Then Exit Sub
    
    
    If Not HayRegParaInforme("conceptos", cadselect) Then Exit Sub
    
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





Private Sub AccionesCSV()
Dim Sql As String

'    'Monto el SQL
    
    Sql = "select codcliecodconce nomconce periodicidad preciocon conceptos.codigiva conceptos.codmacta nommacta gestionadm  "
    Sql = Sql & " licencia,matricula,telefono,telmovil,maiclien,clientes.iban,"
    Sql = Sql & " from conceptos,ariconta" & vParam.Numconta & ".tiposiva,ariconta" & vParam.Numconta & ".cuentas where"
    Sql = Sql & "  conceptos.codigiva=tiposiva.codigiva,conceptos.codmacta =cuentas.codmacta"
    If cadselect <> "" Then
        Sql = Sql & " AND " & cadselect
    End If
    GeneraFicheroCSV Sql, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal()
Dim indRPT As String
Dim nomDocu As String
    
    vMostrarTree = False
    conSubRPT = False
        
    
    
    'If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
 '   If Me.Option1(2).Value Then
        nomDocu = "rConceptos.rpt"
 '   ElseIf Me.Option1(1).Value Then
 '       nomDocu = "rClientesContac.rpt"
 '   Else
 '       nomDocu = "rClientes.rpt"
 '   End If
    cadNomRPT = nomDocu


    
    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then CopiarFicheroASalida False, txtTipoSalida(2).Text, False
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 2
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
End Sub





Private Sub txtconcepto_GotFocus(Index As Integer)
    ConseguirFoco txtConcepto(Index), 3
End Sub

Private Sub txtconcepto_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = teclaBuscar Then
        KeyAscii = 0
        imgCon_Click Index
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub txtconcepto_LostFocus(Index As Integer)
    InformeTxtLostFocus txtConcepto(Index), Me.txtDescConcepto(Index), False
End Sub

