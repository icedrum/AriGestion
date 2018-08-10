VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConceptosListEsta 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12345
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   12345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   2040
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
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
      Width           =   5235
      Begin MSComctlLib.ListView lwConce 
         Height          =   4785
         Left            =   120
         TabIndex        =   23
         Top             =   780
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   8440
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
      Begin VB.Label Label3 
         Caption         =   "Conceptos"
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
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   1140
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   4680
         Picture         =   "frmConceptosListEsta.frx":0000
         ToolTipText     =   "Puntear al haber"
         Top             =   480
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   4320
         Picture         =   "frmConceptosListEsta.frx":014A
         ToolTipText     =   "Quitar al haber"
         Top             =   480
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
         Picture         =   "frmConceptosListEsta.frx":0294
         Top             =   1170
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   930
         Picture         =   "frmConceptosListEsta.frx":031F
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
      Left            =   11010
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
      Left            =   9450
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
Attribute VB_Name = "frmConceptosListEsta"
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



Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1


Private SQL As String
Dim Cad As String
Dim RC As String
Dim I As Integer
Dim IndCodigo As Integer
Dim PrimeraVez As String


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
    
    Cad = ""
    For I = 1 To Me.lwConce.ListItems.Count
        If lwConce.ListItems(I).Checked Then Cad = Cad & ", " & Mid(lwConce.ListItems(I).Key, 2)
    Next
    
    
    If Cad = "" Then
        MsgBox "Seleccione algun concepto", vbExclamation
        Exit Sub
    End If
    Cad = Mid(Cad, 2)
    
    cadNomRPT = "{conceptos.codconce} IN [" & Cad & "]"
    AnyadirAFormula cadFormula, cadNomRPT
    cadNomRPT = Cad
    
    Cad = ""
    
    cadselect = ""
    If txtFecha(0).Text <> "" Then
        Cad = "desde " & txtFecha(0).Text
        cadselect = " AND @f >= " & DBSet(txtFecha(0).Text, "F")
        cadParam = cadParam & "txtfechaIni=""" & txtFecha(0).Text & """|"
        numParam = numParam + 1
    End If
    If txtFecha(1).Text <> "" Then
        Cad = Cad & " hasta " & txtFecha(1).Text
        cadselect = cadselect & " AND @f <= " & DBSet(txtFecha(1).Text, "F")
        cadParam = cadParam & "txtfechafIn=""" & txtFecha(1).Text & """|"
        numParam = numParam + 1
    End If
    If Cad <> "" Then
        Cad = "Fechas: " & Trim(Cad)
        numParam = numParam + 1
        cadParam = cadParam & "pdh1=""" & Cad & """|"
    End If
    
    
    
    
    'en expedientes
    I = 0
    Cad = Replace(cadselect, "@f", "fecexped")
    
    SQL = "  expedientes.TipoRegi = expedientes_lineas.tiporegi And expedientes.numexped = expedientes_lineas.numexped "
    SQL = SQL & " AND  expedientes.anoexped = expedientes_lineas.anoexped  " & Cad
    SQL = SQL & " AND codconce IN (" & cadNomRPT & ") "
    'Para la exportacion a CSV
    Msg = " AND expedientes_lineas.codconce IN (" & cadNomRPT & ") " & Cad & "|"
    If HayRegParaInforme("expedientes,expedientes_lineas", SQL, True) Then I = 1
    cadParam = cadParam & "tieneExpedientes=" & I & "|"
    
    'en facturas
    Cad = Replace(cadselect, "@f", "fecfactu")
    SQL = "codconce IN (" & cadNomRPT & ") " & Cad
    Msg = Msg & " AND factcli_lineas.codconce IN (" & cadNomRPT & ") " & Cad & "|"
    If HayRegParaInforme("factcli_lineas", SQL, True) Then I = 1
    cadParam = cadParam & "tieneFacturas=" & I & "|"
    numParam = numParam + 2
    If I = 0 Then
        MsgBox "No existen datos entre las fechas", vbExclamation
        Exit Sub
    End If
    
    
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
        CargaConceptos
    End If
    Screen.MousePointer = vbDefault
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

    Screen.MousePointer = vbHourglass
    Me.Icon = frmppal.Icon
        
    'Otras opciones
    Me.Caption = "Listado por conceptos"

    
    PrimeraVez = True
     
    
    'txtFecha(0).Text = vParam.fechaini
    'txtFecha(1).Text = vParam.fechafin
     
    
    
    
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
End Sub



Private Sub frmF_Selec(vFecha As Date)
    txtFecha(IndCodigo).Text = Format(vFecha, formatoFechaVer)
End Sub


Private Sub imgCheck_Click(Index As Integer)

    For I = 1 To Me.lwConce.ListItems.Count
        lwConce.ListItems(I).Checked = Index = 1
    Next
End Sub

Private Sub imgFec_Click(Index As Integer)
    
    Screen.MousePointer = vbHourglass
    
    Select Case Index
    Case 0, 1
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
    If Index = 0 Then
        cd1.Filter = "*.csv|*.csv"

    Else
        cd1.Filter = "*.pdf|*.pdf"
    End If
    cd1.InitDir = App.Path & "\Exportar" 'PathSalida
    cd1.FilterIndex = 1
    cd1.ShowSave
    If cd1.FileTitle <> "" Then
        If Dir(cd1.FileName, vbArchive) <> "" Then
            If MsgBox("El archivo ya existe. Reemplazar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
        txtTipoSalida(Index + 1).Text = cd1.FileName
    End If
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
    SQL = "Select conceptos.codconce,conceptos.nomconce,1 EXPE ,expedientes_lineas.numserie,expedientes_lineas.numexped,fecexped,"
    SQL = SQL & "expedientes.codclien,nomclien,importe"
    SQL = SQL & " FROM expedientes,expedientes_lineas,conceptos,clientes WHERE"
    SQL = SQL & " expedientes.TipoRegi = expedientes_lineas.TipoRegi And expedientes.numexped = expedientes_lineas.numexped"
    SQL = SQL & " AND  expedientes.anoexped = expedientes_lineas.anoexped AND expedientes_lineas.codconce=conceptos.codconce"
    SQL = SQL & " AND clientes.codclien=expedientes.codclien"
    SQL = SQL & " " & RecuperaValor(Msg, 1)
    SQL = SQL & " UNION "
    
    
    SQL = SQL & " Select conceptos.codconce,conceptos.nomconce,0 as EXPE,factcli.numserie,"
    SQL = SQL & " FACTCLI.NumFactu , factcli_lineas.Fecfactu, FACTCLI.CodClien, nomclien, Importe"
    SQL = SQL & " FROM factcli,factcli_lineas,conceptos,CLIENTES WHERE factcli.numserie = factcli_lineas.numserie And"
    SQL = SQL & " factcli.numfactu = factcli_lineas.numfactu AND  factcli.fecfactu = factcli_lineas.fecfactu AND"
    SQL = SQL & " factcli_lineas.codconce = conceptos.codconce AND FACTCLI.CODCLIEN=clientes.codclien"
    SQL = SQL & " " & RecuperaValor(Msg, 2)
    
    SQL = SQL & " ORDER BY 1,3,6"
    'LLamos a la funcion
    GeneraFicheroCSV SQL, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal()
Dim indRPT As String
Dim nomDocu As String
    
    vMostrarTree = False
    conSubRPT = False
        
    
    
    'Borrador
    Cad = ""

    Cad = "pdh1= """ + Cad + """|"
    
    Cad = Cad & "pdh2= """ + "" + """|"
    Cad = Cad & "Emp= """ & vEmpresa.nomempre & """|"
    
    
'    If optLog(0).Value Then
'        indRPT = "1412-00" '"rListLogCon.rpt"
'    Else
'        indRPT = "1412-01" '"rListLogConTr.rpt"
'    End If

    
   ' If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = "rEstaConceptos.rpt"


    

    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then CopiarFicheroASalida False, txtTipoSalida(2).Text, False
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 2
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
End Sub








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



Private Sub txtTipoSalida_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub





Private Sub CargaConceptos()
    lwConce.ListItems.Clear
    Set miRsAux = New ADODB.Recordset
    SQL = "Select distinct codconce from expedientes_lineas "
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    While Not miRsAux.EOF
        Cad = Cad & ", " & miRsAux!codconce
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    SQL = ""
    If Cad <> "" Then
        SQL = Mid(Cad, 2)
        SQL = " WHERE NOT codconce in (" & SQL & ")"
    End If
    SQL = "select distinct codconce from factcli_lineas " & SQL
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Cad = Cad & ", " & miRsAux!codconce
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    SQL = ""
    If Cad <> "" Then
        SQL = Mid(Cad, 2)
        SQL = " WHERE  codconce in (" & SQL & ")"
    End If
    SQL = "select  codconce,nomconce ,tipoconcepto from conceptos " & SQL
    SQL = SQL & " ORDER BY 2"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    While Not miRsAux.EOF
        I = I + 1
        lwConce.ListItems.Add , "K" & miRsAux!codconce, Format(miRsAux!nomconce, "0000")
        'lwConce.ListItems(I).SubItems(1) = miRsAux!nomconce
        lwConce.ListItems(I).Checked = Val(miRsAux!tipoconcepto) = 0
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    Set miRsAux = Nothing
    
End Sub
