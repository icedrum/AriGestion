VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImprimir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión listados"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   Icon            =   "frmImprimir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   6780
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pg1 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2340
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdConfigImpre 
      Caption         =   "Sel. &impresora"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   2340
      Width           =   1275
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5340
      TabIndex        =   1
      Top             =   2340
      Width           =   1275
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Default         =   -1  'True
      Height          =   375
      Left            =   3900
      TabIndex        =   0
      Top             =   2340
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   6435
      Begin VB.CheckBox chkEMAIL 
         Caption         =   "Enviar e-mail"
         Height          =   195
         Left            =   4920
         TabIndex        =   8
         Top             =   180
         Width           =   1335
      End
      Begin VB.CheckBox chkSoloImprimir 
         Caption         =   "Previsualizar"
         Height          =   255
         Left            =   420
         TabIndex        =   5
         Top             =   180
         Width           =   1275
      End
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Sin definir"
      Top             =   180
      Width           =   6315
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   240
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   960
      TabIndex        =   7
      Top             =   1320
      Width           =   5535
   End
End
Attribute VB_Name = "frmImprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Integer
    'O .- Conceptos
    '1 .- Cuentas
    '2 .- Tipos diario
    '3 .- Consulta extracto cuentas
    '4 .- Asientos pendientes de actualizar
    '5 .- Cuentas. Informe 1. Listado de cuentas
    '6 .- Cuentas. Informe 2. Ctas ultimo nivel. Apaisado
    '7 .- Asientos predefinidos
    '8 a 11  Listados de consulta de extracto
    '         4 Tipos.- 1- Normal
                        'LAs otras las conseguimos viendo
    '                   2- Con marcas de punteo(punteados y no punteados
    '                   3- Solo punteado
    '                   4- Solo pendiente
    '12 .- Asiento desde historico. Solo 1
    '13 .- Totales por cuenta y concepto
    '14 .-  "  "            "       "    desglosado por mes
    '15 .- Balance de sumas y saldos tipo 1    sin apertura sin movimientos
    '16 .-  " "         "        "    "   2    sin    "     con     "
    '17 .-  " "         "        "    "   3    con    "     sin     "
    '18 .-  " "         "        "    "   4    con    "     sin     "
    '19 .- Cuenta de explotacion. Con movimientos
    '20 .-   "     "       "       Sin   "
    '21 .- Listado facturas clientes
    '22 .- Listado de presupuestos
    '23 .- Balance presupuestario anual
    '24 .- Balance presupuestario mensual
    
    '-------------------------------------------------
    
    '25 .- Simulacion inmovilizado
    '26 .- Estadisticas inmovilizado
    '27 .- Fichas eltos inmov
    '28 .- Estadisticas einmov entre fechas
    '29 .- Certificado de IVA
    '30 .- Liquidacion IVA. Borrador
    '31 .- Listado facturas proveedores
    
    '32 .- Libro diario oficial
    '33 .- Saldos Centros de coste
    '34 .- cta explotacion por centro de coste
    '35 .- cta explotacion por centro de coste con movimietnos post(apaisado
    
    '36 .- Centro de coste por cuenta de explotacion
    '37 .-   "         "                  "           con movimientos periodo posterior. Apaisado
    
    
    
    '39 .-  Simulacion del apunte de pyg , cierre y apertura
    '40 .-  Libro diario resumen
    
    '41.- Centro de coste por cuenta
    
    '42.- Borrador MOD347 del IVA
    '43.- Carta clientes /proveedores 347
        
    '44.- Expltoacion comparativa
    '45.- Expltoacion comparativa porentual
    
    '46.- Balance consolidado
    '47.-   "          "       con desglose
    
    
    '48.- Factura Inmovilizado
    
            'BALANCES
    '49.- Balance con descripcion y solo
    '50.-   "           "         comparativo
    '51.-   "       SIN descripcion y solo
    '52.-   "       SIN    "      comparativo
    
    '53.- Listado HCO inmovilizado a partir de la seleccion
    '54.- Listado CONCEPTOS inmoivlizado
    
    '55.- Configurador memoria ejercicio
    
    '56.- Modelo 349 Operaciones intracomunitarias
    
    '57.- Facturas clientes agrupadas
    '58.- Facturas proveedores agrupadas
    
    '59.- Cta explotacion CONSOLIDADA MOVIMIENTOS sin desglose
    '60.-        "                         "      CON
    '61.- Cta explotacion   "         SIN MOV     sin desglose
    '62.- Cta explotacion   "         SIN MOV     CON desglose
    
    '63.- Liquidacion IVA detallada
    '64.- Listado elementos amortizacion
    
    '65.- Libro diario si mostrar total asiento
    
    '66.- Asientos con errores.
    
    '67.- Listado cuentas por nombre
    '68.-  "         "   2   "
    
    
    '69.- Facturas proveedor consolidada   numreg
    '70.-    "           "                 por fecha
    
    '71.- Facturas clientes conso, numero
    '72.- Facturas clientes conso   fecha
    
    '73.- Configuracion balances
        
    '74.- Total por concepto.... SIN concepto
    '75.-  "     "    "               "         DESGLOSADO
    
    '76.- Evolucion mensaul de Saldo
        
    '77.- Relacion de clientes /proveedores por cuenta de ventas/compras
    '78.-    "          "           "          "            "             CON desglose
        
    '80.- MODELO 347 Agencias de viaje
    '81.- COnsulta extractos EXTENDIDA
    
    
    '82.- Relacion de clientes /proveedores por cuenta de ventas/compras COMPARATIVO
    
    
    
    'Balances personalizables en modo apaisado
    '83.- Balance con descripcion y solo
    '84.-   "           "         comparativo
    '85.-   "       SIN descripcion y solo
    '86.-   "       SIN    "      comparativo
    
    
    
    
    '90 .-  Cta explotacion comparativa MENSUAL
    '91 .-    ""
    '92 .-  Configuracion traspaso PGC 2008
    
    '93 .-  Borrador 340
    '94 .-  Listado centros de coste
    
    
    '96 .- Ratios
    '97 .- Graficas
    '98 .- Graficas resumen
    
    '100 ,101  LOg acciones
    '102  Memoria pagos
    
Public FormulaSeleccion As String
Public SoloImprimir As Boolean
Public OtrosParametros As String   ' El grupo acaba en |
                                   ' param1=valor1|param2=valor2|
Public NumeroParametros As Integer   'Cuantos parametros hay.  EMPRESA(EMP) no es parametro. Es fijo en todos los informes
Public EnvioEMail As Boolean
Public QueEmpresaEs As Byte
    '0 Todas
    '1 Escalon


Private MostrarTree As Boolean
Private Nombre As String
Private MIPATH As String
Private Lanzado As Boolean
Private PrimeraVez As Boolean


'Private ReestableceSoloImprimir As Boolean

Private Sub chkEmail_Click()
    If chkEMAIL.Value = 1 Then Me.chkSoloImprimir.Value = 0
End Sub

Private Sub chkSoloImprimir_Click()
    If Me.chkSoloImprimir.Value = 1 Then Me.chkEMAIL.Value = 0
End Sub

Private Sub cmdConfigImpre_Click()
    Screen.MousePointer = vbHourglass
    'Me.CommonDialog1.Flags = cdlPDPageNums
    CommonDialog1.ShowPrinter
    PonerNombreImpresora
    Screen.MousePointer = vbDefault
End Sub


Private Sub cmdImprimir_Click()
    If Me.chkSoloImprimir.Value = 1 And Me.chkEMAIL.Value = 1 Then
        MsgBox "Si desea enviar por mail no debe marcar vista preliminar", vbExclamation
        Exit Sub
    End If
    'Form2.Show vbModal
    Imprime
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub



Private Sub Form_Activate()
If PrimeraVez Then
    PrimeraVez = False
    espera 0.1
    CommitConexion
    'Primero veo si existe
    If SoloImprimir Then
        Imprime
        Unload Me
    Else
        If Dir(MIPATH & Nombre, vbArchive) = "" Then Me.cmdImprimir.Enabled = False
    End If
End If
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim cad As String

    Me.Icon = frmPpal.Icon

PrimeraVez = True
Lanzado = False
CargaICO
cad = Dir(App.Path & "\impre.dat", vbArchive)

'ReestableceSoloImprimir = False
If cad = "" Then
    chkSoloImprimir.Value = 0
    Else
    chkSoloImprimir.Value = 1
    'ReestableceSoloImprimir = True
End If
cmdImprimir.Enabled = True
If SoloImprimir Then
    chkSoloImprimir.Value = 0
    Me.Frame2.Enabled = False
    chkSoloImprimir.Visible = False
Else
    Frame2.Enabled = True
    chkSoloImprimir.Visible = True
End If
PonerNombreImpresora
MostrarTree = False

'A partir del infome 26, se trabajaba sobre la b de datos de informes(USUARIOS)


    MIPATH = App.Path & "\Informes\"

Select Case Opcion
Case 0
    Text1.Text = "Conceptos"
    Nombre = "Conceptos.rpt"

Case 1
    Text1.Text = "Cuentas"
    Nombre = "colCuentas.rpt"
    
Case 2
    Text1.Text = "Tipos diario"
    Nombre = "Diarios.rpt"
    
Case 3
    Text1.Text = "Consulta extracto de cuentas"
    'Nombre = "ConsExtrac.rpt"  ANTIGUO
    Nombre = "ConsExtracL1.rpt"     'Este tb se usa en la legalizacion. Abra k cambiarlo alli tb
    
Case 4
    Text1.Text = "Asientos pendientes de actualizar"
    Nombre = "DiaPendAct.rpt"
    
Case 5
    Text1.Text = "Listado de cuentas (I)"
    Nombre = "colCuentas.rpt"
    
Case 6
    Text1.Text = "Listado de cuentas (II). Ultimo nivel"
    Nombre = "colCuentas2.rpt"
    
Case 7
    Text1.Text = "Listado de asientos predefinidos."
    Nombre = "asipre.rpt"
    MostrarTree = True
    
Case 8
    Text1.Text = "Listado de extractos de cuentas (Normal)."
    Nombre = "ConsExtracL1.rpt"
    MostrarTree = True

Case 9
    Text1.Text = "Listado de extractos de cuentas (Punteados)."
    Nombre = "ConsExtracL2.rpt"
    MostrarTree = True
    
Case 10
    Text1.Text = "Listado de extractos de cuentas (Solo Pun.)."
    Nombre = "ConsExtracL2.rpt"
    MostrarTree = True
    
Case 11
    Text1.Text = "Listado de extractos de cuentas (Sin puntear)."
    Nombre = "ConsExtracL2.rpt"
    MostrarTree = True
    
Case 12
    Text1.Text = "Listado de asientos."
    Nombre = "AsientoHco.rpt"

Case 13
    Text1.Text = "Totales por cuenta y concepto."
    Nombre = "TotCtaCon.rpt"
    
Case 14
    Text1.Text = "Totales por cuenta y concepto desglosado."
    Nombre = "TotCtaConDesglose.rpt"

Case 15
    Text1.Text = "Balance de sumas y saldos( Sin ape. sin mov.)"
    Nombre = "Sumas1.rpt"
Case 16
    Text1.Text = "Balance de sumas y saldos( Sin ape. con mov.)"
    Nombre = "Sumas2.rpt"
Case 17
    Text1.Text = "Balance de sumas y saldos( Con ape. sin mov.)"
    Nombre = "Sumas3.rpt"
Case 18
    Text1.Text = "Balance de sumas y saldos( Con  ape. con mov.)"
    Nombre = "Sumas4.rpt"
Case 19
    Text1.Text = "Cuenta de explotación. Con movimientos mes."
    Nombre = "ctaexplot1.rpt"
Case 20
    Text1.Text = "Cuenta de explotación."
    Nombre = "ctaexplot2.rpt"
Case 21
    Text1.Text = "Listado facturas clientes"
    'Nombre = "faccli1.rpt"
    Nombre = "faccli2.rpt"
Case 22
    Text1.Text = "Listado presupuestos"
    Nombre = "presu1.rpt"
    MostrarTree = True
Case 23
    Text1.Text = "Balance presupuestario ANUAL"
    Nombre = "presu2.rpt"
Case 24
    Text1.Text = "Balance presupuestario MENSUAL"
    Nombre = "presu3.rpt"
    MostrarTree = True
'-------------------------------- EN NUEVA CARPETA CONTRA ODBC vinculado a BD USAURIOS
Case 25
    Text1.Text = "Simulación amortización"
    Nombre = "simula.rpt"
    MostrarTree = True

Case 26
    Text1.Text = "Estadísticas amortización"
    Nombre = "estadisinmo1.rpt"
    MostrarTree = True
    
Case 27
    Text1.Text = "Fichas elementos inmovilizado"
    Nombre = "fichaelto.rpt"
    MostrarTree = True
    
Case 28
    Text1.Text = "Inmovilizado entre fechas"
    Nombre = "entrefechas.rpt"
    MostrarTree = True

Case 29
    Text1.Text = "Certificado de IVA"
    Nombre = "certiva.rpt"
    
Case 30
    Text1.Text = "Liquidación de IVA"
    Nombre = "liquidiva.rpt"
    
Case 31
    Text1.Text = "Listado facturas proveedores"
    'Nombre = "facprov.rpt"
    Nombre = "facprov2.rpt"
Case 32
    Text1.Text = "Libro diario oficial"
    Nombre = "DiarioOf.rpt"    'Si lo cambiamos hay que cambiar tb en
                                'listado ... Legalizacion libros

Case 33
    Text1.Text = "Saldos centros de coste"
    Nombre = "saldoscc.rpt"

Case 34
    Text1.Text = "Cta explotacion por CC"
    Nombre = "ctaexpcc.rpt"

Case 35
    Text1.Text = "Cta explotacion por CC( con acum. posterior)"
    Nombre = "ctaexpcc2.rpt"

Case 36
    Text1.Text = "Centro de coste por cta. explotacion"
    Nombre = "ccosctaexp.rpt"
    
Case 37
    Text1.Text = "Centro de coste por cta. explotacion( con acum. posterior)"
    Nombre = "ccosctaexp2.rpt"
    
Case 39
    Text1.Text = "Simulación del cierre de ejercicio."
    Nombre = "simcierre.rpt"
    
Case 40
    Text1.Text = "Libro diario resumen."
    Nombre = "resumen.rpt"
    
    
Case 41
    Text1.Text = "Cuentas por centro de coste(Detalle)"
    Nombre = "cc_x_cta.rpt"
    MostrarTree = True
    
Case 42
    Text1.Text = "Modelo 347. Borrador"
    Nombre = "mod347.rpt"
    
Case 43
    Text1.Text = "Carta Clientes/proveedores Mod 347."
    'Nombre = "carta.rpt"
    Nombre = DevNombreInformeCrystal(1) 'El uno es la carta
Case 44
    Text1.Text = "Listado explotación comparativa"
    Nombre = "explocomp.rpt"
    
Case 45
    Text1.Text = "Listado explotación comparativa porcentual"
    Nombre = "explocomp2.rpt"
        
Case 46
    Text1.Text = "Balance consolidado"
    Nombre = "consolidadoA.rpt"
    MostrarTree = True
    
Case 47
    Text1.Text = "Balance consolidado con desglose por empresa"
    Nombre = "consolidado1.rpt"
    MostrarTree = True

    
Case 48
    Text1.Text = "Factura inmovilizado"
    'Nombre = "facturaI.rpt"
    Nombre = DevNombreInformeCrystal(2)
Case 49
    Text1.Text = "Balances"
    Nombre = "Balance1.rpt"


Case 50
    Text1.Text = "Balances. Comparativo"
    Nombre = "Balance2.rpt"


Case 51
    Text1.Text = "Balances SIN descripcion. "
    Nombre = "Balance3.rpt"


Case 52
    Text1.Text = "Balances SIN descripcion. Comp."
    Nombre = "Balance4.rpt"



Case 53
    Text1.Text = "Listado HCO Inmovilizado"
    Nombre = "hcoinmov.rpt"
    MostrarTree = True

Case 54
    Text1.Text = "Conceptos INMOVILIZADO"
    Nombre = "Conceinm.rpt"


Case 55
    Text1.Text = "Configuración memoria ejercicio"
    Nombre = "Memoria.rpt"


Case 56
    Text1.Text = "MODELO 349. Intracomunitarias"
    Nombre = "Mod349.rpt"

Case 57
    Text1.Text = "Listado facturas clientes agrupadas"
    'Nombre = "faccli1g.rpt"
    Nombre = "faccli2g.rpt"
    MostrarTree = True
    
Case 58
    Text1.Text = "Listado facturas proveedores agrupadas"
    'Nombre = "facprovg.rpt"
    Nombre = "facprov2g.rpt"
    MostrarTree = True


Case 59 To 62
    Text1.Text = "Cta explotación conosolidada. "
    Select Case Opcion
    Case 59
        cad = "11"
        Text1.Text = Text1.Text & "SIN Movimientos"
    Case 60
        cad = "12"
        Text1.Text = Text1.Text & "SIN movi. Desglose"
    Case 61
        cad = "21"
        Text1.Text = Text1.Text & "Con movimientos."
    Case Else
        cad = "22"
        Text1.Text = Text1.Text & "Con movi. Desglose"
    End Select
    Nombre = "ctaexplotC" & cad & ".rpt"
    
Case 63
    Text1.Text = "Liquidación IVA Detallada"
    Nombre = "liquidivaD.rpt"
    
Case 64
    Text1.Text = "Listado eltos. amortizacion"
    Nombre = "inmov.rpt"
    
Case 65
   Text1.Text = "Listado de asientos sin totales"
    Nombre = "AsientoHco2.rpt"
    
Case 66
   Text1.Text = "Asientos con errores"
    Nombre = "asierror.rpt"

Case 67
    Text1.Text = "Listado de cuentas (I).Ord. Nombre"
    Nombre = "colCuentas3.rpt"
    
Case 68
    Text1.Text = "Listado de cuentas (II). Ultimo nivel. (Nombre)"
    Nombre = "colCuentas4.rpt"

Case 69, 70
    Text1.Text = "L. Facturas Proveedores Consolidado"
    If Opcion = 69 Then
        Nombre = "facpr_n.rpt"
    Else
        Nombre = "facpr_f.rpt"
    End If
Case 71, 72
    Text1.Text = "L Facturas clientes Consolidado"
    If Opcion = 71 Then
        Nombre = "faccl_n.rpt"
    Else
        Nombre = "faccl_f.rpt"
    End If
    
    
Case 73
    Text1.Text = "Configuración balances"
    Nombre = "Balance1c.rpt"
    
Case 74
    Text1.Text = "Totales por cuenta y concepto(TODOS)."
    Nombre = "TotCtaConTotal.rpt"
    MostrarTree = True
Case 75
    Text1.Text = "Totales por cuenta y concepto. Desglose. (TODOS)."
    Nombre = "TotCtaConTotalD.rpt"
    MostrarTree = True
    
Case 76
    Text1.Text = "Evolución mensual de saldos."
    Nombre = "evolsald.rpt"
    MostrarTree = True
    
Case 77, 78, 82
    Text1.Text = "Client/Proveed por cta ventas/compras"
    Nombre = "ctaxbase"
    If Opcion = 78 Then
        Text1.Text = Text1.Text & " DESGLOSE"
        Nombre = Nombre & "des"
    ElseIf Opcion = 82 Then
        Text1.Text = Text1.Text & " COMPARATIVO"
        Nombre = Nombre & "com"
    End If
    Nombre = Nombre & ".rpt"
    MostrarTree = True


    
Case 80
    Text1.Text = "Modelo 347 Agencias Viaje. Borrador"
    Nombre = "mod347Ag.rpt"
    
Case 81
    Text1.Text = "Listado de extractos de cuentas (EXTENDIDO)."
    Nombre = "ConsExtracExt.rpt"
    MostrarTree = True
    
    
    
    
'Balacnces APAISADOS para el nuevo plan contable 2008
Case 83
    Text1.Text = "Balances (PGC08)"
    Nombre = "Balance1a.rpt"


Case 84
    Text1.Text = "Balances. Comparativo(PGC08)"
    Nombre = "Balance2a.rpt"


Case 85
    Text1.Text = "Balances SIN descripcion. (PGC08)"
    Nombre = "Balance3a.rpt"


Case 86
    Text1.Text = "Balances SIN descripcion. Comp.(PGC08)"
    Nombre = "Balance4a.rpt"

    
Case 90, 91
    Text1.Text = "Analitica comparativo"
    Nombre = "ctaexpcc_c"
    If Opcion = 90 Then
        Text1.Text = Text1.Text & " mensual"
        Nombre = Nombre & "M"
    Else
        Text1.Text = Text1.Text & " anual"
        Nombre = Nombre & "A"
    End If
    Nombre = Nombre & ".rpt"
    
    
Case 92
    Text1.Text = "Conf. traspaso PGC2008"
    Nombre = "traspaso.rpt"
    
Case 93
    Text1.Text = "Borrador modelo 340"
    Nombre = "mod340.rpt"
    
Case 94
    Text1.Text = "Listado centros de coste"
    Nombre = "ccoste.rpt"
    
    
    '96 .- Ratios
    '97 .- Graficas
    '98 .- Graficas resumen
Case 96
    Text1.Text = "Ratios"
    Nombre = "ratios.rpt"
    
Case 97
    Text1.Text = "Gráficas"
    Nombre = "graficas.rpt"
    
Case 98
    Text1.Text = "Grafica resumen"
    Nombre = "GraficaR.rpt"
    
    
Case 100, 101
    Text1.Text = "LOG"
    Nombre = "rListLogCon"
    If Opcion = 101 Then Nombre = Nombre & "Tr"
    Nombre = Nombre & ".rpt"
    
Case 102
    Text1.Text = "Memoria pagos proveedor"
    Nombre = "memopago.rpt"
    
    
Case Else
    Text1.Text = "Opcion incorrecta (" & Opcion & ")"
    Me.cmdImprimir.Enabled = False
End Select



Screen.MousePointer = vbDefault
End Sub




Private Function Imprime() As Boolean
Dim Seguir As Boolean
'Dim Nom As String
'Dim i As Integer
'
'On Error GoTo ErrImprime
'
'
'
'Screen.MousePointer = vbHourglass
'Imprime = False
''Parece ser que hay que esperar
'If Me.chkSoloImprimir.Value = 0 Then
'        If Me.chkEMAIL.Value = 1 Then
'            CR1.Destination = crptToFile
'            Nom = App.path & "\Salida.htm"
'            If Dir(Nom) <> "" Then Kill Nom
'            CR1.PrintFileName = Nom
'        Else
'            CR1.Destination = crptToPrinter
'        End If
'    Else
'        CR1.Destination = crptToWindow
'        CR1.WindowTitle = "ARICONTA -Informes"
'        CR1.WindowState = crptMaximized
'End If
'
'
'
''Modificacion para CR con user pwd
'Nom = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=""DSN=vUsuarios;DESC=MySQL ODBC 3.51 Driver DSN;"
'Nom = Nom & "DATABASE=usuarios;SERVER=" & vConfig.SERVER & ";UID=" & vConfig.User & ";PASSWORD="
'Nom = Nom & vConfig.password & ";PORT=3306;OPTION=3;STMT=;"""
'
'CR1.Connect = Nom
'CR1.LogOnServer "", vConfig.SERVER, "Usuarios", vConfig.User, vConfig.password
'
'
''Anterior a 11 de Abril





'
'
'CR1.ReportFileName = Nom
''Por si hay parametros
'CR1.Formulas(0) = "Emp= """ & vEmpresa.nomempre & """"
'
'If Me.NumeroParametros > 0 Then
'    For i = 1 To Me.NumeroParametros
'        CR1.Formulas(i) = RecuperaValor(Me.OtrosParametros, i)
'    Next i
'End If
'CR1.WindowShowGroupTree = MostrarTree
'
'
''Si hay formulas
'CR1.SelectionFormula = FormulaSeleccion
'CR1.Action = 1
'If Screen.Width < 9100 Then CR1.PageZoom 90
'Lanzado = True
'Imprime = True
'ErrImprime:
'    If Err.Number <> 0 Then _
'        MuestraError Err.Number, Err.Description & vbCrLf & vbCrLf & vbCrLf, Err.Description
 
'
'
'Public FormulaSeleccion As String
'Public SoloImprimir As Boolean
'Public OtrosParametros As String   ' El grupo acaba en |
'                                   ' param1=valor1|param2=valor2|
'Public NumeroParametros As Integer   'Cuantos parametros hay.  EMPRESA(EMP) no es parametro. Es fijo en todos los informes
'
'
'Private MostrarTree As Boolean

    If Dir(MIPATH & Nombre) = "" Then
        MsgBox "Archivo no encontrado: " & MIPATH & Nombre, vbExclamation
        Exit Function
    End If

    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    With frmVisReport
        
        .FormulaSeleccion = Me.FormulaSeleccion
        .SoloImprimir = (Me.chkSoloImprimir.Value = 0)
        .OtrosParametros = OtrosParametros
        .NumeroParametros = NumeroParametros
        .MostrarTree = MostrarTree
        .Informe = MIPATH & Nombre
        .ExportarPDF = (chkEMAIL.Value = 1)
        
        .Show vbModal
    End With
    
    If Me.chkEMAIL.Value = 1 Then
'        If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
    End If
    Unload Me
 
 
 
End Function


Private Sub Form_Unload(Cancel As Integer)
    If Me.chkEMAIL.Value = 1 Then Me.chkSoloImprimir.Value = 1
    'If ReestableceSoloImprimir Then SoloImprimir = False
    OperacionesArchivoDefecto
End Sub

Private Sub OperacionesArchivoDefecto()
Dim Crear  As Boolean
On Error GoTo ErrOperacionesArchivoDefecto

Crear = (Me.chkSoloImprimir.Value = 1)
'crear = crear And ReestableceSoloImprimir
If Not Crear Then
    Kill App.Path & "\impre.dat"
    Else
        FileCopy App.Path & "\Vacio.dat", App.Path & "\impre.dat"
End If
ErrOperacionesArchivoDefecto:
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Text1_DblClick()
Frame2.Tag = Val(Frame2.Tag) + 1
If Val(Frame2.Tag) > 2 Then
    Frame2.Enabled = True
    chkSoloImprimir.Visible = True
End If
End Sub

Private Sub PonerNombreImpresora()
On Error Resume Next
    Label1.Caption = Printer.DeviceName
    If Err.Number <> 0 Then
        Label1.Caption = "No hay impresora instalada"
        Err.Clear
    End If
End Sub

Private Sub CargaICO()
    On Error Resume Next
    Image1.Picture = LoadPicture(App.Path & "\iconos\printer.ico")
    If Err.Number <> 0 Then Err.Clear
End Sub

