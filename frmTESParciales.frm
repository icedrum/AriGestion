VERSION 5.00
Begin VB.Form frmTESParciales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anticipo vto."
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   8430
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
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
      Index           =   1
      Left            =   6960
      TabIndex        =   6
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
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
      Left            =   5520
      TabIndex        =   5
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Frame FrCobro 
      Height          =   5175
      Left            =   60
      TabIndex        =   7
      Top             =   90
      Width           =   8175
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   240
         TabIndex        =   28
         Top             =   1200
         Width           =   7575
      End
      Begin VB.TextBox Text2 
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
         Left            =   6000
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   4290
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
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
         ItemData        =   "frmTESParciales.frx":0000
         Left            =   1590
         List            =   "frmTESParciales.frx":0002
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "Tipo de pago|N|N|||formapago|tipforpa|||"
         Top             =   4260
         Width           =   2475
      End
      Begin VB.TextBox txtCta 
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
         Left            =   1680
         TabIndex        =   0
         Text            =   "Text2"
         Top             =   1470
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtDescCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
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
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "Text2"
         Top             =   1470
         Visible         =   0   'False
         Width           =   4785
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Index           =   5
         Left            =   6030
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   2940
         Width           =   1755
      End
      Begin VB.TextBox Text2 
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
         Left            =   6000
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   3855
         Width           =   1755
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Index           =   4
         Left            =   6030
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   2490
         Width           =   1755
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Index           =   3
         Left            =   6030
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   2010
         Width           =   1755
      End
      Begin VB.TextBox Text3 
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
         Left            =   1590
         TabIndex        =   1
         Top             =   3825
         Width           =   1305
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Gasto Bancario"
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
         Height          =   240
         Index           =   10
         Left            =   4350
         TabIndex        =   27
         Top             =   4335
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   9
         Left            =   240
         TabIndex        =   26
         Top             =   4290
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Datos vto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   56
         Left            =   270
         TabIndex        =   20
         Top             =   360
         Width           =   6150
      End
      Begin VB.Label Label4 
         Caption         =   "Datos vto"
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
         Height          =   240
         Index           =   57
         Left            =   270
         TabIndex        =   19
         Top             =   720
         Width           =   6270
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   7860
         Y1              =   4710
         Y2              =   4710
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   3330
         Width           =   6195
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cta banco"
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
         Height          =   240
         Index           =   7
         Left            =   270
         TabIndex        =   17
         Top             =   1470
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   1
         Left            =   1380
         Picture         =   "frmTESParciales.frx":0004
         Top             =   1530
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Importe"
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
         Height          =   240
         Index           =   6
         Left            =   4380
         TabIndex        =   15
         Top             =   3900
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Pagado"
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
         Height          =   240
         Index           =   5
         Left            =   5100
         TabIndex        =   14
         Top             =   2940
         Width           =   720
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   7860
         Y1              =   3690
         Y2              =   3690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Gastos"
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
         Height          =   240
         Index           =   4
         Left            =   5160
         TabIndex        =   11
         Top             =   2550
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Importe TOTAL"
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
         Height          =   240
         Index           =   2
         Left            =   4380
         TabIndex        =   9
         Top             =   2100
         Width           =   1500
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   3
         Left            =   240
         TabIndex        =   8
         Top             =   3840
         Width           =   600
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   0
         Left            =   1230
         Picture         =   "frmTESParciales.frx":6856
         Top             =   3870
         Width           =   240
      End
   End
   Begin VB.TextBox Text1 
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
      Index           =   2
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   1080
      Width           =   3495
   End
   Begin VB.TextBox Text1 
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
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
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
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   600
      Width           =   4815
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Vencimiento"
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
      Height          =   240
      Index           =   0
      Left            =   360
      TabIndex        =   25
      Top             =   600
      Width           =   1200
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
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
      Height          =   240
      Index           =   1
      Left            =   360
      TabIndex        =   24
      Top             =   1080
      Width           =   675
   End
End
Attribute VB_Name = "frmTESParciales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit





Public Vto As String  'Llevara empipado las claves
Public Cta As String
Public Importes As String 'Empipado los importes
Public FormaPago As Integer

Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1


Dim impo As Currency
Dim Cad As String
Dim PrimeraVez As Boolean
Dim TipForpa As Integer

Dim LineaCobro As Long

Private Sub PonFoco(ByRef T1 As TextBox)
    T1.SelStart = 0
    T1.SelLength = Len(T1.Text)
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub






Private Sub CargarTemporal()
Dim Sql As String

    Sql = "delete from tmppendientes where codusu = " & vUsu.Codigo
    Conn.Execute Sql

    ' en tmppendientes metemos la clave primaria de cobros_recibidos y el importe en letra
                                                      'importe=nro factura,   codforpa=linea de cobros_realizados
    Sql = "insert into tmppendientes (codusu,serie_cta,importe,fecha,numorden,codforpa, observa) values ("
    Sql = Sql & vUsu.Codigo & "," & DBSet(RecuperaValor(Vto, 1), "T") & "," 'numserie
    Sql = Sql & DBSet(RecuperaValor(Vto, 2), "N") & "," 'numfactu
    Sql = Sql & DBSet(RecuperaValor(Vto, 3), "F") & "," 'fecfactu
    Sql = Sql & DBSet(RecuperaValor(Vto, 4), "N") & "," 'numorden
    Sql = Sql & DBSet(LineaCobro, "N") & "," 'numlinea
    Sql = Sql & DBSet(EscribeImporteLetra(ImporteFormateado(Text2(0).Text)), "T") & ") "
    
    Conn.Execute Sql

End Sub


Private Sub Command1_Click(Index As Integer)
Dim B As Boolean
    
    CadenaDesdeOtroForm = ""
    If Index = 0 Then
        'Comprobamos importes. Y fecha de contabilizacioon
        If Not DatosOK Then Exit Sub
        
        CadenaDesdeOtroForm = "cobro"
        CadenaDesdeOtroForm = "Desea generar el " & CadenaDesdeOtroForm & "?"
        B = True
        If MsgBox(CadenaDesdeOtroForm, vbQuestion + vbYesNo) = vbNo Then B = False
        CadenaDesdeOtroForm = ""
        If Not B Then Exit Sub
        
        'UPDATEAMOS EL Vencimiento y CONTABILIZAMOS EL COBRO/PAGO
        Screen.MousePointer = vbHourglass
        B = RealizarAnticipo
        Screen.MousePointer = vbDefault
        If Not B Then Exit Sub
        CadenaDesdeOtroForm = "OK" 'Para que refresque los datos en el form
        

    End If
    
    Unload Me
End Sub



Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
            
        
    
        PosicionarCombo Combo1, FormaPago
    
    End If
        
End Sub

Private Sub Form_Load()
        
    Me.Icon = frmppal.Icon
        
    PrimeraVez = True
    Me.Frame1.BorderStyle = 0
    
        Caption = "Cobro"
        Text1(0).Text = RecuperaValor(Vto, 1) & "/" & RecuperaValor(Vto, 2) & "   Fecha: " & RecuperaValor(Vto, 3) & "   Vto. num: " & RecuperaValor(Vto, 4)
        Text1(1).Text = RecuperaValor(Cta, 1)
        Text1(2).Text = RecuperaValor(Cta, 2)
        'Dos
        txtCta(1).Text = RecuperaValor(Cta, 3)
        Me.txtDescCta(1).Text = RecuperaValor(Cta, 4)
        
        'Importes
        Text1(3).Text = RecuperaValor(Importes, 1)
        Text1(4).Text = RecuperaValor(Importes, 2)
        Text1(5).Text = RecuperaValor(Importes, 3)
        Text3(0).Text = Format(Now, "dd/mm/yyyy")
        Label4(4).Caption = "Gastos"
        Label4(1).Caption = "Cliente"
                
        Label4(57).Caption = Text1(0).Text
        Label4(56).Caption = Text1(2)
        
       
    
    txtCta(1).Visible = True
    txtDescCta(1).Visible = True
    imgCuentas(1).Visible = True
    Label4(7).Visible = True
    'IMPORTE Restante
    
    impo = ImporteFormateado(Text1(3).Text) 'Vto
        'Gastos
        If Text1(4).Text <> "" Then impo = impo + ImporteFormateado(Text1(4).Text)
            
        'Ya cobrado
        If Text1(5).Text <> "" Then impo = impo - ImporteFormateado(Text1(5).Text)
        
    
    Label1.Caption = "Pendiente: " & Format(impo, FormatoImporte)
    
    CargaCombo
    
    Label4(4).Visible = True
    Text1(4).Visible = True
    Me.Height = Me.FrCobro.Height + 1200 '240 + Me.Command1(0).Height + 240
    Text2(0).Text = Format(impo, FormatoImporte)
    Text2(1).Text = "0,00"
    
    Caption = Caption & " de factura"
End Sub

Private Sub frmBa_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtCta(CInt(imgCuentas(1).Tag)).Text = RecuperaValor(CadenaSeleccion, 1)
    Me.txtDescCta(CInt(imgCuentas(1).Tag)).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text3(CInt(Text3(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub Image2_Click(Index As Integer)
    
    Set frmC = New frmCal
    frmC.Fecha = Now
    If Text3(Index).Text <> "" Then frmC.Fecha = CDate(Text3(Index).Text)
    Text3(0).Tag = Index
    frmC.Show vbModal
    Set frmC = Nothing
End Sub





Private Sub Text2_GotFocus(Index As Integer)
    PonFoco Text2(Index)
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text2_LostFocus(Index As Integer)
Dim Valor

    If Text2(Index).Text = "" Then Exit Sub
    If Not IsNumeric(Text2(Index).Text) Then
        MsgBox "Importe debe ser numérico", vbExclamation
        Text2(Index).Text = ""
        PonFoco Text2(Index)
    Else
        If InStr(1, Text2(Index).Text, ",") > 0 Then
            Valor = ImporteFormateado(Text2(Index).Text)
        Else
            Valor = CCur(TransformaPuntosComas(Text2(Index).Text))
        End If
        Text2(Index).Text = Format(Valor, FormatoImporte)
    End If
End Sub

Private Sub Text3_GotFocus(Index As Integer)
    PonFoco Text3(Index)
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text3_LostFocus(Index As Integer)
    Text3(Index).Text = Trim(Text3(Index))
    If Text3(Index) = "" Then Exit Sub
    If Not EsFechaOK(Text3(Index)) Then
        MsgBox "Fecha incorrecta: " & Text3(Index), vbExclamation
        Text3(Index).Text = ""
        Text3(Index).SetFocus
    End If
End Sub


Private Function DatosOK() As Boolean
Dim Im As Currency
Dim CtaBancoGastos As String


    On Error GoTo EDa
    DatosOK = False
    
    
    Cad = ""
    If Text2(0).Text = "" Then Cad = "importe"
    If Text3(0).Text = "" Then Cad = Cad & " fecha"
    If Cad <> "" Then
        MsgBox "Falta: " & Cad, vbExclamation
        Exit Function
    End If
    
    '----------------------------------
    'Junio 2011
    'YA dejamos cobros negativos
    Im = ImporteFormateado(Text2(0).Text)
    'If Im < 0 Then
    If Im = 0 Then
        MsgBox "importes CERO", vbExclamation
        Exit Function
    End If
    
    
    If txtCta(1).Text = "" Then
        MsgBox "Falta cuenta banco", vbExclamation
        Exit Function
    End If
        
    'Fecha dentro ejercicios
    If FechaCorrecta2(CDate(Text3(0).Text)) > 1 Then
        MsgBox "Fecha fuera de ejercicios", vbExclamation
        Exit Function
    End If
    
    If ComprobarCero(Text2(1).Text) <> 0 Then
        CtaBancoGastos = DevuelveDesdeBD("ctagastos", "bancos", "codmacta", txtCta(1), "T")
        If CtaBancoGastos = "" Then
            CtaBancoGastos = DevuelveDesdeBD("ctabenbanc", "paramtesor", "codigo", "1", "N")
        End If
        If CtaBancoGastos = "" Then
            MsgBox "Falta configurar la cuenta de gastos bancarios. Revise.", vbExclamation
            Exit Function
        End If
    End If
    

    impo = ImporteFormateado(Text1(3).Text) 'Vto
    'Gastos
    If Text1(4).Text <> "" Then
        Im = ImporteFormateado(Text1(4).Text)
        impo = impo + Im
    End If
    
    'Ya cobrado
    If Text1(5).Text <> "" Then
        Im = ImporteFormateado(Text1(5).Text)
        impo = impo - Im
    End If

    Im = ImporteFormateado(Text2(0).Text) 'Lo que voy a pagar
    Cad = ""
    If impo < 0 Then
        'Importes negativos
        If Im >= 0 Then
            Cad = "negativo"
        Else
            If Im < impo Then Cad = "X"
        End If
    Else
        If Im <= 0 Then
            Cad = "positivo"
        Else
            If Im > impo Then Cad = "X"
        End If
    End If
        
    If Cad <> "" Then
        
        If Cad = "X" Then
            Cad = "Importe a pagar mayor que el importe restante.(" & Format(Im, FormatoImporte) & " : " & Format(impo, FormatoImporte) & ")"
        Else
            Cad = "El importe debe ser " & Cad
        End If
        MsgBox Cad, vbExclamation
        Exit Function
    End If
        
    'Comprobaremos un par de cosillas
    'If CuentaBloqeada(RecuperaValor(Cta, 1), CDate(Text3(0).Text), True) Then Exit Function
        
    DatosOK = True
    Exit Function
EDa:
    MuestraError Err.Number, "Datos Ok"
End Function


Private Function RealizarAnticipo() As Boolean
Dim B As Boolean

    Conn.BeginTrans
    
    If Me.Combo1.ListIndex = 0 Then
        'A caja
        B = InsertarEnCaja
    Else
        B = Contabilizar
    End If
    
    If B Then
        Conn.CommitTrans
        RealizarAnticipo = True
    Else
        'Conn.RollbackTrans
        TirarAtrasTransaccion
        RealizarAnticipo = False
    End If

End Function


Private Function Contabilizar() As Boolean
Dim Mc As ContadoresConta
Dim FP As Ctipoformapago
Dim Sql As String
Dim Ampliacion As String
Dim Numdocum As String
Dim Conce As Integer
Dim LlevaContr As Boolean
Dim Im As Currency
Dim Debe As Boolean
Dim ElConcepto As Integer
Dim vNumDiari As Integer
Dim Situacion As Integer

Dim Gastos As Currency
Dim Importe As Currency
Dim CtaBancoGastos As String
Dim DescuentaImporteDevolucion As Boolean
Dim Sql5 As String


    On Error GoTo ECon
    Contabilizar = False
    
    Dim mcContador As Long
    '
    Set Mc = New ContadoresConta
    If Mc.ConseguirContador("0", CDate(Text3(0).Text) <= vEmpresa.FechaFinEjercicio, True) = 1 Then Exit Function

    Set FP = New Ctipoformapago
    If FP.Leer(Combo1.ListIndex) Then  ' antes forma de pago
       ' Set Mc = Nothing
        Set FP = Nothing
    End If
    
    
    'importe
    impo = ImporteFormateado(Text2(0).Text)
    
    'Inserto cabecera de apunte
    Sql = "INSERT INTO ariconta" & vParam.Numconta & ".hcabapu (numdiari, fechaent, numasien, obsdiari, feccreacion, usucreacion, desdeaplicacion) VALUES ("
    Sql = Sql & FP.diaricli
    vNumDiari = FP.diaricli
    Sql = Sql & ",'" & Format(Text3(0).Text, FormatoFecha) & "'," & Mc.Contador
    Sql = Sql & ",'"
    Sql = Sql & "Generado desde gestion el " & Format(Now, "dd/mm/yyyy hh:mm") & " por " & DevNombreSQL(vUsu.Nombre)
    If impo < 0 Then Sql = Sql & "  (ABONO)"
    Sql = Sql & "',"
    Sql = Sql & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'Arigestion: Contabilizar Cobros')"
    
    
    Conn.Execute Sql
        
        
    'Inserto en las lineas de apuntes
    Sql = "INSERT INTO ariconta" & vParam.Numconta & ".hlinapu (numdiari, fechaent, numasien, linliapu, "
    Sql = Sql & "codmacta, numdocum, codconce, ampconce,timporteD,"
    Sql = Sql & " timporteH, codccost, ctacontr, idcontab, punteada,"
    
    
    Sql = Sql & "numserie,numfaccl,fecfactu,numorden,tipforpa,reftalonpag,bancotalonpag) VALUES ("
    Sql = Sql & FP.diaricli
    Sql = Sql & ",'" & Format(Text3(0).Text, FormatoFecha) & "'," & Mc.Contador & ","
    
    
    'numdocum
    Numdocum = DevNombreSQL(RecuperaValor(Vto, 2))
    Numdocum = RecuperaValor(Vto, 1) & Format(Numdocum, "0000000")
    
    
    
    'Concepto y ampliacion del apunte
    Ampliacion = ""

    'CLIENTES
    Debe = False
    If impo < 0 Then
        'If Not vParam.Abononeg Then Debe = True
    End If
    If Debe Then
        Conce = FP.ampdecli
        LlevaContr = FP.ctrdecli = 1
        ElConcepto = FP.condecli
    Else
        ElConcepto = FP.conhacli
        Conce = FP.amphacli
        LlevaContr = FP.ctrhacli = 1
    End If

           
    'Si el importe es negativo y no permite abonos negativos
    'como ya lo ha cambiado de lado (dbe <-> haber)
    If impo < 0 Then
        'If Not vParam.Abononeg Then impo = Abs(impo)
    End If
       
           
    If Conce = 2 Then
       Ampliacion = Ampliacion & RecuperaValor(Vto, 3)  'Fecha vto
    ElseIf Conce = 4 Then
        'Contra partida
        Ampliacion = DevNombreSQL(txtDescCta(1).Text)
    Else
        
        If Conce = 1 Then Ampliacion = Ampliacion & FP.siglas & " "
        Ampliacion = Ampliacion & RecuperaValor(Vto, 1) & Format(RecuperaValor(Vto, 2), "0000000") '& "/" & Mid(RecuperaValor(Vto, 2), 1, 9)
    End If
    
    'Fijo en concepto el codconce
    Conce = ElConcepto
    Cad = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(Conce), "N")
    Ampliacion = Cad & " " & Ampliacion
    Ampliacion = Mid(Ampliacion, 1, 35)
    
    
    
    
    'Ahora ponemos linliapu codmacta numdocum codconce ampconce timported timporte codccost ctacontr idcontab punteada
    'Cuenta Cliente/proveedor
    Cad = "1,'" & Text1(1).Text & "','" & Numdocum & "'," & Conce & ",'" & DevNombreSQL(Ampliacion) & "',"
    'Importe cobro-pago
    ' nos lo dire "debe"
    If Not Debe Then
        Cad = Cad & "NULL," & TransformaComasPuntos(CStr(impo))
    Else
        Cad = Cad & TransformaComasPuntos(CStr(impo)) & ",NULL"
    End If
    'Codccost
    Cad = Cad & ",NULL,"
    If LlevaContr Then
        Cad = Cad & "'" & txtCta(1).Text & "'"
    Else
        Cad = Cad & "NULL"
    End If

        Cad = Cad & ",'COBROS',0,"
        Cad = Cad & DBSet(RecuperaValor(Vto, 1), "T") & "," '& RecuperaValor(Vto, 2) & ","

    
    Cad = Cad & DBSet(RecuperaValor(Vto, 2), "T") & "," & DBSet(RecuperaValor(Vto, 3), "F") & ","
    Cad = Cad & DBSet(RecuperaValor(Vto, 4), "N") & "," & DBSet(Combo1.ItemData(Combo1.ListIndex), "N") & "," & ValorNulo & "," & ValorNulo & ")"
    
    Cad = Sql & Cad
    Conn.Execute Cad
    
       
    'El banco    *******************************************************************************
    '---------------------------------------------------------------------------------------------
    
    'Vuelvo a fijar los valores
     'Concepto y ampliacion del apunte
    Ampliacion = ""
    
       'CLIENTES
        'Si el apunte va al debe, el contrapunte va al haber
        If Not Debe Then
            Conce = FP.ampdecli
            LlevaContr = FP.ctrdecli = 1
            ElConcepto = FP.condecli
        Else
            ElConcepto = FP.conhacli
            Conce = FP.amphacli
            LlevaContr = FP.ctrhacli = 1
        End If
           
           
    If Conce = 2 Then
       Ampliacion = Ampliacion & RecuperaValor(Vto, 3)  'Fecha vto
    ElseIf Conce = 4 Then
        'Contra partida
        Ampliacion = DevNombreSQL(Text1(2).Text)
    Else
        If Conce = 1 Then Ampliacion = Ampliacion & FP.siglas & " "
        Ampliacion = Ampliacion & RecuperaValor(Vto, 1) & Format(RecuperaValor(Vto, 2), "0000000") ' "/" & Mid(RecuperaValor(Vto, 2), 1, 9)
        
    End If
    
    
    Conce = ElConcepto
    Cad = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(Conce), "N")
    Ampliacion = Cad & " " & Ampliacion
    Ampliacion = Mid(Ampliacion, 1, 35)
    
    
    Gastos = 0
    If ComprobarCero(Text2(1).Text) <> 0 Then
        Gastos = ImporteFormateado(Text2(1).Text)
    End If
    
    DescuentaImporteDevolucion = False
    If Gastos > 0 Then
        Sql5 = txtCta(1)
       
        Sql5 = DevuelveDesdeBD("GastRemDescontad", "bancos", "codmacta", Sql5, "T")
       
        If Sql5 = "1" Then DescuentaImporteDevolucion = True
    End If
    Importe = impo
    If DescuentaImporteDevolucion Then
        Importe = impo - Gastos
    End If
    
    Cad = "2,'" & txtCta(1).Text & "','" & Numdocum & "'," & Conce & ",'" & Ampliacion & "',"
    'Importe cliente
    'Si el cobro/pago va al debe el contrapunte ira al haber
    If Not Debe Then
        'al debe
        Cad = Cad & TransformaComasPuntos(CStr(Importe)) & ",NULL"
    Else
        'al haber
        Cad = Cad & "NULL," & TransformaComasPuntos(CStr(Importe))
    End If
    
    'Codccost
    Cad = Cad & ",NULL,"
    
    If LlevaContr Then
        Cad = Cad & "'" & Text1(1).Text & "'"
    Else
        Cad = Cad & "NULL"
    End If
    
    Cad = Cad & ",'COBROS',0," ' idcontab
    
    
    ' todo valores a null ????
    Cad = Cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ")"
    
    Cad = Sql & Cad
    Conn.Execute Cad
    
        
    '++
    'Gasto.
    ' Si tiene y no agrupa
    '-------------------------------------------------------
    If Gastos > 0 Then
        If CtaBancoGastos = "" Then CtaBancoGastos = DevuelveDesdeBD("ctagastos", "bancos", "codmacta", txtCta(1), "T")
        If CtaBancoGastos = "" Then
            CtaBancoGastos = DevuelveDesdeBD("ctabenbanc", "paramtesor", "codigo", "1", "N")
        End If

        Cad = "3,'"

        Cad = Cad & CtaBancoGastos & "','" & Numdocum & "'," & Conce
        Cad = Cad & ",'Gastos vto.'"

        'Importe al debe
        Cad = Cad & "," & TransformaComasPuntos(CStr(Gastos)) & ",NULL,"

        'Codccost
        Cad = Cad & "NULL,"

        If LlevaContr Then
            If Not DescuentaImporteDevolucion Then
                Cad = Cad & "'" & txtCta(1).Text & "'"
            Else
                Cad = Cad & "'" & Text1(1).Text & "'"
            End If
        Else
            Cad = Cad & "NULL"
        End If

        Cad = Cad & ",'COBROS',0,"
        
        ' todo valores a null ????
        Cad = Cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ")"
        
        Cad = Sql & Cad
        Conn.Execute Cad
        
        
        If Not DescuentaImporteDevolucion Then
            Cad = "4,'"
    
            Cad = Cad & txtCta(1).Text & "','" & Numdocum & "'," & Conce
            Cad = Cad & ",'Gastos vto.'"
    
            'Importe al debe
            Cad = Cad & ",NULL, " & TransformaComasPuntos(CStr(Gastos)) & ","
    
            'Codccost
            Cad = Cad & "NULL,"
    
            If LlevaContr Then
                Cad = Cad & "'" & CtaBancoGastos & "'"
            Else
                Cad = Cad & "NULL"
            End If
    
            Cad = Cad & ",'COBROS',0,"
            
            ' todo valores a null ????
            Cad = Cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ")"
            
            
            Cad = Sql & Cad
            Conn.Execute Cad
        
        End If
        
    End If
    '++
    
    
    Sql = FP.diaricli
    
    
    'Actualizamos VTO
    ' o lo eliminamos. Segun sea el importe que falte
    'Tomomos prestada LlevaContr
    
    Im = ImporteFormateado(Text2(0).Text)  'lo que voy a anticipar
    
    impo = ImporteFormateado(Text1(3).Text)  'lo que me falta
    
        If Text1(4).Text <> "" Then impo = impo + ImporteFormateado(Text1(4).Text)
        If Text1(5).Text <> "" Then impo = impo - ImporteFormateado(Text1(5).Text)
    
    If impo - Im = 0 Then
        LlevaContr = True  'ELIMINAR VTO ya que esta totalmente pagado
    Else
        LlevaContr = False
    End If
    
    
    impo = ImporteFormateado(Text2(0).Text)
    
        Sql = "cobros"
        Ampliacion = "fecultco"
        Numdocum = "impcobro"
        'El importe es el total. Lo que ya llevaba mas lo de ahora
        If Text1(5).Text <> "" Then impo = impo + ImporteFormateado(Text1(5).Text)
    
    
    '++monica
    Dim NumLin As Long
    
    
        Sql = "update ariconta" & vParam.Numconta & ".cobros set impcobro = coalesce(impcobro,0) + " & DBSet(Text2(0).Text, "N")
        Sql = Sql & ", fecultco = " & DBSet(Text3(0).Text, "F")
        Sql = Sql & " where numserie = " & DBSet(RecuperaValor(Vto, 1), "T") & " and numfactu = " & DBSet(RecuperaValor(Vto, 2), "N")
        Sql = Sql & " and fecfactu = " & DBSet(RecuperaValor(Vto, 3), "F") & " and numorden = " & DBSet(RecuperaValor(Vto, 4), "N")
    
        Conn.Execute Sql
        
        Sql = "select impvenci + coalesce(gastos,0) - coalesce(impcobro,0) from cobros where numserie = " & DBSet(RecuperaValor(Vto, 1), "T") & " and numfactu = " & DBSet(RecuperaValor(Vto, 2), "N")
        Sql = Sql & " and fecfactu = " & DBSet(RecuperaValor(Vto, 3), "F") & " and numorden = " & DBSet(RecuperaValor(Vto, 4), "N")
     
        'ahora es cuando ponemos la situacion
        Situacion = 0
        If DevuelveValor(Sql) = 0 Then
            Situacion = 1
        End If
    
        Sql = "update ariconta" & vParam.Numconta & ".cobros set "
        Sql = Sql & " situacion = " & DBSet(Situacion, "N")
        Sql = Sql & " where numserie = " & DBSet(RecuperaValor(Vto, 1), "T") & " and numfactu = " & DBSet(RecuperaValor(Vto, 2), "N")
        Sql = Sql & " and fecfactu = " & DBSet(RecuperaValor(Vto, 3), "F") & " and numorden = " & DBSet(RecuperaValor(Vto, 4), "N")
    
        Conn.Execute Sql
    
    
    Contabilizar = True

    Set Mc = Nothing
    Set FP = Nothing

    Exit Function
ECon:
    MuestraError Err.Number, "Contabilizar anticipo"
    Set Mc = Nothing
    Set FP = Nothing
End Function
    
Private Sub txtCta_GotFocus(Index As Integer)
    PonFoco txtCta(1)
End Sub

Private Sub txtCta_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCta_LostFocus(Index As Integer)

        txtCta(Index).Text = Trim(txtCta(Index).Text)
        Cad = txtCta(Index).Text
        impo = 0
        If Cad <> "" Then
            If CuentaCorrectaUltimoNivel(Cad, CadenaDesdeOtroForm) Then
                Cad = DevuelveDesdeBD("codmacta", "ariconta" & vParam.Numconta & ".bancos", "codmacta", Cad, "T")
                If Cad = "" Then
                    CadenaDesdeOtroForm = ""
                    MsgBox "La cuenta contable no esta asociada a ninguna cuenta bancaria", vbExclamation
                End If
            Else
                MsgBox CadenaDesdeOtroForm, vbExclamation
                Cad = ""
                CadenaDesdeOtroForm = ""
            End If
            impo = 1
        Else
            CadenaDesdeOtroForm = ""
        End If
        
        
        txtCta(Index).Text = Cad
        txtDescCta(Index).Text = CadenaDesdeOtroForm
        If Cad = "" And impo <> 0 Then
            PonFoco txtCta(Index)
        End If
        CadenaDesdeOtroForm = ""
End Sub


Private Sub CargaCombo()
    Combo1.Clear
    'Conceptos
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "Select * from ariconta" & vParam.Numconta & ".tipofpago order by tipoformapago", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not miRsAux.EOF
        Combo1.AddItem miRsAux!descformapago
        Combo1.ItemData(Combo1.NewIndex) = miRsAux!tipoformapago
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
End Sub


Private Function InsertarEnCaja() As Boolean
Dim TipoRegi As Integer
    On Error GoTo eInsertarEnCaja
    InsertarEnCaja = False
    
    If FechaCorrecta2(CDate(Text3(0).Text)) > 1 Then Err.Raise 513, , "No existe la serie factura: " & Msg
    
    Msg = RecuperaValor(Vto, 1)
    Cad = DevuelveDesdeBD("tiporegi", "contadores", "serfactur", Msg, "T")
    If Cad = "" Then Err.Raise 513, , "No existe la serie factura: " & Msg
    TipoRegi = Cad
    
    Cad = DevuelveDesdeBD("max(feccaja)", "caja_param", "usuario", vUsu.Login, "T")
    If Cad = "" Then Cad = "01/01/1900 12:00:00"
   
    If CDate(Text3(0).Text & " " & Format(Now, "hh:nn:ss")) < CDate(Cad) Then Err.Raise 513, , "Fecha anterior a cierre de caja: " & Cad
    
    Cad = "numserie = " & DBSet(RecuperaValor(Vto, 1), "T") & " AND numfactu =" 'numserie
    Cad = Cad & DBSet(RecuperaValor(Vto, 2), "N") & " AND fecfactu =" 'numfactu
    Cad = Cad & DBSet(RecuperaValor(Vto, 3), "F") & " AND numorden = " 'fecfactu
    Cad = Cad & DBSet(RecuperaValor(Vto, 4), "N")
    Msg = DevuelveDesdeBD("impvenci", "ariconta" & vParam.Numconta & ".cobros", Cad & " AND 1", "1")
    If Msg = "" Then Err.Raise 513, , "No se encuentra el vencimiento en contabilidad"
    
        
        
   
    Msg = "INSERT INTO caja(usuario,feccaja,tipomovi,importe,ampliacion,tiporegi,numserie,numdocum,anoexped) VALUES (" & DBSet(vUsu.Login, "T")
    Msg = Msg & "," & DBSet(Text3(0).Text & " " & Format(Now, "hh:nn:ss"), "FH") & ",0,"
    
    Msg = Msg & DBSet(Text2(0).Text, "N") & "," & DBSet(RecuperaValor(Cta, 2), "T")
    Msg = Msg & "," & TipoRegi & "," & DBSet(RecuperaValor(Vto, 1), "T")
    Msg = Msg & "," & DBSet(RecuperaValor(Vto, 2), "N") & "," & Year(RecuperaValor(Vto, 3)) & ")"
    Conn.Execute Msg
    
    'En el cobro de tesoreria
    Msg = "UPDATE  ariconta" & vParam.Numconta & ".cobros set impcobro =  coalesce(impcobro,0) + " & DBSet(ImporteFormateado((Text2(0).Text)), "N")
    Msg = Msg & " ,fecultco =" & DBSet(Text3(0).Text, "F") & " WHERE " & Cad
    Conn.Execute Msg
   
   

    InsertarEnCaja = True
eInsertarEnCaja:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
End Function



