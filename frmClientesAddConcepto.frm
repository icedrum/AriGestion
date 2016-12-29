VERSION 5.00
Begin VB.Form frmClientesAddConcepto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   8835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
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
      Left            =   7680
      TabIndex        =   6
      Top             =   1680
      Visible         =   0   'False
      Width           =   1035
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
      Left            =   6360
      TabIndex        =   5
      Top             =   1680
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
      Left            =   7680
      TabIndex        =   4
      Top             =   1680
      Width           =   1035
   End
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
      Height          =   360
      Index           =   2
      Left            =   7440
      MaxLength       =   10
      TabIndex        =   3
      Tag             =   "NIF|T|N|||clientes|nifclien|||"
      Text            =   "Text1"
      Top             =   960
      Width           =   1275
   End
   Begin VB.TextBox Text1 
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
      Left            =   1440
      MaxLength       =   100
      TabIndex        =   2
      Tag             =   "NIF|T|N|||clientes|nifclien|||"
      Text            =   "Text1"
      Top             =   960
      Width           =   5835
   End
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
      Height          =   360
      Index           =   0
      Left            =   240
      MaxLength       =   10
      TabIndex        =   1
      Tag             =   "ID concepto|N|N||||codconce|||"
      Text            =   "Text1"
      Top             =   960
      Width           =   1035
   End
   Begin VB.Image imgCC 
      Height          =   480
      Left            =   600
      Top             =   720
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   8475
   End
   Begin VB.Label Label1 
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
      Height          =   240
      Index           =   2
      Left            =   7440
      TabIndex        =   8
      Top             =   720
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Descripción"
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
      Left            =   1440
      TabIndex        =   7
      Top             =   720
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   330
   End
End
Attribute VB_Name = "frmClientesAddConcepto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public IdLinea As Integer '-1 para nueva
Public Nombre As String
Public IdCliente As Long

Private WithEvents frmCon As frmConceptos
Attribute frmCon.VB_VarHelpID = -1
Private Cad As String
Public Tipo As Byte   '0 Cuota      1 Laboral     2Fiscal


Private Sub cmdAceptar_Click()
    If Text1(0).Text = "" Or Text1(1).Text = "" Or Text1(2).Text = "" Then Exit Sub
    
    
    'Veremos si el concepto es valido
    Cad = DevuelveDesdeBD("tipoconcepto", "conceptos", "codconce", Text1(0).Text)
    If Tipo = 0 Then
        If Cad = "1" Or Cad = "2" Then Cad = ""
    ElseIf Tipo = 1 Then
        If Cad = 3 Then Cad = ""
    Else
        If Cad = 4 Then Cad = ""
    End If
    
    If Cad <> "" Then
        MsgBox "El tipo de concepto no es valido para el tipo de cuota", vbExclamation
        Exit Sub
    End If
    
    'clientes_cuotas clientes_fiscal clientes_laboral
    Msg = IIf(Tipo = 0, "clientes_cuotas", IIf(Tipo = 1, "clientes_laboral", "clientes_fiscal"))
    
    If IdLinea < 0 Then
        Cad = DevuelveDesdeBD("numlinea", Msg, "codclien", CStr(IdCliente))
        Cad = Val(Cad) + 1
        
        Cad = Msg & "( codclien , numlinea, codconce, Importe) VALUES (" & IdCliente & "," & Cad
        Cad = "INSERT INTO " & Cad & "," & Text1(0).Text & "," & DBSet(Text1(2).Text, "N") & ")"
    Else
        Cad = "UPDATE " & Msg & " SET codconce=" & Text1(0).Text
        Cad = Cad & ", importe =" & DBSet(Text1(2).Text, "N")
        Cad = Cad & "  WHERE codclien=" & IdCliente & " AND numlinea=" & IdLinea
    End If
    If Ejecuta(Cad) Then
        CadenaDesdeOtroForm = "Ok"
        Unload Me
    End If
End Sub

Private Sub cmdCancelar_Click()
    CadenaDesdeOtroForm = ""
    Unload Me
End Sub

Private Sub Form_Activate()
    If Me.Tag = 1 Then
        Me.Tag = 2
        Limpiar Me
        
        If IdLinea = -1 Then

            imgCC_Click
            Me.Tag = 0
        Else
            
            Msg = IIf(Tipo = 0, "clientes_cuotas", IIf(Tipo = 1, "clientes_laboral", "clientes_fiscal"))
            Cad = "tabla.codconce = conceptos.codconce AND codclien= " & IdCliente & " AND numlinea"
            Cad = DevuelveDesdeBD("concat(nomconce,'|',conceptos.codconce,'|',importe,'|')", Msg & " tabla,conceptos ", Cad, CStr(IdLinea), "N")
            If Cad = "" Then
                MsgBox "Error leyendo BD. Valor no encontrado.  " & IdLinea, vbExclamation
                Me.cmdAceptar.Enabled = False
            Else
                Text1(1).Text = RecuperaValor(Cad, 1)
                Text1(0).Text = RecuperaValor(Cad, 2)
                Cad = Replace(RecuperaValor(Cad, 3), ".", ",")
                
                Text1(2).Text = Format(Cad, FormatoImporte)
                PonFoco Text1(2)
                Me.Tag = 0
            End If
        End If
        CadenaDesdeOtroForm = ""
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Me.Icon = frmppal.Icon
    Me.Label1(3).Caption = Nombre
    Me.Tag = 1
    Text1(2).Alignment = 1
    Select Case Tipo
    Case 0
        Caption = "Cuota"
    Case 1
        Caption = "Laboral"
    Case Else
        Caption = "Fiscal"
    End Select
    Me.imgCC.Picture = frmppal.imgIcoForms.ListImages(1).Picture
End Sub

Private Sub frmCon_DatoSeleccionado(CadenaSeleccion As String)
    Msg = CadenaSeleccion
End Sub

Private Sub imgCC_Click()
    
    Set frmCon = New frmConceptos
    frmCon.DatosADevolverBusqueda = "0"
    Msg = ""
    frmCon.Show vbModal
    Set frmCon = Nothing
    If Msg <> "" Then
    
        Text1(0).Text = RecuperaValor(Msg, 1)
        Text1(1).Text = RecuperaValor(Msg, 2)
        Msg = RecuperaValor(Msg, 3)
        Text1(2).Text = Format(Msg, FormatoImporte)
        
    End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), 3
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
    
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim Procesar As Boolean
    'Esta cargando datos iniciales
    If Me.Tag = 2 Then Exit Sub

    If Index = 0 Then
        Msg = ""
        
        
        Procesar = True
        If IdLinea > 0 Then If CStr(IdLinea) = Text1(0).Text Then Procesar = False
        
        If Procesar Then
            If Text1(0).Text <> "" Then
                If PonerFormatoEntero(Text1(0)) Then
                    Cad = "preciocon"
                    Msg = DevuelveDesdeBD("nomconce", "conceptos", "codconce", Text1(0).Text, "N", Cad)
                    If Msg <> "" Then
                        Text1(1).Text = Msg
                        Text1(2).Text = Cad
                    End If
                End If
            End If
                If Msg = "" Then
                    If Text1(0).Text <> "" Then
                        Text1(0).Text = ""
                        Text1(1).Text = ""
                       ' PonerFoco  Text1(0)
                    End If
                End If
                
             Msg = ""
        End If
    Else
        If Index = 2 Then
            If Not PonerFormatoDecimal(Text1(2), 1) Then Text1(2).Text = ""
        End If
    End If
End Sub
