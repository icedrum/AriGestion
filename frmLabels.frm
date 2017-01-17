VERSION 5.00
Begin VB.Form frmLabels 
   BorderStyle     =   0  'None
   Caption         =   "Identificacion"
   ClientHeight    =   5640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9705
   ForeColor       =   &H8000000B&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   9705
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Versión 6.0.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   4350
      TabIndex        =   1
      Top             =   2340
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cargando ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00765341&
      Height          =   285
      Index           =   2
      Left            =   360
      TabIndex        =   0
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   5790
      Left            =   0
      Top             =   0
      Width           =   9750
   End
End
Attribute VB_Name = "frmLabels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PrimeraVez As Boolean
Dim T1 As Single
Dim CodPC As Long
Dim UsuarioOK As String
Private Sub Form_Activate()

    Screen.MousePointer = vbHourglass
    If PrimeraVez Then
        PrimeraVez = False
        'La madre de todas las batallas
        pLabel "Cargando principal"
    
        
    End If
    
End Sub



Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    
    Label1(3).Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
    Label1(2).Caption = "Cargando datos de usuarios"
    
    PrimeraVez = True
    CargaImagen
    Me.Height = 5625 '5535
    Me.Width = 9705 ' 7935
    
'    '?????????????? QUITAR ESTO
'    If Combo1.Text = "root" Then Text1(1).Text = "aritel"
    
End Sub


Private Sub CargaImagen()
    On Error Resume Next
    Me.Image1 = LoadPicture(App.Path & "\arifon6.dat")
    If Err.Number <> 0 Then
        MsgBox Err.Description & vbCrLf & vbCrLf & "Error cargando", vbCritical

        Set Conn = Nothing
        End
    End If
End Sub



Public Sub pLabel(TEXTO As String)

    Me.Label1(2).Caption = TEXTO
    Label1(2).Refresh
    espera 0.3
End Sub

