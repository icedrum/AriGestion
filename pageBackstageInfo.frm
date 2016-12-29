VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#17.2#0"; "Codejock.CommandBars.v17.2.0.ocx"
Begin VB.Form pageBackstageInfo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12540
   LinkTopic       =   "Form1"
   ScaleHeight     =   454
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   836
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeCommandBars.BackstageButton btnManageVersions 
      Height          =   1230
      Left            =   360
      TabIndex        =   10
      Top             =   3120
      Width           =   1290
      _Version        =   1114114
      _ExtentX        =   2275
      _ExtentY        =   2170
      _StockProps     =   79
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EnableMarkup    =   -1  'True
      TextImageRelation=   1
      ShowShadow      =   -1  'True
      IconWidth       =   32
      Icon            =   "pageBackstageInfo.frx":0000
   End
   Begin XtremeCommandBars.BackstageButton btnCheckForIssues 
      Height          =   1230
      Left            =   360
      TabIndex        =   9
      Top             =   5160
      Visible         =   0   'False
      Width           =   1290
      _Version        =   1114114
      _ExtentX        =   2275
      _ExtentY        =   2170
      _StockProps     =   79
      Caption         =   "Check for  Issues"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EnableMarkup    =   -1  'True
      TextImageRelation=   1
      ShowShadow      =   -1  'True
      IconWidth       =   32
      Icon            =   "pageBackstageInfo.frx":106A
   End
   Begin XtremeCommandBars.BackstageButton btnProtectDocument 
      Height          =   1230
      Left            =   360
      TabIndex        =   8
      Top             =   960
      Width           =   1290
      _Version        =   1114114
      _ExtentX        =   2275
      _ExtentY        =   2170
      _StockProps     =   79
      Caption         =   $"pageBackstageInfo.frx":20D4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EnableMarkup    =   -1  'True
      TextImageRelation=   1
      PushButtonStyle =   4
      ShowShadow      =   -1  'True
      IconWidth       =   32
      Icon            =   "pageBackstageInfo.frx":21A8
   End
   Begin XtremeCommandBars.BackstageSeparator lblBackstageSeparator4 
      Height          =   6615
      Left            =   7200
      TabIndex        =   14
      Top             =   120
      Width           =   135
      _Version        =   1114114
      _ExtentX        =   238
      _ExtentY        =   11668
      _StockProps     =   2
      Vertical        =   -1  'True
      MarkupText      =   ""
   End
   Begin XtremeCommandBars.BackstageSeparator lblBackstageSeparator3 
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   4560
      Width           =   6615
      _Version        =   1114114
      _ExtentX        =   11668
      _ExtentY        =   450
      _StockProps     =   2
      MarkupText      =   ""
   End
   Begin XtremeCommandBars.BackstageSeparator lblBackstageSeparator2 
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2520
      Width           =   6615
      _Version        =   1114114
      _ExtentX        =   11668
      _ExtentY        =   450
      _StockProps     =   2
      MarkupText      =   ""
   End
   Begin XtremeCommandBars.BackstageSeparator lblBackstageSeparator1 
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   600
      Width           =   6615
      _Version        =   1114114
      _ExtentX        =   11668
      _ExtentY        =   450
      _StockProps     =   2
      MarkupText      =   ""
   End
   Begin VB.Image Image1 
      Height          =   1350
      Left            =   7560
      Picture         =   "pageBackstageInfo.frx":3212
      Top             =   120
      Width           =   3840
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tel: +34 963 805 579"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   7800
      TabIndex        =   17
      Top             =   3600
      Width           =   4215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "46007 Valencia"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   7800
      TabIndex        =   16
      Top             =   3120
      Width           =   4215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pasaje Ventura Feliu 13, Entresuelo 2 Izquierda"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   7800
      TabIndex        =   15
      Top             =   2640
      Width           =   4215
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   10200
      Top             =   960
      _Version        =   1114114
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      ScaleMode       =   1
      VisualTheme     =   6
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Version"
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
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "There are no previous versions of this file"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   1920
      TabIndex        =   6
      Top             =   3480
      Width           =   4215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Document properties and author's name"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   6000
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Prepare for Sharing"
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
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   5280
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Before sharing this file, be aware that it contains:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   5640
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      TabIndex        =   2
      Top             =   1440
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Licencia usuario final"
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
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label lblCaption 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Acerca de Ariconta 6"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003B3B3B&
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "pageBackstageInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Type RECT
        Left As Long
        top As Long
        Right As Long
        Bottom As Long
End Type

Private Function AddButton(Controls As CommandBarControls, ControlType As XTPControlType, Id As Long, Caption As String, Optional BeginGroup As Boolean = False, Optional DescriptionText As String = "", Optional ButtonStyle As XTPButtonStyle = xtpButtonAutomatic, Optional Category As String = "Controls") As CommandBarControl
    Dim Control As CommandBarControl
    Set Control = Controls.Add(ControlType, Id, Caption)
    
    Control.BeginGroup = BeginGroup
    Control.DescriptionText = DescriptionText
    Control.Style = ButtonStyle
    Control.Category = Category
    
    Set AddButton = Control
    
End Function

Private Sub btnManageVersions_Click()
    frmppal.OpcionesMenuInformacion ID_Ver_Version_operativa_web
End Sub


Private Sub btnProtectDocument_DropDown()
        Dim Popup As CommandBar
        Set Popup = CommandBars.Add("Popup", xtpBarPopup)
             
   
        AddButton Popup.Controls, xtpControlButton, ID_Licencia_Usuario_Final_txt, "Ver", False, "Ver licencia en formato texto."
        AddButton Popup.Controls, xtpControlButton, ID_Licencia_Usuario_Final_web, "Abrir licencia en navegador", False, "Ver licencia en www.ariadnasw.com"
        'AddButton Popup.Controls, xtpControlButton, 0, "Restrict Editing", False, "Control what types of changes people can make to this document."
        'AddButton Popup.Controls, xtpControlButton, 0, "Restrict Permission by People", False, "Grant peole access while removing their ability to edit, copy, or print."
        'AddButton Popup.Controls, xtpControlButton, 0, "Add a Digital Signature", False, "Ensure the integrity of the document by adding an invisible digital signature."
        
        Popup.ShowGripper = False
        Popup.SetIconSize 32, 32
        Popup.DefaultButtonStyle = xtpButtonCaptionAndDescription
        
        CommandBars.Icons.LoadBitmap App.Path + "\res\ProtectDocument.png", Array(12400), xtpImageNormal
    
    
        Dim RC As RECT
        GetWindowRect btnProtectDocument.hwnd, RC

        Popup.ShowPopup 0, RC.Left, RC.Bottom

End Sub

Private Sub LanzaAccion(Id As Long)
    frmppal.OpcionesMenuInformacion Id
    
End Sub


Private Sub CommandBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
Dim L As Long
    L = 0
    Select Case Control.Id
    Case ID_Licencia_Usuario_Final_txt
        L = Control.Id
    Case ID_Licencia_Usuario_Final_web
        L = Control.Id
    
    End Select
    If L > 0 Then LanzaAccion L
End Sub

Private Sub Form_Load()
    CommandBars.ActiveMenuBar.Delete
    Label6(0).Caption = App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_Resize()
On Error Resume Next
    lblBackstageSeparator4.Height = Me.ScaleHeight
End Sub

Private Sub ImagePreview_Click()
    Dim BackstageView As RibbonBackstageView
    Set BackstageView = frmppal.RibbonBar.SystemButton.CommandBar
    BackstageView.Close
End Sub

