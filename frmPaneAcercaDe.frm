VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#17.2#0"; "Codejock.Controls.v17.2.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#17.2#0"; "Codejock.ShortcutBar.v17.2.0.ocx"
Begin VB.Form frmPaneAcercaDe 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.TreeView tree 
      Height          =   4215
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   4095
      _Version        =   1114114
      _ExtentX        =   7223
      _ExtentY        =   7435
      _StockProps     =   77
      ForeColor       =   -2147483640
      Appearance      =   6
      IconSize        =   16
      ShowBorder      =   0   'False
   End
   Begin XtremeShortcutBar.ShortcutCaption MainCaption 
      Height          =   360
      Left            =   4080
      TabIndex        =   2
      Top             =   0
      Width           =   360
      _Version        =   1114114
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
      Expandable      =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption ItemCaption 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   4695
      _Version        =   1114114
      _ExtentX        =   8281
      _ExtentY        =   503
      _StockProps     =   14
      Caption         =   "Acerca de"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      VisualTheme     =   3
      Expandable      =   -1  'True
   End
End
Attribute VB_Name = "frmPaneAcercaDe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Set tree.Icons = frmShortBar.wndShortcutBar.Icons
    tree.IconSize = 24
     tree.Font.SIZE = 10
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "select codigo,descripcion,imagen from menus where aplicacion='introcon' and padre=2", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    While Not miRsAux.EOF
        NumRegElim = NumRegElim + 1
        i = CInt(RecuperaValor("12|11|2|", CInt(NumRegElim)))
        tree.Nodes.Add , , "C" & miRsAux!Codigo, miRsAux!Descripcion, i
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    UpdateLayout
End Sub

Public Sub SetFlatStyle(FlatStyle As Boolean)
      
    Me.BackColor = frmShortBar.wndShortcutBar.PaintManager.PaneBackgroundColor
    tree.BackColor = Me.BackColor
    tree.ForeColor = frmShortBar.wndShortcutBar.PaintManager.PaneTextColor
    
    MainCaption.GradientColorDark = frmShortBar.wndShortcutBar.PaintManager.PaneBackgroundColor
    MainCaption.GradientColorLight = frmShortBar.wndShortcutBar.PaintManager.PaneBackgroundColor

End Sub

Private Sub ItemCaption_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    ItemCaption.Expanded = Not ItemCaption.Expanded
    UpdateLayout
End Sub

Private Sub Form_Resize()
    ItemCaption.Width = Me.ScaleWidth
    MainCaption.Left = Me.ScaleWidth - MainCaption.Width
End Sub


Sub UpdateLayout()

    Dim top As Long
    
    top = ItemCaption.top + ItemCaption.Height
    If ItemCaption.Expanded Then
        tree.Visible = True
        tree.top = 80 + top
        top = 80 + top + tree.Height
    Else
        tree.Visible = False
    End If

End Sub

Private Sub MainCaption_ExpandButtonClicked()
    Call frmPpalNuevooo.ExpandButtonClicked
End Sub


Private Sub tree_NodeClick(ByVal Node As XtremeSuiteControls.TreeViewNode)
'    AbrirFormulariosAyuda CLng(Mid(Node.Key, 2))
End Sub




Private Sub AbrirFormulariosAyuda(Accion As Long)

    Select Case Accion
        Case 4
            'Zona ARIADNA
            LanzaVisorMimeDocumento Me.hwnd, vParam.WebSoporte '"http://www.ariadnasw.com/"
        
        Case 7
            'licencia de usuario final
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & "/Licenciadeuso.html" ' "http://www.ariadnasw.com/clientes/"
            
            
        Case Else
            Dim Lanza As String
            Dim Aux As String


            Lanza = vParam.MailSoporte & "||"

            'Aqui pondremos lo del texto del BODY
            Lanza = Lanza & "|"
            'Envio o mostrar
            Lanza = Lanza & "0"   '0. Display   1.  send

            'Campos reservados para el futuro
            Lanza = Lanza & "||||"

            'El/los adjuntos
            Lanza = Lanza & "|"

            Aux = App.Path & "\ARIMAILGES.EXE" & " " & Lanza  '& vParamAplic.ExeEnvioMail & " " & Lanza
            Shell Aux, vbNormalFocus

        End Select
End Sub
