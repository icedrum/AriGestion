VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#17.2#0"; "Codejock.Controls.v17.2.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#17.2#0"; "Codejock.ShortcutBar.v17.2.0.ocx"
Begin VB.Form frmPaneInformacion 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
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
      Caption         =   "Informacion"
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
Attribute VB_Name = "frmPaneInformacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private frmMens As frmMensajes

Private Sub Form_Load()
Dim i As Integer
    Set tree.Icons = SuiteControlsGlobalSettings.Icons
    tree.IconSize = 16
     tree.Font.SIZE = 10
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "select codigo,descripcion,imagen from menus where aplicacion='introcon' and padre=3", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    While Not miRsAux.EOF
        NumRegElim = NumRegElim + 1
        i = CInt(RecuperaValor("2|16|10|9|7|2|", CInt(NumRegElim)))
        tree.Nodes.Add , , "C" & miRsAux!Codigo, miRsAux!Descripcion, i
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing

    Set tree.SelectedItem = Nothing
    
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
        Case 6
            'CAlendario del contribuyente
            LanzaVisorMimeDocumento Me.hwnd, "http://www.agenciatributaria.es/AEAT.internet/Bibl_virtual/folletos/calendario_contribuyente.shtml"
 
    
        Case 8 ' documentos
            frmVarios.Opcion = Accion - 2
            frmVarios.Show vbModal
        Case 9 ' ayuda
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & "/Ariconta-6.html"  ' "http://www.ariadnasw.com/clientes/"


        Case 12 ' Informacion de la base de datos
            If CargarInformacionBBDD Then
                Set frmMens = New frmMensajes

                frmMens.Opcion = 25
                frmMens.Show vbModal

                Set frmMens = Nothing
            End If

        Case 13
            ' Panel de control donde seleccionamos los iconos que vamos a mostrar
            frmMensajes.Opcion = 24
            frmMensajes.Show vbModal



        Case 14 'Usuarios activos
            'mnUsuariosActivos_Click
            Set frmMens = New frmMensajes

            frmMens.Opcion = 26
            frmMens.Show vbModal

            Set frmMens = Nothing


    End Select
    
End Sub



Private Function CargarInformacionBBDD() As String
Dim Sql As String
Dim Sql2 As String
Dim CadValues As String
Dim NroRegistros As Long
Dim NroRegistrosSig As Long
Dim NroRegistrosTot As Long
Dim NroRegistrosTotSig As Long
Dim FecIniSig As Date
Dim FecFinSig As Date
Dim Porcen1 As Currency
Dim Porcen2 As Currency
Dim Rs As ADODB.Recordset

    On Error GoTo eCargarInformacionBBDD
    
    CargarInformacionBBDD = False
    
    Sql = "delete from tmpinfbbdd where codusu = " & vUsu.Codigo
    Conn.Execute Sql
    
    FecIniSig = DateAdd("yyyy", 1, vParam.fechaini)
    FecFinSig = DateAdd("yyyy", 1, vParam.fechafin)
    
    Sql2 = "insert into tmpinfbbdd (codusu,posicion,concepto,nactual,poractual,nsiguiente,porsiguiente) values "
    
    'asientos
    Sql = "select count(*) from hcabapu where fechaent between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
    NroRegistros = DevuelveValor(Sql)
    Sql = "select count(*) from hcabapu where fechaent between " & DBSet(FecIniSig, "F") & " and " & DBSet(FecFinSig, "F")
    NroRegistrosSig = DevuelveValor(Sql)
    
    CadValues = "(" & vUsu.Codigo & ",1,'Asientos'," & DBSet(NroRegistros, "N") & ",0," & DBSet(NroRegistrosSig, "N") & ",0)"
    Conn.Execute Sql2 & CadValues
    
    'apuntes
    Sql = "select count(*) from hlinapu where fechaent between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
    NroRegistros = DevuelveValor(Sql)
    Sql = "select count(*) from hlinapu where fechaent between " & DBSet(FecIniSig, "F") & " and " & DBSet(FecFinSig, "F")
    NroRegistrosSig = DevuelveValor(Sql)
    
    CadValues = "(" & vUsu.Codigo & ",2,'Apuntes'," & DBSet(NroRegistros, "N") & ",0," & DBSet(NroRegistrosSig, "N") & ",0)"
    Conn.Execute Sql2 & CadValues
    
    'facturas de venta
    Sql = "select count(*) from factcli where "
    Sql = Sql & " fecfactu between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
    
    NroRegistrosTot = DevuelveValor(Sql)
    
    
    Sql = "select count(*) from factcli where "
    Sql = Sql & " fecfactu between " & DBSet(FecIniSig, "F") & " and " & DBSet(FecFinSig, "F")
    
    NroRegistrosTotSig = DevuelveValor(Sql)
    
    i = 3
    
    Sql = "select * from contadores where not tiporegi in ('0','1')"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
    
        Sql = "select count(*) from factcli where numserie = " & DBSet(Rs!tiporegi, "T")
        Sql = Sql & " and fecfactu between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
    
        NroRegistros = DevuelveValor(Sql)
        Porcen1 = 0
        If NroRegistrosTot <> 0 Then
            Porcen1 = Round(NroRegistros * 100 / NroRegistrosTot, 2)
        End If
        
        Sql = "select count(*) from factcli where numserie = " & DBSet(Rs!tiporegi, "T")
        Sql = Sql & " and fecfactu between " & DBSet(FecIniSig, "F") & " and " & DBSet(FecFinSig, "F")
        
        NroRegistrosSig = DevuelveValor(Sql)
        Porcen2 = 0
        If NroRegistrosTotSig <> 0 Then
            Porcen2 = Round(NroRegistrosSig * 100 / NroRegistrosTotSig, 2)
        End If
    
        CadValues = "(" & vUsu.Codigo & "," & DBSet(i, "N") & "," & DBSet(Rs!nomregis, "T") & "," & DBSet(NroRegistros, "N") & "," & DBSet(Porcen1, "N") & ","
        CadValues = CadValues & DBSet(NroRegistrosSig, "N") & "," & DBSet(Porcen2, "N") & ")"
        Conn.Execute Sql2 & CadValues
        
        i = i + 1
    
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    'facturas de proveedor
    i = i + 1
    
    Sql = "select count(*) from factpro where fecharec between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
    NroRegistros = DevuelveValor(Sql)
    Sql = "select count(*) from factpro where fecharec between " & DBSet(FecIniSig, "F") & " and " & DBSet(FecFinSig, "F")
    NroRegistrosSig = DevuelveValor(Sql)
    
    CadValues = "(" & vUsu.Codigo & "," & DBSet(i, "N") & ",'Facturas Proveedores'," & DBSet(NroRegistros, "N") & ",0,"
    CadValues = CadValues & DBSet(NroRegistrosSig, "N") & ",0)"
    
    Conn.Execute Sql2 & CadValues
    CargarInformacionBBDD = True
    Exit Function


eCargarInformacionBBDD:
    MuestraError Err.Number, "Cargar Temporal de BBDD", Err.Description
End Function


