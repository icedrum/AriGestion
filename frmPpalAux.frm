VERSION 5.00
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#17.2#0"; "Codejock.SkinFramework.v17.2.0.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#17.2#0"; "Codejock.CommandBars.v17.2.0.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#17.2#0"; "Codejock.DockingPane.v17.2.0.ocx"
Begin VB.Form frmPpalAux 
   Caption         =   "Form1"
   ClientHeight    =   3555
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   ScaleHeight     =   3555
   ScaleWidth      =   8715
   Begin VB.Label Label1 
      Caption         =   "AUX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      TabIndex        =   0
      Top             =   2400
      Width           =   4095
   End
   Begin XtremeDockingPane.DockingPane DockingPaneManager 
      Left            =   0
      Top             =   0
      _Version        =   1114114
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeCommandBars.ImageManager ImageManager 
      Left            =   3480
      Top             =   240
      _Version        =   1114114
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmPpalAux.frx":0000
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   3120
      Top             =   840
      _Version        =   1114114
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeCommandBars.ImageManager ImageManagerGalleryStyles 
      Left            =   840
      Top             =   1080
      _Version        =   1114114
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmPpalAux.frx":001A
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   1440
      Top             =   240
      _Version        =   1114114
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPpalAux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FontSizes(4) As Integer
Dim RibbonSeHaCreado As Boolean
Dim RN2 As ADODB.Recordset

Private Sub Form_Load()
    frmLabels.pLabel "Carga DLL"
    CargaIconosDlls
   
    CommandBarsGlobalSettings.App = App
            
    frmLabels.pLabel "Leyendo menus usuario"
    CargaDatosMenusDemas
    
    ShowEventInPane = False
       
    FontSizes(0) = 0
    FontSizes(1) = 11
    FontSizes(2) = 13
    FontSizes(3) = 16
               
    DockingPaneManager.SetCommandBars Me.CommandBars
              
    Set frmShortBar = New frmShortcutBar2
    Set frmInbox = New frmInbox
        
    Dim A As Pane, B As Pane, C As Pane, d As Pane
    
    frmLabels.pLabel "Creando paneles"
    Set A = DockingPaneManager.CreatePane(PANE_SHORTCUTBAR, 170, 120, DockLeftOf, Nothing)
    A.Tag = PANE_SHORTCUTBAR
  '  A.MinTrackSize.Width = MinimizedShortcutBarWidth
    
    Set B = DockingPaneManager.CreatePane(PANE_REPORT_CONTROL, 700, 400, DockRightOf, A)
    B.Tag = PANE_REPORT_CONTROL
   
   

End Sub

Private Sub CargaIconosDlls()

    'ImageList1 .ImageHeight = 48
    'ImageList1 .ImageWidth = 48
    'GetIconsFromLibrary App.Path & "\styles\icoconppal.dll", 1, 48
    With formIcon

    .ImageListPPal48.ImageHeight = 48
    .ImageListPPal48.ImageWidth = 48
    GetIconsFromLibrary App.Path & "\styles\icoconppal2.dll", 8, 48


    .ImageListPpal16.ImageHeight = 16
    .ImageListPpal16.ImageWidth = 16
    GetIconsFromLibrary App.Path & "\styles\icoconppal2.dll", 9, 16



    .ImgListComun.ImageHeight = 24
    .ImgListComun.ImageWidth = 24
    GetIconsFromLibrary App.Path & "\styles\iconosconta.dll", 2, 24 'antes icolistcon
    
    '++
    .imgListComun_BN.ImageHeight = 24
    .imgListComun_BN.ImageWidth = 24
    GetIconsFromLibrary App.Path & "\styles\iconosconta_BN.dll", 3, 24
    
    .imgListComun_OM.ImageHeight = 24
    .imgListComun_OM.ImageWidth = 24
    GetIconsFromLibrary App.Path & "\styles\iconosconta_OM.dll", 4, 24
    
    .imgListComun16.ImageHeight = 16
    .imgListComun16.ImageWidth = 16
    GetIconsFromLibrary App.Path & "\styles\iconosconta.dll", 5, 16
    
'    GetIconsFromLibrary App.Path & "\styles\iconosconta_BN.dll", 6, 16
'    GetIconsFromLibrary App.Path & "\styles\iconosconta_OM.dll", 7, 16

    End With
End Sub
Public Sub GetIconsFromLibrary(ByVal sLibraryFilePath As String, ByVal op As Integer, ByVal tam As Integer)
    Dim I As Integer
    Dim tRes As ResType, iCount As Integer
        
    opcio = op
    tamany = tam
    ghmodule = LoadLibraryEx(sLibraryFilePath, 0, DONT_RESOLVE_DLL_REFERENCES)

    If ghmodule = 0 Then
        MsgBox "Invalid library file.", vbCritical
        Exit Sub
    End If
        
    For tRes = RT_FIRST To RT_LAST
        DoEvents
        EnumResourceNames ghmodule, tRes, AddressOf EnumResNameProc, 0
    Next
    FreeLibrary ghmodule
             
End Sub


Private Sub CargaDatosMenusDemas()
    
    CreateRibbon
    CreateBackstage
    CreateRibbonOptions
    CargaMenu
'    CreateStatusBar
'    PonerCaption
'    CreateCalendarTabOriginal
    RibbonSeHaCreado = True
End Sub
Private Sub CreateRibbon()

    
    If RibbonSeHaCreado Then Exit Sub
    Dim RibbonBar As RibbonBar
    
    
    Set RibbonBar = CommandBars.AddRibbonBar("The Ribbon")
    RibbonBar.EnableDocking xtpFlagStretched
    
    RibbonBar.AllowQuickAccessCustomization = False
    RibbonBar.ShowQuickAccessBelowRibbon = False
    RibbonBar.ShowGripper = False
    
    RibbonBar.AllowMinimize = False
    RibbonBar.AddSystemButton
    
    RibbonBar.SystemButton.IconId = ID_SYSTEM_ICON
    RibbonBar.SystemButton.Caption = "&Menu"
    RibbonBar.SystemButton.Style = xtpButtonCaption
End Sub

Private Sub CreateBackstage()

    
    Dim RibbonBar As RibbonBar
    Set RibbonBar = CommandBars.ActiveMenuBar
    
    Dim BackstageView As RibbonBackstageView
    Set BackstageView = CommandBars.CreateCommandBar("CXTPRibbonBackstageView")
    
    BackstageView.SetTheme xtpThemeRibbon


    CommandBars.Icons.LoadBitmap App.Path & "\styles\BackstageIcons.png", _
    Array(1, 1, 1002, 1, 1, ID_APP_EXIT), xtpImageNormal

    Set RibbonBar.AddSystemButton.CommandBar = BackstageView
    
    'BackstageView.AddCommand ID_FILE_SAVE, "Cambiar empresa"
    'BackstageView.AddCommand ID_FILE_SAVE_AS, "Personalizar"
    'BackstageView.AddCommand ID_FILE_OPEN, "Open"
    'BackstageView.AddCommand ID_FILE_CLOSE, "Close"
    
    If (pageBackstageInfo Is Nothing) Then Set pageBackstageInfo = New pageBackstageInfo
    If (pageBackstageSend Is Nothing) Then Set pageBackstageSend = New pageBackstageSend
    If (pageBackstageHelp Is Nothing) Then Set pageBackstageHelp = New pageBackstageHelp
    
    Dim ControlInfo As RibbonBackstageTab
    Set ControlInfo = BackstageView.AddTab(1000, "Info", pageBackstageHelp.hwnd)
    
    BackstageView.AddTab 1002, "Empresas", pageBackstageSend.hwnd

    ' Los menus de informacion...
    BackstageView.AddTab 1001, "Acerca de", pageBackstageInfo.hwnd
    
    
    
    
    
    
    
    
    
    
    'BackstageView.AddCommand ID_FILE_OPTIONS, "Options"
    BackstageView.AddCommand ID_APP_EXIT, "Salir"
    
    ControlInfo.DefaultItem = True
    

End Sub

Private Sub CreateRibbonOptions()

    CommandBars.EnableActions
    If RibbonSeHaCreado Then Exit Sub
    
    CommandBars.Actions.Add ID_OPTIONS_STYLEBLUE2010, "Office 2010 Blue", "Office 2010 Blue", "Office 2010 Blue", "Themes"
    CommandBars.Actions.Add ID_OPTIONS_STYLESILVER2010, "Office 2010 Silver", "Office 2010 Silver", "Office 2010 Silver", "Themes"
    CommandBars.Actions.Add ID_OPTIONS_STYLEBLACK2010, "Office 2010 Black", "Office 2010 Black", "Office 2010 Black", "Themes"

    Dim Control As CommandBarControl, ControlAbout As CommandBarControl
    Dim ControlPopup As CommandBarPopup, ControlOptions As CommandBarPopup
         
    Set ControlOptions = RibbonBar.Controls.Add(xtpControlPopup, 0, "Opciones")
    ControlOptions.Flags = xtpFlagRightAlign
    
    Set Control = ControlOptions.CommandBar.Controls.Add(xtpControlPopup, 0, "Styles")
    Control.CommandBar.Controls.Add xtpControlButton, ID_OPTIONS_STYLEBLUE2010, "Office 2010 Blue"
    Control.CommandBar.Controls.Add xtpControlButton, ID_OPTIONS_STYLESILVER2010, "Office 2010 Silver"
    Control.CommandBar.Controls.Add xtpControlButton, ID_OPTIONS_STYLEBLACK2010, "Office 2010 Black"
    
    Set ControlPopup = ControlOptions.CommandBar.Controls.Add(xtpControlPopup, 0, "Tamaño fuente", -1, False)
    ControlPopup.CommandBar.Controls.Add xtpControlRadioButton, ID_OPTIONS_FONT_SYSTEM, "Sistema", -1, False
    Set Control = ControlPopup.CommandBar.Controls.Add(xtpControlRadioButton, ID_OPTIONS_FONT_NORMAL, "Normal", -1, False)
    Control.BeginGroup = True
    ControlPopup.CommandBar.Controls.Add xtpControlRadioButton, ID_OPTIONS_FONT_LARGE, "Grande", -1, False
    ControlPopup.CommandBar.Controls.Add xtpControlRadioButton, ID_OPTIONS_FONT_EXTRALARGE, "Extra grande", -1, False
    Set Control = ControlPopup.CommandBar.Controls.Add(xtpControlButton, ID_OPTIONS_FONT_AUTORESIZEICONS, "Ajustar Icons", -1, False)
    Control.BeginGroup = True
    
    'ControlOptions.CommandBar.Controls.Add xtpControlButton, ID_OPTIONS_RTL, "Right To Left"
    ControlOptions.CommandBar.Controls.Add xtpControlButton, ID_OPTIONS_ANIMATION, "Animation   "
    
    Set Control = AddButton(RibbonBar.Controls, xtpControlButton, ID_RIBBON_MINIMIZE, "Minimizar la barra", False, "Muestra solo los titulos del menu principal.")
    Control.Flags = xtpFlagRightAlign
    
    Set Control = AddButton(RibbonBar.Controls, xtpControlButton, ID_RIBBON_EXPAND, "Expandir la barra", False, "Muestra todos los elementos del menu.")
    Control.Flags = xtpFlagRightAlign
        
    Set ControlAbout = RibbonBar.Controls.Add(xtpControlButton, ID_APP_ABOUT, "&Acerca de")
    ControlAbout.Flags = xtpFlagRightAlign Or xtpFlagManualUpdate
    

        
End Sub


Public Function RibbonBar() As RibbonBar
    Set RibbonBar = CommandBars.ActiveMenuBar
    
End Function
Public Function AddButton(Controls As CommandBarControls, ControlType As XTPControlType, id As Long, Caption As String, Optional BeginGroup As Boolean = False, Optional DescriptionText As String = "", Optional ButtonStyle As XTPButtonStyle = xtpButtonAutomatic, Optional Category As String = "Controls") As CommandBarControl
    Dim Control As CommandBarControl
    Set Control = Controls.Add(ControlType, id, Caption)
    
    Control.BeginGroup = BeginGroup
    Control.DescriptionText = DescriptionText
    Control.Style = ButtonStyle
    Control.Category = Category
    
    Set AddButton = Control
    
End Function
'*************************************************************************
'*************************************************************************
'*************************************************************************
'
'       CARGA menus en Ribbon
'
'




Public Sub CargaMenu()
Dim RN As ADODB.Recordset




    Set RN = New ADODB.Recordset
    Set RN2 = New ADODB.Recordset
    On Error GoTo eCargaMenu
        
    idTabPpal = 0
    If RibbonSeHaCreado Then RibbonBar.RemoveAllTabs
    
    Cad = "Select * from menus where aplicacion = 'arigestion' and padre =0 ORDER BY padre,orden "
    RN.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
    
        
        If Not BloqueaPuntoMenu(RN!Codigo, "arigestion") Then
             Habilitado = True
             
             If Not MenuVisibleUsuario(DBLet(RN!Codigo), "arigestion") Then
                 Habilitado = False
             Else
         
                 If (MenuVisibleUsuario(DBLet(RN!Padre), "arigestion") And DBLet(RN!Padre) <> 0) Or DBLet(RN!Padre) = 0 Then
                     'OK todo habilitado
                 Else
                     Habilitado = False
                 End If
             End If
             
      
                
            If Habilitado Then
                
                Select Case RN!Codigo
                Case 1
                    '1   "CONFIGURACION"
                    CargaMenuConfiguracion RN!Codigo
                    
                    
                ' ****  Iran todos juntos en un tab
                Case 2
                    '2 Datos generales
                    CargaMenuDatosGenerales RN!Codigo
                Case 3
                    '3   "TRABAJO DIARIO"
                    CargaMenuTrabajoDiario RN!Codigo
                Case 4
                    '4   "FACTURACION"
                    CargaMenuFacturacion RN!Codigo
                Case 5
                    '5   "ESTADISTICAS"
                    CargaMenuEstadistica RN!Codigo


                Case Else
                    MsgBox "Menu no tratado"
                    End
                End Select
                
            End If
                                                 
        End If  'de habilitado el padre
    
        RN.MoveNext
    Wend
    RN.Close
                        
               
    
        RibbonBar.Tab(idTabPpal).Visible = True
        RibbonBar.Tab(idTabPpal).Selected = True
        Set RibbonBar.SelectedTab = RibbonBar.Tab(idTabPpal)
      
    
    
eCargaMenu:
    If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation
    Set TabNuevo = Nothing
    Set GroupNew = Nothing
    Set Control = Nothing
    Set RN = Nothing
    Set RN2 = Nothing
End Sub



Private Sub CargaMenuConfiguracion(IdMenu As Integer)

        'Creamos la TAB
        Set TabNuevo = RibbonBar.InsertTab(CLng(IdMenu), "Configuracion")
        TabNuevo.id = CLng(IdMenu)
        Set GroupNew = TabNuevo.Groups.AddGroup("", 1000000)
        
       
        
        'todos los hijos que cuelgan en la tab
        Cad = "Select * from menus where aplicacion = 'arigestion' and padre =" & IdMenu & " ORDER BY padre,orden"
        RN2.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        While Not RN2.EOF
         
           If Not BloqueaPuntoMenu(RN2!Codigo, "arigestion") Then
                Habilitado = True
    
                If Not MenuVisibleUsuario(DBLet(RN2!Codigo), "arigestion") Then
                    Habilitado = False
                Else
                    If Not MenuVisibleUsuario(DBLet(RN2!Padre), "arigestion") Then Habilitado = False
                End If
           
           
                    
                Set Control = GroupNew.Add(xtpControlButton, RN2!Codigo, RN2!Descripcion)
                Control.Enabled = Habilitado
             
            End If
            RN2.MoveNext
        Wend
        RN2.Close

         Set GroupNew = Nothing
End Sub



Private Sub CrearTabPPal()
    
    If idTabPpal = 0 Then
        Set TabNuevo = RibbonBar.InsertTab(9999, "Diario")
        idTabPpal = TabNuevo.Index
    Else
        Set TabNuevo = RibbonBar.Tab(idTabPpal)
    End If
End Sub


Private Sub CargaMenuDatosGenerales(IdMenu As Integer)

        'Creamos la TAB
        CrearTabPPal
        
        'En este llevaremos dos solapas, tesoreria y contabilidad (no le ponemos nombres)
        Cad = CStr(IdMenu * 100000)
        
        Set GroupNew = TabNuevo.Groups.AddGroup("General", Cad & "0")
        
        'todos los hijos que cuelgan en la tab
        Cad = "Select * from menus where aplicacion = 'arigestion' and padre =" & IdMenu & " ORDER BY padre,orden"
        RN2.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        While Not RN2.EOF
         
           If Not BloqueaPuntoMenu(RN2!Codigo, "arigestion") Then
                Habilitado = True
    
                If Not MenuVisibleUsuario(DBLet(RN2!Codigo), "arigestion") Then
                    Habilitado = False
                Else
                    If Not MenuVisibleUsuario(DBLet(RN2!Padre), "arigestion") Then Habilitado = False
                End If
           
           
                    
              
                Set Control = GroupNew.Add(xtpControlButton, RN2!Codigo, RN2!Descripcion)
                
                 
                Control.Enabled = Habilitado
             
            End If
            RN2.MoveNext
        Wend
        RN2.Close

         Set GroupNew = Nothing
End Sub


Private Sub CargaMenuFacturacion(IdMenu As Integer)


        'Creamos la TAB
        CrearTabPPal
        
        
        Cad = CStr(IdMenu * 100000)
        Set GroupNew = TabNuevo.Groups.AddGroup("Facturación", Cad & "0")
        
    
        
        'todos los hijos que cuelgan en la tab
        Cad = "Select * from menus where aplicacion = 'arigestion' and padre =" & IdMenu & " ORDER BY padre,orden"
        RN2.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        While Not RN2.EOF
        
           If Not BloqueaPuntoMenu(RN2!Codigo, "arigestion") Then
                Habilitado = True
    
                If Not MenuVisibleUsuario(DBLet(RN2!Codigo), "arigestion") Then
                    Habilitado = False
                Else
                    If Not MenuVisibleUsuario(DBLet(RN2!Padre), "arigestion") Then Habilitado = False
                End If
                

                
                
                Set Control = GroupNew.Add(xtpControlButton, RN2!Codigo, RN2!Descripcion)
                
                Control.Enabled = Habilitado
                
              
              
              
            End If
            RN2.MoveNext
        Wend
        RN2.Close


End Sub


Private Sub CargaMenuEstadistica(IdMenu As Integer)
'Dim GropCli As RibbonGroup
'Dim GrupPag As RibbonGroup
        

        
        'Creamos la TAB
        CrearTabPPal
        
        Cad = CStr(IdMenu * 100000)
        Set GroupNew = TabNuevo.Groups.AddGroup("Estadística", Cad & "2")
    


'
        
        'todos los hijos que cuelgan en la tab
        Cad = "Select * from menus where aplicacion = 'arigestion' and padre =" & IdMenu & " ORDER BY padre,orden"
        RN2.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        While Not RN2.EOF
        
           If Not BloqueaPuntoMenu(RN2!Codigo, "arigestion") Then
                Habilitado = True
    
                If Not MenuVisibleUsuario(DBLet(RN2!Codigo), "arigestion") Then
                    Habilitado = False
                Else
                    If Not MenuVisibleUsuario(DBLet(RN2!Padre), "arigestion") Then Habilitado = False
                End If
            End If
            
            
'            Select Case Rn2!Codigo
'            Case 401, 402, 403
'                Set Control = GropCli.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
 '           Case 404, 405, 406
 '               Set Control = GrupPag.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
 '           Case Else
                Set Control = GroupNew.Add(xtpControlButton, RN2!Codigo, RN2!Descripcion)
 '           End Select
            
            
            Cad = "NO"
            Control.Enabled = Habilitado
           ' ControlNew_NewItems.KeyboardTip = "V"
         
            RN2.MoveNext
        Wend
        RN2.Close


End Sub








Private Sub CargaMenuTrabajoDiario(IdMenu As Integer)
Dim Col As Collection

        
        
        
        'Este veremos si tiene alguna utilidad activa. Si es asi, crearemos la solapa, si no nada
        '.......................................................................
        
        
        'todos los hijos que cuelgan en la tab
        Cad = "Select * from menus where aplicacion = 'arigestion' and padre =" & IdMenu & " ORDER BY padre,orden"
        RN2.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        Set Col = New Collection
        While Not RN2.EOF
           I = I + 1
           If Not BloqueaPuntoMenu(RN2!Codigo, "arigestion") Then
                Habilitado = True
    
                If Not MenuVisibleUsuario(DBLet(RN2!Codigo), "arigestion") Then
                    Habilitado = False
                Else
                    If Not MenuVisibleUsuario(DBLet(RN2!Padre), "arigestion") Then Habilitado = False
                End If
            End If
            
            Col.Add Abs(Habilitado) & "|" & RN2!Codigo & "|" & RN2!Descripcion & "|"
            If Habilitado Then Cad = "S"
            
            RN2.MoveNext
        Wend
        RN2.Close
        
            '1408    "Traspaso de cuentas en apuntes"
            '1409    "Renumerar registros proveedor"
            '1410    "Aumentar dígitos contables"
            '1411    "Traspaso códigos de I.V.A."
            '1412    "Acciones realizadas"
            '1413    Importar fras cliente
            
        'Ya puedo utilizar numregelim
        If Cad <> "" Then
            'OK creamos solapa y demas
            'Creamos la TAB
            'Set TabNuevo = RibbonBar.InsertTab(CLng(IdMenu), "Trabajo diario")
            'TabNuevo.id = CLng(IdMenu)
            CrearTabPPal
            Set GroupNew = TabNuevo.Groups.AddGroup("Trabajo diario", 14000001)
            For NumRegElim = 1 To Col.Count
                Habilitado = CStr(RecuperaValor(Col.Item(NumRegElim), 1)) = "1"
                Set Control = GroupNew.Add(xtpControlButton, CLng(RecuperaValor(Col.Item(NumRegElim), 2)), CStr(RecuperaValor(Col.Item(NumRegElim), 3)))
                Control.Enabled = Habilitado
            Next
                
            
        End If
        

Set Col = Nothing
End Sub



