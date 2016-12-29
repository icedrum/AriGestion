VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEditorMenus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Editor de Menús"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   9795
   Icon            =   "frmEditorMenus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   9795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   2220
      MaxLength       =   20
      TabIndex        =   14
      Tag             =   "Aplicacion|T|N|||menus_usuarios|aplicacion||S|"
      Top             =   4890
      Width           =   855
   End
   Begin VB.CheckBox chkAux 
      BackColor       =   &H80000005&
      Height          =   255
      Index           =   4
      Left            =   6960
      TabIndex        =   13
      Tag             =   "Especial|N|N|0|1|menus_usuarios|especial|0|N|"
      Top             =   4920
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CheckBox chkAux 
      BackColor       =   &H80000005&
      Height          =   255
      Index           =   3
      Left            =   6270
      TabIndex        =   12
      Tag             =   "Imprimir|N|N|0|1|menus_usuarios|imprimir|0|N|"
      Top             =   4920
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CheckBox chkAux 
      BackColor       =   &H80000005&
      Height          =   255
      Index           =   2
      Left            =   5580
      TabIndex        =   11
      Tag             =   "Modificar|N|N|0|1|menus_usuarios|modificar|0|N|"
      Top             =   4920
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CheckBox chkAux 
      BackColor       =   &H80000005&
      Height          =   255
      Index           =   1
      Left            =   5010
      TabIndex        =   4
      Tag             =   "CrearEliminar|N|N|0|1|menus_usuarios|creareliminar|0|N|"
      Top             =   4920
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CheckBox chkAux 
      BackColor       =   &H80000005&
      Height          =   255
      Index           =   0
      Left            =   4440
      TabIndex        =   3
      Tag             =   "Ver|N|N|0|1|menus_usuarios|ver|0|N|"
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtAux2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   3180
      MaxLength       =   50
      TabIndex        =   2
      Top             =   4890
      Width           =   1155
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
      Left            =   7305
      TabIndex        =   5
      Top             =   5415
      Visible         =   0   'False
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
      Left            =   8475
      TabIndex        =   6
      Top             =   5415
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   1
      Left            =   900
      MaxLength       =   10
      TabIndex        =   1
      Tag             =   "Codigo|N|N|||menus_usuarios|codigo||S|"
      Top             =   4890
      Width           =   1230
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   60
      MaxLength       =   6
      TabIndex        =   0
      Tag             =   "Código|N|N|0|999999|menus_usuarios|codusu|000000|S|"
      Top             =   4890
      Width           =   800
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmEditorMenus.frx":000C
      Height          =   4950
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   9450
      _ExtentX        =   16669
      _ExtentY        =   8731
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   23
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
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
      Left            =   8460
      TabIndex        =   10
      Top             =   5430
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   5340
      Width           =   2385
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
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
         Left            =   45
         TabIndex        =   8
         Top             =   180
         Width           =   2295
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   375
      Left            =   2790
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         HelpContextID   =   2
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   2
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         HelpContextID   =   2
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnCalculomenus_usuariosProd 
         Caption         =   "&Cálculo menus_usuarios Prod."
         Shortcut        =   ^C
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnExportacion 
         Caption         =   "Exportar Excel"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnFiltro 
      Caption         =   "Filtro"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnFiltro1 
         Caption         =   "Año actual"
      End
      Begin VB.Menu mnFiltro2 
         Caption         =   "Año actual y anterior"
      End
      Begin VB.Menu mnBarra4 
         Caption         =   "-"
      End
      Begin VB.Menu mnFiltro3 
         Caption         =   "Sin Filtro"
      End
   End
End
Attribute VB_Name = "frmEditorMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: MONICA  +-+-
' +-+-+-+-+-+-+-+-+-+-+-+-

' **************** PER A QUE FUNCIONE EN UN ATRE MANTENIMENT ********************
' 0. Posar-li l'atribut Datasource a "adodc1" del Datagrid1. Canviar el Caption
'    del formulari
' 1. Canviar els TAGs i els Maxlength de TextAux(0) i TextAux(1)
' 2. En PonerModo(vModo) repasar els indexs del botons, per si es canvien
' 3. En la funció BotonAnyadir() canviar la taula i el camp per a SugerirCodigoSiguienteStr
' 4. En la funció BotonBuscar() canviar el nom de la clau primaria
' 5. En la funció BotonEliminar() canviar la pregunta, les descripcions de la
'    variable SQL i el contingut del DELETE
' 6. En la funció PonerLongCampos() posar els camps als que volem canviar el MaxLength quan busquem
' 7. En Form_Load() repasar la barra d'iconos (per si es vol canviar algún) i
'    canviar la consulta per a vore tots els registres
' 8. En Toolbar1_ButtonClick repasar els indexs de cada botó per a que corresponguen
' 9. En la funció CargaGrid canviar l'ORDER BY (normalment per la clau primaria);
'    canviar ademés els noms dels camps, el format i si fa falta la cantitat;
'    repasar els index dels botons modificar i eliminar.
'    NOTA: si en Form_load ya li he posat clausula WHERE, canviar el `WHERE` de
'    `SQL = CadenaConsulta & " WHERE " & vSQL` per un `AND`
' 10. En txtAux_LostFocus canviar el mensage i el format del camp
' 11. En la funció DatosOk() canviar els arguments de DevuelveDesdeBD i el mensage
'    en cas d'error
' 12. En la funció SepuedeBorrar() canviar les comprovacions per a vore si es pot
'    borrar el registre
' *******************************SI N'HI HA COMBO*******************************
' 0. Comprovar que en el SQL de Form_Load() es faça referència a la taula del Combo
' 1. Pegar el Combo1 al  costat dels TextAux. Canviar-li el TAG
' 2. En BotonModificar() canviar el camp del Combo
' 3. En CargaCombo() canviar la consulta i els noms del camps, o posar els valor
'    a ma si no es llig de cap base de datos els valors del Combo

Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'codi per al registe que s'afegix al cridar des d'atre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
Public NuevoCodigo As String
Public CodigoActual As String
Public DeConsulta As Boolean

Private CadenaConsulta As String
Private CadB As String

Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Dim Modo As Byte
'----------- MODOS --------------------------------
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'--------------------------------------------------
Dim PrimeraVez As Boolean
Dim indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim I As Integer

' utilizado para buscar por checks
Private BuscaChekc As String




Private Sub PonerModo(vModo)
Dim B As Boolean
Dim I As Integer

    Modo = vModo
    
    B = (Modo = 2)
    If B Then
        PonerContRegIndicador
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    For I = 0 To txtaux.Count - 1
        txtaux(I).Visible = False
    Next I
    txtAux2(0).Visible = False
    
    For I = 0 To chkAux.Count - 1
        chkAux(I).Visible = Not B
    Next I

    cmdAceptar.Visible = Not B
    cmdcancelar.Visible = Not B
    DataGrid1.Enabled = B
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.Visible = B
    
    PonerLongCampos
    
    'Si estamos modo Modificar bloquear clave primaria
    txtaux(0).Enabled = False
    txtaux(1).Enabled = False
    txtaux(2).Enabled = False
    
    chkAux(0).Enabled = (Modo = 4)
    chkAux(1).Enabled = (Modo = 4)
    chkAux(2).Enabled = (Modo = 4)
    chkAux(3).Enabled = (Modo = 4)
    chkAux(4).Enabled = (Modo = 4)
    
End Sub

Private Sub PonerContRegIndicador()
'si estamos en modo ver registros muestra el numero de registro en el que estamos
'situados del total de registros mostrados: 1 de 24
Dim cadReg As String

    If (Modo = 2 Or Modo = 0) Then
        cadReg = PonerContRegistros(Me.Adodc1)
        If CadB = "" Then
            lblIndicador.Caption = cadReg
        Else
            lblIndicador.Caption = "BUSQUEDA: " & cadReg
        End If
    End If
End Sub

Private Sub BotonAnyadir()
    Dim NumF As String
    Dim anc As Single
    
    CargaGrid CadB

    DataGrid1.AllowAddNew = True
    If Adodc1.Recordset.RecordCount > 0 Then
        DataGrid1.HoldFields
        Adodc1.Recordset.MoveLast
        DataGrid1.Row = DataGrid1.Row + 1
    End If
    
    If DataGrid1.Row < 0 Then
        anc = DataGrid1.top + 260
    Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.top
    End If
         
    For I = 0 To txtaux.Count - 1
        txtaux(I).Text = ""
    Next I
    txtaux(1).Text = Format(Now, formatoFechaVer)
    chkAux(0).Value = 0
    chkAux(1).Value = 0

    txtaux(2).Text = 0
    txtaux(3).Text = 0

    LLamaLineas anc, 3 'Pone el form en Modo=3, Insertar
       
    'Ponemos el foco
    PonFoco txtaux(0)
End Sub

Private Sub BotonVerTodos()
    CadB = ""
    CargaGrid CadB
    PonerModo 2
End Sub

Private Sub BotonBuscar()
    ' ***************** canviar per la clau primaria ********
    CargaGrid "menus_usuarios.codtraba = -1"
    '*******************************************************************************
    'Buscar
    For I = 0 To txtaux.Count - 1
        txtaux(I).Text = ""
    Next I
    chkAux(0).Value = 0
    chkAux(1).Value = 0
    
    
    LLamaLineas DataGrid1.top + 206, 1 'Pone el form en Modo=1, Buscar
    PonFoco txtaux(0)
End Sub

Private Sub BotonModificar()
    Dim anc As Single
    Dim I As Integer
    
    Screen.MousePointer = vbHourglass
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 320
    Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 360  '545
    End If

    'Llamamos al form
    txtaux(0).Text = CodigoActual
    txtaux(1).Text = DataGrid1.Columns(1).Text
    txtaux(2).Text = Adodc1.Recordset!aplicacion
    txtAux2(0).Text = DataGrid1.Columns(3).Text
    
    Me.chkAux(0).Value = Me.Adodc1.Recordset!Ver
    Me.chkAux(1).Value = Me.Adodc1.Recordset!creareliminar
    Me.chkAux(2).Value = Me.Adodc1.Recordset!Modificar
    Me.chkAux(3).Value = Me.Adodc1.Recordset!Imprimir
    Me.chkAux(4).Value = Me.Adodc1.Recordset!especial

    LLamaLineas anc, 4 'Pone el form en Modo=4, Modificar
   
    'Como es modificar
    PonleFoco chkAux(0)
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    
    'Fijamos el ancho
    For I = 0 To txtaux.Count - 1
        txtaux(I).top = alto
    Next I
    
    Me.chkAux(0).top = alto
    Me.chkAux(1).top = alto
    Me.chkAux(2).top = alto
    Me.chkAux(3).top = alto
    Me.chkAux(4).top = alto
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los txtAux
    PonerLongCamposGnral Me, Modo, 3
End Sub

Private Sub chkAux_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "chkAux(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "chkAux(" & Index & ")|"
    End If
End Sub

Private Sub chkAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
    Dim I As Long
    Dim vWhere As String
    

    Select Case Modo
        Case 1 'BUSQUEDA
            CadB = ObtenerBusqueda(Me)
            If CadB <> "" Then
                CargaGrid CadB
                PonerModo 2
                PonerFocoGrid Me.DataGrid1
            End If
            
        Case 3 'INSERTAR
            If DatosOK Then
                If InsertarDesdeForm(Me) Then
                    CargaGrid CadB
                    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
                        cmdCancelar_Click
                        If Not Adodc1.Recordset.EOF Then
                            Adodc1.Recordset.Find (Adodc1.Recordset.Fields(0).Name & " =" & NuevoCodigo)
                        End If
                        cmdRegresar_Click
                    Else
                        BotonAnyadir
                    End If
                    CadB = ""
                End If
            End If
            
        Case 4 'MODIFICAR
            If DatosOK Then
                If ModificaDesdeFormulario(Me) Then
                    ' en caso de modificar el padre, modificamos todos los hijos
                    If Adodc1.Recordset.Fields(0) Mod 10000 = 0 Then
                        ModificarHijos
                    End If
                    
                    TerminaBloquear
                    I = Adodc1.Recordset.Fields(1)
                    PonerModo 2
                    CargaGrid CadB
                    Adodc1.Recordset.Find "codigo=" & I

                    PonerFocoGrid Me.DataGrid1
                End If
            End If
    End Select
End Sub

Private Sub ModificarHijos()
Dim SQL As String

    SQL = "update menus_usuarios set ver = " & DBSet(chkAux(0).Value, "N")
    SQL = SQL & ", creareliminar = " & DBSet(chkAux(1).Value, "N")
    SQL = SQL & ", modificar = " & DBSet(chkAux(2).Value, "N")
    SQL = SQL & ", imprimir = " & DBSet(chkAux(3).Value, "N")
    SQL = SQL & ", especial = " & DBSet(chkAux(4).Value, "N")
    SQL = SQL & " where codusu = " & DBSet(CodigoActual, "N")
    SQL = SQL & " and aplicacion = 'arigestion'"
    SQL = SQL & " and codigo in (select codigo from menus where aplicacion = 'arigestion' and padre = " & DBSet(Adodc1.Recordset!Codigo, "N") & ")"

    Conn.Execute SQL
End Sub


Private Sub cmdCancelar_Click()
    On Error Resume Next
    
    Select Case Modo
        Case 1 'búsqueda
            CargaGrid CadB
        Case 3 'insertar
            DataGrid1.AllowAddNew = False
            CargaGrid
            If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveFirst
        Case 4 'modificar
            TerminaBloquear
    End Select
    
    PonerModo 2
    
    
    PonerFocoGrid Me.DataGrid1
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub cmdRegresar_Click()
Dim Cad As String
Dim I As Integer
Dim J As Integer
Dim Aux As String

    If Adodc1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    Cad = ""
    I = 0
    Do
        J = I + 1
        I = InStr(J, DatosADevolverBusqueda, "|")
        If I > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, I - J)
            J = Val(Aux)
            Cad = Cad & Adodc1.Recordset.Fields(J) & "|"
        End If
    Loop Until I = 0
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub

Private Sub DataGrid1_Click()
    BotonModificar
End Sub

Private Sub DataGrid1_DblClick()
'    If cmdRegresar.Visible Then cmdRegresar_Click

End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Modo = 2 Then PonerContRegIndicador
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault

    If PrimeraVez Then
        PrimeraVez = False
        If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
            BotonAnyadir
        Else
            PonerModo 2
             If Me.CodigoActual <> "" Then
                SituarData Me.Adodc1, "codusu=" & CodigoActual, "", True
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    PrimeraVez = True

    
    '****************** canviar la consulta *********************************
    CadenaConsulta = "SELECT cast(concat(right(concat('0000',menus.codigo),4),'0000') as signed),menus.codigo, menus_usuarios.aplicacion, if(cast(concat(right(concat('0000',menus.codigo),4),'0000') as signed) mod 10000<>0,concat('     ', menus.descripcion), menus.descripcion),menus_usuarios.ver, IF(menus_usuarios.ver=1,'*','') as pver, menus_usuarios.creareliminar,  IF(menus_usuarios.creareliminar=1,'*','') as pcreareliminar, menus_usuarios.modificar,  IF(menus_usuarios.modificar=1,'*','') as pmodificar, menus_usuarios.imprimir,  IF(menus_usuarios.imprimir=1,'*','') as pimprimir, menus_usuarios.especial, IF(menus_usuarios.especial=1,'*','') as pespecial "
    CadenaConsulta = CadenaConsulta & " from menus, menus_usuarios "
    CadenaConsulta = CadenaConsulta & " where menus.aplicacion = 'arigestion' "
    CadenaConsulta = CadenaConsulta & " and menus.padre = 0 "
    CadenaConsulta = CadenaConsulta & " and menus.codigo > 1 "
    CadenaConsulta = CadenaConsulta & " and menus.aplicacion = menus_usuarios.aplicacion and menus.codigo = menus_usuarios.codigo and menus_usuarios.codusu = " & DBSet(CodigoActual, "N")
    
    ' si no tiene tesoreria no los muestro
'    If Not vEmpresa.TieneTesoreria Then
'        CadenaConsulta = CadenaConsulta & " and not menus.codigo in (select codigo from menus where aplicacion = 'arigestion' and tipo = 1) "
'    End If
    
    CadenaConsulta = CadenaConsulta & " UNION "
    CadenaConsulta = CadenaConsulta & " select cast(concat(right(concat('0000',hh.padre),4), right(concat('0000',hh.orden),4)) as signed), hh.codigo, hh.aplicacion, if(cast(concat(right(concat('0000',hh.padre),4), right(concat('0000',hh.orden),4)) as signed) mod 10000<>0,concat('     ', hh.descripcion), hh.descripcion), uu.ver, IF(uu.ver=1,'*','') as pver, uu.creareliminar,  IF(uu.creareliminar=1,'*','') as pcreareliminar, uu.modificar,  IF(uu.modificar=1,'*','') as pmodificar, uu.imprimir,  IF(uu.imprimir=1,'*','') as pimprimir, uu.especial, IF(uu.especial=1,'*','') as pespecial  "
    CadenaConsulta = CadenaConsulta & " from menus pp, menus hh, menus_usuarios uu "
    CadenaConsulta = CadenaConsulta & " where pp.aplicacion = 'arigestion' and  "
    CadenaConsulta = CadenaConsulta & " hh.padre > 1 and "
    CadenaConsulta = CadenaConsulta & " pp.aplicacion = hh.aplicacion And hh.Padre = pp.Codigo and "
    CadenaConsulta = CadenaConsulta & " hh.aplicacion = uu.aplicacion and hh.codigo = uu.codigo and uu.codusu = " & DBSet(CodigoActual, "N")
    
    ' si no tiene tesoreria no los muestro
    'If Not vEmpresa.TieneTesoreria Then
    '    CadenaConsulta = CadenaConsulta & " and not hh.codigo in (select codigo from menus where aplicacion = 'arigestion' and tipo = 1) "
    'End If
    
    '************************************************************************
    
    CadB = ""
    CargaGrid ""
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Modo = 4 Then TerminaBloquear
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    If indice = 1 Then
        txtaux(1).Text = Format(vFecha, formatoFechaVer) '<===
    Else
        txtaux(5).Text = Format(vFecha, formatoFechaVer) '<===
    End If
    ' ********************************************
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub


Private Sub mnModificar_Click()
    'Comprobaciones
    '--------------
    If Adodc1.Recordset.EOF Then Exit Sub
    
    If Adodc1.Recordset.RecordCount < 1 Then Exit Sub
    
    'Preparamos para modificar
    '-------------------------
    If BLOQUEADesdeFormulario(Me) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub CargaGrid(Optional vSQL As String)
    Dim SQL As String
    Dim tots As String
    Dim cadFiltro As String
    
    
'    adodc1.ConnectionString = Conn
    If vSQL <> "" Then
        SQL = CadenaConsulta & " AND " & vSQL
    Else
        SQL = CadenaConsulta
    End If
    SQL = SQL & " ORDER BY 1 "
    
    CargaGridGnral Me.DataGrid1, Me.Adodc1, SQL, PrimeraVez
    
    ' *******************canviar els noms i si fa falta la cantitat********************
    tots = "N||||0|;N||||0|;N||||0|;S|txtAux2(0)|T|Descripcion|5000|;"
    tots = tots & "N||||0|;S|chkAux(0)|CB|Ver|760|;N||||0|;S|chkAux(1)|CB|Ins.Eli|760|;"
    tots = tots & "N||||0|;S|chkAux(2)|CB|Modif|760|;N||||0|;S|chkAux(3)|CB|Impr|760|;"
    tots = tots & "N||||0|;S|chkAux(4)|CB|Esp|760|;"
    arregla tots, DataGrid1, Me
    
    DataGrid1.ScrollBars = dbgAutomatic
    DataGrid1.Columns(0).Alignment = dbgLeft
    
    DataGrid1.Columns(5).Alignment = dbgCenter
    DataGrid1.Columns(7).Alignment = dbgCenter
    DataGrid1.Columns(9).Alignment = dbgCenter
    DataGrid1.Columns(11).Alignment = dbgCenter
    DataGrid1.Columns(13).Alignment = dbgCenter
    
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    If Index = 6 And Modo = 3 Then
        txtaux(Index).Enabled = False
    End If

    ConseguirFoco txtaux(Index), Modo
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String

    If Not PerderFocoGnral(txtaux(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 0
    End Select
    
End Sub

Private Function DatosOK() As Boolean
'Dim Datos As String
Dim B As Boolean
Dim SQL As String
Dim Mens As String


    B = CompForm(Me)
    If Not B Then Exit Function
    
    
    DatosOK = B
End Function


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 2: KEYBusqueda KeyAscii, 0 'cuenta contable
            End Select
        End If
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvançar/Retrocedir els camps en les fleches de desplaçament del teclat.
    KEYdown KeyCode
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub



Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
End Sub



