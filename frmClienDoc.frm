VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Begin VB.Form frmClienDoc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Visor"
   ClientHeight    =   11100
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11100
   ScaleWidth      =   13350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSaveAs 
      Caption         =   "&Guadar a disco"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   10320
      Width           =   1815
   End
   Begin VB.CommandButton cmdEliminar 
      Height          =   375
      Left            =   9720
      Picture         =   "frmClienDoc.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Eliminar documento"
      Top             =   480
      Width           =   375
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3000
      Top             =   10560
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.CommandButton cmdGuardar 
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
      Height          =   495
      Left            =   9960
      TabIndex        =   5
      Top             =   10320
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   4080
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAbrir 
      Height          =   375
      Left            =   9240
      Picture         =   "frmClienDoc.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Abrir"
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11520
      TabIndex        =   2
      Top             =   10320
      Width           =   1455
   End
   Begin AcroPDFLibCtl.AcroPDF AcroPDF1 
      Height          =   8895
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   12615
      _cx             =   22251
      _cy             =   15690
   End
   Begin VB.TextBox Text2 
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
      Height          =   390
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Width           =   8895
   End
   Begin VB.PictureBox PicContenedor 
      Height          =   8895
      Left            =   240
      ScaleHeight     =   8835
      ScaleWidth      =   12555
      TabIndex        =   8
      Top             =   1080
      Width           =   12615
   End
   Begin VB.Label Label1 
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
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmClienDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public IdCliente As Long
Public IdLinea As Integer


Dim descrip As String
Dim pathNombreFichero As String
Dim extension As String



' \\ Declaraciones Apis ( para el manifest y los temas de windows)
' ------------------------------------------------------------------------------------------
Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Private Declare Sub InitCommonControls Lib "Comctl32" ()
  
' \\ -- Declarar variable con evento para la clase
Private WithEvents mcPicScroll As cPicScroll
Attribute mcPicScroll.VB_VarHelpID = -1
  





Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If MsgBox("Desea eliminar el documento de la base de datos?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    If EjecutaSQL("DELETE FROM clientes_doc WHERE codclien =" & IdCliente & " AND id=" & IdLinea) Then Unload Me
End Sub

Private Sub cmdGuardar_Click()
Dim C As String
Dim id As Integer
Dim CargarIMgEnBD As Boolean
    On Error GoTo EcmdGuardar_Click
    
    If Text2.Text = "" Then Exit Sub
    If extension = "" Then Exit Sub
    If Not cmdEliminar.Visible Then
        If Val(cmdEliminar.Tag) = 0 Then Exit Sub   'No has modificado la imagen
    End If
    
    
    CargarIMgEnBD = True
    If cmdEliminar.Visible Then
        'MODIFICAR
        C = "UPDATE clientes_doc SET descDoc=" & DBSet(Text2.Text, "T")
        If cmdEliminar.Tag = 1 Then
            'Ha cambiado la imagen
            C = C & ", ext ='" & extension & "'"
        Else
            'NO ha cambiado la imagen
            CargarIMgEnBD = False
        End If
        C = C & " WHERE codclien =" & IdCliente & " AND id=" & IdLinea
        
    Else
        'NUEVO documento
        
        C = DevuelveDesdeBD("max(id)", "clientes_doc", "codclien", CStr(IdCliente))
        id = CInt(Val(C)) + 1
        
        C = "Insert into clientes_doc(codclien,id,descDoc,ext,fecha) VALUES (" & IdCliente & "," & id & "," & DBSet(Text2.Text, "T") & ",'" & extension & "',"
        C = C & DBSet(Now, "FH") & ")"
        

    End If
    Conn.Execute C
    
    espera 0.2
    
    If CargarIMgEnBD Then
        'Abro parar guardar el binary
        C = "Select * from clientes_doc where codclien =" & IdCliente & " AND ID=" & id
        Adodc1.ConnectionString = Conn
        Adodc1.RecordSource = C
        Adodc1.Refresh
        '
        If Adodc1.Recordset.EOF Then
            'MAAAAAAAAAAAAL
            MsgBox "Recordse EOF guardando imagen", vbCritical
        Else
            'Guardar
            ' InsertandoImg = True
            'CargarIMG lw1.ListItems(k).SubItems(2)
            GuardarBinary Adodc1.Recordset!ImgDOC, pathNombreFichero
            Adodc1.Recordset.Update
            
        End If
    End If
    Unload Me
    Exit Sub
EcmdGuardar_Click:
    MuestraError Err.Number, Err.Description
End Sub

Private Sub cmdAbrir_Click()
Dim J As Integer
Dim cad As String
         
    If Me.cmdEliminar.Visible Then
        If MsgBox("¿Desea modificar el documento?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
    
    
    
    cd1.FileName = ""
    'cd1.InitDir = "c:\"
    cd1.Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer
    cd1.MaxFileSize = 1024 * 30
    cd1.Filter = "Archivos JPG|*.jpg|Archivos TIFF|*.tif|Archivos PDF|*.pdf"
    cd1.CancelError = False
    cd1.ShowOpen
    
    If cd1.FileName = "" Then Exit Sub
            
    J = InStrRev(cd1.FileTitle, ".")
    If J = 0 Then
        MsgBox "archivo sin extension", vbExclamation
        Exit Sub
    End If
    
    extension = LCase(Mid(cd1.FileTitle, J + 1))
    cad = Mid(cd1.FileTitle, 1, J - 1)
    J = 0
    If extension = "jpg" Or extension = "tif" Or extension = "pdf" Then J = 1
    
    If J = 0 Then
        MsgBox "Extension no reconocida", vbExclamation
        Exit Sub
    End If
    
    Text2.Text = cad
    Text2.Enabled = True
    pathNombreFichero = cd1.FileName
    CargaFichero
    cmdEliminar.Tag = 1
            
End Sub


Private Sub TraeFicheroDesdedBD()
Dim C As String
     'Abro parar guardar el binary
     Text2.Text = "Leyendo BD..."
     Text2.Refresh
     
    Screen.MousePointer = vbHourglass
    C = "Select ImgDOC,descDoc,ext from clientes_doc where codclien =" & IdCliente & " AND ID=" & IdLinea
    Adodc1.ConnectionString = Conn
    Adodc1.RecordSource = C
    Adodc1.Refresh
    
    
    If Not Adodc1.Recordset.EOF Then
       Text2.Text = "Abriendo fichero"
       Text2.Refresh
       
       
       
        extension = Adodc1.Recordset!ext
        pathNombreFichero = App.Path & "\temp\" & Format(IdCliente, "000000") & Format(IdLinea, "000") & "." & extension
        If Dir(pathNombreFichero, vbArchive) <> "" Then Kill pathNombreFichero
        descrip = Adodc1.Recordset!descDoc
        If LeerBinary(Adodc1.Recordset!ImgDOC, pathNombreFichero) Then
            CargaFichero
            Text2.Text = descrip
            Text2.Enabled = True
        Else
            Text2.Text = "ERROR "
        End If
        
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub CargaFichero()
    Screen.MousePointer = vbHourglass
    If extension = "pdf" Then
        CargaPDF
    Else
        CargaIMG
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaPDF()
    On Error GoTo eCargaPDF
    
    Me.AcroPDF1.LoadFile pathNombreFichero
    AcroPDF1.Visible = True
    
    
    Exit Sub
eCargaPDF:
    MuestraError Err.Number, Err.Description
End Sub
Private Sub CargaIMG()
    On Error GoTo eCargaIMG
    
     mcPicScroll.SetPicture (pathNombreFichero)
    PicContenedor.Visible = True

    
    Exit Sub
eCargaIMG:
    MuestraError Err.Number, Err.Description
End Sub

Private Sub Command2_Click()
    
    
    
    
    
End Sub

Private Sub cmdSaveAs_Click()

On Error GoTo ecmdSaveAs
    
    cd1.FileName = "" ' Mid(pathNombreFichero, InStrRev(pathNombreFichero, "\") + 1)
    'cd1.InitDir = "c:\"
    cd1.Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer
    cd1.MaxFileSize = 1024 * 30
    If extension = "pdf" Then
        cd1.Filter = "Archivos PDF|*.pdf"
    Else
         cd1.Filter = "Archivos JPG|*.jpg"
    End If
    cd1.FilterIndex = 1
    cd1.CancelError = True
    cd1.ShowSave
    
    If cd1.FileName = "" Then Exit Sub
    
    If Dir(cd1.FileName, vbArchive) <> "" Then
        If MsgBox("Ya existe el archivo: " & vbCrLf & cd1.FileName & vbCrLf & "¿Sobreescribir?", vbQuestion + vbYesNoCancel) = vbYes Then
            Kill cd1.FileName
        End If
    End If
    
    FileCopy pathNombreFichero, cd1.FileName
     
     
ecmdSaveAs:
    If Err.Number <> 0 Then
        If Err.Number <> 32755 Then MuestraError Err.Number, Err.Description
        
        Err.Clear
    End If
End Sub

Private Sub Form_Activate()
    If Me.Tag = 1 Then
        Me.Tag = 0
        If IdLinea < 0 Then
            cmdEliminar.Visible = False
            cmdSaveAs.Visible = False
            Text2.Text = "Seleccionar fichero"
            'Boton nuevo para añadir
            cmdAbrir_Click
            
        Else
            'Abirir imagen
            cmdEliminar.Tag = 0 'Documento original
            cmdEliminar.Visible = True
            cmdSaveAs.Visible = True
            TraeFicheroDesdedBD
            
        End If
    End If
End Sub

Private Sub Form_Initialize()
    Call SetErrorMode(2)
    Call InitCommonControls
End Sub

Private Sub Form_Load()
    Me.Icon = frmppal.Icon
    Me.Tag = 1
    cmdEliminar.Tag = 0
        AcroPDF1.Visible = False
     PicContenedor.Visible = False
    
    
    ' -- Crear nueva instancia
    Set mcPicScroll = New cPicScroll
    ' -- Indicar al mÃ©todo Init, el control PicBox que serÃ¡ el contenedor
    Call mcPicScroll.Init(PicContenedor, Me)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mcPicScroll = Nothing
End Sub
