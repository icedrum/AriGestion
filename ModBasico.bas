Attribute VB_Name = "ModBasico"
Option Explicit

Public Sub arregla(ByRef tots As String, ByRef grid As DataGrid, ByRef formu As Form)
    'Dim tots As String
    Dim camp As String
    Dim Mens As String
    Dim difer As Integer
    Dim I As Integer
    Dim k As Integer
    Dim posi As Integer
    Dim posi2 As Integer
    Dim fil As Integer
    Dim C As Integer
    Dim o As Integer
    Dim A() As Variant 'per als 5 parametres
    'Dim grid As DataGrid
    Dim Obj As Object
    Dim obj_ant As Object
    Dim primer As Boolean
    Dim TotalAncho As Integer
    
    grid.AllowRowSizing = False
    grid.RowHeight = 350
    
    '***********
    difer = 563 'dirència recomanda entre l'ample del Datagrid i la suma dels amples de les columnes
    '***********
    
    TotalAncho = 0
    primer = False
'    Set grid = DataGrid1 'nom del DataGrid
    fil = -1 'fila a -1
    C = -1 'columna del datagrid a 0
    'tots = "S|txtAux(0)|T|Código|700|;S|txtAux(1)|T|Descripción|3000|;"
    
    While (tots <> "") 'bucle per a recorrer els distins camps
        Set Obj = Nothing
        Set obj_ant = Nothing
    
        fil = fil + 1
        'ReDim Preserve A(6, fil)
        ReDim Preserve A(5, fil)
        'fila i columna a 0 (NOTA: les files es numeren a partir d'1 i les columnes a partir de 0)
        posi = InStr(tots, ";") '1ª posicio del ;
        camp = Left(tots, posi - 1)
        tots = Right(tots, Len(tots) - posi) 'lleve el camp actual
        'For k = 0 To 5
        For k = 0 To 4
          posi2 = InStr(camp, "|") '1ª posició del |
          A(k, fil) = Left(camp, posi2 - 1)
          camp = Right(camp, Len(camp) - posi2) 'lleve l'argument actual
        Next k 'quan acabe el for tinc en A el camp actual
        
        'només incremente el nº de la columna si no es un boto
        If A(2, fil) <> "B" Then C = C + 1
        
        If A(0, fil) = "N" Then 'no visible
            grid.Columns(C).Visible = False
            grid.Columns(C).Width = 0 'si no es visible, pose a 0 l'ample
        ElseIf A(0, fil) = "S" Then 'visible
            ' ********* CAPTION I WIDTH DE L'OBJECTE ************
            
            Select Case A(2, fil) 'tipo (T, C o B) (o CB=CheckBox ) (DT=DTPicker)
                Case "T"
                    grid.Columns(C).Visible = True
                    If A(3, fil) <> "" Then grid.Columns(C).Caption = A(3, fil)
                    If A(4, fil) <> "" Then grid.Columns(C).Width = CInt(A(4, fil))
'                    If A(5, fil) <> "" Then
'                        grid.Columns(c).NumberFormat = A(5, fil)
'                    Else
'                        grid.Columns(c).NumberFormat = ""
'                    End If
                    TotalAncho = TotalAncho + CInt(A(4, fil))
                Case "C"
                    grid.Columns(C).Visible = True
                    If A(3, fil) <> "" Then grid.Columns(C).Caption = A(3, fil)
                    If A(4, fil) <> "" Then grid.Columns(C).Width = CInt(A(4, fil)) - 10
'                    If A(5, fil) <> "" Then
'                        grid.Columns(c).NumberFormat = A(5, fil)
'                    Else
'                        grid.Columns(c).NumberFormat = ""
'                    End If
                    TotalAncho = TotalAncho + CInt(A(4, fil))
                Case "B"
                
               '=== LAURA (07/04/06): añadir tipo CB=CheckBox
                Case "CB"
                    grid.Columns(C).Visible = True
                    If A(3, fil) <> "" Then grid.Columns(C).Caption = A(3, fil)
                    If A(4, fil) <> "" Then grid.Columns(C).Width = CInt(A(4, fil))
                    TotalAncho = TotalAncho + CInt(A(4, fil))
               '===============================================
               '=== LAURA (14/07/06): añadir tipo DT=DTPicker
                Case "DT"
                    grid.Columns(C).Visible = True
                    If A(3, fil) <> "" Then grid.Columns(C).Caption = A(3, fil)
                    If A(4, fil) <> "" Then grid.Columns(C).Width = CInt(A(4, fil))
                    TotalAncho = TotalAncho + CInt(A(4, fil))
                '==============================================
            End Select
                       
            ' ********* CARREGUE L'OBJECTE ************
            Set Obj = eval(formu, CStr(A(1, fil)))
            
            ' ********* NUMBERFORMAT i ALIGNMENT DE L'OBJECTE ************
            If (A(2, fil) = "T") Or (A(2, fil) = "C") Or (A(2, fil) = "DT") Then 'el numberformat només es per a text o combo
                If Obj.Tag <> "" Then
                    grid.Columns(C).NumberFormat = FormatoCampo2(Obj)
                    If TipoCamp(Obj) = "N" Then
                        If (A(2, fil) = "T") Then _
                            grid.Columns(C).Alignment = dbgRight ' el Alignment només per a Text
                        grid.Columns(C).NumberFormat = grid.Columns(C).NumberFormat & " "
                    End If
                Else
                    grid.Columns(C).NumberFormat = ""
                End If
            End If
            
            ' ********* WIDTH I LEFT DE L'OBJECTE ************
            Select Case A(2, fil) 'tipo (T, C o B)
                Case "T"
                    If Not primer Then 'es el primer objecte visible
                        Obj.Width = grid.Columns(C).Width - 60
                        'obj.Width = grid.Columns(c).Width - 8
                        Obj.Left = grid.Left + 340
                        'obj.Left = grid.Left + 308
                    Else
                        o = 0
                        While obj_ant Is Nothing
                            o = o + 1
                            If A(0, fil - o) = "S" Then
                                Set obj_ant = eval(formu, CStr(A(1, fil - o)))
                            End If
                        Wend
                        'en obj_ant tindré el 1r objecte per darrere que siga visible
                        Select Case A(2, fil - o)
                            Case "T" 'objecte anterior a text es text
                                Obj.Width = grid.Columns(C).Width - 60
                                'obj.Width = grid.Columns(c).Width - 38
                                Obj.Left = obj_ant.Left + obj_ant.Width + 60
                                'obj.Left = obj_ant.Left + obj_ant.Width + 38
                            Case "C" 'objecte anterior a text es combo
                                Obj.Width = grid.Columns(C).Width - 60
                                Obj.Left = obj_ant.Left + obj_ant.Width + 30
                            Case "B" 'objecte anterior a text es un boto
                                Obj.Width = grid.Columns(C).Width - 60
                                Obj.Left = obj_ant.Left + obj_ant.Width + 30
                                
                             '=== LAURA (07/04/06): añadir tipo CB=CheckBox
                            Case "CB" 'anterior es CheckBox
                                Obj.Width = grid.Columns(C).Width - 60
                                Obj.Left = obj_ant.Left + obj_ant.Width + 60
                            '=== LAURA (14/07/06): añadir tipo DT=DTPicker
                            Case "DT" 'anterior es un DTPicker
                                Obj.Width = grid.Columns(C).Width - 60
                                Obj.Left = obj_ant.Left + obj_ant.Width + 60
                        End Select
                    End If
                    
                Case "C"
                    If Not primer Then 'es el primer objecte visible
                        Obj.Width = grid.Columns(C).Width - 10
                        Obj.Left = grid.Left + 320
                    Else
                        o = 0
                        While obj_ant Is Nothing
                            o = o + 1
                            If A(0, fil - o) = "S" Then
                                Set obj_ant = eval(formu, CStr(A(1, fil - o)))
                            End If
                        Wend
                        'en obj_ant tindré el 1r objecte per darrere que siga visible
                        Select Case A(2, fil - o)
                            Case "T" 'objecte anterior a combo es text
                                Obj.Width = grid.Columns(C).Width - 20
                                Obj.Left = obj_ant.Left + obj_ant.Width + 40
                            Case "C" 'objecte anterior a combo es combo
                                Obj.Width = grid.Columns(C).Width
                                Obj.Left = obj_ant.Left + obj_ant.Width
                            Case "B" 'objecte anterior a combo es un boto
                                ' *** FALTA PER A QUAN L'OBJECTE ANTERIOR A UN COMBO ES UN BOTO
'                                mens = "Falta programar en arreglaGrid per al cas que un l'objecte anterior a un ComboBox es un Button"
'                                MsgBox "MÒDUL arreglaGrid:" & vbCrLf & "-----------------------" & vbCrLf & vbCrLf & mens
                                '=== LAURA (14/09/06): añadir este caso
                                Obj.Width = grid.Columns(C).Width
                                Obj.Left = obj_ant.Left + obj_ant.Width + 10
                            
                            '=== LAURA (07/04/06): añadir tipo CB=CheckBox
                            Case "CB" 'anterior es CheckBox (falta comprobar)
                                Obj.Width = grid.Columns(C).Width
                                Obj.Left = obj_ant.Left + obj_ant.Width
                        End Select
                    End If
                    
                Case "B"
                    If Not primer Then 'es el primer objecte visible
                        ' *** FALTA PER A QUAN UN BOTO ES EL PRIMER OBJECTE VISIBLE
                        Mens = "Falta programar en arreglaGrid per al cas que un Button es el primer objete visible d'un Datagrid"
                        MsgBox "MÒDUL arreglaGrid:" & vbCrLf & "-----------------------" & vbCrLf & vbCrLf & Mens
                    Else
                        o = 0
                        While obj_ant Is Nothing
                            o = o + 1
                            If A(0, fil - o) = "S" Then
                                Set obj_ant = eval(formu, CStr(A(1, fil - o)))
                            End If
                        Wend
                        'en obj_ant tindré el 1r objecte per darrere que siga visible
                        Select Case A(2, fil - o)
                            Case "T" 'objecte anterior a boto es text
                                obj_ant.Width = obj_ant.Width - Obj.Width + 30 '1r faig més curt l'objecte de text
                                Obj.Left = obj_ant.Left + obj_ant.Width
                                'obj.Left = obj_ant.Left + obj_ant.Width - obj.Width
                            Case "C" 'objecte anterior a boto es combo
                                ' *** FALTA PER A QUAN L'OBJECTE ANTERIOR A UN BOTO ES UN COMBO
                                Mens = "Falta programar en arreglaGrid per al cas que un l'objecte anterior a un Button es un ComboBox"
                                MsgBox "MÒDUL arreglaGrid:" & vbCrLf & "-----------------------" & vbCrLf & vbCrLf & Mens
                            Case "B" 'objecte anterior a combo es un boto
                                ' *** FALTA PER A QUAN L'OBJECTE ANTERIOR A UN BOTO ES UN BOTO
                                Mens = "Falta programar en arreglaGrid per al cas que un l'objecte anterior a un Button es un Button"
                                MsgBox "MÒDUL arreglaGrid:" & vbCrLf & "-----------------------" & vbCrLf & vbCrLf & Mens
                        End Select
                    End If
                    
                 '=== LAURA (07/04/06): añadir tipo CB=CheckBox
                Case "CB"
                    If Not primer Then 'es el primer objecte visible
                        Obj.Width = grid.Columns(C).Width - 10
                        Obj.Left = grid.Left + 320
                    Else
                        o = 0
                        While obj_ant Is Nothing
                            o = o + 1
                            If A(0, fil - o) = "S" Then
                                Set obj_ant = eval(formu, CStr(A(1, fil - o)))
                            End If
                        Wend
                        'en obj_ant tindré el 1r objecte per darrere que siga visible
                        Select Case A(2, fil - o)
                            Case "T" 'objecte anterior a combo es text
                                Obj.Width = grid.Columns(C).Width - (grid.Columns(C).Width / 3)
                                Obj.Left = obj_ant.Left + obj_ant.Width + (grid.Columns(C).Width / 3) - 10
                            Case "C" 'objecte anterior a combo es combo
                                Obj.Width = grid.Columns(C).Width
                                Obj.Left = obj_ant.Left + obj_ant.Width
                            Case "B" 'objecte anterior a combo es un boto
                                ' *** FALTA PER A QUAN L'OBJECTE ANTERIOR A UN COMBO ES UN BOTO
'Laura: 140508
'                                mens = "Falta programar en arreglaGrid per al cas que un l'objecte anterior a un ComboBox es un Button"
'                                MsgBox "MÒDUL arreglaGrid:" & vbCrLf & "-----------------------" & vbCrLf & vbCrLf & mens
                                
                                Obj.Width = grid.Columns(C).Width
                                Obj.Left = obj_ant.Left + obj_ant.Width + 10
                                
                             '=== LAURA (07/04/06): añadir tipo CB=CheckBox
                            Case "CB" 'anterior es un ChekBox
                                Obj.Width = grid.Columns(C).Width - (grid.Columns(C).Width / 3)
                                Obj.Left = obj_ant.Left + obj_ant.Width + (grid.Columns(C).Width / 3)
                        End Select
                    End If
                
                
                 '=== LAURA (14/07/06): añadir tipo DT=DTPicker
                Case "DT"
                    If Not primer Then 'es el primer objecte visible
                        Obj.Width = grid.Columns(C).Width - 10
                        Obj.Left = grid.Left + 320
                    Else
                        o = 0
                        While obj_ant Is Nothing
                            o = o + 1
                            If A(0, fil - o) = "S" Then
                                Set obj_ant = eval(formu, CStr(A(1, fil - o)))
                            End If
                        Wend
                        'en obj_ant tindré el 1r objecte per darrere que siga visible
                        Select Case A(2, fil - o)
                            Case "T" 'objecte anterior a combo es text
                                Obj.Width = grid.Columns(C).Width - 40
                                Obj.Left = obj_ant.Left + obj_ant.Width + 40
                            Case "C" 'objecte anterior a combo es combo
                                Obj.Width = grid.Columns(C).Width
                                Obj.Left = obj_ant.Left + obj_ant.Width
                            Case "B" 'objecte anterior a combo es un boto
                                ' *** FALTA PER A QUAN L'OBJECTE ANTERIOR A UN COMBO ES UN BOTO
                                Mens = "Falta programar en arreglaGrid per al cas que un l'objecte anterior a un ComboBox es un Button"
                                MsgBox "MÒDUL arreglaGrid:" & vbCrLf & "-----------------------" & vbCrLf & vbCrLf & Mens
                                
                             '=== LAURA (07/04/06): añadir tipo CB=CheckBox
                            Case "CB" 'anterior es un ChekBox
                                Obj.Width = grid.Columns(C).Width - (grid.Columns(C).Width / 3)
                                Obj.Left = obj_ant.Left + obj_ant.Width + (grid.Columns(C).Width / 3)
                            Case "DT" 'anterior es un DTPicker
                                Obj.Width = grid.Columns(C).Width - 40
                                Obj.Left = obj_ant.Left + obj_ant.Width + 40
                        End Select
                    End If
                Case Else
                    MsgBox "No existix el tipo de control " & A(2, fil)
            End Select
            
        primer = True
        End If
                
    Wend

    'No permitir canviar tamany de columnes
    For I = 0 To grid.Columns.Count - 1
         grid.Columns(I).AllowSizing = False
    Next I

'    If grid.Width - TotalAncho <> difer Then
'        mens = "Es recomana que el total d'amples de les columnes per a este DataGrid siga de "
'        mens = mens & CStr(grid.Width - difer)
'        MsgBox "MÒDUL arreglaGrid:" & vbCrLf & "-----------------------" & vbCrLf & vbCrLf & mens
'    End If
End Sub

Public Function eval(ByRef formu As Form, nom_camp As String) As Control
Dim Ctrl As Control
Dim nom_camp2 As String
Dim nou_i As Integer
Dim J As Integer

    Set eval = Nothing
    J = InStr(1, nom_camp, "(")
    If J = 0 Then
        nou_i = -1
    Else
        nom_camp = Left(nom_camp, Len(nom_camp) - 1)
        nou_i = Val(Mid(nom_camp, J + 1))
        nom_camp = Left(nom_camp, J - 1)
    End If
    
    For Each Ctrl In formu.Controls
        If Ctrl.Name = nom_camp Then
            If nou_i >= 0 Then
                If nou_i = Ctrl.Index Then
                    J = 1 'coincidix el nom i l'index
                Else
                    J = 0 'coincidix el nom però no l'index
                End If
            Else
                J = 1 'coincidix el nom i no te index
            End If
        Else
            J = -1 'no coincidix el nom
        End If
        
        If J > 0 Then
            Set eval = Ctrl
            Exit For
        End If
    Next Ctrl
End Function


Public Function PerderFocoGnral(ByRef Text As TextBox, Modo As Byte) As Boolean
Dim Comprobar As Boolean
'Dim mTag As CTag

    On Error Resume Next

    If Screen.ActiveForm.ActiveControl.Name = "cmdCancelar" Then
        PerderFocoGnral = False
        Exit Function
    End If

    With Text
        'Quitamos blancos por los lados
        .Text = Trim(.Text)
        
        
         If .BackColor = vbLightBlue Then
            If .Locked Then
                .BackColor = vbLightBlue '&H80000018
            Else
                .BackColor = vbWhite
            End If
        End If
        
        
        'Si no estamos en modo: 3=Insertar o 4=Modificar o 1=Busqueda, no hacer ninguna comprobacion
        If (Modo <> 3 And Modo <> 4 And Modo <> 1 And Modo <> 5) Then
            PerderFocoGnral = False
            Exit Function
        End If
        
        If Modo = 1 Then
            'Si estamos en modo busqueda y contiene un caracter especial no realizar
            'las comprobaciones
            Comprobar = ContieneCaracterBusqueda(.Text)
            If Comprobar Then
                PerderFocoGnral = False
                Exit Function
            End If
        End If
        PerderFocoGnral = True
    End With
    
    If Err.Number <> 0 Then Err.Clear
End Function



Public Function ExisteCP(T As TextBox) As Boolean
'comprueba para un campo de texto que sea clave primaria, si ya existe un
'registro con ese valor
Dim vTag As CTag
Dim Devuelve As String

    On Error GoTo ErrExiste

    ExisteCP = False
    If T.Text <> "" Then
        If T.Tag <> "" Then
            Set vTag = New CTag
            If vTag.Cargar(T) Then
'                If vtag.EsClave Then
                    Devuelve = DevuelveDesdeBD(vTag.Columna, vTag.tabla, vTag.Columna, T.Text, vTag.TipoDato)
                    If Devuelve <> "" Then
    '                    MsgBox "Ya existe un registro para " & vtag.Nombre & ": " & T.Text, vbExclamation
                        MsgBox "Ya existe el " & vTag.Nombre & ": " & T.Text, vbExclamation
                        ExisteCP = True
                        PonFoco T
                    End If
'                End If
            End If
            Set vTag = Nothing
        End If
    End If
    Exit Function
    
ErrExiste:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar código.", Err.Description
End Function



Public Function FormatoCampo2(ByRef objec As Object) As String
'Devuelve el formato del campo en el TAg: "0000"
Dim mTag As CTag
Dim Cad As String

    On Error GoTo EFormatoCampo2

    Set mTag = New CTag
    mTag.Cargar objec
    If mTag.Cargado Then
        FormatoCampo2 = mTag.Formato
    End If
    
EFormatoCampo2:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Function

Public Function TipoCamp(ByRef objec As Object) As String
Dim mTag As CTag
Dim Cad As String

    On Error GoTo ETipoCamp

    Set mTag = New CTag
    mTag.Cargar objec
    If mTag.Cargado Then
        TipoCamp = mTag.TipoDato
    End If

ETipoCamp:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Function


Public Function ContieneCaracterBusqueda(CADENA As String) As Boolean
'Comprueba si la cadena contiene algun caracter especial de busqueda
' >,>,>=,: , ....
'si encuentra algun caracter de busqueda devuelve TRUE y sale
Dim B As Boolean
Dim I As Integer
Dim Ch As String

    'For i = 1 To Len(cadena)
    I = 1
    B = False
    Do
        Ch = Mid(CADENA, I, 1)
        Select Case Ch
            Case "<", ">", ":", "="
                B = True
            Case "*", "%", "?", "_", "\", ":" ', "."
                B = True
            Case Else
                B = False
        End Select
    'Next i
        I = I + 1
    Loop Until (B = True) Or (I > Len(CADENA))
    ContieneCaracterBusqueda = B
End Function



'-----------------------------------------------------------------------

Public Sub AyudaFormaPago(frmBas As frmBasico, Optional CodActual As String, Optional cWhere As String)

    frmBas.CadenaTots = "S|txtAux(0)|T|Código|870|;S|txtAux(1)|T|Descripción|5230|;"
    frmBas.CadenaConsulta = "SELECT formapago.codforpa, formapago.nomforpa "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM ariconta" & vParam.Numconta & ".formapago"
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "Código|T|N|||formapago|codforpa||S|"
    frmBas.Tag2 = "Descripción|T|N|||formapago|nomforpa|||"
    
    frmBas.Maxlen1 = 4
    frmBas.Maxlen2 = 30
    
    frmBas.tabla = "ariconta" & vParam.Numconta & ".formapago"
    frmBas.CampoCP = "codforpa"
    frmBas.Caption = "Formas de pago"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
    
End Sub


Public Sub AyudaClientesBasico(frmBas As frmBasico, Optional CodActual As String, Optional cWhere As String)

    frmBas.CadenaTots = "S|txtAux(0)|T|Código|870|;S|txtAux(1)|T|Descripción|5230|;"
    frmBas.CadenaConsulta = "SELECT codclien,nomclien FROm clientes  WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "Código|T|N|||clientes|codclien||S|"
    frmBas.Tag2 = "Descripción|T|N|||clientes|nomclien|||"
    
    frmBas.Maxlen1 = 4
    frmBas.Maxlen2 = 30
    
    frmBas.tabla = "clientes"
    frmBas.CampoCP = "codclien"
    frmBas.Caption = "Clientes"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
    
End Sub



Public Sub AyudaIVA(frmBas As frmBasico, Optional CodActual As String, Optional cWhere As String)

    frmBas.CadenaTots = "S|txtAux(0)|T|Código|870|;S|txtAux(1)|T|Descripción|5230|;"
    frmBas.CadenaConsulta = "SELECT codigiva,concat(replace(nombriva,'(',''),' ',format(porceiva,2)) nombriva"
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM ariconta" & vParam.Numconta & ".tiposiva"
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "Código|T|N|||tiposiva|codigiva||S|"
    frmBas.Tag2 = "Descripción|T|N|||tiposiva|nombriva|||"
    
    frmBas.Maxlen1 = 4
    frmBas.Maxlen2 = 30
    
    frmBas.tabla = "ariconta" & vParam.Numconta & ".tiposiva"
    frmBas.CampoCP = "codigiva"
    frmBas.Caption = "Tipos de IVA"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
    
End Sub


Public Sub AyudaCtasContabilidad(frmBas As frmBasico, Optional CodActual As String, Optional cWhere As String)

    frmBas.CadenaTots = "S|txtAux(0)|T|Código|1470|;S|txtAux(1)|T|Descripción|4630|;"
    frmBas.CadenaConsulta = "SELECT codmacta,nommacta "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM ariconta" & vParam.Numconta & ".cuentas"
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE apudirec='S' "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "Código|T|N|||cuentas|codmacta||S|"
    frmBas.Tag2 = "Descripción|T|N|||cuentas|nommacta|||"
    
    frmBas.Maxlen1 = 10
    frmBas.Maxlen2 = 30
    
    frmBas.tabla = "ariconta" & vParam.Numconta & ".cuentas"
    frmBas.CampoCP = "codmacta"
    frmBas.Caption = "Cuentas contables"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
    
End Sub




Public Sub AyudaAccionesCliente(frmBas As frmBasico, Optional CodActual As String, Optional cWhere As String)

    frmBas.CadenaTots = "S|txtAux(0)|T|Código|2300|;S|txtAux(1)|T|Descripción|3730|;"
    frmBas.CadenaConsulta = "SELECT concat( right(concat('000000',id),6),' ',fechahora ) c1,concat(clientes_historial.codclien,'-',nomclien) c2"
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM clientes_historial,clientes"
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE clientes_historial.codclien=clientes.codclien "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "Código|T|N||||c1||S|"
    frmBas.Tag2 = "Descripción|T|N||||c2|||"
    
    frmBas.Maxlen1 = 30
    frmBas.Maxlen2 = 30
    
    frmBas.tabla = "clientes_historial,clientes"
    frmBas.CampoCP = "id"
    frmBas.Caption = "Acciones clientes"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
    
End Sub

Public Sub AyudaContadoresBasico(frmBas As frmBasico, Optional CodActual As String, Optional cWhere As String)
    'contadores   serfactur nomregis
    frmBas.CadenaTots = "S|txtAux(0)|T|Código|870|;S|txtAux(1)|T|Descripción|5230|;"
    frmBas.CadenaConsulta = "SELECT serfactur,nomregis FROm contadores  WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "Código|T|N|||contadores|serfactur||S|"
    frmBas.Tag2 = "Descripción|T|N|||contadores|nomregis|||"
    
    frmBas.Maxlen1 = 4
    frmBas.Maxlen2 = 30
    
    frmBas.tabla = "contadores"
    frmBas.CampoCP = "tiporegi"
    frmBas.Caption = "Contadores"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
    
End Sub


