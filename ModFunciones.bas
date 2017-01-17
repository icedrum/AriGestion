Attribute VB_Name = "ModFunciones"
'////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////
'   En este modulo estan las funciones que recorren el form
'   usando el each for
'   Estas son
'
'   CamposSiguiente -> Nos devuelve el el text siguiente en
'           el orden del tabindex
'
'   CompForm -> Compara los valores con su tag
'
'   InsertarDesdeForm - > Crea el sql de insert e inserta
'
'   Limpiar -> Pone a "" todos los objetos text de un form
'
'   ObtenerBusqueda -> A partir de los text crea el sql a
'       partir del WHERE ( sin el).
'
'   ModifcarDesdeFormulario -> Opcion modificar. Genera el SQL
'
'   PonerDatosForma -> Pone los datos del RECORDSET en el form
'////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////
Option Explicit

Public Const FormatoFechaHora = "yyyy-mm-dd hh:nn:ss"
Public Const ValorNulo = "Null"

Public Function CompForm(ByRef formulario As Form) As Boolean
    Dim Control As Object
    Dim mTag As CTag
    Dim Carga As Boolean
    Dim Correcto As Boolean
       
    Dim HayCamposIncorrectos As Boolean
    Dim CampoIncorrecto As String
    
    HayCamposIncorrectos = False
    CampoIncorrecto = ""
       
    CompForm = False
    Set mTag = New CTag
    For Each Control In formulario.Controls
        'TEXT BOX
        If TypeOf Control Is TextBox Then
            Carga = mTag.Cargar(Control)
            If Carga = True Then
                Correcto = mTag.Comprobar(Control, True)
'                If Not Correcto Then Exit Function
                If Not Correcto Then
                    Control.BackColor = vbErrorColor
                    HayCamposIncorrectos = True
                    CampoIncorrecto = Control.Name
                    If IsArray(Control) Then CampoIncorrecto = CampoIncorrecto & "(" & Control.Index & ")"
                Else
                    Control.BackColor = vbWhite
                End If
            Else
                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                Exit Function
            End If
        'COMBOBOX
        ElseIf TypeOf Control Is ComboBox Then
            'Comprueba que los campos estan bien puestos
            If Control.Tag <> "" Then
                Carga = mTag.Cargar(Control)
                If Carga = False Then
                    MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                    Exit Function
                Else
                    If mTag.Vacio = "N" And Control.ListIndex < 0 Then
'                            MsgBox "Seleccione una dato para: " & mTag.Nombre, vbExclamation
'                            Exit Function
                        Control.BackColor = vbErrorColor
                        HayCamposIncorrectos = True
                        CampoIncorrecto = Control.Name
                        If IsArray(Control) Then CampoIncorrecto = CampoIncorrecto & "(" & Control.Index & ")"
                    Else
                        Control.BackColor = vbWhite
                    End If
                End If
            End If
        End If
    Next Control
    
    If HayCamposIncorrectos Then
        MsgBox "Revise datos obligatorios o incorrectos", vbExclamation
    End If
    CompForm = Not HayCamposIncorrectos
    
End Function

Public Function CompForm2(ByRef formulario As Form, Optional opcio As Integer, Optional nom_frame As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Carga As Boolean
Dim Correcto As Boolean

    Dim HayCamposIncorrectos As Boolean
    Dim CampoIncorrecto As String
    
    HayCamposIncorrectos = False
    CampoIncorrecto = ""



    CompForm2 = False
    Set mTag = New CTag
    For Each Control In formulario.Controls
        'TEXT BOX
        If TypeOf Control Is TextBox Then
            If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                Carga = mTag.Cargar(Control)
                If Carga = True Then
                    Correcto = mTag.Comprobar(Control, True)
'                    If Not Correcto Then Exit Function
                    If Not Correcto Then
                        Control.BackColor = vbErrorColor
                        HayCamposIncorrectos = True
                        CampoIncorrecto = Control.Name
                        If IsArray(Control) Then CampoIncorrecto = CampoIncorrecto & "(" & Control.Index & ")"
                    Else
                        If Control.BackColor <> -2147483624 Then Control.BackColor = vbWhite
                    End If
    
                Else
                    MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                    Exit Function
                End If
            End If
        'COMBOBOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Visible = True Then
                'Comprueba que los campos estan bien puestos
                If Control.Tag <> "" Then
                    If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                        Carga = mTag.Cargar(Control)
                        If Carga = False Then
                            MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                            Exit Function
        
                        Else
                            If mTag.Vacio = "N" And Control.ListIndex < 0 Then
'                                    MsgBox "Seleccione una dato para: " & mTag.Nombre, vbExclamation
'                                    Exit Function
                                Control.BackColor = vbErrorColor
                                HayCamposIncorrectos = True
                                CampoIncorrecto = Control.Name
                                If IsArray(Control) Then CampoIncorrecto = CampoIncorrecto & "(" & Control.Index & ")"
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next Control
    
    If HayCamposIncorrectos Then
        MsgBox "Revise datos obligatorios o incorrectos", vbExclamation
    End If
    CompForm2 = Not HayCamposIncorrectos
'    CompForm2 = True
End Function

Public Sub Limpiar(ByRef formulario As Form)
    Dim Control As Object
    
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Control.Text = ""
        End If
    Next Control
End Sub


Public Function CampoSiguiente(ByRef formulario As Form, Valor As Integer) As Control
Dim Fin As Boolean
Dim Control As Object

On Error GoTo ECampoSiguiente

    'Debug.Print "Llamada:  " & Valor
    'Vemos cual es el siguiente
    Do
        Valor = Valor + 1
        For Each Control In formulario.Controls
            'Debug.Print "-> " & Control.Name & " - " & Control.TabIndex
            'Si es texto monta esta parte de sql
            If Control.TabIndex = Valor Then
                    Set CampoSiguiente = Control
                    Fin = True
                    Exit For
            End If
        Next Control
        If Not Fin Then
            Valor = -1
        End If
    Loop Until Fin
    Exit Function
ECampoSiguiente:
    Set CampoSiguiente = Nothing
    Err.Clear
End Function


Private Function ValorParaSQL(Valor, ByRef vTag As CTag) As String
Dim Dev As String
Dim d As Single
Dim I As Integer
Dim V
    Dev = ""
    If Valor <> "" Then
        Select Case vTag.TipoDato
        Case "N"
            V = Valor
            If InStr(1, Valor, ",") Then
                If InStr(1, Valor, ".") Then
                    'ABRIL 2004
                
                    'Ademas de la coma lleva puntos
                    V = ImporteFormateado(CStr(Valor))
                    Valor = V
                Else
                
                    V = CSng(Valor)
                    Valor = V
                End If
            Else
         
            End If
            Dev = TransformaComasPuntos(CStr(Valor))
            
        Case "F"
            Dev = "'" & Format(Valor, FormatoFecha) & "'"
        Case "T"
            Dev = CStr(Valor)
            NombreSQL Dev
            Dev = "'" & Dev & "'"
            
        Case "FH"
            Dev = "'" & Format(Valor, "yyyy-mm-dd hh:mm:ss") & "'"
        
        Case Else
            Dev = "'" & Valor & "'"
        End Select
        
    Else
        'Si se permiten nulos, la "" ponemos un NULL
        If vTag.Vacio = "S" Then Dev = ValorNulo
    End If
    ValorParaSQL = Dev
End Function

Public Function InsertarDesdeForm(ByRef formulario As Form) As Boolean
    Dim Control As Object
    Dim mTag As CTag
    Dim Izda As String
    Dim Der As String
    Dim Cad As String
    
    On Error GoTo EInsertarF
    'Exit Function
    Set mTag = New CTag
    InsertarDesdeForm = False
    Der = ""
    Izda = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If mTag.Columna <> "" Then
                        If Izda <> "" Then Izda = Izda & ","
                        'Access
                        'Izda = Izda & "[" & mTag.Columna & "]"
                        Izda = Izda & "" & mTag.Columna & ""
                    
                        'Parte VALUES
                        Cad = ValorParaSQL(Control.Text, mTag)
                        If Der <> "" Then Der = Der & ","
                        Der = Der & Cad
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Izda <> "" Then Izda = Izda & ","
                'Access
                'Izda = Izda & "[" & mTag.Columna & "]"
                Izda = Izda & "" & mTag.Columna & ""
                If Control.Value = 1 Then
                    Cad = "1"
                    Else
                    Cad = "0"
                End If
                If Der <> "" Then Der = Der & ","
                If mTag.TipoDato = "N" Then Cad = Abs(CBool(Cad))
                Der = Der & Cad
            End If
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Izda <> "" Then Izda = Izda & ","
                    'Izda = Izda & "[" & mTag.Columna & "]"
                    Izda = Izda & "" & mTag.Columna & ""
                    If Control.ListIndex = -1 Then
                        Cad = ValorNulo
                        Else
                        Cad = Control.ItemData(Control.ListIndex)
                    End If
                    If Der <> "" Then Der = Der & ","
                    Der = Der & Cad
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo
    'INSERT INTO Empleados (Nombre,Apellido, Cargo) VALUES ('Carlos', 'Sesma', 'Prácticas');
    
    Cad = "INSERT INTO " & mTag.tabla
    Cad = Cad & " (" & Izda & ") VALUES (" & Der & ");"
    
   
    Conn.Execute Cad, , adCmdText
    
    
    InsertarDesdeForm = True
Exit Function
EInsertarF:
    MuestraError Err.Number, "Inserta. "
    
End Function

Public Function InsertarDesdeForm2(ByRef formulario As Form, Optional opcio As Integer, Optional nom_frame As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Izda As String
Dim Der As String
Dim Cad As String
    
    On Error GoTo EInsertarF
    
    'Exit Function
    Set mTag = New CTag
    InsertarDesdeForm2 = False
    Der = ""
    Izda = ""
    
    For Each Control In formulario.Controls
    
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If mTag.Columna <> "" Then
                            If Izda <> "" Then Izda = Izda & ","
                            'Access
                            'Izda = Izda & "[" & mTag.Columna & "]"
                            Izda = Izda & "" & mTag.Columna & ""
                        
                            'Parte VALUES
                            Cad = ValorParaSQL(Control.Text, mTag)
                            If Der <> "" Then Der = Der & ","
                            Der = Der & Cad
                        End If
                    End If
                End If
            End If
            
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If Izda <> "" Then Izda = Izda & ","
                    'Access
                    'Izda = Izda & "[" & mTag.Columna & "]"
                    Izda = Izda & "" & mTag.Columna & ""
                    If Control.Value = 1 Then
                        Cad = "1"
                        Else
                        Cad = "0"
                    End If
                    If Der <> "" Then Der = Der & ","
                    If mTag.TipoDato = "N" Then Cad = Abs(CBool(Cad))
                    Der = Der & Cad
                End If
            End If
            
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Izda <> "" Then Izda = Izda & ","
                        'Izda = Izda & "[" & mTag.Columna & "]"
                        Izda = Izda & "" & mTag.Columna & ""
                        If Control.ListIndex = -1 Then
                            Cad = ValorNulo
                        ElseIf mTag.TipoDato = "N" Then
                            Cad = Control.ItemData(Control.ListIndex)
                        Else
                            Cad = ValorParaSQL(Control.List(Control.ListIndex), mTag)
                        End If
                        If Der <> "" Then Der = Der & ","
                        Der = Der & Cad
                    End If
                End If
            End If
            
        'OPTION BUTTON
        ElseIf TypeOf Control Is OptionButton Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Control.Value Then
                            If Izda <> "" Then Izda = Izda & ","
                            Izda = Izda & "" & mTag.Columna & ""
                            Cad = Control.Index
                            If Der <> "" Then Der = Der & ","
                            Der = Der & Cad
                        End If
                    End If
                End If
            End If
            
        End If
        
    Next Control
    'Construimos el SQL
    'Ejemplo
    'INSERT INTO Empleados (Nombre,Apellido, Cargo) VALUES ('Carlos', 'Sesma', 'Prácticas');
    
    Cad = "INSERT INTO " & mTag.tabla
    Cad = Cad & " (" & Izda & ") VALUES (" & Der & ");"
    Conn.Execute Cad, , adCmdText
    
     ' ### [Monica] 18/12/2006
    CadenaCambio = Cad
   
    InsertarDesdeForm2 = True
Exit Function

EInsertarF:
    MuestraError Err.Number, "Inserta. "
End Function




Public Function PonerCamposForma(ByRef formulario As Form, ByRef vData As Adodc) As Boolean
    Dim Control As Object
    Dim mTag As CTag
    Dim Cad As String
    Dim Valor As Variant
    Dim Campo As String  'Campo en la base de datos
    Dim I As Integer

    On Error GoTo EPonerCamposForma


    Set mTag = New CTag
    PonerCamposForma = False

    For Each Control In formulario.Controls
        'TEXTO
        'Debug.Print Control.Tag
        If TypeOf Control Is TextBox Then
            'Comprobamos que tenga tag
            mTag.Cargar Control
            If Control.Tag <> "" Then
                If mTag.Cargado Then
                    'Columna en la BD
                    If mTag.Columna <> "" Then
                        Campo = mTag.Columna
                        If mTag.Vacio = "S" Then
                            Valor = DBLet(vData.Recordset.Fields(Campo))
                        Else
                            Valor = vData.Recordset.Fields(Campo)
                        End If
                        If mTag.Formato <> "" And CStr(Valor) <> "" Then
                            If mTag.TipoDato = "N" Then
                                'Es numerico, entonces formatearemos y sustituiremos
                                ' La coma por el punto
                                Cad = Format(Valor, mTag.Formato)
                                'Antiguo
                                'Control.Text = TransformaComasPuntos(cad)
                                'nuevo
                                Control.Text = Cad
                            Else
                                Control.Text = Format(Valor, mTag.Formato)
                            End If
                        Else
                            Control.Text = Valor
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Columna en la BD
                    Campo = mTag.Columna
                    If mTag.Vacio = "S" Then
                        Valor = DBLet(vData.Recordset.Fields(Campo), mTag.TipoDato)
                    Else
                        Valor = vData.Recordset.Fields(Campo)
                    End If
                    Else
                        Valor = 0
                End If
                Control.Value = Valor
            End If
            
         'COMBOBOX
         ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    Campo = mTag.Columna
                    Valor = vData.Recordset.Fields(Campo)
                    I = 0
                    For I = 0 To Control.ListCount - 1
                        If Control.ItemData(I) = Val(Valor) Then
                            Control.ListIndex = I
                            Exit For
                        End If
                    Next I
                    If I = Control.ListCount Then Control.ListIndex = -1
                End If 'de cargado
            End If 'de <>""
        End If
    Next Control
    
    'Veremos que tal
    PonerCamposForma = True
Exit Function
EPonerCamposForma:
    MuestraError Err.Number, "Poner campos formulario. "
End Function

Public Function PonerCamposForma2(ByRef formulario As Form, ByRef vData As Adodc, Optional opcio As Integer, Optional nom_frame As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Cad As String
Dim Valor As Variant
Dim Campo As String  'Campo en la base de datos
Dim I As Integer
    On Error GoTo EPonerCamposForma2
    
    Set mTag = New CTag
    PonerCamposForma2 = False
    For Each Control In formulario.Controls
        'TEXTO
        If (TypeOf Control Is TextBox) Then
            'Comprobamos que tenga tag
            mTag.Cargar Control
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    If mTag.Cargado Then
                        'Columna en la BD
                        If mTag.Columna <> "" Then
                            Campo = mTag.Columna
                            If mTag.Vacio = "S" Then
                                Valor = DBLet(vData.Recordset.Fields(Campo))
                            Else
                                Valor = vData.Recordset.Fields(Campo)
                            End If
                            If mTag.Formato <> "" And CStr(Valor) <> "" Then
                                If mTag.TipoDato = "N" Then
                                    'Es numerico, entonces formatearemos y sustituiremos
                                    ' La coma por el punto
                                    Cad = Format(Valor, mTag.Formato)
                                    'Antiguo
                                    'Control.Text = TransformaComasPuntos(cad)
                                    'nuevo
                                    Control.Text = Cad
                                Else
                                    Control.Text = Format(Valor, mTag.Formato)
                                End If
                            Else
                                Control.Text = Valor
                            End If
                        End If
                    End If
                End If
            End If
            
        'CheckBOX
        ElseIf (TypeOf Control Is CheckBox) Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        'Columna en la BD
                        Campo = mTag.Columna
                        Valor = vData.Recordset.Fields(Campo)
                    Else
                        Valor = 0
                    End If
                    If IsNull(Valor) Then Valor = 0
                    If Val(Valor) > 1 Then Valor = 1
                    Control.Value = Valor
                End If
            End If

         'COMBOBOX
         ElseIf (TypeOf Control Is ComboBox) Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        Campo = mTag.Columna
                        Valor = DBLet(vData.Recordset.Fields(Campo))
                        I = 0
                        For I = 0 To Control.ListCount - 1
                            If Control.ItemData(I) = Val(Valor) Then
                                Control.ListIndex = I
                                Exit For
                            End If
                        Next I
                        If I = Control.ListCount Then Control.ListIndex = -1
                    End If 'de cargado
                End If
            End If 'de <>""
            
        ElseIf TypeOf Control Is OptionButton Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        'Columna en la BD
                        Campo = mTag.Columna
                        Valor = vData.Recordset.Fields(Campo)
                        If IsNull(Valor) Then Valor = 0
                        If Control.Index = Valor Then
                            Control.Value = True
                        Else
                            Control.Value = False
                        End If
                    End If
                End If
            End If
            
        End If
    Next Control

    'Veremos que tal
    PonerCamposForma2 = True
Exit Function
EPonerCamposForma2:
    MuestraError Err.Number, "Poner campos formulario 2. "
End Function





Private Function ObtenerMaximoMinimo(ByRef vSQL As String) As String
Dim Rs As Recordset
ObtenerMaximoMinimo = ""
Set Rs = New ADODB.Recordset
Rs.Open vSQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
If Not Rs.EOF Then
    If Not IsNull(Rs.EOF) Then
        ObtenerMaximoMinimo = CStr(Rs.Fields(0))
    End If
End If
Rs.Close
Set Rs = Nothing
End Function


Public Function ObtenerBusqueda(ByRef formulario As Form) As String
    Dim Control As Object
    Dim Carga As Boolean
    Dim mTag As CTag
    Dim Aux As String
    Dim Cad As String
    Dim SQL As String
    Dim tabla As String
    Dim RC As Byte

    On Error GoTo EObtenerBusqueda

    'Exit Function
    Set mTag = New CTag
    ObtenerBusqueda = ""
    SQL = ""

    'Recorremos los text en busca de ">>" o "<<"
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = Trim(Control.Text)
            If Aux = ">>" Or Aux = "<<" Then
                Carga = mTag.Cargar(Control)
                If Carga Then
                    If Aux = ">>" Then
                        Cad = " MAX(" & mTag.Columna & ")"
                    Else
                        Cad = " MIN(" & mTag.Columna & ")"
                    End If
                    SQL = "Select " & Cad & " from " & mTag.tabla
                    SQL = ObtenerMaximoMinimo(SQL)
                    Select Case mTag.TipoDato
                    Case "N"
                        SQL = mTag.tabla & "." & mTag.Columna & " = " & TransformaComasPuntos(SQL)
                    Case "F"
                        SQL = mTag.tabla & "." & mTag.Columna & " = '" & Format(SQL, "yyyy-mm-dd") & "'"
                    Case Else
                        SQL = mTag.tabla & "." & mTag.Columna & " = '" & SQL & "'"
                    End Select
                    SQL = "(" & SQL & ")"
                End If
            End If
        End If
    Next

    
    
    'Recorremos los text en busca del NULL
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = Trim(Control.Text)
            'If Control.Text <> "" Then Stop
            If UCase(Aux) = "NULL" Then
                Carga = mTag.Cargar(Control)
                If Carga Then

                    SQL = mTag.tabla & "." & mTag.Columna & " is NULL"
                    SQL = "(" & SQL & ")"
                    Control.Text = ""
                End If
            End If
        End If
    Next

    

    'Recorremos los textbox
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            
            'Cargamos el tag
            Carga = mTag.Cargar(Control)
            If Carga Then
                Aux = Trim(Control.Text)
                'If Control.Text <> "" Then Stop
                If Aux <> "" Then
                    If mTag.tabla <> "" Then
                        tabla = mTag.tabla & "."
                        Else
                        tabla = ""
                    End If
                    RC = SeparaCampoBusqueda(mTag.TipoDato, tabla & mTag.Columna, Aux, Cad)
                    If RC = 0 Then
                        If SQL <> "" Then SQL = SQL & " AND "
                        SQL = SQL & "(" & Cad & ")"
                    End If
                End If
            Else
                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                Exit Function
            End If
        
        
        
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            mTag.Cargar Control
            If mTag.Cargado Then
                If Control.ListIndex > -1 Then
                    Cad = Control.ItemData(Control.ListIndex)
                    Cad = mTag.tabla & "." & mTag.Columna & " = " & Cad
                    If SQL <> "" Then SQL = SQL & " AND "
                    SQL = SQL & "(" & Cad & ")"
                End If
            End If
        
        
        'CHECK
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.Value = 1 Then
                        Cad = mTag.tabla & "." & mTag.Columna & " = 1"
                        If SQL <> "" Then SQL = SQL & " AND "
                        SQL = SQL & "(" & Cad & ")"
                    End If
                End If
            End If
        End If

        
    Next Control
    ObtenerBusqueda = SQL
Exit Function
EObtenerBusqueda:
    ObtenerBusqueda = ""
    MuestraError Err.Number, "Obtener búsqueda. "
End Function


Public Function ObtenerBusqueda2(ByRef formulario As Form, Optional CHECK As String, Optional opcio As Integer, Optional nom_frame As String) As String
    Dim Control As Object
    Dim Carga As Boolean
    Dim mTag As CTag
    Dim Aux As String
    Dim Cad As String
    Dim SQL As String
    Dim tabla As String
    Dim RC As Byte

    On Error GoTo EObtenerBusqueda

    'Exit Function
    Set mTag = New CTag
    ObtenerBusqueda2 = ""
    SQL = ""

    'Recorremos los text en busca de ">>" o "<<"
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = Trim(Control.Text)
            If Aux = ">>" Or Aux = "<<" Then
                Carga = mTag.Cargar(Control)
                If Carga Then
                    If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                        If Aux = ">>" Then
                            Cad = " MAX(" & mTag.Columna & ")"
                        Else
                            Cad = " MIN(" & mTag.Columna & ")"
                        End If
                        SQL = "Select " & Cad & " from " & mTag.tabla
                        SQL = ObtenerMaximoMinimo(SQL)
                        Select Case mTag.TipoDato
                        Case "N"
                            SQL = mTag.tabla & "." & mTag.Columna & " = " & TransformaComasPuntos(SQL)
                        Case "F"
                            SQL = mTag.tabla & "." & mTag.Columna & " = '" & Format(SQL, "yyyy-mm-dd") & "'"
                        Case Else
                            '[Monica]04/03/2013: quito las comillas y pongo el dbset
                            SQL = mTag.tabla & "." & mTag.Columna & " = " & DBSet(SQL, "T") ' & "'"
                        End Select
                        SQL = "(" & SQL & ")"
                    End If
                End If
            End If
        End If
    Next

'++monica: lo he añadido del anterior obtenerbusqueda
    'Recorremos los text en busca del NULL
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = Trim(Control.Text)
            If UCase(Aux) = "NULL" Then
                Carga = mTag.Cargar(Control)
                If Carga And mTag.Columna <> "" Then

                    SQL = mTag.tabla & "." & mTag.Columna & " is NULL"
                    SQL = "(" & SQL & ")"
                    Control.Text = ""
                End If
            End If
        End If
    Next

    'Recorremos los textbox
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
          If Control.Tag <> "" Then
            'Cargamos el tag
            Carga = mTag.Cargar(Control)
            If Carga Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    Aux = Trim(Control.Text)
                    If Aux <> "" Then
                        If mTag.tabla <> "" Then
                            tabla = mTag.tabla & "."
                            Else
                            tabla = ""
                        End If
                        RC = SeparaCampoBusqueda(mTag.TipoDato, tabla & mTag.Columna, Aux, Cad)
                        If RC = 0 Then
                            If SQL <> "" Then SQL = SQL & " AND "
                            SQL = SQL & "(" & Cad & ")"
                        End If
                    End If
                End If
            Else
                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                Exit Function
            End If
        End If
        
        
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then ' +-+- 12/05/05: canvi de Cèsar, no te sentit passar-li un control que no té TAG +-+-
                mTag.Cargar Control
                If mTag.Cargado Then
                    If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                        If Control.ListIndex > -1 Then
                            Cad = Control.ItemData(Control.ListIndex)
                            Cad = mTag.tabla & "." & mTag.Columna & " = " & Cad
                            If SQL <> "" Then SQL = SQL & " AND "
                            SQL = SQL & "(" & Cad & ")"
                        End If
                    End If
                End If
            End If
            
         ElseIf TypeOf Control Is CheckBox Then
            '=============== Añade: Laura, 27/04/05
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    ' añadido 12022007
                    Aux = ""
                    If CHECK <> "" Then
                        tabla = DBLet(Control.Index, "T")
                        If tabla <> "" Then tabla = "(" & tabla & ")"
                        tabla = Control.Name & tabla & "|"
                        If InStr(1, CHECK, tabla, vbTextCompare) > 0 Then Aux = Control.Value
                    Else
                        If Control.Value = 1 Then Aux = "1"
                    End If
                    If Aux <> "" Then
'                    If Control.Value = 1 Then
                        Cad = Control.Value
                        Cad = mTag.tabla & "." & mTag.Columna & " = " & Cad
                        If SQL <> "" Then SQL = SQL & " AND "
                        SQL = SQL & "(" & Cad & ")"
                    End If
                End If
            End If
            '===================
        End If
    Next Control
    ObtenerBusqueda2 = SQL
Exit Function
EObtenerBusqueda:
    ObtenerBusqueda2 = ""
    MuestraError Err.Number, "Obtener búsqueda. " & vbCrLf & Err.Description
End Function



Public Function ModificaDesdeFormulario(ByRef formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim CadWhere As String
Dim cadUPDATE As String

On Error GoTo EModificaDesdeFormulario
    ModificaDesdeFormulario = False
    Set mTag = New CTag
    Aux = ""
    CadWhere = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
              
                mTag.Cargar Control
                If mTag.Cargado Then
                    If mTag.Columna <> "" Then
                        'Sea para el where o para el update esto lo necesito
                        Aux = ValorParaSQL(Control.Text, mTag)
                        'Si es campo clave NO se puede modificar y se utiliza como busqueda
                        'dentro del WHERE
                        If mTag.EsClave Then
                            'Lo pondremos para el WHERE
                             If CadWhere <> "" Then CadWhere = CadWhere & " AND "
                             CadWhere = CadWhere & "(" & mTag.Columna & " = " & Aux & ")"
                             
                        Else
                            If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                            cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Control.Value = 1 Then
                    Aux = "TRUE"
                    Else
                    Aux = "FALSE"
                End If
                If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                'Esta es para access
                'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
            End If
            
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.ListIndex = -1 Then
                        Aux = ValorNulo
                        Else
                        Aux = Control.ItemData(Control.ListIndex)
                    End If
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                    cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo:
    'Update Pedidos
    'SET ImportePedido = ImportePedido * 1.1,
    'Cargo = Cargo * 1.03
    'WHERE PaísDestinatario = 'México';
    If CadWhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & CadWhere
    Conn.Execute Aux, , adCmdText



ModificaDesdeFormulario = True
Exit Function
EModificaDesdeFormulario:
    MuestraError Err.Number, "Modificar. " & Err.Description
End Function


Public Function ModificaDesdeFormulario2(ByRef formulario As Form, Optional opcio As Integer, Optional nom_frame As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim CadWhere As String
Dim cadUPDATE As String

On Error GoTo EModificaDesdeFormulario
    ModificaDesdeFormulario2 = False
    Set mTag = New CTag
    Aux = ""
    CadWhere = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If mTag.Columna <> "" Then
                            'Sea para el where o para el update esto lo necesito
                            Aux = ValorParaSQL(Control.Text, mTag)
                            'Si es campo clave NO se puede modificar y se utiliza como busqueda
                            'dentro del WHERE
                            If mTag.EsClave Then
                                'Lo pondremos para el WHERE
                                 If CadWhere <> "" Then CadWhere = CadWhere & " AND "
                                 CadWhere = CadWhere & "(" & mTag.Columna & " = " & Aux & ")"
    
                            Else
                                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                                cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
                            End If
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    If Control.Visible Then
                        mTag.Cargar Control
                        If Control.Value = 1 Then
                            Aux = "TRUE"
                        Else
                            Aux = "FALSE"
                        End If
                        If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                        If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                        'Esta es para access
                        'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                        cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
                    End If
                End If
            End If

        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Control.ListIndex = -1 Then
                            Aux = ValorNulo
                        ElseIf mTag.TipoDato = "N" Then
                            Aux = Control.ItemData(Control.ListIndex)
                        Else
                            Aux = ValorParaSQL(Control.List(Control.ListIndex), mTag)
                        End If
                        
                        'Si es campo clave NO se puede modificar y se utiliza como busqueda
                        'dentro del WHERE
                        If mTag.EsClave Then
                            'Lo pondremos para el WHERE
                             If CadWhere <> "" Then CadWhere = CadWhere & " AND "
                             CadWhere = CadWhere & "(" & mTag.Columna & " = " & Aux & ")"
                        Else
                            If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                            cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
                        End If
'
'
'                        If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
'                        'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
'                        cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                    End If
                End If
            End If
            
        ElseIf TypeOf Control Is OptionButton Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Control.Value Then
                            Aux = Control.Index
                            'Si es campo clave NO se puede modificar y se utiliza como busqueda
                            'dentro del WHERE
                              If mTag.EsClave Then
                                  'Lo pondremos para el WHERE
                                   If CadWhere <> "" Then CadWhere = CadWhere & " AND "
                                   CadWhere = CadWhere & "(" & mTag.Columna & " = " & Aux & ")"
                              Else
                                  If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                                  cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
                              End If
                        End If
                    End If
                End If
            End If
            
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo:
    'Update Pedidos
    'SET ImportePedido = ImportePedido * 1.1,
    'Cargo = Cargo * 1.03
    'WHERE PaísDestinatario = 'México';
    If CadWhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & CadWhere
    Conn.Execute Aux, , adCmdText

    ' ### [Monica] 18/12/2006
    CadenaCambio = cadUPDATE

    ModificaDesdeFormulario2 = True
    Exit Function
    
EModificaDesdeFormulario:
    MuestraError Err.Number, "Modificar 2. " & Err.Description
End Function





Public Function ParaGrid(ByRef Control As Control, AnchoPorcentaje As Integer, Optional Desc As String) As String
Dim mTag As CTag
Dim Cad As String

'Montamos al final: "Cod Diag.|idDiag|N|10·"

ParaGrid = ""
Cad = ""
Set mTag = New CTag
mTag.Cargar Control
If mTag.Cargado Then
    If Control.Tag <> "" Then
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Desc <> "" Then
                Cad = Desc
            Else
                Cad = mTag.Nombre
            End If
            Cad = Cad & "|"
            Cad = Cad & mTag.Columna & "|"
            Cad = Cad & mTag.TipoDato & "|"
            Cad = Cad & AnchoPorcentaje & "·"
            
                
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            
        ElseIf TypeOf Control Is ComboBox Then
        
        
        End If 'De los elseif
    End If
Set mTag = Nothing
ParaGrid = Cad
End If



End Function

'////////////////////////////////////////////////////
' Monta a partir de una cadena devuelta por el formulario
'de busqueda el sql para situar despues el datasource
Public Function ValorDevueltoFormGrid(ByRef Control As Control, ByRef CadenaDevuelta As String, Orden As Integer) As String
Dim mTag As CTag
Dim Cad As String
Dim Aux As String
'Montamos al final: " columnatabla = valordevuelto "

ValorDevueltoFormGrid = ""
Cad = ""
Set mTag = New CTag
mTag.Cargar Control
If mTag.Cargado Then
    If Control.Tag <> "" Then
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            Aux = RecuperaValor(CadenaDevuelta, Orden)
            If Aux <> "" Then Cad = mTag.Columna & " = " & ValorParaSQL(Aux, mTag)
                
            
            
                
        'CheckBOX
       ' ElseIf TypeOf Control Is CheckBox Then
       '
       ' ElseIf TypeOf Control Is ComboBox Then
       '
       '
        End If 'De los elseif
    End If
End If
Set mTag = Nothing
ValorDevueltoFormGrid = Cad
End Function


Public Sub FormateaCampo(vTex As TextBox)
    Dim mTag As CTag
    Dim Cad As String
    On Error GoTo EFormateaCampo
    Set mTag = New CTag
    mTag.Cargar vTex
    If mTag.Cargado Then
        If vTex.Text <> "" Then
            If mTag.Formato <> "" Then
                Cad = TransformaPuntosComas(vTex.Text)
                Cad = Format(Cad, mTag.Formato)
                vTex.Text = Cad
            End If
        End If
    End If
EFormateaCampo:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Sub


'recupera valor desde una cadena con pipes(acabada en pipes)
'Para ello le decimos el orden  y ya ta
Public Function RecuperaValor(ByRef Cadena As String, Orden As Integer) As String
Dim I As Integer
Dim J As Integer
Dim Cont As Integer
Dim Cad As String

I = 0
Cont = 1
Cad = ""
Do
    J = I + 1
    I = InStr(J, Cadena, "|")
    If I > 0 Then
        If Cont = Orden Then
            Cad = Mid(Cadena, J, I - J)
            I = Len(Cadena) 'Para salir del bucle
            Else
                Cont = Cont + 1
        End If
    End If
Loop Until I = 0
RecuperaValor = Cad
End Function

'recupera valor desde una cadena con pipes(acabada en pipes)
'Para ello le decimos el orden  y ya ta
Public Function RecuperaValorNew(ByRef Cadena As String, Separador As String, Orden As Integer) As String
Dim I As Integer
Dim J As Integer
Dim Cont As Integer
Dim Cad As String

    I = 0
    Cont = 1
    Cad = ""
    Do
        J = I + 1
        I = InStr(J, Cadena, Separador)
        If I > 0 Then
            If Cont = Orden Then
                Cad = Mid(Cadena, J, I - J)
                I = Len(Cadena) 'Para salir del bucle
                Else
                    Cont = Cont + 1
            End If
        End If
    Loop Until I = 0
    RecuperaValorNew = Cad
End Function





'recupera valor desde una cadena con pipes(acabada en pipes)
'Para ello le decimos el orden  y ya ta
Public Function InsertaValor(ByRef Cadena As String, Orden As Integer, SubCadena As String) As String
Dim I As Integer
Dim J As Integer
Dim Cont As Integer
Dim Cad As String
Dim Cad2 As String
I = 0
Cont = 1
Cad = ""
Do
    J = I + 1
    I = InStr(J, Cadena, "|")
    If I > 0 Then
        If Cont = Orden Then
            Cad = Mid(Cadena, J, I - J)
            
            Cad2 = Mid(Cadena, 1, J - 1) & SubCadena & Mid(Cadena, I, Len(Cadena))
            
            I = Len(Cadena) 'Para salir del bucle
            Else
                Cont = Cont + 1
        End If
    End If
Loop Until I = 0
InsertaValor = Cad2
End Function






'-----------------------------------------------------------------------
'Deshabilitar ciertas opciones del menu
'EN funcion del nivel de usuario
'Esto es a nivel general, cuando el Toolba es el mismo

'Para ello en el tag del button tendremos k poner un numero k nos diara hasta k nivel esta permitido

Public Sub PonerOpcionesMenuGeneral(ByRef formulario As Form)
Dim I As Integer
Dim J As Integer


On Error GoTo EPonerOpcionesMenuGeneral


'Añadir, modificar y borrar deshabilitados si no nivel
With formulario

    'LA TOOLBAR  .--> Requisito, k se llame toolbar1
    For I = 1 To .Toolbar1.Buttons.Count
        If .Toolbar1.Buttons(I).Tag <> "" Then
            J = Val(.Toolbar1.Buttons(I).Tag)
            If J < vUsu.Nivel Then
                .Toolbar1.Buttons(I).Enabled = False
            End If
        End If
    Next I
    
    'Esto es un poco salvaje. Por si acaso , no existe en este trozo pondremos los errores on resume next
    
    On Error Resume Next
    
    'Los MENUS
    'K sean mnAlgo
    J = Val(.mnNuevo.HelpContextID)
    If J < vUsu.Nivel Then .mnNuevo.Enabled = False
    
    J = Val(.mnModificar.HelpContextID)
    If J < vUsu.Nivel Then .mnModificar.Enabled = False
    
    J = Val(.mnEliminar.HelpContextID)
    If J < vUsu.Nivel Then .mnEliminar.Enabled = False
    On Error GoTo 0
End With




Exit Sub
EPonerOpcionesMenuGeneral:
    MuestraError Err.Number, "Poner opciones usuario generales"
End Sub



'Este modifica las claves prinipales y todo
'la sentenca del WHERE cod=1 and .. viene en claves
Public Function ModificaDesdeFormularioClaves(ByRef formulario As Form, Claves As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim CadWhere As String
Dim cadUPDATE As String
Dim I As Integer

On Error GoTo EModificaDesdeFormulario
    ModificaDesdeFormularioClaves = False
    Set mTag = New CTag
    Aux = ""
    CadWhere = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Sea para el where o para el update esto lo necesito
                    Aux = ValorParaSQL(Control.Text, mTag)
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Control.Value = 1 Then
                    Aux = "TRUE"
                    Else
                    Aux = "FALSE"
                End If
                If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                'Esta es para access
                'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
            End If
            
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.ListIndex = -1 Then
                        Aux = ValorNulo
                        Else
                        Aux = Control.ItemData(Control.ListIndex)
                    End If
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                    cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
                End If
            End If
        End If
    Next Control
    CadWhere = Claves
    'Construimos el SQL
    If CadWhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & CadWhere
    Conn.Execute Aux, , adCmdText






ModificaDesdeFormularioClaves = True
Exit Function
EModificaDesdeFormulario:
    MuestraError Err.Number, "Modificar. " & Err.Description
End Function







Public Function BLOQUEADesdeFormulario(ByRef formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim CadWhere As String
Dim AntiguoCursor As Byte

On Error GoTo EBLOQUEADesdeFormulario
    BLOQUEADesdeFormulario = False
    Set mTag = New CTag
    Aux = ""
    CadWhere = ""
    AntiguoCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
              
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Sea para el where o para el update esto lo necesito
                    Aux = ValorParaSQL(Control.Text, mTag)
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    'dentro del WHERE
                    If mTag.EsClave Then
                        'Lo pondremos para el WHERE
                         If CadWhere <> "" Then CadWhere = CadWhere & " AND "
                         CadWhere = CadWhere & "(" & mTag.Columna & " = " & Aux & ")"
                    End If
                End If
            End If
        End If
    Next Control
    
    If CadWhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        
    Else
        Aux = "select * FROM " & mTag.tabla
        Aux = Aux & " WHERE " & CadWhere & " FOR UPDATE"
        
        'Intenteamos bloquear
        PreparaBloquear
        Conn.Execute Aux, , adCmdText
        BLOQUEADesdeFormulario = True
    End If
EBLOQUEADesdeFormulario:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Bloqueo tabla"
        TerminaBloquear
    End If
    Screen.MousePointer = AntiguoCursor
End Function


Public Function BLOQUEADesdeFormulario2(ByRef formulario As Form, ByRef ado As Adodc, Optional opcio As Integer, Optional nom_frame As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim CadWhere As String
Dim AntiguoCursor As Byte
Dim nomcamp As String

    On Error GoTo EBLOQUEADesdeFormulario2
    
    BLOQUEADesdeFormulario2 = False
    Set mTag = New CTag
    Aux = ""
    CadWhere = ""
    AntiguoCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If (TypeOf Control Is TextBox) Or (TypeOf Control Is ComboBox) Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        'Sea para el where o para el update esto lo necesito
                        'Aux = ValorParaSQL(Control.Text, mTag)
                        'Si es campo clave NO se puede modificar y se utiliza como busqueda
                        'dentro del WHERE
                        If mTag.EsClave Then
                            Aux = ValorParaSQL(CStr(ado.Recordset.Fields(mTag.Columna)), mTag)
                            'Lo pondremos para el WHERE
                             If CadWhere <> "" Then CadWhere = CadWhere & " AND "
                             CadWhere = CadWhere & "(" & mTag.Columna & " = " & Aux & ")"
                        End If
                    End If
                End If
            End If
        End If
    Next Control

    If CadWhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
    Else
        Aux = "select * FROM " & mTag.tabla
        Aux = Aux & " WHERE " & CadWhere & " FOR UPDATE"

        'Intenteamos bloquear
        PreparaBloquear
        Conn.Execute Aux, , adCmdText
        BLOQUEADesdeFormulario2 = True
    End If
    
EBLOQUEADesdeFormulario2:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Bloqueo tabla 2"
'        BLOQUEADesdeFormulario2 = False
        TerminaBloquear
    End If
    Screen.MousePointer = AntiguoCursor
End Function


Public Function BloqueaRegistroForm(ByRef formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim AuxDef As String
Dim AntiguoCursor As Byte

On Error GoTo EBLOQ
    BloqueaRegistroForm = False
    Set mTag = New CTag
    Aux = ""
    AuxDef = ""
    AntiguoCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
              
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    'dentro del WHERE
                    If mTag.EsClave Then
                        Aux = ValorParaSQL(Control.Text, mTag)
                        AuxDef = AuxDef & Aux & "|"
                    End If
                End If
            End If
        End If
    Next Control
    
    If AuxDef = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        
    Else
        Aux = "Insert into zBloqueos(codusu,tabla,clave) VALUES(" & vUsu.Codigo & ",'" & mTag.tabla
        Aux = Aux & "',""" & AuxDef & """)"
        Conn.Execute Aux
        BloqueaRegistroForm = True
    End If
EBLOQ:
    If Err.Number <> 0 Then
        Aux = ""
        If Conn.Errors.Count > 0 Then
            If Conn.Errors(0).NativeError = 1062 Then
                '¡Ya existe el registro, luego esta bloqueada
                Aux = "BLOQUEO"
            End If
        End If
        If Aux = "" Then
            MuestraError Err.Number, "Bloqueo tabla"
        Else
            MsgBox "Registro bloqueado por otro usuario", vbExclamation
        End If
    End If
    Screen.MousePointer = AntiguoCursor
End Function


Public Function DesBloqueaRegistroForm(ByRef TextBoxConTag As TextBox) As Boolean
Dim mTag As CTag
Dim SQL As String

'Solo me interesa la tabla
On Error Resume Next
    Set mTag = New CTag
    mTag.Cargar TextBoxConTag
    If mTag.Cargado Then
        SQL = "DELETE from zBloqueos where codusu=" & vUsu.Codigo & " and tabla='" & mTag.tabla & "'"
        Conn.Execute SQL
        If Err.Number <> 0 Then
            Err.Clear
        End If
    End If
    Set mTag = Nothing
End Function



'----------------------------------------------------------
'----------------------------------------------------------
'
'       Funciones comunes en todos los formularios
'           del tipo KEY PRESS, gotfocus
'

Public Sub KEYpressGnral(KeyAscii As Integer, Modo As Byte, cerrar As Boolean)
'IN: codigo keyascii tecleado, y modo en que esta el formulario
'OUT: si se tiene que cerrar el formulario o no
    cerrar = False
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        CreateObject("WScript.Shell").SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then 'ESC
        If (Modo = 0 Or Modo = 2) Then cerrar = True
        End If
End Sub



Public Sub KEYdown(KeyCode As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
On Error Resume Next
    Select Case KeyCode
        Case 38 'Desplazamieto Fecha Hacia Arriba
            CreateObject("WScript.Shell").SendKeys "+{tab}"
        Case 40 'Desplazamiento Flecha Hacia Abajo
            CreateObject("WScript.Shell").SendKeys "{tab}"
    End Select
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub PonFoco(ByRef T As TextBox)
On Error Resume Next
'        T.SelStart = 0
'        T.SelLength = Len(T.Text)
    T.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub PonerFocoGrid(ByRef DGrid As DataGrid)
    On Error Resume Next
    DGrid.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub PonerFocoCmb(ByRef combo As ComboBox)
On Error Resume Next
    combo.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub PonerFocoChk(ByRef chk As CheckBox)
On Error Resume Next
    chk.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub PonerFocoLw(ByRef Lw As ListView)
On Error Resume Next
    Lw.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub PonerFocoBtn(ByRef Btn As CommandButton)
    On Error Resume Next
    Btn.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Sub DeseleccionaGrid(ByRef DataGrid1 As DataGrid)
    On Error GoTo EDeseleccionaGrid
        
    While DataGrid1.SelBookmarks.Count > 0
        DataGrid1.SelBookmarks.Remove 0
    Wend
    Exit Sub
EDeseleccionaGrid:
        Err.Clear

End Sub


Public Sub PonLblIndicador(ByRef L As Label, ByRef AdodcX As Adodc)
    On Error Resume Next
    L.Caption = AdodcX.Recordset.AbsolutePosition & " de " & AdodcX.Recordset.RecordCount
    If Err.Number <> 0 Then
        Err.Clear
        L.Caption = ""
    End If
End Sub


Public Sub PonleFoco(Ob As Object)
    On Error Resume Next
    Ob.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub



'Devuelve la variable parafijar la altura donde empiezan los txtaux
' y cuadren con el datagrid
Public Function FijarVariableAnc(ByRef DTGRD1 As DataGrid) As Single
Dim I As Integer

    If DTGRD1.Row < 0 Then
        I = 0
        Else
        I = DTGRD1.Row
    End If
    FijarVariableAnc = DTGRD1.RowTop(I) + DTGRD1.top + 15
    
End Function


Public Function BloqueaRegistro(cadTabla As String, CadWhere As String) As Boolean
Dim Aux As String

    On Error GoTo EBloqueaRegistro
        
    BloqueaRegistro = False
    Aux = "select * FROM " & cadTabla
    Aux = Aux & " WHERE " & CadWhere & " FOR UPDATE"

    'Intenteamos bloquear
    PreparaBloquear
    Conn.Execute Aux, , adCmdText
    BloqueaRegistro = True

EBloqueaRegistro:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Bloqueo tabla"
        TerminaBloquear
    End If
End Function

'++++++++++++++++++++++++++++++++++++
'++       FUNCIONES AÑADIDAS
'++++++++++++++++++++++++++++++++++++
Public Function SituarData(ByRef vData As Adodc, vWhere As String, ByRef Indicador As String, Optional NoRefresca As Boolean) As Boolean
'Situa un DataControl en el registo que cumple vwhere
    On Error GoTo ESituarData

        'Actualizamos el recordset
        If Not NoRefresca Then vData.Refresh
        
        'El sql para que se situe en el registro en especial es el siguiente
        vData.Recordset.Find vWhere
        If vData.Recordset.EOF Then
            If vData.Recordset.RecordCount > 0 Then vData.Recordset.MoveFirst
            GoTo ESituarData
        End If
        Indicador = vData.Recordset.AbsolutePosition & " de " & vData.Recordset.RecordCount
        SituarData = True
        Exit Function

ESituarData:
        If Err.Number <> 0 Then Err.Clear
        SituarData = False
End Function

Public Function SituarDataMULTI(ByRef vData As Adodc, vWhere As String, ByRef Indicador As String, Optional NoRefresca As Boolean) As Boolean
'Situa un DataControl en el registo que cumple vwhere
On Error GoTo ESituarData
        'Actualizamos el recordset
        If Not NoRefresca Then vData.Refresh
        'El sql para que se situe en el registro en especial es el siguiente
        Multi_Find vData.Recordset, vWhere
        'vData.Recordset.Find vWhere
        If vData.Recordset.EOF Then GoTo ESituarData
        Indicador = vData.Recordset.AbsolutePosition & " de " & vData.Recordset.RecordCount
        SituarDataMULTI = True
        Exit Function
ESituarData:
        If Err.Number <> 0 Then Err.Clear
        SituarDataMULTI = False
End Function


Public Sub Multi_Find(ByRef oRs As ADODB.Recordset, sCriteria As String)

    Dim clone_rs As ADODB.Recordset
    Set clone_rs = oRs.Clone
    
    clone_rs.Filter = sCriteria
    
    If clone_rs.EOF Or clone_rs.BOF Then
     oRs.MoveLast
     oRs.MoveNext
    Else
     oRs.Bookmark = clone_rs.Bookmark
    End If
    
    clone_rs.Close
    Set clone_rs = Nothing

End Sub


Public Sub PonerIndicador(ByRef lblIndicador As Label, Modo As Byte, Optional ModoLineas As Byte)
'Pone el titulo del label lblIndicador
    Select Case Modo
        Case 0    'Modo Inicial
            lblIndicador.Caption = ""
        Case 1 'Modo Buscar
            lblIndicador.Caption = "BUSQUEDA"
        Case 2    'Preparamos para que pueda Modificar
            lblIndicador.Caption = ""

        Case 3 'Modo Insertar
            lblIndicador.Caption = "INSERTAR"
        Case 4 'MODIFICAR
            lblIndicador.Caption = "MODIFICAR"
            
        Case 5 'Modo Lineas
            If ModoLineas = 1 Then
                lblIndicador.Caption = "INSERTAR LINEA"
            ElseIf ModoLineas = 2 Then
                lblIndicador.Caption = "MODIFICAR LINEA"
            End If
        Case Else
            lblIndicador.Caption = ""
    End Select
End Sub


Public Function DBSet(vData As Variant, Tipo As String, Optional EsNulo As String) As Variant
'Establece el valor del dato correcto antes de Insertar en la BD
'Tipos
'       T
'       N
'       F
'       H
'       FH
'       B
'       S   single O DOUBLE. sINGLE DE MOMENTO.    MAYO 2009
Dim Cad As String
Dim ValorNumericoCero As Boolean

    On Error GoTo Error1

        If IsNull(vData) Then
            DBSet = ValorNulo
            Exit Function
        End If
    
        If Tipo <> "" Then
            Select Case Tipo
                Case "T"    'Texto
                    If vData = "" Then
                        If EsNulo = "N" Then
                            DBSet = "''"
                        Else
                            DBSet = ValorNulo
                        End If
                    Else
                        Cad = (CStr(vData))
                        NombreSQL Cad
                        DBSet = "'" & Cad & "'"
                    End If
                    
                Case "N", "S"   'Numero  y  SINGLE
                    
                    If CStr(vData) = "" Then
                        ValorNumericoCero = True
                    
                    Else
                        If Tipo = "S" Then
                            ValorNumericoCero = CSng(vData) = 0
                        Else
                            ValorNumericoCero = CCur(vData) = 0
                        End If
                    End If
                    
                    If ValorNumericoCero Then
                        If EsNulo <> "" Then
                            If EsNulo = "S" Then
                                DBSet = ValorNulo
                            Else
                                DBSet = 0
                            End If
                        Else
                            DBSet = 0
                        End If
                    Else
                        If Tipo = "N" Then
                            Cad = CStr(ImporteFormateado(CStr(vData)))
                        Else
                            'Sngle
                            Cad = CStr(ImporteFormateadoSingle(CStr(vData)))
                        End If
                        DBSet = TransformaComasPuntos(Cad)
                    End If
                    
                Case "F"    'Fecha
'                     '==David
''                    DBLet = "0:00:00"
'                     '==Laura
                    If vData = "" Then
                        If EsNulo = "S" Then
                            DBSet = ValorNulo
                        Else
                            DBSet = "'1900-01-01'"
                        End If
                    Else
                        DBSet = "'" & Format(vData, FormatoFecha) & "'"
                    End If

                Case "FH" 'Fecha/Hora
                    If vData = "" Then
                        If EsNulo = "S" Then DBSet = ValorNulo
                    Else
                        DBSet = "'" & Format(vData, "yyyy-mm-dd hh:mm:ss") & "'"
                    End If

                Case "H" 'Hora
                    If vData = "" Then
                    Else
                        DBSet = "'" & Format(vData, "hh:mm:ss") & "'"
                    End If
                
                Case "B"  'Boolean
                    If vData Then
                        DBSet = 1
                    Else
                        DBSet = 0
                    End If
            End Select
        End If
Error1:
    If Err.Number <> 0 Then MuestraError Err.Number, "Formato para la BD.", Err.Description
End Function

Public Function ImporteFormateadoSingle(Importe As String) As Single
Dim I As Integer

    If Importe = "" Then
        ImporteFormateadoSingle = 0
    Else
        'Primero quitamos los puntos
        Do
            I = InStr(1, Importe, ".")
            If I > 0 Then Importe = Mid(Importe, 1, I - 1) & Mid(Importe, I + 1)
        Loop Until I = 0
        ImporteFormateadoSingle = Importe
    End If
End Function

Public Function DevuelveValor(vSQL As String) As Variant
'Devuelve el valor de la SQL
Dim Rs As ADODB.Recordset

    On Error Resume Next

    Set Rs = New ADODB.Recordset
    Rs.Open vSQL, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    DevuelveValor = 0
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0).Value) Then DevuelveValor = Rs.Fields(0).Value   'Solo es para saber que hay registros que mostrar
    End If
    Rs.Close
    Set Rs = Nothing

    If Err.Number <> 0 Then
        DevuelveValor = 0
        Err.Clear
    End If
End Function

Public Sub PosicionarCombo(ByRef Combo1 As ComboBox, Valor As Integer)
'Situa el combo en la posicion de un valor concreto
Dim J As Integer

    On Error GoTo EPosCombo
    
    For J = 0 To Combo1.ListCount - 1
        If Combo1.ItemData(J) = Valor Then
            Combo1.ListIndex = J
            Exit For
        End If
    Next J

EPosCombo:
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub PosicionarCombo2(ByRef Combo1 As ComboBox, Valor As String)
'Situa el combo en la posicion de un valor concreto
Dim J As Integer

    On Error GoTo EPosCombo
    
    For J = 0 To Combo1.ListCount - 1
        If Trim(Combo1.List(J)) = Trim(Valor) Then
            Combo1.ListIndex = J
            Exit For
        End If
    Next J

EPosCombo:
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub PosicionarCombo3(ByRef Combo1 As ComboBox, Valor As String)
'Situa el combo en la posicion de un valor concreto
Dim J As Integer

    On Error GoTo EPosCombo
    
    For J = 0 To Combo1.ListCount - 1
        If Mid(Trim(Combo1.List(J)), 1, 3) = Trim(Valor) Then
            Combo1.ListIndex = J
            Exit For
        End If
    Next J

EPosCombo:
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Function ValorCombo(ByRef Cbo As ComboBox) As Integer
'obtiene el valor del combo de la posicion en la q se encuentra

    On Error GoTo EValCombo
    
    If Cbo.ListIndex < 0 Then
        ValorCombo = -1
    Else
        ValorCombo = Cbo.ItemData(Cbo.ListIndex)
    End If
    Exit Function

EValCombo:
    If Err.Number <> 0 Then Err.Clear
End Function


Public Function TextoCombo(ByRef Cbo As ComboBox) As String
'obtiene la descripcion del combo de la posicion en la q se encuentra

    On Error GoTo ErrTexCombo
    
    If Cbo.ListIndex < 0 Then
        TextoCombo = ""
    Else
        TextoCombo = Cbo.List(Cbo.ListIndex)
    End If
    Exit Function

ErrTexCombo:
    If Err.Number <> 0 Then Err.Clear
End Function

Public Sub SituarItemList(ByRef LView As ListView)
'Subir el item seleccionado del listview una posicion
Dim I As Byte, Item As Byte
Dim Aux As String
On Error Resume Next
   
    For I = 1 To LView.ListItems.Count
        If LView.ListItems(I).ToolTipText = vUsu.CadenaConexion Then
            LView.ListItems(I).Selected = True
            LView.ListItems(I).Bold = True
            LView.ListItems(I).ListSubItems(1).Bold = True
            LView.ListItems(I).ListSubItems(2).Bold = True
            LView.ListItems(I).EnsureVisible
            Exit For
        End If
    Next I
    LView.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Function QuitarCaracterNULL(vCad As String) As String
Dim I As Integer

    Do
        I = InStr(1, vCad, vbNullChar)
        If I > 0 Then 'Hay null
            vCad = Mid(vCad, 1, I - 1) & Mid(vCad, I + 2)
        End If
    Loop Until I = 0
    QuitarCaracterNULL = vCad
End Function

   

Public Sub ConseguirFoco(ByRef Text As TextBox, Modo As Byte)
'Acciones que se realizan en el evento:GotFocus de los TextBox:Text1
'en los formularios de Mantenimiento
On Error Resume Next

    If (Modo <> 0 And Modo <> 2) Then
        If Modo = 1 Then 'Modo 1: Busqueda
            Text.BackColor = vbLightBlue 'vbYellow
        End If
        Text.SelStart = 0
        Text.SelLength = Len(Text.Text)
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub ConseguirFocoLin(ByRef Text As TextBox)
'Acciones que se realizan en el evento:GotFocus de los TextBox:TxtAux para LINEAS
'en los formularios de Mantenimiento
On Error Resume Next

    With Text
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub ConseguirFocoChk(Modo As Byte)
'Acciones que se realizan en el evento:GotFocus de los TextBox:TxtAux para LINEAS
'en los formularios de Mantenimiento
On Error Resume Next

    If Modo = 0 Or Modo = 2 Then
'        KEYpress 13
        CreateObject("WScript.Shell").SendKeys "{tab}"

        
    End If
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Function CadenaDesdeHasta(ByRef TD As TextBox, TH As TextBox, Campo As String, TipoCampo As String, Optional NomCampo As String) As String
'Devuelve la cadena de seleccion: " (campo >= cadDesde and campo<=cadHasta) "
'para Crystal Report
Dim CadAux As String
Dim cadDesde As String, cadHasta As String
On Error GoTo ErrDH

    cadDesde = "": cadHasta = ""
    If Not TD Is Nothing Then cadDesde = TD.Text
    If Not TH Is Nothing Then cadHasta = TH.Text
    
    Campo = "{" & Campo & "}"
    If Trim(cadDesde) = "" And Trim(cadHasta) = "" Then
        'Campo Desde y Hasta no tiene valor
            CadAux = ""
    Else
        'Campo DESDE
        If cadDesde <> "" Then
            Select Case TipoCampo
                Case "N"
                    CadAux = Campo & " >= " & Val(cadDesde)
                Case "T"
                    CadAux = Campo & " >= """ & cadDesde & """"
                Case "F"
                    CadAux = Campo & " >= Date(" & Year(cadDesde) & "," & Month(cadDesde) & "," & Day(cadDesde) & ")"
                Case "FH"
                    CadAux = Campo & " >= DateTime(" & Year(cadDesde) & "," & Month(cadDesde) & "," & Day(cadDesde) & "," & Hour(cadDesde) & "," & Minute(cadDesde) & "," & Second(cadDesde) & ")"
                    
            End Select
        End If
        
        'Campo HASTA
        If cadHasta <> "" Then
            If CadAux <> "" Then 'Hay campo Desde y campo Hasta
                'Comprobar Desde <= Hasta
                Select Case TipoCampo
                    Case "N"
                        If CSng(cadDesde) > CSng(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            CadAux = "Error"
                        Else
                            CadAux = CadAux & " and " & Campo & " <= " & Val(cadHasta)
                        End If
                        
                    Case "T"
                        If cadDesde > cadHasta Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            CadAux = "Error"
                        Else
                            CadAux = CadAux & " and " & Campo & " <= """ & cadHasta & """"
                        End If
                    
                    Case "F"
                        If CDate(cadDesde) > CDate(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            CadAux = "Error"
                        Else
                            CadAux = CadAux & " and " & Campo & " <= Date(" & Year(cadHasta) & "," & Month(cadHasta) & "," & Day(cadHasta) & ")"
                        End If
                        
                    Case "FH"
                        If CDate(cadDesde) > CDate(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            CadAux = "Error"
                        Else
                                   
                            CadAux = CadAux & " AND " & Campo & " <= DateTime(" & Year(cadHasta) & "," & Month(cadHasta) & "," & Day(cadHasta) & ","
                            If Len(cadHasta) = 10 Then
                                CadAux = CadAux & "23,59,59"
                            Else
                                CadAux = CadAux & Hour(cadHasta) & "," & Minute(cadHasta) & "," & Second(cadHasta)
                            End If
                            CadAux = CadAux & ")"
                        End If
                End Select
            Else 'No hay campo Desde. Solo hay campo Hasta
                Select Case TipoCampo
                    Case "N"
                        CadAux = Campo & " <= " & Val(cadHasta)
                    Case "T"
                        CadAux = Campo & " <= """ & cadHasta & """"
                    Case "F"
                        CadAux = Campo & " <= Date(" & Year(cadHasta) & "," & Month(cadHasta) & "," & Day(cadHasta) & ")"
                    Case "FH"
                            CadAux = Campo & " <= DateTime(" & Year(cadHasta) & "," & Month(cadHasta) & "," & Day(cadHasta) & ","
                            If Len(cadHasta) = 10 Then
                                CadAux = CadAux & "23,59,59"
                            Else
                                CadAux = CadAux & Hour(cadHasta) & "," & Minute(cadHasta) & "," & Second(cadHasta)
                            End If
                            CadAux = CadAux & ")"
                        
                End Select
            End If
        End If
    End If
    If CadAux <> "" And CadAux <> "Error" Then CadAux = "(" & CadAux & ")"
    CadenaDesdeHasta = CadAux
ErrDH:
    If Err.Number <> 0 Then CadenaDesdeHasta = "Error"
End Function


Public Function CadenaDesdeHastaBD(TDes As TextBox, THas As TextBox, Campo As String, TipoCampo As String) As String
'Devuelve la cadena de seleccion: " (campo >= valor1 and campo<=valor2) "
'Para MySQL
Dim CadAux As String
Dim cadHasta As String
Dim cadDesde As String

    cadDesde = ""
    cadHasta = ""
    If Not TDes Is Nothing Then cadDesde = TDes.Text
    If Not THas Is Nothing Then cadHasta = THas.Text

    If Trim(cadDesde) = "" And Trim(cadHasta) = "" Then
        'Campo Desde y Hasta no tiene valor
            CadAux = ""
    Else
        'Campo DESDE
        If cadDesde <> "" Then
            Select Case TipoCampo
                Case "N"
                    CadAux = Campo & " >= " & Val(cadDesde)
                Case "T"
                    CadAux = Campo & " >= """ & cadDesde & """"
                Case "F"
                    CadAux = "(" & Campo & " >= '" & Format(cadDesde, FormatoFecha) & "')"
                Case "FH"
                    If Len(cadDesde) = 10 Then cadDesde = cadDesde & " 00:00:00"
                    CadAux = "(" & Campo & " >= '" & Format(cadDesde, FormatoFechaHora) & "')"
            End Select
        End If
        
        'Campo HASTA
        If cadHasta <> "" Then
            If CadAux <> "" Then 'Hay campo Desde y campo Hasta
                'Comprobar Desde <= Hasta
                Select Case TipoCampo
                    Case "N"
                        If CSng(cadDesde) > CSng(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            CadAux = "Error"
                        Else
                            CadAux = CadAux & " and " & Campo & " <= " & Val(cadHasta)
                        End If
                        
                    Case "T"
                        If CSng(cadDesde) > CSng(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            CadAux = "Error"
                        Else
                            CadAux = CadAux & " and " & Campo & " <= """ & cadHasta & """"
                        End If
                    
                    Case "F"
                        If CDate(cadDesde) > CDate(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            CadAux = "Error"
                        Else
                            CadAux = CadAux & " and (" & Campo & " <= '" & Format(cadHasta, FormatoFecha) & "')"
                        End If
                    Case "FH"
                        If Len(cadHasta) = 10 Then cadHasta = cadHasta & " 23:59:59"
                        If CDate(cadDesde) > CDate(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            CadAux = "Error"
                        Else
                            CadAux = CadAux & " AND (" & Campo & " <= '" & Format(cadHasta, FormatoFechaHora) & "')"
                        End If

                    

                End Select
            Else 'No hay campo Desde. Solo hay campo Hasta
                Select Case TipoCampo
                    Case "N"
                        CadAux = Campo & " <= " & Val(cadHasta)
                    Case "T"
                        CadAux = Campo & " <= """ & cadHasta & """"
                    Case "F"
                        CadAux = Campo & " <= '" & Format(cadHasta, FormatoFecha) & "'"
                End Select
            End If
        End If
    End If
    If CadAux <> "" And CadAux <> "Error" Then CadAux = "(" & CadAux & ")"
    CadenaDesdeHastaBD = CadAux
End Function

Public Function TotalRegistros(vSQL As String) As Long
'Devuelve el valor de la SQL
'para obtener COUNT(*) de la tabla
Dim Rs As ADODB.Recordset

    On Error Resume Next

    Set Rs = New ADODB.Recordset
    Rs.Open vSQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    TotalRegistros = 0
    If Not Rs.EOF Then
        If Rs.Fields(0).Value > 0 Then TotalRegistros = Rs.Fields(0).Value  'Solo es para saber que hay registros que mostrar
    End If
    Rs.Close
    Set Rs = Nothing

    If Err.Number <> 0 Then
        TotalRegistros = 0
        Err.Clear
    End If
End Function

Public Function TotalRegistrosConsulta(cadSQL) As Long
Dim Cad As String
Dim Rs As ADODB.Recordset

    On Error GoTo ErrTotReg
    Cad = "SELECT count(*) FROM (" & cadSQL & ") x"
    Set Rs = New ADODB.Recordset
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    If Not Rs.EOF Then
        TotalRegistrosConsulta = DBLet(Rs.Fields(0).Value, "N")
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
ErrTotReg:
    MuestraError Err.Number, "", Err.Description
End Function


Public Sub PonerLongCamposGnral(ByRef formulario As Form, Modo As Byte, Opcion As Byte)
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'ya que en busqueda se permite introducir criterios más largos del tamaño del campo
'en busqueda permitimos escribir: "0001:0004"
'en cambio al insertar o modificar la longitud solo debe permitir ser: "0001"
'(IN) formulario y Modo en que se encuentra el formulario
'(IN) Opcion : 1 para los TEXT1, 3 para los txtAux

    Dim I As Integer
    
    On Error Resume Next

    With formulario
        If Modo = 1 Then 'BUSQUEDA
            Select Case Opcion
                Case 1 'Para los TEXT1
                    For I = 0 To .Text1.Count - 1
                        With .Text1(I)
                            If .MaxLength <> 0 Then
                               .HelpContextID = .MaxLength 'guardamos es maxlenth para reestablecerlo despues
                                .MaxLength = 0 'tamaño infinito
                            End If
                        End With
                    Next I
                
                Case 3 'para los TXTAUX
                    For I = 0 To .txtaux.Count - 1
                        With .txtaux(I)
                            If .MaxLength <> 0 Then
                               .HelpContextID = .MaxLength 'guardamos es maxlenth para reestablecerlo despues
                                .MaxLength = 0 'tamaño infinito
                            End If
                        End With
                    Next I
            End Select
            
        Else 'resto de modos
            Select Case Opcion
                Case 1 'par los Text1
                    For I = 0 To .Text1.Count - 1
                        With .Text1(I)
                            If .HelpContextID <> 0 Then
                                .MaxLength = .HelpContextID 'volvemos a poner el valor real del maxlenth
                                .HelpContextID = 0
                            End If
                        End With
                    Next I
                Case 3 'para los txtAux
                    For I = 0 To .txtaux.Count - 1
                        With .txtaux(I)
                            If .HelpContextID <> 0 Then
                                .MaxLength = .HelpContextID 'volvemos a poner el valor real del maxlenth
                                .HelpContextID = 0
                            End If
                        End With
                    Next I
            End Select
        End If
    End With
    
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Sub CargaComboTabla(ByRef Cbo As ComboBox, Codigo As String, Descripcion As String, tabla As String, whereyorden As String)
 Dim RN As ADODB.Recordset
 Dim Cad As String
    Set RN = New ADODB.Recordset
    Cad = "Select " & Codigo & "," & Descripcion & " FROM " & tabla
    Cad = Cad & whereyorden
    Cbo.Clear
    RN.Open Cad, Conn, adOpenForwardOnly, adCmdText
    While Not RN.EOF
        Cbo.AddItem RN.Fields(1)
        Cbo.ItemData(Cbo.NewIndex) = RN.Fields(0)
        RN.MoveNext
    Wend
    RN.Close
    Set RN = Nothing
End Sub


Public Sub CargaGridGnral(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, SQL As String, PrimeraVez As Boolean)
    On Error GoTo ECargaGRid

    vDataGrid.Enabled = True
    '    vdata.Recordset.Cancel
    vData.ConnectionString = Conn
    vData.RecordSource = SQL
    vData.CursorType = adOpenDynamic
    vData.LockType = adLockPessimistic
    vDataGrid.ScrollBars = dbgNone
    vData.Refresh
    
    Set vDataGrid.DataSource = vData
    vDataGrid.AllowRowSizing = False
    vDataGrid.RowHeight = 350 '350
    
    If PrimeraVez Then
        vDataGrid.ClearFields
        vDataGrid.ReBind
        vDataGrid.Refresh
    End If
    
ECargaGRid:
    If Err.Number <> 0 Then MuestraError Err.Number, "CargaGrid", Err.Description
End Sub

'Esto es para que cuando pincha en siguiente le sugerimos
'Se puede comentar todo y asi no hace nada ni da error
'El SQL es propio de cada tabla
Public Function SugerirCodigoSiguienteStr(NomTabla As String, NomCodigo As String, Optional CondLineas As String) As String
Dim SQL As String
Dim Rs As ADODB.Recordset

    On Error GoTo ESugerirCodigo

    'SQL = "Select Max(codtipar) from stipar"
    SQL = "Select Max(" & NomCodigo & ") from " & NomTabla
    If CondLineas <> "" Then
        SQL = SQL & " WHERE " & CondLineas
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, , , adCmdText
    SQL = "1"
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            If IsNumeric(Rs.Fields(0)) Then
                SQL = CStr(Rs.Fields(0) + 1)
            Else
                If Asc(Left(Rs.Fields(0), 1)) <> 122 Then 'Z
                SQL = Left(Rs.Fields(0), 1) & CStr(Asc(Right(Rs.Fields(0), 1)) + 1)
                End If
            End If
        End If
    End If
    Rs.Close
    Set Rs = Nothing
    SugerirCodigoSiguienteStr = SQL
ESugerirCodigo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Public Sub CargarValoresAnteriores(formulario As Form, Optional opcio As Integer, Optional nom_frame As String)
Dim Control As Object
Dim mTag As CTag
Dim Izda As String
Dim Cad As String
    Set mTag = New CTag

    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If mTag.Columna <> "" And Not mTag.EsClave Then
                            If Izda <> "" Then Izda = Izda & " , "
                            'Access
                            'Izda = Izda & "[" & mTag.Columna & "]"
                            Izda = Izda & "" & mTag.Columna & " = "
                            'Parte VALUES
                            Cad = ValorParaSQL(Control.Text, mTag)
                            Izda = Izda & Cad
                        End If
                    End If
                End If
            End If
            
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If Not mTag.EsClave Then
                        If Izda <> "" Then Izda = Izda & " , "
                        'Access
                        'Izda = Izda & "[" & mTag.Columna & "]"
                        Izda = Izda & "" & mTag.Columna & " = "
                        If Control.Value = 1 Then
                            Cad = "1"
                            Else
                            Cad = "0"
                        End If
                        If mTag.TipoDato = "N" Then Cad = Abs(CBool(Cad))
                        Izda = Izda & Cad
                    End If
                End If
            End If
            
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado And Not mTag.EsClave Then
                        If Izda <> "" Then Izda = Izda & " , "
                        'Izda = Izda & "[" & mTag.Columna & "]"
                        Izda = Izda & "" & mTag.Columna & " = "
                        If Control.ListIndex = -1 Then
                            Cad = ValorNulo
                        ElseIf mTag.TipoDato = "N" Then
                            Cad = Control.ItemData(Control.ListIndex)
                        Else
                            Cad = ValorParaSQL(Control.List(Control.ListIndex), mTag)
                        End If
                        Izda = Izda & Cad
                    End If
                End If
            End If
            
        'OPTION BUTTON
        ElseIf TypeOf Control Is OptionButton Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
                    mTag.Cargar Control
                    If mTag.Cargado And Not mTag.EsClave Then
                        If Control.Value Then
                            If Izda <> "" Then Izda = Izda & " , "
                            Izda = Izda & "" & mTag.Columna & " = "
                            Cad = Control.Index
                            Izda = Izda & Cad
                        End If
                    End If
                End If
            End If
            
'        ElseIf TypeOf Control Is DTPicker Then
'            If Control.Tag <> "" Then
'                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
'                    mTag.Cargar Control
'                    If mTag.Cargado And Not mTag.EsClave Then
'                        If Izda <> "" Then Izda = Izda & " , "
'                        Izda = Izda & "" & mTag.Columna & " = "
'
'                        'Parte VALUES
'                        If Control.Visible Then
'                            cad = ValorParaSQL(Control.Value, mTag)
'                        Else
'                            cad = ValorNulo
'                        End If
'                        Izda = Izda & cad
'                    End If
'                End If
'            End If
        End If
        
    Next Control

    ValorAnterior = Izda

End Sub

Public Function SituarDataTrasEliminar(ByRef vData As Adodc, NumReg, Optional no_refre As Boolean) As Boolean
    On Error GoTo ESituarDataElim

    If Not no_refre Then vData.Refresh 'quan siga False o no es passe a la funció, es refrescarà. Hi ha que passar-lo com a True quan el manteniment siga Grid per a que no refresque
    
    If Not vData.Recordset.EOF Then    'Solo habia un registro
        If NumReg > vData.Recordset.RecordCount Then
            vData.Recordset.MoveLast
        Else
            vData.Recordset.MoveFirst
            vData.Recordset.Move NumReg - 1
        End If
        SituarDataTrasEliminar = True
    Else
        SituarDataTrasEliminar = False
    End If
        
ESituarDataElim:
    If Err.Number <> 0 Then
        Err.Clear
        SituarDataTrasEliminar = False
    End If
End Function


Public Sub SituarCombo(Cbo As ComboBox, Valor As Integer)
Dim I As Integer
    For I = 0 To Cbo.ListCount - 1
        If Cbo.ItemData(I) = Valor Then
            Cbo.ListIndex = I
            Exit For
        End If
    Next
    If I > Cbo.ListCount - 1 Then Cbo.ListIndex = -1
End Sub


Public Sub AnyadirLinea(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
On Error Resume Next

    vDataGrid.AllowAddNew = True
    If vData.Recordset.RecordCount > 0 Then
        vDataGrid.HoldFields
        vData.Recordset.MoveLast
        vDataGrid.Row = vDataGrid.Row + 1
    End If
    vDataGrid.Enabled = False
    
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub LimpiarLin(ByRef formulario As Form, nomframe As String)
'Limpiar los controles Text que esten dentro del frame nomFrame
    Dim Control As Object

    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            If Control.Container.Name = nomframe Then
                Control.Text = ""
            End If
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Container.Name = nomframe Then
                Control.ListIndex = -1
            End If
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Container.Name = nomframe Then
                Control.Value = 0
            End If
        End If
    Next Control
End Sub


Public Function PonerFormatoFecha(ByRef T As TextBox) As Boolean
Dim Cad As String

    Cad = T.Text
    If Cad <> "" Then
        If Not EsFechaOK(T) Then
            MsgBox "Fecha incorrecta. (dd/MM/yyyy)", vbExclamation
            Cad = "mal"
        End If
        If Cad <> "" And Cad <> "mal" Then
'            T.Text = Cad
            PonerFormatoFecha = True
        Else
            PonFoco T
        End If
    End If
End Function

Public Function PonerFormatoFechaHora(ByRef T As TextBox) As Boolean
Dim Cad As String

    Cad = T.Text
    If Cad <> "" Then
        If Not EsFechaHoraOK(T) Then
            MsgBox "Fecha/hora incorrecta. (dd/MM/yyyy hh:mm:ss)", vbExclamation
            Cad = "mal"
        End If
        If Cad <> "" And Cad <> "mal" Then
'            T.Text = Cad
            PonerFormatoFechaHora = True
        Else
            PonFoco T
        End If
    End If
End Function



Public Function ComprobarCero(Valor As String) As String
    If Valor = "" Then
        ComprobarCero = "0"
    Else
        ComprobarCero = Valor
    End If
End Function

Public Function PonerFormatoDecimal(ByRef T As TextBox, tipoF As Single) As Boolean
'tipoF: tipo de Formato a aplicar
'  1 -> Decimal(12,2)
'  2 -> Decimal(8,3)
'  3 -> Decimal(10,2)
'  4 -> Decimal(5,2)

Dim Valor As Double
Dim PEntera As Currency
Dim NoOK As Boolean
Dim I As Byte
Dim cadEnt As String
'Dim mTas As CTag

    If T.Text = "" Then Exit Function
    PonerFormatoDecimal = False
    NoOK = False
    With T
        If Not EsNumerico(CStr(.Text)) Then
            PonFoco T
            Exit Function
        End If


        If InStr(1, .Text, ",") > 0 Then
            Valor = ImporteFormateado(.Text)
        Else
            cadEnt = .Text
            I = InStr(1, cadEnt, ".")
            If I > 0 Then cadEnt = Mid(cadEnt, 1, I - 1)
            If tipoF = 1 And Len(cadEnt) > 10 Then
                MsgBox "El valor no puede ser mayor de 9999999999,99", vbExclamation
                NoOK = True
            End If
            If NoOK Then
'                    .Text = ""
                T.SetFocus
                Exit Function
            End If
            Valor = CDbl(TransformaPuntosComas(.Text))
        End If
            
        'Comprobar la longitud de la Parte Entera
        PEntera = Int(Valor)
        Select Case tipoF 'Comprobar longitud
            Case 1 'Decimal(12,2)
                If Len(CStr(PEntera)) > 10 Then
                    MsgBox "El valor no puede ser mayor de 9999999999,99", vbExclamation
                    NoOK = True
                End If
            Case 2 'Decimal(8,3)
                If Len(CStr(PEntera)) > 5 Then
                    MsgBox "El valor no puede ser mayor de 99999,999", vbExclamation
                    NoOK = True
                End If
            Case 3 'Decimal(10,2)
                If Len(CStr(PEntera)) > 8 Then
                    MsgBox "El valor no puede ser mayor de 99999999,99", vbExclamation
                    NoOK = True
                End If
            Case 4 'Decimal(5,2)
                If Len(CStr(PEntera)) > 3 Then
                    MsgBox "El valor no puede ser mayor de 999,99", vbExclamation
                    NoOK = True
                End If
            
        End Select

            
            If NoOK Then
                PonerFormatoDecimal = False
                T.SetFocus
                Exit Function
            End If
            
            'Poner el Formato
            Select Case tipoF
                Case 1 'Formato Decimal(12,2)
                    .Text = Format(Valor, FormatoImporte)
                Case 2 'Formato Decimal(8,3)
                    .Text = Format(Valor, FormatoPrecio)
                Case 3 'Formato Decimal(10,2)
                    .Text = Format(Valor, FormatoDec10d2)
                Case 4 'Formato Decimal(5,2)
                    .Text = Format(Valor, FormatoPorcen)
            
            
            End Select
            PonerFormatoDecimal = True
    End With
End Function

Public Function PonerNombreDeCod(ByRef Txt As TextBox, tabla As String, Campo As String, Optional Codigo As String, Optional Tipo As String, Optional cBD As Byte, Optional codigo2 As String, Optional Valor2 As String, Optional tipo2 As String) As String
'Devuelve el nombre/Descripción asociado al Código correspondiente
'Además pone formato al campo txt del código a partir del Tag
Dim SQL As String
Dim Devuelve As String
Dim vTag As CTag
Dim ValorCodigo As String

    On Error GoTo EPonerNombresDeCod

    ValorCodigo = Txt.Text
    If ValorCodigo <> "" Then
        Set vTag = New CTag
        If vTag.Cargar(Txt) Then
            If Codigo = "" Then Codigo = vTag.Columna
            If Tipo = "" Then Tipo = vTag.TipoDato
            
            SQL = DevuelveDesdeBDNew(cConta, tabla, Campo, Codigo, ValorCodigo, Tipo, , codigo2, Valor2, tipo2)
            If vTag.TipoDato = "N" Then ValorCodigo = Format(ValorCodigo, vTag.Formato)
            Txt.Text = ValorCodigo 'Valor codigo formateado
            If SQL = "" Then
            
            Else
                PonerNombreDeCod = SQL 'Descripcion del codigo
            End If
        End If
        Set vTag = Nothing
    Else
        PonerNombreDeCod = ""
    End If
'    Exit Function
EPonerNombresDeCod:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Nombre asociado a código: " & Codigo, Err.Description
End Function



Public Sub CargarProgres(ByRef PBar As ProgressBar, Valor As Long)
On Error Resume Next
    PBar.Max = 100
    PBar.Value = 0
    PBar.Tag = Valor
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub IncrementarProgres(ByRef PBar As ProgressBar, Veces As Integer)
On Error Resume Next
    PBar.Value = PBar.Value + ((Veces * PBar.Max) / CInt(PBar.Tag))
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Function PonerContRegistros(ByRef vData As Adodc) As String
'indicador del registro donde nos encontramos: "1 de 20"
    On Error GoTo EPonerReg
    
    If Not vData.Recordset.EOF Then
        PonerContRegistros = vData.Recordset.AbsolutePosition & " de " & vData.Recordset.RecordCount
    Else
        PonerContRegistros = ""
    End If
    
EPonerReg:
    If Err.Number <> 0 Then
        Err.Clear
        PonerContRegistros = ""
    End If
End Function



Public Function Ejecuta(ByRef SQL As String) As Boolean

    On Error Resume Next
    Conn.Execute SQL
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cadena: " & SQL & vbCrLf
        Ejecuta = False
    Else
        Ejecuta = True
    End If
End Function




Public Sub BloquearTxt(ByRef Text As TextBox, B As Boolean, Optional EsContador As Boolean)
'Bloquea un control de tipo TextBox
'Si lo bloquea lo pone de color amarillo claro sino lo pone en color blanco (sino es contador)
'pero si es contador lo pone color azul claro
On Error Resume Next

    Text.Locked = B
    If Not B And Text.Enabled = False Then Text.Enabled = True
    If B Then
        If EsContador Then
            'Si Es un campo que se obtiene de un contador poner color azul
'            Text.BackColor = &H80000013 'Azul Claro
            Text.BackColor = &HFFFFC0   'Azul claro con vista
        Else
            Text.BackColor = &H80000018 'Amarillo Claro
        End If
    Else
        Text.BackColor = vbWhite
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub



'........................................................................................
'........................................................................................
'........................................................................................
'........................................................................................

' Cuenta contable

'........................................................................................
'........................................................................................
'........................................................................................
'........................................................................................


'Public Function CuentaCorrectaUltimoNivel(ByRef Cuenta As String, ByRef Devuelve As String) As Boolean
Public Function CuentaCorrectaUltimoNivelTextBox(ByRef TCta As TextBox, ByRef TDesCta As TextBox) As Boolean
'Comprueba si es numerica
Dim Error As String
Dim SQL As String

    CuentaCorrectaUltimoNivelTextBox = False
    Error = ""
    If TCta.Text = "" Then
        TDesCta.Text = ""
        
    Else
        If Not IsNumeric(TCta.Text) Then
            Error = "La cuenta debe de ser numérica: " & TCta.Text
        Else
        
            'Rellenamos si procede
            TCta.Text = RellenaCodigoCuenta(TCta.Text)
            
            If Not EsCuentaUltimoNivel(TCta.Text) Then
                Error = "No es cuenta de último nivel: " & TCta.Text
                
            Else
                SQL = DevuelveDesdeBD("nommacta", "ariconta" & vParam.Numconta & ".cuentas", "codmacta", TCta.Text, "T")
                If SQL = "" Then
                    Error = "No existe la cuenta : " & TCta.Text
                Else
                    'Llegados aqui, si que existe la cuenta
                    CuentaCorrectaUltimoNivelTextBox = True
                    TDesCta.Text = SQL
                End If
            End If
        End If
        If Error <> "" Then
            MsgBox Error, vbExclamation
            TCta.Text = ""
            TDesCta.Text = ""
            PonFoco TCta
        End If
    End If
End Function


Private Function RellenaCodigoCuenta(vCodigo As String) As String
    Dim I As Integer
    Dim J As Integer
    Dim Cont As Integer
    Dim Cad As String
    
    RellenaCodigoCuenta = vCodigo
    If Len(vCodigo) > vEmpresa.DigitosUltimoNivel Then Exit Function
    I = 0: Cont = 0
    Do
        I = I + 1
        I = InStr(I, vCodigo, ".")
        If I > 0 Then
            If Cont > 0 Then Cont = 1000
            Cont = Cont + I
        End If
    Loop Until I = 0
    
    'Habia mas de un punto
    If Cont > 1000 Or Cont = 0 Then Exit Function
    
    'Cambiamos el punto por 0's  .-Utilizo la variable maximocaracteres, para no tener k definir mas
    I = Len(vCodigo) - 1 'el punto lo quito
    J = vEmpresa.DigitosUltimoNivel - I
    Cad = ""
    For I = 1 To J
        Cad = Cad & "0"
    Next I
    
    Cad = Mid(vCodigo, 1, Cont - 1) & Cad
    Cad = Cad & Mid(vCodigo, Cont + 1)
    RellenaCodigoCuenta = Cad
End Function






'--------------------------------------


Public Function BloqueoManual(cadTabla As String, CadWhere As String, Optional OcultarMsg As Boolean) As Boolean
Dim Aux As String

On Error GoTo EBLOQ
    BloqueoManual = False
    If CadWhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
    Else
        Aux = "INSERT INTO zbloqueos(codusu,tabla,clave) VALUES(" & vUsu.Codigo & ",'" & cadTabla
        Aux = Aux & "',""" & CadWhere & """)"
        Conn.Execute Aux
        BloqueoManual = True
    End If
EBLOQ:
    If Err.Number <> 0 Then
        Aux = ""
        If Conn.Errors.Count > 0 Then
            If Conn.Errors(0).NativeError = 1062 Then
                '¡Ya existe el registro, luego esta bloqueada
                Aux = "BLOQUEO"
            End If
        End If
        
        If Aux = "" Then
            MuestraError Err.Number, "Bloqueo tabla"
        Else
            If Not OcultarMsg Then MsgBox "Registro bloqueado por otro usuario", vbExclamation
        End If
    End If
'    Screen.MousePointer = AntiguoCursor
End Function



Public Function DesBloqueoManual(cadTabla As String) As Boolean
Dim SQL As String

'Solo me interesa la tabla
On Error Resume Next

        SQL = "DELETE FROM zbloqueos WHERE codusu=" & vUsu.Codigo & " and tabla='" & cadTabla & "'"
        Conn.Execute SQL
        If Err.Number <> 0 Then
            Err.Clear
        End If
End Function


Public Sub BorrarZBloqueos()
    On Error GoTo eBorrarZBloqueos
    
    If Not App.PrevInstance Then Conn.Execute "DELETE FROM zbloqueos WHERE codusu=" & vUsu.Codigo
    Exit Sub
eBorrarZBloqueos:
    Err.Clear
    Conn.Errors.Clear
End Sub


'---------------------------------
