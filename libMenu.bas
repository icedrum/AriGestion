Attribute VB_Name = "libMenu"
Option Explicit


Private ultNodo As Long
Private Const QueAplicacion = "arigestion"


Private Const CoordX = "450,14182|2085,166|3720,189|5355,213|6990,236|8625,261|10260,28|11894,74|"
Private Const CoordY = "30,04725|1665,071|3300,095|4935,118|6570,142|"
  
  
Public Sub CargaMenu(aplicacion As String, ByRef Tr1 As TreeView)
Dim cad As String
Dim RN As ADODB.Recordset
Dim N As Node
Dim NodoPadre As String
Dim ClaveNodo As String
Dim Sql As String

    Set RN = New ADODB.Recordset
    
    On Error GoTo eCargaMenu
    
    
    Tr1.Nodes.Clear
    
    
    cad = "Select * from menus where aplicacion = '" & aplicacion & "' ORDER BY padre,orden"
    RN.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
    
        If aplicacion = "introcon" Then
                ClaveNodo = "TR" & Format(RN!Codigo, "000000")
                If RN!Padre = 0 Then
                    'EL es padre de nivel 0
                    Set N = Tr1.Nodes.Add(, , ClaveNodo)
                    N.Bold = True
                Else
                    NodoPadre = "TR" & Format(RN!Padre, "000000")
                    Set N = Tr1.Nodes.Add(NodoPadre, tvwChild, ClaveNodo)
                End If
                
                N.Text = Trim(RN!Descripcion)
    '            If RN!oculto = 1 Then N.BackColor = vbRed  'Eso es que esta coluto
            
                'EN EL TAG lleva
                'los siguientes campos enpipados
                'imagen
                N.Tag = DBLet(RN!imagen, "N") & "|"
        Else
           If Not BloqueaPuntoMenu(RN!Codigo, QueAplicacion) Then
            If MenuVisibleUsuario(DBLet(RN!Codigo), aplicacion) Then
             If (MenuVisibleUsuario(DBLet(RN!Padre), aplicacion) And DBLet(RN!Padre) <> 0) Or DBLet(RN!Padre) = 0 Then
            
                ClaveNodo = "TR" & Format(RN!Codigo, "000000")
                If RN!Padre = 0 Then
                    'EL es padre de nivel 0
                    Set N = Tr1.Nodes.Add(, , ClaveNodo)
                    N.Bold = True
                Else
                    NodoPadre = "TR" & Format(RN!Padre, "000000")
                    Set N = Tr1.Nodes.Add(NodoPadre, tvwChild, ClaveNodo)
                End If
                
                N.Text = Trim(RN!Descripcion)
    '            If RN!oculto = 1 Then N.BackColor = vbRed  'Eso es que esta coluto
            
                'EN EL TAG lleva
                'los siguientes campos enpipados
                'imagen
                N.Tag = DBLet(RN!imagen, "N") & "|"
            
                If False Then
                
                End If
             End If
            End If
           End If
        End If
        
        RN.MoveNext
    Wend
    RN.Close
    
    
eCargaMenu:
    If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation
    Set RN = Nothing
End Sub



Public Function IntercambiarNodo(ByRef Tv1 As TreeView, ByRef NBorrarCrear As Node, ByRef NodoRelativo As Node, Anterior As Boolean, SeleccionadoElDeBorrar As Boolean)
Dim Col As Collection
Dim poscion
Dim N2 As Node
Dim N3 As Node
Dim Aux As String
Dim J As Integer
Dim EstabaDesplegado As Boolean
Dim Clave As String
Dim Padre As String

    ' tvwNext  tvwPrevious
    If Anterior Then
        poscion = tvwPrevious
    Else
        poscion = tvwNext
    End If
    
    Set Col = New Collection
    
    
    EstabaDesplegado = False
    If NBorrarCrear.Children > 0 Then
        If NBorrarCrear.Child.Visible Then EstabaDesplegado = True
    End If
    
    GuardarNodo True, False, Col, NBorrarCrear
    Tv1.Nodes.Remove NBorrarCrear.Index
    
    Aux = Col.Item(1)
    Aux = RecuperaValor(Aux, 1)  'key
    
    'Creamos el nodo
    
    Set N2 = Tv1.Nodes.Add(NodoRelativo, poscion, Aux)
   
    
    'Asignamos los valores
    Aux = Col.Item(1)
    TextoANodo N2, Aux
    
    
    For J = 2 To Col.Count
        
    
        Aux = Col.Item(J)
        Clave = RecuperaValor(Aux, 1)
        Padre = RecuperaValor(Aux, 2)
        If Padre = "" Then Padre = N2.Key
            
        Set N3 = Tv1.Nodes.Add(Padre, tvwChild, Clave)
    
        Aux = Col.Item(J)
        TextoANodo N3, Aux

    Next
    
    If SeleccionadoElDeBorrar Then
        Set Tv1.SelectedItem = N2
    Else
        Set Tv1.SelectedItem = NodoRelativo
    End If
    If EstabaDesplegado Then Tv1.SelectedItem.Child.EnsureVisible
    Set Col = Nothing
End Function




Private Sub GuardarNodo(EsNodoAMover As Boolean, PrimerNodoHijo As Boolean, ByRef Cl As Collection, N As Node)
Dim N1 As Node
    
        Debug.Print N.Text
        Cl.Add NodoATexto(N)
    
        If N.Children > 0 Then
            Set N1 = N.Child
            GuardarNodo False, True, Cl, N1
        End If
        
        If Not EsNodoAMover Then
            If PrimerNodoHijo Then
                Set N1 = N.Next
                While Not N1 Is Nothing
                    GuardarNodo False, False, Cl, N1
                    Set N1 = N1.Next
                Wend
            End If
        End If
    
End Sub





Public Function SubirNivelNodo(ByRef Tv1 As TreeView, ByRef NodoASubir As Node, Subir As Boolean)
Dim EstabaDesplegado As Boolean
Dim Col As Collection
Dim N2 As Node
Dim N3 As Node
Dim Aux As String
Dim Padre As String
Dim J As Integer
Dim Clave As String
Dim posicion
    
    EstabaDesplegado = False
    If NodoASubir.Children > 0 Then
        If NodoASubir.Child.Visible Then EstabaDesplegado = True
    End If
    
    Set Col = New Collection
    
    If Subir Then
        
            
        Padre = NodoASubir.Parent.Key
            
        posicion = tvwPrevious
    Else
        
        If NodoASubir.Next.Children > 0 Then
            Padre = NodoASubir.Next.Child.Key
            posicion = tvwFirst
        Else
            Padre = NodoASubir.Next.Key
            posicion = tvwChild
        End If
        
    End If
    
    GuardarNodo True, False, Col, NodoASubir
    Tv1.Nodes.Remove NodoASubir.Index
    
    Aux = Col.Item(1)
    Aux = RecuperaValor(Aux, 1)  'key
    
    
    
    If Subir Then
        If Padre = "" Then
            Set N2 = Tv1.Nodes.Add(, , Aux)
        Else
            Set N2 = Tv1.Nodes.Add(Padre, posicion, Aux)
        End If
    Else
        'Bajar al nodo de abajo
        Set N2 = Tv1.Nodes.Add(Padre, posicion, Aux)

    End If
    
    Aux = Col.Item(1)
    TextoANodo N2, Aux
      
    
    
    For J = 2 To Col.Count
        
    
        Aux = Col.Item(J)
        Clave = RecuperaValor(Aux, 1)
        Padre = RecuperaValor(Aux, 2)
        If Padre = "" Then Padre = N2.Key
            
        Set N3 = Tv1.Nodes.Add(Padre, tvwChild, Clave)
    
        Aux = Col.Item(J)
        TextoANodo N3, Aux

    Next
    
    N2.EnsureVisible
    Set Tv1.SelectedItem = N2
    
    
    
    Set Col = Nothing
    


End Function





Private Function NodoATexto(N As Node) As String

    'Por si acaso mas adelante cambiamos cosas
    NodoATexto = N.Key & "|"
    If Not N.Parent Is Nothing Then NodoATexto = NodoATexto & N.Parent.Key
    NodoATexto = NodoATexto & "|" & N.Text & "|" & N.ForeColor & "|" & N.Tag & "|"
End Function


Private Sub TextoANodo(ByRef N As Node, Linea As String)
    'Por si acaso mas adelante cambiamos cosas
    
    N.Text = RecuperaValor(Linea, 3)
    
    N.ForeColor = RecuperaValor(Linea, 4)
    N.Tag = RecuperaValor(Linea, 5)
    
    
    
End Sub


Private Function NodoASQL(ByRef N As Node) As String
Dim Codigo As Long
Dim Sql As String

    'Aux = "INSERT INTO appnuevomenus(aplicacion,codigo,padre,descripcion,orden,ocultar,imagen) VALUES " & SQL
    ' apli,
    If Mid(N.Key, 1, 2) = "##" Then
        'ES NUEVO
        ultNodo = ultNodo + 1
        Codigo = ultNodo
    Else
        Codigo = Val(Mid(N.Key, 3))
        
    End If
    If N.Parent Is Nothing Then
        Sql = "0"
    Else
        Sql = Mid(N.Parent.Key, 3)
    End If
    Sql = Codigo & "," & Sql & ",'" & N.Text & "',"
    If N.ForeColor = vbRed Then
        Sql = Sql & "1"
    Else
        Sql = Sql & "0"
    End If
    NodoASQL = Sql & ",0"
    
End Function

Public Sub ActualizarExpansionMenus(Usuario As Long, ByRef Tr1 As TreeView, aplicacion As String)
Dim Sql As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset
Dim I As Long

    On Error GoTo eActualizarExpansionMenus
    
    'seleccionamos todos los que sean padres
    For I = 1 To Tr1.Nodes.Count
    
        Sql2 = "update menus_usuarios set expandido = "
        
        If Tr1.Nodes(I).Expanded Then
            Sql3 = Sql2 & "1" 'expandido
        Else
            Sql3 = Sql2 & "0" 'no expandido
        End If
        
        Sql3 = Sql3 & "  where codigo = " & DBSet(Mid(Tr1.Nodes(I).Key, 3, 6), "N") & " and codusu = " & DBSet(Usuario, "N") & " and aplicacion = " & DBSet(aplicacion, "T")
    
        Conn.Execute Sql3
    Next I
        
        
    Exit Sub
    
eActualizarExpansionMenus:
    MuestraError Err.Number, "Actualizar personalización de menus", Err.Description
End Sub



'******  Cada vez que hace mouseup, actualiza TODOS los iconos?
Public Sub ActualizarItems(Usuario As Long, ByRef Lv1 As ListView, aplicacion As String)
Dim Sql As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset
Dim I As Long

    On Error GoTo eActualizarItems
    
    Sql = "update menus_usuarios set posx = 0, posy = 0 where codusu = " & Usuario & " and aplicacion = " & DBSet(aplicacion, "T") & " and "
    Sql = Sql & " codigo in (select codigo from menus where aplicacion = " & DBSet(aplicacion, "T") & ")"
    
    Conn.Execute Sql
    
    For I = 1 To Lv1.ListItems.Count
        Debug.Print I & "  x,y   " & Lv1.ListItems(I).Left & ", " & Lv1.ListItems(I).top
        Sql = "update menus_usuarios set posx = " & DBSet(Lv1.ListItems(I).Left, "N") & ", posy = " & DBSet(Lv1.ListItems(I).top, "N") & " where codusu = " & Usuario & " and aplicacion = " & DBSet(aplicacion, "T") & " and "
        Sql = Sql & " codigo in (select codigo from menus where aplicacion = " & DBSet(aplicacion, "T") & " and descripcion =  " & DBSet(Lv1.ListItems(I).Text, "T") & ")"
    
        Conn.Execute Sql

    Next I
    
        
    Exit Sub
    
eActualizarItems:
    MuestraError Err.Number, "Actualizar Items de menus", Err.Description
End Sub


Public Sub RecolocarItems(Usuario As Long, ByRef Lv1 As ListView, aplicacion As String)
Dim CoordX As String
Dim CoordY As String
Dim X As Currency
Dim Y As Currency
Dim Sql As String
Dim J As Integer

    On Error GoTo eRecolocarItems


    CoordX = "450,14182|2085,166|3720,189|5355,213|6990,236|8625,261|10260,28|11894,74|"
    CoordY = "30,04725|1665,071|3300,095|4935,118|6570,142|"


    For I = 1 To Lv1.ListItems.Count
        If I <= 8 Then
            Y = RecuperaValor(CoordY, 1)
            J = I
        ElseIf I <= 16 Then
            Y = RecuperaValor(CoordY, 2)
            J = I - 8
        ElseIf I <= 24 Then
            Y = RecuperaValor(CoordY, 3)
            J = I - 16
        ElseIf I <= 32 Then
            Y = RecuperaValor(CoordY, 4)
            J = I - 32
        End If
        
        X = RecuperaValor(CoordX, J)
            
        Lv1.ListItems(I).Left = X
        Lv1.ListItems(I).top = Y
            
            
        Debug.Print I & "  x,y   " & Lv1.ListItems(I).Left & ", " & Lv1.ListItems(I).top
        
        Sql = "update menus_usuarios set posx = " & DBSet(X, "N") & ", posy = " & DBSet(Y, "N") & " where codusu = " & vUsu.Id & " and aplicacion = " & DBSet(aplicacion, "T") & " and "
        Sql = Sql & " codigo in (select codigo from menus where aplicacion = " & DBSet(aplicacion, "T") & " and descripcion =  " & DBSet(Lv1.ListItems(I).Text, "T") & ")"
    
        Conn.Execute Sql

    Next I


    Exit Sub


eRecolocarItems:
    MuestraError Err.Number, "Recolocar Items de menus", Err.Description
End Sub


Public Sub DevuelCoordenadasCuadricula(ColX As Integer, ColY As Integer, ByRef PosX As Single, ByRef PosY As Single)
    PosX = RecuperaValor(CoordX, ColX)
    PosY = RecuperaValor(CoordY, ColY)
End Sub

Public Sub ActualizarItemCuadricula(Usuario As Long, ByRef Lv1 As ListView, aplicacion As String, X As Single, Y As Single, ByRef VolverACargarLw As Boolean)
Dim Sql As String
Dim I As Integer
Dim Valor As Currency

    On Error GoTo eActualizarItems
    
    If Lv1.SelectedItem Is Nothing Then Exit Sub
    
    
        'Cuadricula X e y
       
        VolverACargarLw = False
    
        Valor = RecuperaValor(CoordX, 1)
        If X <= Valor + 817 Then
            Lv1.SelectedItem.Left = Valor
            If X < 0 Then VolverACargarLw = True
        Else
            For I = 2 To 8
                Valor = RecuperaValor(CoordX, I)
                If X <= Valor + 817 Then
                    Lv1.SelectedItem.Left = Valor
                    Exit For
                Else
                    'Es la utlima. No puede ir a mas
                    If I = 8 Then Lv1.SelectedItem.Left = Valor
                End If
            Next
        End If
          
        Valor = RecuperaValor(CoordY, 1)
        If Y <= Valor + 817 Then
            Lv1.SelectedItem.top = Valor
            If Y < 0 Then VolverACargarLw = True
        Else
            For I = 1 To 5
            Valor = RecuperaValor(CoordY, I)
                If Y <= Valor + 817 Then
                    Lv1.SelectedItem.top = Valor
                    Exit For
                Else
                    'Es la utlima. No puede ir a mas
                    If I = 5 Then Lv1.SelectedItem.top = Valor
                End If
            Next
        End If
        
        Sql = "select count(*) from menus_usuarios where posx = " & DBSet(Lv1.SelectedItem.Left, "N") & " and posy = " & DBSet(Lv1.SelectedItem.top, "N") & " and  codusu = " & Usuario & " and aplicacion = 'ariconta'  and "
        Sql = Sql & " codigo in (select codigo from menus where aplicacion = 'ariconta')"
        If TotalRegistros(Sql) = 0 Then
            Sql = "update menus_usuarios set posx = " & DBSet(Lv1.SelectedItem.Left, "N") & ", posy = " & DBSet(Lv1.SelectedItem.top, "N") & " where codusu = " & Usuario & " and aplicacion = " & DBSet(aplicacion, "T") & " and "
            Sql = Sql & " codigo in (select codigo from menus where aplicacion = " & DBSet(aplicacion, "T") & " and descripcion =  " & DBSet(Lv1.SelectedItem, "T") & ")"
        
            Conn.Execute Sql
        Else
        
' no hace falta actualizarlo es donde estaba
'            Sql = "update menus_usuarios set posx = " & DBSet(XAnt, "N") & ", posy = " & DBSet(YAnt, "N") & " where codusu = " & vUsu.Id & " and aplicacion = 'ariconta' and "
'            Sql = Sql & " codigo in (select codigo from menus where aplicacion = 'ariconta' and descripcion =  " & DBSet(Lv1.SelectedItem, "T") & ")"
'
'            Conn.Execute Sql
            
            Lv1.SelectedItem.top = YAnt
            Lv1.SelectedItem.Left = XAnt
            
        End If

    
        
    Exit Sub
    
eActualizarItems:
    MuestraError Err.Number, "Actualizar Items de menus", Err.Description
End Sub



Public Function MenuVisibleUsuario(Proceso As Long, aplicacion As String) As Boolean
Dim Sql As String
Dim Excepcion As String


'    MenuVisibleUsuario = True
'Exit Function


    Sql = "select ver from menus_usuarios where codigo = " & DBSet(Proceso, "N") & " and aplicacion = " & DBSet(aplicacion, "T")
    Sql = Sql & " and codusu = " & DBSet(vUsu.Id, "N")
    
   ' If Not vEmpresa.TieneTesoreria Then
   '     Sql = Sql & " and not codigo in (select codigo from menus where aplicacion = " & DBSet(aplicacion, "T") & " and tipo = 1)"
   ' End If
   '
   ' If Not vEmpresa.TieneContabilidad Then
   '     Sql = Sql & " and not codigo in (select codigo from menus where aplicacion = " & DBSet(aplicacion, "T") & " and tipo = 0)"
   ' End If
    
    MenuVisibleUsuario = (DevuelveValor(Sql) = 1)

End Function

Public Function BloqueaPuntoMenu(IdProg As Long, aplicacion As String) As Boolean
Dim EsdeAnalitica As Boolean

    BloqueaPuntoMenu = False

    If aplicacion = QueAplicacion Then
        ' programas de analitica
      '  EsdeAnalitica = (IdProg = 10 Or IdProg = 1001 Or IdProg = 1002 Or IdProg = 1003 Or IdProg = 1004 Or IdProg = 1005)
      '  BloqueaPuntoMenu = (Not vParam.autocoste And EsdeAnalitica)
    End If
    
End Function

