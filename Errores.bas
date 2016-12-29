Attribute VB_Name = "Errores"

'---------------------------------------------------------------
' Si tiene error el objeto conn entonces lo mostramos
'---------------------------------------------------------------

Public Sub ControlamosError(ByRef CADENA As String)

Select Case Conn.Errors(0).NativeError
Case 0
    CADENA = "El controlador ODBC no admite las propiedades solicitadas."
Case 1044
    CADENA = "Acceso denegado para usuario: " & CadenaDesde(15, Conn.Errors(0).Description, ":")
Case 1045
    CADENA = "Acceso denegado para usuario: " & CadenaDesde(15, Conn.Errors(0).Description, ":")
Case 1048
    CADENA = "Columna no puede ser nula: " & CadenaDesde(1, Conn.Errors(0).Description, ":")
Case 1049
    CADENA = "Base de datos desconocida: " & CadenaDesde(1, Conn.Errors(0).Description, "'")
Case 1052
    CADENA = "La columna :" & CadenaDesde(1, Conn.Errors(0).Description, "'") & " tiene un nombre ambiguo "
Case 1054
    CADENA = "Columna desconocida en cadena SQL."
Case 1062
    CADENA = "Entrada duplicada en BD." & vbCrLf & CadenaDesde(60, Conn.Errors(0).Description, "'")
Case 1064
    CADENA = "Error en el SQL."
Case 1109
    CADENA = "Tabla desconocida:  " & CadenaDesde(1, Conn.Errors(0).Description, "'")
Case 1110
    CADENA = "Columna : " & CadenaDesde(1, Conn.Errors(0).Description, "'") & " especificada dos veces"
Case 1146
    CADENA = "Tabla no existe:  " & CadenaDesde(1, Conn.Errors(0).Description, "'")
Case 1136
    CADENA = "Nº de columnas en el SQL incorrectos."
Case 1205
    CADENA = "Tabla bloqueada. Tiempo espera excedido"
Case 1216
    CADENA = "Imposible añadir una columna hija. Fallo en la clave referencial"
Case 1217
    CADENA = "El registro es clave referencial en otras tablas"
Case 2003
    CADENA = "Imposible conectar con el servidor " & CadenaDesde(15, Conn.Errors(0).Description, "'")
Case 2005
    CADENA = "Servidor host MYSQL desconocido:  " & CadenaDesde(1, Conn.Errors(0).Description, "'")
Case 2013
    CADENA = "Se ha perdido la conexión con el servidor MySQL durante la ejecución."
Case Else
    CADENA = ""
End Select
End Sub


Private Function CadenaDesde(Inicio As Integer, CADENA As String, Caracter As String) As String
Dim i, J
CadenaDesde = ""
i = InStr(Inicio, CADENA, Caracter)
If i >= Inicio Then
    J = InStr(i + 1, CADENA, Caracter)
    i = i + 1
    If J > 0 Then CadenaDesde = Mid(CADENA, i, J - i)
End If
End Function


