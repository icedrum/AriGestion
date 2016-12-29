Attribute VB_Name = "libArigestion"
Option Explicit



Public Function PuedeBorrarCliente(IdCliente As Long, mostrarMsg As Boolean) As Boolean
Dim C As String
Dim Aux As String


    C = ""
    
    'Si tiene xpediente
    Aux = DevuelveDesdeBD("count(*)", "expedientes  ", "codclien", CStr(IdCliente))
    If Val(Aux) > 0 Then C = C & vbCrLf & "- Tiene expedientes"
    
    Aux = DevuelveDesdeBD("count(*)", "factcli", "codclien", CStr(IdCliente))
    If Val(Aux) > 0 Then C = C & vbCrLf & "- Tiene facturas"
    
    Aux = DevuelveDesdeBD("count(*)", "clientes_doc", "codclien", CStr(IdCliente))
    If Val(Aux) > 0 Then C = C & vbCrLf & "- Tiene documentos asociados"
    
    
    
    If C <> "" Then
        PuedeBorrarCliente = False
        C = "No se puede eliminar el cliente: " & vbCrLf & C
        If mostrarMsg Then MsgBox C, vbExclamation
    Else
        PuedeBorrarCliente = True
    End If

End Function






Public Function EliminarCliente(IdCliente As Long) As Boolean
Dim C As String
    On Error GoTo eEliminarCliente
    EliminarCliente = False
    
    Conn.BeginTrans
    
    For I = 1 To 6
        C = "DELETE from " & RecuperaValor("clientes_cuotas|clientes_fiscal|clientes_doc|clientes_historial|clientes_laboral|clientes|", 4)
        C = C & " WHERE codclien=" & IdCliente
        Conn.Execute C
    Next
    'Llega aqui, todo bien
    EliminarCliente = True
    
    
eEliminarCliente:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
        Conn.RollbackTrans
    Else
        Conn.CommitTrans
    End If
End Function

