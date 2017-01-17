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


'En cada guardar el importe vencido
Public Function TieneCobrosPendientes(IdCliente As Long, ByRef Cad As String) As Boolean
Dim Aux As String
    TieneCobrosPendientes = False
    Aux = "  codmacta IN ('" & DevuelveCuentaContableCliente(True, CStr(IdCliente))
    Aux = Aux & " ','" & DevuelveCuentaContableCliente(False, CStr(IdCliente)) & "')"
    Aux = Aux & " AND (ImpVenci - coalesce(gastos, 0) - coalesce(impcobro, 0)) <> 0 and now()>fecvenci AND 1"
    Cad = DevuelveDesdeBD("sum(ImpVenci - coalesce(gastos, 0) - coalesce(impcobro, 0))", "ariconta" & vParam.Numconta & ".cobros", Aux, "1")
    If Cad <> "" Then
        If CCur(Cad) = 0 Then Cad = ""
    End If
    If Cad <> "" Then TieneCobrosPendientes = True


   
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




Public Function FechaFacturaOK(FecFac As Date) As Boolean
Dim F As Date
Dim Cad As String
    
    
    
    F = DateAdd("yyyy", 1, vEmpresa.FechaInicioEjercicio)
    Cad = ""
    If FecFac < vEmpresa.FechaInicioEjercicio Then
        Cad = "Fecha anterior inicio ejercicio"
    Else
        If FecFac >= F Then
            Cad = "Fecha posterior fin ejercicio"
        Else
            If FecFac <= vEmpresa.UltimoDiaPeriodoLiquidado Then Cad = "Periodo IVA liquidado"
        End If
    End If
    If Cad <> "" Then
        Cad = "Error en fecha factura: " & Cad
        MsgBox Cad, vbExclamation
        FechaFacturaOK = False
    Else
        FechaFacturaOK = True
    End If

End Function



Public Function NumeroFactura_y_Fecha_OK(Serie As String, NumFac As Long, FecFac As Date) As Boolean
Dim Cad As String
Dim RN As ADODB.Recordset
Dim Maxi As Long

    NumeroFactura_y_Fecha_OK = False
    Cad = "Select max(fecfactu),max(numfactu) from factcli where numserie = " & DBSet(Serie, "T")
    Cad = Cad & " AND year(fecfactu)=" & Year(vEmpresa.FechaInicioEjercicio)
    Set RN = New ADODB.Recordset
    RN.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    If Not RN.EOF Then
        If Not IsNull(RN.Fields(0)) Then
            If FecFac < RN.Fields(0) Then
                Cad = "Fecha anterior a ultima fecha facturada para la serie seleccionada"
            Else
                Maxi = DBLet(RN.Fields(1), "N")
                If Maxi > 0 Then
                    If Maxi > NumFac Then Cad = "Hay una factura mayor para la serie seleccionada (" & Maxi & ")"
                End If
            End If
        End If
    End If
    RN.Close
    Set RN = Nothing
    If Cad <> "" Then
        MsgBox Cad, vbExclamation
        Exit Function
    Else
        NumeroFactura_y_Fecha_OK = True
    End If

End Function




