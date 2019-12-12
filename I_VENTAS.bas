Attribute VB_Name = "I_VENTAS"
Option Explicit

Public Function fBuklAlicuotas(ByVal Datos As String, ByVal Formato As String)
Dim ls_SP As String
Dim result As Integer
Dim desResult As String
Dim return_value As Integer
On Error GoTo tratar

Dim oCom As ADODB.Command
Dim Param As ADODB.Parameter
    
    Set oCom = New ADODB.Command

    ls_SP = "[dbo].[sp_crearTMPI_ALICUOTAS]"
    
    With oCom
        .ActiveConnection = oCon
        .CommandType = adCmdStoredProc
        .CommandText = ls_SP
        .Prepared = True
        
        .Parameters.Append .CreateParameter(, adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@Datos", adVarChar, adParamInput, Len(Datos), Datos)
        .Parameters.Append .CreateParameter("@Formato", adVarChar, adParamInput, Len(Formato), Formato)
        .Parameters.Append .CreateParameter("@result", adInteger, adParamOutput, 8)
        .Parameters.Append .CreateParameter("@desResult", adVarChar, adParamOutput, 255)

    End With

    oCom.Execute
    
    return_value = oCom.Parameters(0)
    result = oCom.Parameters(3)
    desResult = oCom.Parameters(4)
    
    If result <> 0 Then
        MsgBox "Error Número: " & result & " - " & desResult, vbCritical, "fBuklAlicuotas"
        return_value = result
    End If
    fBuklAlicuotas = return_value
    
    Exit Function

tratar:
    fBuklAlicuotas = -1
    If oCon.Errors.Count >= 1 Then
        MsgBox oCon.Errors(0).Description, vbCritical, "Nro de Error: " & oCon.Errors(0).Number
        oCon.Errors.Clear
    Else
        MsgBox Err.Number & "-" & Err.Description, vbCritical, "fBuklAlicuotas"
        Err.Clear
    End If

End Function

Public Function fBuklVentas(ByVal Datos As String, ByVal Formato As String)
Dim ls_SP As String
Dim result As Integer
Dim desResult As String
Dim return_value As Integer
On Error GoTo tratar

Dim oCom As ADODB.Command
Dim Param As ADODB.Parameter
    
    Set oCom = New ADODB.Command

    ls_SP = "[dbo].[sp_crearTMPI_VENTAS]"
    
    With oCom
        .ActiveConnection = oCon
        .CommandType = adCmdStoredProc
        .CommandText = ls_SP
        .Prepared = True
        
        .Parameters.Append .CreateParameter(, adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@Datos", adVarChar, adParamInput, Len(Datos), Datos)
        .Parameters.Append .CreateParameter("@Formato", adVarChar, adParamInput, Len(Formato), Formato)
        .Parameters.Append .CreateParameter("@result", adInteger, adParamOutput, 8)
        .Parameters.Append .CreateParameter("@desResult", adVarChar, adParamOutput, 255)

    End With

    oCom.Execute
    
    return_value = oCom.Parameters(0)
    result = oCom.Parameters(3)
    desResult = oCom.Parameters(4)
    
    If result <> 0 Then
        MsgBox "Error Número: " & result & " - " & desResult, vbCritical, "fBuklVentas"
        return_value = result
    End If
    fBuklVentas = return_value
    
    Exit Function

tratar:
    fBuklVentas = -1
    If oCon.Errors.Count >= 1 Then
        MsgBox oCon.Errors(0).Description, vbCritical, "Nro de Error: " & oCon.Errors(0).Number
        oCon.Errors.Clear
    Else
        MsgBox Err.Number & "-" & Err.Description, vbCritical, "fBuklVentas"
        Err.Clear
    End If

End Function

Public Function cantidadTmpI_Ventas() As Integer

Dim ls_sql As String
Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
    ls_sql = "select count(*) from I_VENTAS"

    Set rs = New ADODB.Recordset
    
    
    With rs
        .ActiveConnection = oCon
        .CursorType = adOpenForwardOnly
        .Source = ls_sql
        .Open
    End With

    cantidadTmpI_Ventas = rs(0)
    
    Set rs = Nothing

End Function

Public Function cantidadTmpI_Alicuotas() As Integer

Dim ls_sql As String
Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
    ls_sql = "select count(*) from I_ALICUOTAS"

    Set rs = New ADODB.Recordset
    
    
    With rs
        .ActiveConnection = oCon
        .CursorType = adOpenForwardOnly
        .Source = ls_sql
        .Open
    End With

    cantidadTmpI_Alicuotas = rs(0)
    
    Set rs = Nothing

End Function

Public Sub BorrarTmpI_Ventas()
Dim ls_sql As String
Dim oCom As ADODB.Command

    Set oCom = New ADODB.Command
    ls_sql = "DELETE From I_VENTAS"
    With oCom
        .ActiveConnection = oCon
        .CommandType = adCmdText
        .CommandText = ls_sql
        .Prepared = True
    End With
    oCom.Execute
    
End Sub


Public Sub BorrarTmpI_Alicuotas()
Dim ls_sql As String
Dim oCom As ADODB.Command

    Set oCom = New ADODB.Command
    ls_sql = "DELETE From I_ALICUOTAS"
    With oCom
        .ActiveConnection = oCon
        .CommandType = adCmdText
        .CommandText = ls_sql
        .Prepared = True
    End With
    oCom.Execute
    
End Sub

Public Function cantidadTipoComprobanteI_Ventas()
Dim ls_sql As String
Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
    ls_sql = "select count (distinct Tipocomprobante) from i_ventas"

    Set rs = New ADODB.Recordset
    
    
    With rs
        .ActiveConnection = oCon
        .CursorType = adOpenForwardOnly
        .Source = ls_sql
        .Open
    End With
      
    cantidadTipoComprobanteI_Ventas = rs(0)
    
    Set rs = Nothing
End Function


Public Function cantidadTipoComprobanteHabilitados_IVentas()
Dim ls_sql As String
Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
    ls_sql = "select count(*) from TIPO_COMPROBANTE " & " where codigo in (select distinct Tipocomprobante from i_ventas) and VISIBLE = 1"

    Set rs = New ADODB.Recordset
    
    
    With rs
        .ActiveConnection = oCon
        .CursorType = adOpenForwardOnly
        .Source = ls_sql
        .Open
    End With
      
    cantidadTipoComprobanteHabilitados_IVentas = rs(0)
    
    Set rs = Nothing
End Function





Public Function cantidadPuntosDeVenta_IVentas()
Dim ls_sql As String
Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
    ls_sql = "select count (distinct PuntoVenta) from i_ventas"

    Set rs = New ADODB.Recordset
    
    
    With rs
        .ActiveConnection = oCon
        .CursorType = adOpenForwardOnly
        .Source = ls_sql
        .Open
    End With
      
    cantidadPuntosDeVenta_IVentas = rs(0)
    
    Set rs = Nothing
End Function

Public Function cantidadPuntosDeVentaHabilitados_IVentas(empresa_id As String)
Dim ls_sql As String
Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
    ls_sql = "select count(*) from PUNTO_VENTA where EMPRESA_ID=" & empresa_id & " and ACTIVO = 1 and PUNTO_VENTA in (select distinct PuntoVenta from i_ventas)"
    Set rs = New ADODB.Recordset
    
    
    With rs
        .ActiveConnection = oCon
        .CursorType = adOpenForwardOnly
        .Source = ls_sql
        .Open
    End With
      
    cantidadPuntosDeVentaHabilitados_IVentas = rs(0)
    
    Set rs = Nothing

End Function


Public Function todosLosComprobantesTieneAlicuotas()
Dim ls_sql As String
Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
    ls_sql = "select COUNT(*) from i_ventas  v where not exists(select 1 from i_alicuotas  a " & " where a.TipoComprobante = v.tipoComprobante and a.PuntoVenta = v.PuntoVenta and a.NumeroComprobante = v.NumeroComprobante)  "
    Set rs = New ADODB.Recordset
    
    
    With rs
        .ActiveConnection = oCon
        .CursorType = adOpenForwardOnly
        .Source = ls_sql
        .Open
    End With
      
    todosLosComprobantesTieneAlicuotas = rs(0)
    
    Set rs = Nothing
End Function

Public Function ExistenComprobantesIngresados(empresa_id As String)
Dim ls_sql As String
Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
    ls_sql = "select count(*) from i_ventas v where exists (select 1 from CABECERA_VENTA C " & " Where c.empresa_id = " & empresa_id + _
             "AND c.tipo_comprobante_id= v.tipoComprobante " & " AND c.punto_venta_id in " & "  (select PUNTO_VENTA_ID from PUNTO_VENTA where EMPRESA_ID=" + _
             empresa_id + "   and ACTIVO = 1 and PUNTO_VENTA = v.PuntoVenta) " & "   and NRO_COMPROBANTE_DESDE = RIGHT([NumeroComprobante] ,8)) "
    Set rs = New ADODB.Recordset
    
    
    With rs
        .ActiveConnection = oCon
        .CursorType = adOpenForwardOnly
        .Source = ls_sql
        .Open
    End With
      
    ExistenComprobantesIngresados = rs(0)
    
    Set rs = Nothing
End Function


Public Function fImportarAfip(empresa_id As String, concepto_id As String)
Dim ls_SP As String
Dim result As Integer
Dim desResult As String
Dim return_value As Integer
On Error GoTo tratar

Dim oCom As ADODB.Command
Dim Param As ADODB.Parameter
    
    Set oCom = New ADODB.Command

    ls_SP = "[dbo].[SP_IMPORTAAFIP]"
    
    With oCom
        .ActiveConnection = oCon
        .CommandType = adCmdStoredProc
        .CommandText = ls_SP
        .Prepared = True
        
        .Parameters.Append .CreateParameter(, adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@ID_EMPRESA", adVarChar, adParamInput, Len(empresa_id), empresa_id)
        .Parameters.Append .CreateParameter("@CONCEPTO_ID", adVarChar, adParamInput, Len(concepto_id), concepto_id)
        .Parameters.Append .CreateParameter("@result", adInteger, adParamOutput, 8)
        .Parameters.Append .CreateParameter("@desResult", adVarChar, adParamOutput, 255)

    End With

    oCom.Execute
    
    return_value = oCom.Parameters(0)
    result = oCom.Parameters(3)
    desResult = oCom.Parameters(4)
    
    If result <> 0 Then
        MsgBox "Error Número: " & result & " - " & desResult, vbCritical, "fImportarAfip"
        return_value = result
    End If
    fImportarAfip = return_value
    
    Exit Function

tratar:
    fImportarAfip = -1
    Screen.MousePointer = vbDefault
    If oCon.Errors.Count >= 1 Then
        MsgBox oCon.Errors(0).Description, vbCritical, "Nro de Error: " & oCon.Errors(0).Number
        oCon.Errors.Clear
    Else
        MsgBox Err.Number & "-" & Err.Description, vbCritical, "fImportarAfip"
        Err.Clear
    End If
End Function
