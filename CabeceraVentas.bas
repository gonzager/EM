Attribute VB_Name = "CabeceraVentas"
Option Explicit

Public Type tCabecera_Venta
    Empresa As tEmpresa
    tipoComprobabte As tTipoComprobante
    moneda As tMoneda
    PuntoDeVenta As tPuntoDeVenta
    nroComprobanteDesde As String
    nroComprobanteHasta As String
    fechaVenta As Date
    tipoDocumento As tDocumento
    compradorId As String
    razonSocialComprador As String
    codigoOperacionVenta As tCodigoOperacion
    existe As Boolean
    tipoCambio As Double
End Type


Public Function insertarCabeceraVenta(cv As tCabecera_Venta) As Double
On Error GoTo errores
Dim ls_SP As String
Dim venta_id As Double
Dim oCom As ADODB.Command
Dim oRec As ADODB.Recordset


    Set oCom = New ADODB.Command
    Set oRec = New ADODB.Recordset
 
    
    ls_SP = "sp_insertCabeceraVenta"
    
    With oCom
        .ActiveConnection = oCon
        .CommandType = adCmdStoredProc
        .CommandText = ls_SP
        .Prepared = True

        .Parameters.Append .CreateParameter("@EMPRESA_ID", adDouble, adParamInput, , cv.Empresa.Identificador)
        .Parameters.Append .CreateParameter("@TIPO_COMPROBANTE_ID", adVarChar, adParamInput, 3, cv.tipoComprobabte.codigo)
        .Parameters.Append .CreateParameter("@TIPO_MONEDA", adVarChar, adParamInput, 3, cv.moneda.codigo)
        .Parameters.Append .CreateParameter("@PUNTO_VENTA_ID", adInteger, adParamInput, , cv.PuntoDeVenta.puntoVentaId)
        .Parameters.Append .CreateParameter("@NRO_COMPROBANTE_DESDE", adVarChar, adParamInput, 8, Format(cv.nroComprobanteDesde, formatoNroComprobante))
        .Parameters.Append .CreateParameter("@NRO_COMPROBANTE_HASTA", adVarChar, adParamInput, 8, Format(cv.nroComprobanteHasta, formatoNroComprobante))
        .Parameters.Append .CreateParameter("@FECHA_VENTA", adDBDate, adParamInput, , cv.fechaVenta)
        .Parameters.Append .CreateParameter("@TIPO_DOCUMENTO_ID", adVarChar, adParamInput, 2, cv.tipoDocumento.codigo)
        .Parameters.Append .CreateParameter("@COMPRADOR_ID", adVarChar, adParamInput, Len(cv.compradorId), cv.compradorId)
        .Parameters.Append .CreateParameter("@RAZON_SOCIAL_COMPRADOR", adVarChar, adParamInput, Len(cv.razonSocialComprador), cv.razonSocialComprador)
        .Parameters.Append .CreateParameter("@CODIGO_OPERACION_ID", adVarChar, adParamInput, 1, cv.codigoOperacionVenta.codigo)
        .Parameters.Append .CreateParameter("@TIPO_CAMBIO", adDouble, adParamInput, , cv.tipoCambio)
        
    End With
    
    Set oRec = oCom.Execute
    
    insertarCabeceraVenta = oRec(0)
    
    Set oRec = Nothing
    
    Exit Function
    
errores:
    insertarCabeceraVenta = 0
    If oCon.Errors.Count >= 1 Then
        MsgBox oCon.Errors(0).Description, vbCritical, "Nro de Error: " & oCon.Errors(0).Number
        oCon.Errors.Clear
    Else
        MsgBox Err.Number & "-" & Err.Description, vbCritical, "Error: Al insertar cabecera de ventas"
        Err.Clear
    End If

    Set oRec = Nothing

End Function


Public Function borrarCabeceraDetalle(venta_id As Double) As Double


On Error GoTo errores
Dim ls_SP As String

Dim oCom As ADODB.Command
Dim oRec As ADODB.Recordset


    Set oCom = New ADODB.Command
    Set oRec = New ADODB.Recordset
 
    
    ls_SP = "sp_BorrarCabeceraDetalleVentas"
    
    With oCom
        .ActiveConnection = oCon
        .CommandType = adCmdStoredProc
        .CommandText = ls_SP
        .Prepared = True
        .Parameters.Append .CreateParameter("@VENTA_ID", adDouble, adParamInput, , venta_id)
    End With

    Set oRec = oCom.Execute
    
    borrarCabeceraDetalle = oRec(0)
    
    Set oRec = Nothing
    
    Exit Function

errores:
    borrarCabeceraDetalle = 0
    If oCon.Errors.Count >= 1 Then
        MsgBox oCon.Errors(0).Description, vbCritical, "Nro de Error: " & oCon.Errors(0).Number
        oCon.Errors.Clear
    Else
        MsgBox Err.Number & "-" & Err.Description, vbCritical, "Error: Al insertar cabecera de ventas"
        Err.Clear
    End If

    Set oRec = Nothing
End Function
