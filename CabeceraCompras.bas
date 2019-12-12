Attribute VB_Name = "CabeceraCompras"
Option Explicit

Public Type tCabecera_Compra
    Empresa As tEmpresa
    tipoComprobabte As tTipoComprobante
    moneda As tMoneda
    fechaImputacion As Date
    fechaCompra As Date
    tipoDocumento As tDocumento
    PuntoDeVenta As String
    nroComprobante As String
    vendedorId As String
    razonSocialVendedor As String
    codigoOperacion As tCodigoOperacion
    tipoCambio As Double
    totalOperacion As Double
    existe As Boolean
    
End Type


Public Function insertarCabeceraCompra(cc As tCabecera_Compra) As Double
On Error GoTo errores
Dim ls_SP As String
Dim compra_id As Double
Dim oCom As ADODB.Command
Dim oRec As ADODB.Recordset


    Set oCom = New ADODB.Command
    Set oRec = New ADODB.Recordset
 
    
    ls_SP = "sp_insertCabeceraCompra"
    
    With oCom
        .ActiveConnection = oCon
        .CommandType = adCmdStoredProc
        .CommandText = ls_SP
        .Prepared = True

        .Parameters.Append .CreateParameter("@EMPRESA_ID", adDouble, adParamInput, , cc.Empresa.Identificador)
        .Parameters.Append .CreateParameter("@TIPO_COMPROBANTE_ID", adVarChar, adParamInput, 3, cc.tipoComprobabte.CODIGO)
        .Parameters.Append .CreateParameter("@TIPO_MONEDA", adVarChar, adParamInput, 3, cc.moneda.CODIGO)
        .Parameters.Append .CreateParameter("@FECHA_IMPUTACION", adDBDate, adParamInput, , cc.fechaImputacion)
        .Parameters.Append .CreateParameter("@FECHA_COMPRA", adDBDate, adParamInput, , cc.fechaCompra)
        .Parameters.Append .CreateParameter("@TIPO_DOCUMENTO_ID", adVarChar, adParamInput, 2, cc.tipoDocumento.CODIGO)
        .Parameters.Append .CreateParameter("@PUNTO_VENTA", adVarChar, adParamInput, 5, cc.PuntoDeVenta)
        .Parameters.Append .CreateParameter("@NRO_COMPROBANTE", adVarChar, adParamInput, 8, Format(cc.nroComprobante, formatoNroComprobante))
        .Parameters.Append .CreateParameter("@VENDEDOR_ID", adVarChar, adParamInput, Len(cc.vendedorId), cc.vendedorId)
        .Parameters.Append .CreateParameter("@RAZON_SOCIAL_VENDEDOR", adVarChar, adParamInput, Len(cc.razonSocialVendedor), cc.razonSocialVendedor)
        .Parameters.Append .CreateParameter("@CODIGO_OPERACION_ID", adVarChar, adParamInput, 1, cc.codigoOperacion.CODIGO)
        .Parameters.Append .CreateParameter("@TIPO_CAMBIO", adDouble, adParamInput, , cc.tipoCambio)
        .Parameters.Append .CreateParameter("@TOTAL_OPERACION", adDouble, adParamInput, , cc.totalOperacion)
        
        
    End With
    
    Set oRec = oCom.Execute
    
    insertarCabeceraCompra = oRec(0)
    
    Set oRec = Nothing
    
    Exit Function
    
errores:
    insertarCabeceraCompra = 0
    If oCon.Errors.Count >= 1 Then
        MsgBox oCon.Errors(0).Description, vbCritical, "Nro de Error: " & oCon.Errors(0).Number
        oCon.Errors.Clear
    Else
        MsgBox Err.Number & "-" & Err.Description, vbCritical, "Error: Al insertar cabecera de COMPRAS"
        Err.Clear
    End If

    Set oRec = Nothing

End Function


Public Function borrarCabeceraDetalleCompra(compra_id As Double) As Double


On Error GoTo errores
Dim ls_SP As String

Dim oCom As ADODB.Command
Dim oRec As ADODB.Recordset


    Set oCom = New ADODB.Command
    Set oRec = New ADODB.Recordset
 
    
    ls_SP = "sp_BorrarCabeceraDetalleCompras"
    
    With oCom
        .ActiveConnection = oCon
        .CommandType = adCmdStoredProc
        .CommandText = ls_SP
        .Prepared = True
        .Parameters.Append .CreateParameter("@COMPRA_ID", adDouble, adParamInput, , compra_id)
    End With

    Set oRec = oCom.Execute
    
    borrarCabeceraDetalleCompra = oRec(0)
    
    Set oRec = Nothing
    
    Exit Function

errores:
    borrarCabeceraDetalleCompra = 0
    If oCon.Errors.Count >= 1 Then
        MsgBox oCon.Errors(0).Description, vbCritical, "Nro de Error: " & oCon.Errors(0).Number
        oCon.Errors.Clear
    Else
        MsgBox Err.Number & "-" & Err.Description, vbCritical, "Error: Al Borrar la compra"
        Err.Clear
    End If

    Set oRec = Nothing
End Function


Public Function ExisteComprobanteDeCompra(cc As tCabecera_Compra) As Double

On Error GoTo errores
Dim ls_SP As String

Dim oCom As ADODB.Command
Dim oRec As ADODB.Recordset


    Set oCom = New ADODB.Command
    Set oRec = New ADODB.Recordset
 
    
    ls_SP = "sp_ExisteComprobanteDeCompra"
    
    With oCom
        .ActiveConnection = oCon
        .CommandType = adCmdStoredProc
        .CommandText = ls_SP
        .Prepared = True
        .Parameters.Append .CreateParameter("@VENDEDOR_ID", adVarChar, adParamInput, 11, cc.vendedorId)
        .Parameters.Append .CreateParameter("@TIPO_COMPROBANTE_ID", adVarChar, adParamInput, 3, cc.tipoComprobabte.CODIGO)
        .Parameters.Append .CreateParameter("@PUNTO_VENTA", adVarChar, adParamInput, 5, cc.PuntoDeVenta)
        .Parameters.Append .CreateParameter("@NRO_COMPROBANTE", adVarChar, adParamInput, 8, cc.nroComprobante)
        .Parameters.Append .CreateParameter("@EMPRESA_ID", adDouble, adParamInput, , cc.Empresa.Identificador)
    End With

    Set oRec = oCom.Execute
    
    ExisteComprobanteDeCompra = oRec(0)
    
    Set oRec = Nothing
    
    Exit Function

errores:
    ExisteComprobanteDeCompra = 0
    If oCon.Errors.Count >= 1 Then
        MsgBox oCon.Errors(0).Description, vbCritical, "Nro de Error: " & oCon.Errors(0).Number
        oCon.Errors.Clear
    Else
        MsgBox Err.Number & "-" & Err.Description, vbCritical, "Error: Al controlar Comprobantes de Compras"
        Err.Clear
    End If

    Set oRec = Nothing
End Function
