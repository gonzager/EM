Attribute VB_Name = "DetalleVentas"
Option Explicit


Public Type tDetalle_Venta
    venta_id As Double
    concepto As tConcepto
    alicuota As tAlicuota
    neto_gravado As Currency
    exento As Currency
    iva As Currency
    total As Currency
    existe As Boolean
End Type


Public Function insertarDetalleVenta(dv As tDetalle_Venta) As Boolean

Dim ls_SP As String
Dim oCom As ADODB.Command
Dim oRec As ADODB.Recordset
On Error GoTo errores

    Set oCom = New ADODB.Command
    Set oRec = New ADODB.Recordset


    ls_SP = "sp_insertDetalleVenta"

    With oCom
        .ActiveConnection = oCon
        .CommandType = adCmdStoredProc
        .CommandText = ls_SP
        .Prepared = True
        .Parameters.Append .CreateParameter("@CABECERA_VENTA_ID", adDouble, adParamInput, , dv.venta_id)
        .Parameters.Append .CreateParameter("@CONCEPTO_ID", adVarChar, adParamInput, 3, dv.concepto.codigo)
        .Parameters.Append .CreateParameter("@ALICUOTA_ID", adVarChar, adParamInput, 4, dv.alicuota.alicuta)
        .Parameters.Append .CreateParameter("@NETO_GRAVADO", adCurrency, adParamInput, , dv.neto_gravado)
        .Parameters.Append .CreateParameter("@EXENTO", adCurrency, adParamInput, , dv.exento)
        .Parameters.Append .CreateParameter("@IVA", adCurrency, adParamInput, , dv.iva)
        .Parameters.Append .CreateParameter("@TOTAL", adCurrency, adParamInput, , dv.total)
        
    End With
    
    insertarDetalleVenta = False
    Set oRec = oCom.Execute
    If oRec(0) = 1 Then insertarDetalleVenta = True
    Exit Function
    
errores:
    If oCon.Errors.Count > 0 Then
        MsgBox oCon.Errors(0).Description, vbCritical, "Nro de Error: " & oCon.Errors(0).Number
        oCon.Errors.Clear
    Else
        MsgBox Err.Description, vbCritical, "Error: AL INSERTAR EL DETALLE DE VENTAS"
    End If
    
End Function
