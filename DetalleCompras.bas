Attribute VB_Name = "DetalleCompras"
Option Explicit


Public Type tDetalle_compra
    compra_id As Double
    concepto As tConcepto
    alicuota As tAlicuota
    neto_gravado As Currency
    exento As Currency
    iva As Currency
    total As Currency
    existe As Boolean
End Type


Public Function insertarDetalleCompra(dc As tDetalle_compra) As Boolean

Dim ls_SP As String
Dim oCom As ADODB.Command
Dim oRec As ADODB.Recordset
On Error GoTo errores

    Set oCom = New ADODB.Command
    Set oRec = New ADODB.Recordset


    ls_SP = "sp_insertDetalleCompra"

    With oCom
        .ActiveConnection = oCon
        .CommandType = adCmdStoredProc
        .CommandText = ls_SP
        .Prepared = True
        .Parameters.Append .CreateParameter("@CABECERA_COMPRA_ID", adDouble, adParamInput, , dc.compra_id)
        .Parameters.Append .CreateParameter("@CONCEPTO_ID", adVarChar, adParamInput, 3, dc.concepto.codigo)
        .Parameters.Append .CreateParameter("@ALICUOTA_ID", adVarChar, adParamInput, 4, dc.alicuota.alicuta)
        .Parameters.Append .CreateParameter("@NETO_GRAVADO", adCurrency, adParamInput, , dc.neto_gravado)
        .Parameters.Append .CreateParameter("@EXENTO", adCurrency, adParamInput, , dc.exento)
        .Parameters.Append .CreateParameter("@IVA", adCurrency, adParamInput, , dc.iva)
        .Parameters.Append .CreateParameter("@TOTAL", adCurrency, adParamInput, , dc.total)
        
    End With
    
    insertarDetalleCompra = False
    Set oRec = oCom.Execute
    If oRec(0) = 1 Then insertarDetalleCompra = True
    Exit Function
    
errores:
    If oCon.Errors.Count > 0 Then
        MsgBox oCon.Errors(0).Description, vbCritical, "Nro de Error: " & oCon.Errors(0).Number
        oCon.Errors.Clear
    Else
        MsgBox Err.Description, vbCritical, "Error: AL INSERTAR EL DETALLE DE COMPRA"
    End If
    
End Function

