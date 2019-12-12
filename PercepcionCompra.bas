Attribute VB_Name = "PercepcionCompra"
Option Explicit

Public Type tPercepcion_compra
    compra_id As Double
    tuTipoPercepcion As tTipoPercepcion
    tuJurisdiccion As tJurisdiccion
    totalPercepcion As Currency
    existe As Boolean
End Type


Public Function insertarPercepcionCompra(pc As tPercepcion_compra) As Boolean

Dim ls_SP As String
Dim oCom As ADODB.Command
Dim oRec As ADODB.Recordset
On Error GoTo errores

    Set oCom = New ADODB.Command
    Set oRec = New ADODB.Recordset


    ls_SP = "sp_insertPercepcionCompra"

    With oCom
        .ActiveConnection = oCon
        .CommandType = adCmdStoredProc
        .CommandText = ls_SP
        .Prepared = True
        .Parameters.Append .CreateParameter("@CABECERA_COMPRA_ID", adDouble, adParamInput, , pc.compra_id)
        .Parameters.Append .CreateParameter("@TIPO_PERCEPCION_ID", adVarChar, adParamInput, 2, pc.tuTipoPercepcion.percipcion)
        .Parameters.Append .CreateParameter("@JURISDICCION_ID", adVarChar, adParamInput, 2, pc.tuJurisdiccion.jurisdiccion)
        .Parameters.Append .CreateParameter("@IMPORTE_PERCECPION", adCurrency, adParamInput, , pc.totalPercepcion)
        
    End With
    
    insertarPercepcionCompra = False
    Set oRec = oCom.Execute
    If oRec(0) = 1 Then insertarPercepcionCompra = True
    Exit Function
    
errores:
    If oCon.Errors.Count > 0 Then
        MsgBox oCon.Errors(0).Description, vbCritical, "Nro de Error: " & oCon.Errors(0).Number
        oCon.Errors.Clear
    Else
        MsgBox Err.Description, vbCritical, "Error: AL INSERTAR EL PERCEPCION DE COMPRA"
    End If
    
End Function

