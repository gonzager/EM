Attribute VB_Name = "Concepto"
Option Explicit


Public Type tConceptos
    codigo As String
    descripcion As String
    defecto As Boolean
    visible As Boolean
    
End Type

Public Function insertConcepto(ltConcepto As tConceptos, CompraVenta As String) As Integer
On Error GoTo errores
Dim ls_SP As String

Dim oCom As ADODB.Command
Dim oRec As ADODB.Recordset


    Set oCom = New ADODB.Command
    Set oRec = New ADODB.Recordset
 
    
    If CompraVenta = "COMPRAS" Then
        ls_SP = "sp_insertConceptoCompra"
    ElseIf CompraVenta = "VENTAS" Then
        ls_SP = "sp_insertConceptoVenta"
    End If
       
    With oCom
        .ActiveConnection = oCon
        .CommandType = adCmdStoredProc
        .CommandText = ls_SP
        .Prepared = True
        .Parameters.Append .CreateParameter("@CODIGO", adVarChar, adParamInput, 3, ltConcepto.codigo)
        .Parameters.Append .CreateParameter("@DESCRIPCION ", adVarChar, adParamInput, Len(ltConcepto.descripcion), ltConcepto.descripcion)
        .Parameters.Append .CreateParameter("@DEFECTO", adBoolean, adParamInput, , ltConcepto.defecto)
        .Parameters.Append .CreateParameter("@VISIBLE", adBoolean, adParamInput, , ltConcepto.visible)

    End With

    Set oRec = oCom.Execute
    
    insertConcepto = oRec(0)
    
    Set oRec = Nothing
    
    Exit Function

errores:
    insertConcepto = 0
    If oCon.Errors.Count >= 1 Then
        MsgBox oCon.Errors(0).Description, vbCritical, "Nro de Error: " & oCon.Errors(0).Number
        oCon.Errors.Clear
    Else
        MsgBox Err.Number & "-" & Err.Description, vbCritical, "Error: Al insertar concepto."
        Err.Clear
    End If

    Set oRec = Nothing
End Function



Public Function updateConcepto(ltConcepto As tConceptos, CompraVenta As String) As Integer
On Error GoTo errores
Dim ls_SP As String

Dim oCom As ADODB.Command
Dim oRec As ADODB.Recordset


    Set oCom = New ADODB.Command
    Set oRec = New ADODB.Recordset
 
    If CompraVenta = "COMPRAS" Then
        ls_SP = "sp_UPDATEConceptoCompra"
    ElseIf CompraVenta = "VENTAS" Then
         ls_SP = "sp_UPDATEConceptoVenta"
    End If
    
    With oCom
        .ActiveConnection = oCon
        .CommandType = adCmdStoredProc
        .CommandText = ls_SP
        .Prepared = True
        .Parameters.Append .CreateParameter("@CODIGO", adVarChar, adParamInput, 3, ltConcepto.codigo)
        .Parameters.Append .CreateParameter("@DESCRIPCION ", adVarChar, adParamInput, Len(ltConcepto.descripcion), ltConcepto.descripcion)
        .Parameters.Append .CreateParameter("@DEFECTO", adBoolean, adParamInput, , ltConcepto.defecto)
        .Parameters.Append .CreateParameter("@VISIBLE", adBoolean, adParamInput, , ltConcepto.visible)

    End With

    Set oRec = oCom.Execute
    
    updateConcepto = oRec(0)
    
    Set oRec = Nothing
    
    Exit Function

errores:
    updateConcepto = 0
    If oCon.Errors.Count >= 1 Then
        MsgBox oCon.Errors(0).Description, vbCritical, "Nro de Error: " & oCon.Errors(0).Number
        oCon.Errors.Clear
    Else
        MsgBox Err.Number & "-" & Err.Description, vbCritical, "Error: Al modificar concepto."
        Err.Clear
    End If

    Set oRec = Nothing
End Function


Public Function fMaxCodigo(CompraVenta As String) As String

    Dim ls_sql As String
    Dim oRec As ADODB.Recordset

    Set oRec = New ADODB.Recordset

    
    If CompraVenta = "COMPRAS" Then
        ls_sql = "SELECT ISNULL(MAX(CODIGO),0) + 1 FROM CONCEPTOCPA "
    ElseIf CompraVenta = "VENTAS" Then
        ls_sql = "SELECT ISNULL(MAX(CODIGO),0) + 1 FROM CONCEPTO WHERE CODIGO <> '999'"
    End If
    With oRec
        .ActiveConnection = oCon
        .CursorType = adOpenStatic
        .CursorLocation = adUseClient
        .Source = ls_sql
        .Open
    End With
    
    fMaxCodigo = Format(oRec(0), "000")

End Function

