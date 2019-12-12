Attribute VB_Name = "Empresa"
Option Explicit

Public Type tEmpresa
    Identificador As Double
    razonSocial As String
    domicilio As String
    vendedor As Boolean
    comprador As Boolean
    Activo As Boolean
    Monotributista As Boolean
    existe As Boolean
End Type


Public Function updateEmpresa(ltEmpresa As tEmpresa) As Integer
On Error GoTo errores
Dim ls_SP As String

Dim oCom As ADODB.Command
Dim oRec As ADODB.Recordset


    Set oCom = New ADODB.Command
    Set oRec = New ADODB.Recordset
 
    
    ls_SP = "sp_updateEmpresa"
    
    With oCom
        .ActiveConnection = oCon
        .CommandType = adCmdStoredProc
        .CommandText = ls_SP
        .Prepared = True
        .Parameters.Append .CreateParameter("@IDENTIFICADOR", adDouble, adParamInput, , ltEmpresa.Identificador)
        .Parameters.Append .CreateParameter("@RAZONSOCIAL", adVarChar, adParamInput, Len(ltEmpresa.razonSocial), ltEmpresa.razonSocial)
        .Parameters.Append .CreateParameter("@VENDEDOR", adBoolean, adParamInput, , ltEmpresa.vendedor)
        .Parameters.Append .CreateParameter("@COMPRADOR", adBoolean, adParamInput, , ltEmpresa.comprador)
        .Parameters.Append .CreateParameter("@DOMICILIO", adVarChar, adParamInput, Len(ltEmpresa.domicilio), ltEmpresa.domicilio)
        .Parameters.Append .CreateParameter("@ACTIVO", adBoolean, adParamInput, , ltEmpresa.Activo)
        .Parameters.Append .CreateParameter("@MONOTRIBUTISTA", adBoolean, adParamInput, , ltEmpresa.Monotributista)
    End With

    Set oRec = oCom.Execute
    
    updateEmpresa = oRec(0)
    
    Set oRec = Nothing
    
    Exit Function

errores:
    updateEmpresa = 0
    If oCon.Errors.Count >= 1 Then
        MsgBox oCon.Errors(0).Description, vbCritical, "Nro de Error: " & oCon.Errors(0).Number
        oCon.Errors.Clear
    Else
        MsgBox Err.Number & "-" & Err.Description, vbCritical, "Error: Al insertar cabecera de ventas"
        Err.Clear
    End If

    Set oRec = Nothing
End Function


Public Function insertEmpresa(ltEmpresa As tEmpresa) As Integer
On Error GoTo errores
Dim ls_SP As String

Dim oCom As ADODB.Command
Dim oRec As ADODB.Recordset


    Set oCom = New ADODB.Command
    Set oRec = New ADODB.Recordset
 
    
    ls_SP = "sp_insertEmpresa"
    
    With oCom
        .ActiveConnection = oCon
        .CommandType = adCmdStoredProc
        .CommandText = ls_SP
        .Prepared = True
        .Parameters.Append .CreateParameter("@IDENTIFICADOR", adDouble, adParamInput, , ltEmpresa.Identificador)
        .Parameters.Append .CreateParameter("@RAZONSOCIAL", adVarChar, adParamInput, Len(ltEmpresa.razonSocial), ltEmpresa.razonSocial)
        .Parameters.Append .CreateParameter("@VENDEDOR", adBoolean, adParamInput, , ltEmpresa.vendedor)
        .Parameters.Append .CreateParameter("@COMPRADOR", adBoolean, adParamInput, , ltEmpresa.comprador)
        .Parameters.Append .CreateParameter("@DOMICILIO", adVarChar, adParamInput, Len(ltEmpresa.domicilio), ltEmpresa.domicilio)
        .Parameters.Append .CreateParameter("@ACTIVO", adBoolean, adParamInput, , ltEmpresa.Activo)
        .Parameters.Append .CreateParameter("@MONOTRIBUTISTA", adBoolean, adParamInput, , ltEmpresa.Monotributista)
    End With

    Set oRec = oCom.Execute
    
    insertEmpresa = oRec(0)
    
    Set oRec = Nothing
    
    Exit Function

errores:
    insertEmpresa = 0
    If oCon.Errors.Count >= 1 Then
        MsgBox oCon.Errors(0).Description, vbCritical, "Nro de Error: " & oCon.Errors(0).Number
        oCon.Errors.Clear
    Else
        MsgBox Err.Number & "-" & Err.Description, vbCritical, "Error: Al insertar cabecera de ventas"
        Err.Clear
    End If

    Set oRec = Nothing
End Function
