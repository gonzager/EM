Attribute VB_Name = "PuntoDeVenta"
Option Explicit

Public Type tPuntoDeVenta
    puntoVentaId As Integer
    empresa_id As Double
    PuntoDeVenta As String
    Activo As Boolean
    existe As Boolean
End Type

Public Function recuperarPuntosDeVenta(Identificador As Double, ByRef lb_tiene As Boolean, Optional Filtro As String = "T") As tPuntoDeVenta()
    Dim ls_Sql As String
    Dim oRec As ADODB.Recordset
    Dim i As Integer
    Set oRec = New ADODB.Recordset
    Dim ltPuntoDeVenta() As tPuntoDeVenta
    
    lb_tiene = False
    ls_Sql = "SELECT PUNTO_VENTA FROM PUNTO_VENTA WHERE EMPRESA_ID=" & Identificador
    If Filtro = "A" Then
        ls_Sql = ls_Sql + " AND ACTIVO = 1 "
    ElseIf Filtro = "I" Then
        ls_Sql = ls_Sql + " AND ACTIVO = 0 "
    End If
    
    
    
    With oRec
        .ActiveConnection = oCon
        .CursorType = adOpenStatic
        .CursorLocation = adUseClient
        .Source = ls_Sql
        .Open
    End With
    
    If oRec.RecordCount > 0 Then lb_tiene = True
    
    
    i = 0
    Do While Not oRec.EOF
        ReDim Preserve ltPuntoDeVenta(i)
        ltPuntoDeVenta(i).empresa_id = Identificador
        ltPuntoDeVenta(i).PuntoDeVenta = oRec!PUNTO_VENTA
        i = i + 1
        oRec.MoveNext
    Loop
   
    oRec.Close
    Set oRec = Nothing
    recuperarPuntosDeVenta = ltPuntoDeVenta
End Function


Public Function recuperarIDPuntoDeVenta(Identificador As Double, PuntoDeVenta As String) As Integer
    Dim ls_Sql As String
    Dim oRec As ADODB.Recordset
    Set oRec = New ADODB.Recordset
    
    ls_Sql = "SELECT PUNTO_VENTA_ID FROM PUNTO_VENTA WHERE EMPRESA_ID=" & Identificador & " AND PUNTO_VENTA = '" & PuntoDeVenta & "' "
    
    With oRec
        .ActiveConnection = oCon
        .CursorType = adOpenStatic
        .CursorLocation = adUseClient
        .Source = ls_Sql
        .Open
    End With
    
    recuperarIDPuntoDeVenta = 0
    If oRec.RecordCount > 0 Then
        recuperarIDPuntoDeVenta = oRec!PUNTO_VENTA_ID
    End If

   
    oRec.Close
    Set oRec = Nothing
    
End Function

Public Function recuperarRegistroPuntoVtas(Identificador As Double, ByRef lb_tiene As Boolean, Optional Filtro As String = "T") As tPuntoDeVenta()
    Dim ls_Sql As String
    Dim oRec As ADODB.Recordset
    Dim i As Integer
    Set oRec = New ADODB.Recordset
    Dim ltPuntoDeVenta() As tPuntoDeVenta
    
    lb_tiene = False
    ls_Sql = "SELECT PUNTO_VENTA_ID, EMPRESA_ID, PUNTO_VENTA, ACTIVO FROM PUNTO_VENTA WHERE EMPRESA_ID=" & Identificador
    If Filtro = "A" Then
        ls_Sql = ls_Sql + " AND ACTIVO = 1 "
    ElseIf Filtro = "I" Then
        ls_Sql = ls_Sql + " AND ACTIVO = 0 "
    End If
    
    ls_Sql = ls_Sql & " ORDER BY PUNTO_VENTA"
    
    
    With oRec
        .ActiveConnection = oCon
        .CursorType = adOpenStatic
        .CursorLocation = adUseClient
        .Source = ls_Sql
        .Open
    End With
    
    If oRec.RecordCount > 0 Then lb_tiene = True
    
    
    i = 0
    Do While Not oRec.EOF
        ReDim Preserve ltPuntoDeVenta(i)
        ltPuntoDeVenta(i).puntoVentaId = oRec!PUNTO_VENTA_ID
        ltPuntoDeVenta(i).empresa_id = oRec!empresa_id
        ltPuntoDeVenta(i).PuntoDeVenta = oRec!PUNTO_VENTA
        ltPuntoDeVenta(i).Activo = oRec!Activo
        ltPuntoDeVenta(i).existe = True
        i = i + 1
        oRec.MoveNext
    Loop
   
    oRec.Close
    Set oRec = Nothing
    recuperarRegistroPuntoVtas = ltPuntoDeVenta
End Function


Public Function updatePuntoDeVenta(ltPuntoDeVenta As tPuntoDeVenta) As Integer
On Error GoTo errores
Dim ls_SP As String

Dim oCom As ADODB.Command
Dim oRec As ADODB.Recordset


    Set oCom = New ADODB.Command
    Set oRec = New ADODB.Recordset
 
    
    ls_SP = "sp_updatePuntoDeVenta"
    
    With oCom
        .ActiveConnection = oCon
        .CommandType = adCmdStoredProc
        .CommandText = ls_SP
        .Prepared = True
        .Parameters.Append .CreateParameter("@PUNTO_VENTA_ID", adInteger, adParamInput, , ltPuntoDeVenta.puntoVentaId)
        .Parameters.Append .CreateParameter("@EMPRESA_ID", adDouble, adParamInput, , ltPuntoDeVenta.empresa_id)
        .Parameters.Append .CreateParameter("@PUNTO_VENTA", adVarChar, adParamInput, Len(ltPuntoDeVenta.PuntoDeVenta), ltPuntoDeVenta.PuntoDeVenta)
        .Parameters.Append .CreateParameter("@ACTIVO", adBoolean, adParamInput, , ltPuntoDeVenta.Activo)

    End With

    Set oRec = oCom.Execute
    
    updatePuntoDeVenta = oRec(0)
    
    Set oRec = Nothing
    
    Exit Function

errores:
    updatePuntoDeVenta = 0
    If oCon.Errors.Count >= 1 Then
        MsgBox oCon.Errors(0).Description, vbCritical, "Nro de Error: " & oCon.Errors(0).Number
        oCon.Errors.Clear
    Else
        MsgBox Err.Number & "-" & Err.Description, vbCritical, "Error: Al insertar cabecera de ventas"
        Err.Clear
    End If

    Set oRec = Nothing
End Function


Public Function insertPuntoDeVenta(ltPuntoDeVenta As tPuntoDeVenta) As Integer
On Error GoTo errores
Dim ls_SP As String

Dim oCom As ADODB.Command
Dim oRec As ADODB.Recordset


    Set oCom = New ADODB.Command
    Set oRec = New ADODB.Recordset
 
    
    ls_SP = "sp_insertPuntoDeVenta"
    
    With oCom
        .ActiveConnection = oCon
        .CommandType = adCmdStoredProc
        .CommandText = ls_SP
        .Prepared = True
        .Parameters.Append .CreateParameter("@EMPRESA_ID", adDouble, adParamInput, , ltPuntoDeVenta.empresa_id)
        .Parameters.Append .CreateParameter("@PUNTO_VENTA", adVarChar, adParamInput, Len(ltPuntoDeVenta.PuntoDeVenta), ltPuntoDeVenta.PuntoDeVenta)
        .Parameters.Append .CreateParameter("@ACTIVO", adBoolean, adParamInput, , ltPuntoDeVenta.Activo)

    End With

    Set oRec = oCom.Execute
    
    insertPuntoDeVenta = oRec(0)
    
    Set oRec = Nothing
    
    Exit Function

errores:
    insertPuntoDeVenta = 0
    If oCon.Errors.Count >= 1 Then
        MsgBox oCon.Errors(0).Description, vbCritical, "Nro de Error: " & oCon.Errors(0).Number
        oCon.Errors.Clear
    Else
        MsgBox Err.Number & "-" & Err.Description, vbCritical, "Error: Al insertar cabecera de ventas"
        Err.Clear
    End If

    Set oRec = Nothing
End Function
