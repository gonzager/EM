Attribute VB_Name = "AsoProFarma"
Option Explicit
  
  
Public Function fBuklAsoProFarma(ByVal Datos As String, ByVal Formato As String) As Integer

Dim ls_SP As String
Dim result As Integer
Dim desResult As String
Dim return_value As Integer
On Error GoTo tratar

Dim Param As ADODB.Parameter


    Dim oCom As ADODB.Command
    Set oCom = New ADODB.Command

    ls_SP = "sp_crearTMPAsoProFarma"
    
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
        MsgBox "Error Número: " & result & " - " & desResult, vbCritical, "fBuklAsoProFarma"
        return_value = result
    End If
    fBuklAsoProFarma = return_value
    
    Exit Function

tratar:
    fBuklAsoProFarma = -1
    If oCon.Errors.Count >= 1 Then
        MsgBox oCon.Errors(0).Description, vbCritical, "Nro de Error: " & oCon.Errors(0).Number
        oCon.Errors.Clear
    Else
        MsgBox Err.Number & "-" & Err.Description, vbCritical, "fBuklAsoProFarma"
        Err.Clear
    End If

End Function

Public Function cantidadTmpAsoprofarma() As Integer

Dim ls_sql As String
Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
    ls_sql = "select count(*) from TMP_ASOPROFARMA"

    Set rs = New ADODB.Recordset
    
    
    With rs
        .ActiveConnection = oCon
        .CursorType = adOpenForwardOnly
        .Source = ls_sql
        .Open
    End With
      
    

    cantidadTmpAsoprofarma = rs(0)
    
    Set rs = Nothing

End Function


Public Function fImportarAsoProForma() As Integer

Dim ls_SP As String
Dim result As Integer
Dim desResult As String
Dim return_value As Integer
On Error GoTo tratar

Dim Param As ADODB.Parameter


    Dim oCom As ADODB.Command
    Set oCom = New ADODB.Command

    ls_SP = "SP_IMPORTAASOPROFORMA"
    
    With oCom
        .ActiveConnection = oCon
        .CommandType = adCmdStoredProc
        .CommandText = ls_SP
        .Prepared = True
        
        .Parameters.Append .CreateParameter(, adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@RESULT", adInteger, adParamOutput, 8)
        .Parameters.Append .CreateParameter("@DESRESULT", adVarChar, adParamOutput, 255)

    End With

    oCom.Execute
    
    return_value = oCom.Parameters(0)
    result = oCom.Parameters(1)
    desResult = oCom.Parameters(2)
    
    If result <> 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "Error Número: " & result & " - " & desResult, vbCritical, "fImportarAsoProForma"
        return_value = result
    End If
    fImportarAsoProForma = return_value
    
    Exit Function

tratar:
    fImportarAsoProForma = -1
    If oCon.Errors.Count >= 1 Then
        Screen.MousePointer = vbDefault
        MsgBox oCon.Errors(0).Description, vbCritical, "Nro de Error: " & oCon.Errors(0).Number
        oCon.Errors.Clear
    Else
        MsgBox Err.Number & "-" & Err.Description, vbCritical, "fImportarAsoProForma"
        Err.Clear
    End If
End Function

