Attribute VB_Name = "ExportacionArchivos"
Option Explicit

  
  
' --------------------------------------------------------------------------------
' \\ -- Función que exporta el recordset a un archivo de texto
' --------------------------------------------------------------------------------
Public Function Exportar_Recordset( _
    sFileName As String, _
    ls_Sql As String, _
    Optional sDelimiter As String = "", _
    Optional bPrintField As Boolean = False, Optional iFieldsMenos As Integer = 0) As Boolean
  

  
    Dim iFreeFile   As Integer
    Dim iField      As Long
    Dim i           As Long
    
    Dim obj_Field   As ADODB.Field
    Dim rs As ADODB.Recordset
  
    On Error GoTo error_handler:
    
    Set rs = New ADODB.Recordset
    
    With rs
        .ActiveConnection = oCon
        .CursorType = adOpenForwardOnly
        .Source = ls_Sql
        .Open
    End With
      
    Screen.MousePointer = vbHourglass
    ' -- Otener número de archivo disponible
    iFreeFile = FreeFile
    ' -- Crear el archivo
    Open sFileName For Output As #iFreeFile
  
    With rs
        iField = .Fields.Count - 1 - iFieldsMenos
        On Error Resume Next
        ' -- Primer registro
        .MoveFirst
        On Error GoTo error_handler
        ' -- Recorremos campo por campo y los registros de cada uno
        Do While Not .EOF
            For i = 0 To iField
                  
                ' -- Asigna el objeto Field
                Set obj_Field = .Fields(i)
                ' -- Verificar que el campo no es de ipo bunario o  un tipo no válido para grabar en el archivo
                If isValidField(obj_Field) Then
                    If i < iField Then
                        If bPrintField Then
                            ' -- Escribir el campo y el valor
                            Print #iFreeFile, obj_Field.Name & ":" & obj_Field.value & sDelimiter;
                        Else
                            ' -- Guardar solo el valor sin el campo
                            Print #iFreeFile, obj_Field.value & sDelimiter;
                        End If
                    Else
                        If bPrintField Then
                            ' -- Escribir el nombre del campo y el valor de la última columna ( Sin delimitador y sin punto y coma para añadir nueva línea )
                            Print #iFreeFile, obj_Field.Name & ": " & obj_Field.value
                        Else
                            ' -- Guardar solo el valor sin el campo
                            Print #iFreeFile, obj_Field.value
                        End If
                    End If
                End If
            Next
            ' -- Mover el cursor al siguiente registro
            .MoveNext
        Loop
    End With
      
    ' -- Cerrar el recordset
    rs.Close
    Exportar_Recordset = True
    Screen.MousePointer = vbDefault
    Close #iFreeFile
    Exit Function
error_handler:
 On Error Resume Next
 Close #iFreeFile
 rs.Close
 Screen.MousePointer = vbDefault
End Function
  
' ----------------------------------------------------------------------------------------------
' -- Si el campo es nulo ( binario, o tipo desconocido etc..) devuelve False para no añadir el dato
' ----------------------------------------------------------------------------------------------
Private Function isValidField(obj_Field As ADODB.Field) As Boolean
      
    With obj_Field
        On Error GoTo error_handler
        Select Case obj_Field.Type
            Case adBinary, adIDispatch, adIUnknown, adUserDefined
                isValidField = False
            ' -- Campo válido
            Case Else
                isValidField = True
        End Select
    End With
Exit Function
error_handler:
End Function

