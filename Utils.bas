Attribute VB_Name = "Utils"
Option Explicit

  
'Constantes
Const LOCALE_SDECIMAL = &HE
  
'Funciones api
Private Declare Function GetUserDefaultLangID Lib "kernel32" () As Integer
Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
  
Private Declare Function GetLocaleInfo _
    Lib "kernel32" _
    Alias "GetLocaleInfoA" ( _
        ByVal Locale As Long, _
        ByVal LCType As Long, _
        ByVal lpLCData As String, _
        ByVal cchData As Long) As Long
        
        'Función api que recupera un valor-dato de un archivo Ini
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long
  
'Función api que Escribe un valor - dato en un archivo Ini
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpString As String, _
    ByVal lpFileName As String) As Long
Public Const SERVERFILESHARE As String = "CONFIGUAR SERVERFILESHARE EN EL INI"
Public Const AFIPSERVERFILESHARE As String = "CONFIGUAR SERVERFILESHARE EN EL INI"
  
' Estructura SHFILEOPSTRUCT o para usar con el Api
Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String
End Type
  
'Declaración Api SHFileOperation
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" _
                                                (lpFileOp As SHFILEOPSTRUCT) As Long
  
'Constantes
Private Const FO_COPY = &H2
Private Const FOF_ALLOWUNDO = &H40


'Declaración del Api GetFileTitle
Private Declare Function GetFileTitle _
    Lib "comdlg32.dll" _
    Alias "GetFileTitleA" ( _
        ByVal lpszFile As String, _
        ByVal lpszTitle As String, _
        ByVal cbBuf As Integer) As Integer
  
  
Public Function Obtener_Nombre_Archivo(p As String)
  
      
    Dim Buffer As String
    'Buffer de caracteres
    Buffer = String(255, 0)
    'Llamada a GetFileTitle, pasandole el path, el buffer y el tamaño
    GetFileTitle p, Buffer, Len(Buffer)
      
    'Retornamos el nombre eliminando los espacios nulos
    Obtener_Nombre_Archivo = Left$(Buffer, InStr(1, Buffer, Chr$(0)) - 1)
      
  
End Function
  
  
' Subrutina que copia el archivo
Public Sub Copiar_Archivo(ByVal Origen As String, ByVal Destino As String)
  
Dim t_Op As SHFILEOPSTRUCT
  
    With t_Op
        .hwnd = 0
        .wFunc = FO_COPY
        .pFrom = Origen & vbNullChar & vbNullChar
        .pTo = Destino & vbNullChar & vbNullChar
        .fFlags = FOF_ALLOWUNDO
    End With
  
    ' Se ejecuta la función Api pasandole la estructura
    SHFileOperation t_Op
      
      
End Sub
  
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  
'Función quie retorna la cadena con el resultado
'************************************************
Public Function Obtener_Separador_Decimal() As String
  
    Dim Buffer As String, ret As Long
  
    Buffer = String(255, " ")
  
    'Ejecutamos el Api. En el Buffer obtenermos el separador
    ret = GetLocaleInfo(GetUserDefaultLCID, LOCALE_SDECIMAL, Buffer, 255)
  
    'Quitamos los espacios nulos
    Obtener_Separador_Decimal = Trim$(Replace$(Buffer, Chr(0), ""))
  
End Function

'Recibe la ruta del archivo, la clave a leer y _
 el valor por defecto en caso de que la Key no exista
Public Function Leer_Ini(Path_INI As String, Key As String, Default As Variant) As String
  
Dim bufer As String * 256
Dim Len_Value As Long
  
        Len_Value = GetPrivateProfileString(APPLICATION, _
                                         Key, _
                                         Default, _
                                         bufer, _
                                         Len(bufer), _
                                         Path_INI)
          
        Leer_Ini = Left$(bufer, Len_Value)
  
End Function

Public Function Grabar_Ini(Path_INI As String, Key As String, Valor As Variant) As String
  
    WritePrivateProfileString APPLICATION, _
                                         Key, _
                                         Valor, _
                                         Path_INI
  
End Function
Public Sub cargarCombo(cmb_generico As ComboBox, ls_sql As String)

Dim oRec As ADODB.Recordset
Dim i As Integer
Dim iDefault As Integer

    Set oRec = New ADODB.Recordset
    
    With oRec
        .ActiveConnection = oCon
        .CursorType = adOpenForwardOnly
        .Source = ls_sql
        .Open
    End With
    
    cmb_generico.Clear
    i = 0
    iDefault = 0
    Do While Not oRec.EOF
        cmb_generico.AddItem oRec(0) & "-" & IIf(separadorDecimal = ".", oRec(1), Replace(oRec(1), ".", separadorDecimal)), i
        If oRec!defecto Then
            iDefault = i
        End If
        i = i + 1
        oRec.MoveNext
    Loop
    
    If cmb_generico.ListCount > 0 Then cmb_generico.ListIndex = iDefault
    
    oRec.Close
    Set oRec = Nothing
    
End Sub

Public Sub setLisIndexCombo(cmb_generico As ComboBox, codigo As String)
    Dim ls_tmp() As String
    Dim i As Integer
    
    For i = 0 To cmb_generico.ListCount - 1
        cmb_generico.ListIndex = i
        ls_tmp = separarCodigoDescripcion(cmb_generico.text)
        If ls_tmp(0) = codigo Then Exit For
    Next
End Sub


Public Function recuperarEmpresaPorIdentificador(Identificador As Double) As tEmpresa
    Dim ltEmpresa As tEmpresa
    Dim ls_sql As String
    Dim oRec As ADODB.Recordset
    
    Set oRec = New ADODB.Recordset
    ltEmpresa.existe = False
    
    ls_sql = "SELECT IDENTIFICADOR, RAZONSOCIAL, MONOTRIBUTISTA FROM EMPRESA WHERE IDENTIFICADOR = " & Identificador & " and vendedor = 1"
    
    With oRec
        .ActiveConnection = oCon
        .CursorType = adOpenStatic
        .CursorLocation = adUseClient
        .Source = ls_sql
        .Open
    End With
    
    If oRec.RecordCount = 1 Then
        ltEmpresa.existe = True
        ltEmpresa.Identificador = oRec!Identificador
        ltEmpresa.razonSocial = oRec!razonSocial
        ltEmpresa.Monotributista = oRec!Monotributista
    End If
    
    oRec.Close
    Set oRec = Nothing
    recuperarEmpresaPorIdentificador = ltEmpresa
End Function



Public Function separarCodigoDescripcion(aSeparar As String) As String()
    Dim retorno(2) As String
    Dim i As Integer
    retorno(0) = ""
    retorno(1) = ""
    
    i = InStr(1, aSeparar, "-")
    If i > 0 Then
        retorno(0) = Left(aSeparar, i - 1)
        retorno(1) = Mid(aSeparar, i + 1)
    Else
        retorno(0) = aSeparar
        retorno(1) = ""
    End If
    
    separarCodigoDescripcion = retorno
End Function



Public Function fechaActualServer() As Date
Dim ls_sql As String
Dim oRec As ADODB.Recordset

    Set oRec = New ADODB.Recordset
    ls_sql = "Select GETDATE()"
    
    With oRec
        .ActiveConnection = oCon
        .CursorType = adOpenForwardOnly
        .Source = ls_sql
        .Open
    End With
    fechaActualServer = Now
    If oRec.RecordCount = 1 Then fechaActualServer = oRec(0)
End Function

Public Function calculoIVATipoDirecto(importe As Currency, alicuota As Currency) As Currency()
Dim retorno(1) As Currency

    retorno(0) = CCur(convertirCurrencyAString(importe * alicuota / 100))
    retorno(1) = importe + retorno(0)
  
    calculoIVATipoDirecto = retorno
    
End Function


Public Function calculoIVATipoInverso(importe As Currency, alicuota As Currency) As Currency()
Dim retorno(1) As Currency

    retorno(0) = CCur(convertirCurrencyAString((importe / (1 + alicuota / 100)) * (alicuota / 100)))
    retorno(1) = importe - retorno(0)
    calculoIVATipoInverso = retorno
    
End Function

Public Function convertirAdouble(numero As String) As Double
    numero = Val(Replace(numero, ",", "."))
End Function


Public Function ValidarCuit(ByVal Cuit As String) As Boolean
    If Len(Cuit) = 11 Then
        Dim Ponderador As Integer
        Dim Acumulado As Integer
        Dim Digito As Integer

        Ponderador = 2
        Acumulado = 0
        
        Dim Posicion As Integer
        'Recorro la cadena de atrás para adelante
        For Posicion = 10 To 1 Step -1
            'Sumo las multiplicaciones de cada dígito x su ponderador
            Acumulado = Acumulado + Val(Mid$(Cuit, Posicion, 1)) * Ponderador
            Ponderador = Ponderador + 1

            If Ponderador > 7 Then Ponderador = 2
        Next
    
        Digito = 11 - (Acumulado Mod 11)
        If Digito = 11 Then Digito = 0
        If Digito = 10 Then Digito = 9

        ValidarCuit = (Digito = Right(Cuit, 1))
    Else
        ValidarCuit = False
    End If
End Function

Public Function convertirCurrencyAString(value As Currency) As String

Dim cantidadDecimales As Integer

    cantidadDecimales = 0
    If InStr(CStr(value), separadorDecimal) > 0 Then
        cantidadDecimales = Len(CStr(value)) - InStr(CStr(value), separadorDecimal)
    End If
    If cantidadDecimales > 0 Then
        If cantidadDecimales > 2 Then cantidadDecimales = 2
        convertirCurrencyAString = Left(CStr(value), InStr(CStr(value), separadorDecimal) + cantidadDecimales)
    Else
        convertirCurrencyAString = CStr(value)
    End If
  
End Function


Public Function recuperarTipoCalculo(ltTipoComprobante As tTipoComprobante) As String
    Dim ls_sql As String
    Dim oRec As ADODB.Recordset
    On Error GoTo errores
    Set oRec = New ADODB.Recordset
    
    ls_sql = "SELECT ISNULL(CALCULO,'D') as CALCULO FROM TIPO_COMPROBANTE WHERE CODIGO ='" & ltTipoComprobante.codigo & "'"
    With oRec
        .ActiveConnection = oCon
        .CursorType = adOpenStatic
        .CursorLocation = adUseClient
        .Source = ls_sql
        .Open
    End With
    
    recuperarTipoCalculo = "D"
    If oRec.RecordCount = 1 Then
        recuperarTipoCalculo = oRec!CALCULO
    End If
    Set oRec = Nothing
    Exit Function
    

errores:
    If oCon.Errors.Count > 0 Then
        MsgBox oCon.Errors(0).Description, vbCritical, "Nro de Error: " & oCon.Errors(0).Number
        oCon.Errors.Clear
    End If
    Set oRec = Nothing
    
End Function


Public Function existeRegistradoNroComprobanteEmpresa(ltPuntoDeVenta As tPuntoDeVenta, nroComprobante As String, tipoDeComprobante As String) As Boolean
    Dim ls_sql As String
    Dim oRec As ADODB.Recordset
    On Error GoTo errores
    Set oRec = New ADODB.Recordset
    
    ls_sql = "SELECT VENTA_ID from CABECERA_VENTA WHERE PUNTO_VENTA_ID = " & ltPuntoDeVenta.puntoVentaId & _
             " AND NRO_COMPROBANTE_DESDE='" & nroComprobante & "' " & _
             " AND TIPO_COMPROBANTE_ID ='" & tipoDeComprobante & "' "
    
    With oRec
        .ActiveConnection = oCon
        .CursorType = adOpenStatic
        .CursorLocation = adUseClient
        .Source = ls_sql
        .Open
    End With
    
    existeRegistradoNroComprobanteEmpresa = False
    
    If oRec.RecordCount > 0 Then
        existeRegistradoNroComprobanteEmpresa = True
    End If
    
    
    Exit Function
errores:
    If oCon.Errors.Count > 0 Then
        MsgBox oCon.Errors(0).Description, vbCritical, "Nro de Error: " & oCon.Errors(0).Number
        oCon.Errors.Clear
        existeRegistradoNroComprobanteEmpresa = False
    End If
    Set oRec = Nothing
End Function


Public Sub Sendkeys(text As Variant, Optional wait As Boolean = False)
   Dim WshShell As Object
   Set WshShell = CreateObject("wscript.shell")
   WshShell.Sendkeys CStr(text), wait
   Set WshShell = Nothing
End Sub

