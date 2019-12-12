Attribute VB_Name = "Principal"
Option Explicit

Private Sub main()
    Dim Path_Archivo_Ini As String
    Dim provider As String
    Dim dataSource As String
    Dim dataBaseName As String
    Dim uid As String
    Dim pwd As String
    
    Path_Archivo_Ini = App.Path & "\EM_config.ini"
    
    provider = Leer_Ini(Path_Archivo_Ini, "provider", "")
    dataSource = Leer_Ini(Path_Archivo_Ini, "dataSource", "")
    dataSource = "Data Source=" & dataSource & ";"
    dataBaseName = Leer_Ini(Path_Archivo_Ini, "dataBaseName", "")
    dataBaseName = "Database=" & dataBaseName & ";"
    uid = Leer_Ini(Path_Archivo_Ini, "uid", "")
    uid = "UID=" & uid & ";"
    pwd = Leer_Ini(Path_Archivo_Ini, "pwd", "1qaz!QAZ")
    pwd = "pwd=" & pwd & ";"
    
    separadorDecimal = Obtener_Separador_Decimal
    
    If abrirConexion(provider, dataSource, dataBaseName, uid, pwd) Then
        frmPrincipal.Show
    End If


End Sub


Private Function abrirConexion(provider As String, dataSource, dataBaseName As String, uid As String, pwd As String) As Boolean
    Set oCon = New ADODB.Connection

    On Error GoTo errores
    
    
    With oCon
        .ConnectionTimeout = 30
        .provider = provider
        .ConnectionString = dataSource & dataBaseName & uid & pwd
        
        .Open
    End With
    
    abrirConexion = oCon.State = adStateOpen
    Exit Function
errores:
    MsgBox oCon.Errors(0).Description, vbCritical, "Nro de Error: " & oCon.Errors(0).Number
    oCon.Errors.Clear
    abrirConexion = False
    
End Function
