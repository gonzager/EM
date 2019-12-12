VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImporVentasAfip 
   Caption         =   "Importar Archivos de Ventas AFIP"
   ClientHeight    =   7530
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   8130
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmDatos 
      Height          =   6855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8085
      Begin VB.CommandButton cmdVerImportados 
         Caption         =   "Ver Comprobantes Importados"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2880
         TabIndex        =   20
         Top             =   6240
         Width           =   2415
      End
      Begin VB.Frame frmControl 
         Caption         =   "Control de Integridad"
         Height          =   2415
         Left            =   120
         TabIndex        =   15
         Top             =   3720
         Width           =   7875
         Begin VB.CommandButton cmdValidaciones 
            Caption         =   "Validar Datos"
            Height          =   375
            Left            =   5520
            TabIndex        =   17
            Top             =   1920
            Width           =   2175
         End
         Begin MSFlexGridLib.MSFlexGrid grillaControl 
            Height          =   1575
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   2778
            _Version        =   393216
            Rows            =   5
            FixedCols       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame frmCriterios 
         Caption         =   "Datos de Generación"
         Height          =   1815
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   7875
         Begin VB.ComboBox cmbPuntoVenta 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   1200
            Width           =   915
         End
         Begin VB.ComboBox cmbConcepto 
            Height          =   315
            Left            =   4560
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1200
            Width           =   3135
         End
         Begin VB.TextBox txtRazonSocial 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1200
            TabIndex        =   9
            Top             =   720
            Width           =   6435
         End
         Begin MSMask.MaskEdBox txtNroIdVendedor 
            Height          =   315
            Left            =   2220
            TabIndex        =   10
            Top             =   300
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   13
            Mask            =   "##-########-#"
            PromptChar      =   "_"
         End
         Begin VB.Label lblPuntoVenta 
            Caption         =   "Punto Venta:"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   1260
            Width           =   1215
         End
         Begin VB.Label lblConcepto 
            Caption         =   "Concepto de Importación:"
            Height          =   195
            Left            =   2520
            TabIndex        =   14
            Top             =   1260
            Width           =   1995
         End
         Begin VB.Label lblIdentificadorVendedor 
            Caption         =   " Nro. Identificador Vendedor:"
            Height          =   255
            Left            =   60
            TabIndex        =   12
            Top             =   360
            Width           =   2115
         End
         Begin VB.Label lblRazonSocial 
            Caption         =   "Razón Social:"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   780
            Width           =   1035
         End
      End
      Begin VB.CommandButton cmdAlicuotas 
         Caption         =   "Copiar y Leer Archivo ALICUOTAS"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   3120
         Width           =   2895
      End
      Begin VB.CommandButton cmpImportar 
         Caption         =   "Importar Datos"
         Enabled         =   0   'False
         Height          =   495
         Left            =   5520
         TabIndex        =   3
         Top             =   6240
         Width           =   2415
      End
      Begin VB.CommandButton cmdVentas 
         Caption         =   "Copiar y Leer Archivo VENTAS"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   2520
         Width           =   2895
      End
      Begin VB.Label lblCantidad1 
         Caption         =   "Cantidad de Comprobantes a  Procesar : 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3240
         TabIndex        =   7
         Top             =   2640
         Width           =   4695
      End
      Begin VB.Label lblCantidad2 
         Caption         =   "Cantidad de Alicuotas a  Procesar : 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3240
         TabIndex        =   6
         Top             =   3240
         Width           =   4695
      End
      Begin VB.Label lblServerFile 
         Caption         =   "lblServerFile"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   120
         Width           =   7125
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6360
      TabIndex        =   4
      Top             =   7080
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmImporVentasAfip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tTip As ToolTip

Private Sub cmdAlicuotas_Click()

Dim Linea As String
Dim i As Integer
Dim errTam As Boolean
Dim fileName As String
Dim serverFile As String
Dim serverPath As String
Dim Path_Archivo_Ini As String
Dim formatBulkXML  As String
Dim errTram As Boolean

On Error GoTo tratarError
    
    Me.cmpImportar.Enabled = False
    Me.cmdVerImportados.Enabled = False
    errTram = False
    'Titulo del CommonDialog
    CommonDialog1.fileName = ""
    CommonDialog1.DialogTitle = "Seleccione el archivo a Importar"
    
    'Extension del CommonDialog. Archivos txt
    CommonDialog1.Filter = "Archivos tipo txt|ALICUOTAS_" + Me.txtNroIdVendedor.text + ".txt"

    'Abrimos el CommonDialog
    CommonDialog1.ShowOpen

    If CommonDialog1.fileName = "" Then
        'salimos de la rutina ya que no se ha seleccionado ningún archivo
        Exit Sub
    Else
        i = 0
        fileName = CommonDialog1.fileName
        'Abrimos el archivo para leerlo, pasándole la ruta con la propiedad FileName del Commondialog
        Open fileName For Input As #1
  
        While Not EOF(1) And Not errTam
            'Leemos la línea
            Line Input #1, Linea
            If Len(Linea) <> 62 Then
                errTam = True
            Else
                If Mid(Linea, 4, 5) <> Me.cmbPuntoVenta.text Then
                    MsgBox "El punto de venta selecionado no se corresponde con el punto de venta del archivo", vbCritical, "Error: Punto de Venta"
                    Close #1
                    Exit Sub
                End If
            End If
            i = i + 1
            
        Wend
  
        'Cerramos el archivo abierto anteriormente
        Close #1
        
        If errTam Then
            MsgBox "Error de formato en la linea " & i & " del Archivo: " & vbCrLf & fileName, vbCritical, "Error en el Archivo"
            Exit Sub
        End If
        
        Path_Archivo_Ini = App.Path & "\EM_config.ini"
        serverPath = Leer_Ini(Path_Archivo_Ini, "AFIPSERVERFILESHARE", AFIPSERVERFILESHARE) + "\"
        serverFile = serverPath & Obtener_Nombre_Archivo(fileName)
        formatBulkXML = serverPath & Leer_Ini(Path_Archivo_Ini, "ALICUOTAS_XML_FILE_FORMAT", "I_ALICUOTAS.xml")
        Copiar_Archivo fileName, serverFile
        
        If fBuklAlicuotas(serverFile, formatBulkXML) = 0 Then
            If (i = cantidadTmpI_Alicuotas) Then
                lblCantidad2 = "Cantidad de Alicuotas a  Procesar : " & i
            End If
        End If
    
    End If
    
    Exit Sub
tratarError:
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdValidaciones_Click()
    Dim activarBoton As Boolean
    activarBoton = False
    
    Dim l_cantidadTipoComprobanteI_Ventas As Integer
    Dim l_cantidadPuntosDeVenta_IVentas As Integer
    l_cantidadTipoComprobanteI_Ventas = cantidadTipoComprobanteI_Ventas
    If l_cantidadTipoComprobanteI_Ventas > 0 And l_cantidadTipoComprobanteI_Ventas = cantidadTipoComprobanteHabilitados_IVentas Then
        Me.grillaControl.TextMatrix(1, 2) = "OK"
        activarBoton = True
    Else
        Me.grillaControl.TextMatrix(1, 2) = "ERROR"
        activarBoton = False
    End If
    
    l_cantidadPuntosDeVenta_IVentas = cantidadPuntosDeVenta_IVentas
    If l_cantidadPuntosDeVenta_IVentas > 0 And l_cantidadPuntosDeVenta_IVentas = cantidadPuntosDeVentaHabilitados_IVentas(IIf(Me.txtNroIdVendedor.ClipText <> "", Me.txtNroIdVendedor.ClipText, 0)) Then
         Me.grillaControl.TextMatrix(2, 2) = "OK"
         activarBoton = activarBoton And True
    Else
        Me.grillaControl.TextMatrix(2, 2) = "ERROR"
        activarBoton = False
    End If
    
    If 0 = todosLosComprobantesTieneAlicuotas And cantidadTmpI_Ventas > 0 Then
         activarBoton = activarBoton And True
         Me.grillaControl.TextMatrix(3, 2) = "OK"
         activarBoton = activarBoton And True
    Else
        Me.grillaControl.TextMatrix(3, 2) = "ERROR"
        activarBoton = False
    End If
    
    If ExistenComprobantesIngresados(IIf(Me.txtNroIdVendedor.ClipText <> "", Me.txtNroIdVendedor.ClipText, 0)) = 0 And cantidadTmpI_Ventas > 0 Then
        Me.grillaControl.TextMatrix(4, 2) = "OK"
        activarBoton = activarBoton And True
    Else
        Me.grillaControl.TextMatrix(4, 2) = "ERROR"
        activarBoton = False
    End If
    
    cmpImportar.Enabled = activarBoton
    
End Sub

Private Sub cmdVentas_Click()
Dim Linea As String
Dim i As Integer
Dim errTam As Boolean
Dim fileName As String
Dim serverFile As String
Dim serverPath As String
Dim Path_Archivo_Ini As String
Dim formatBulkXML  As String
Dim errTram As Boolean
On Error GoTo tratarError
    
    Me.cmpImportar.Enabled = False
    Me.cmdVerImportados.Enabled = False
    errTram = False
    'Titulo del CommonDialog
    CommonDialog1.DialogTitle = "Seleccione el archivo a Importar"
    
    'Extension del CommonDialog. Archivos txt
    CommonDialog1.Filter = "Archivos tipo txt|VENTAS_" + Me.txtNroIdVendedor.text + ".txt"

    'Abrimos el CommonDialog
    CommonDialog1.ShowOpen

    If CommonDialog1.fileName = "" Then
        'salimos de la rutina ya que no se ha seleccionado ningún archivo
        Exit Sub
    Else
        i = 0
        fileName = CommonDialog1.fileName
        'Abrimos el archivo para leerlo, pasándole la ruta con la propiedad FileName del Commondialog
        Open fileName For Input As #1
  
        While Not EOF(1) And Not errTam
            'Leemos la línea
            Line Input #1, Linea
            If Len(Linea) <> 266 Then
                errTam = True
            Else
                If Mid(Linea, 12, 5) <> Me.cmbPuntoVenta.text Then
                    MsgBox "El punto de venta selecionado no se corresponde con el punto de venta del archivo", vbCritical, "Error: Punto de Venta"
                    Close #1
                    Exit Sub
                End If
            End If
            
            i = i + 1
            
        Wend
  
        'Cerramos el archivo abierto anteriormente
        Close #1
        
        If errTam Then
            MsgBox "Error de formato en la linea " & i & " del Archivo: " & vbCrLf & fileName, vbCritical, "Error en el Archivo"
            Exit Sub
        End If
        
        Path_Archivo_Ini = App.Path & "\EM_config.ini"
        serverPath = Leer_Ini(Path_Archivo_Ini, "AFIPSERVERFILESHARE", AFIPSERVERFILESHARE) + "\"
        serverFile = serverPath & Obtener_Nombre_Archivo(fileName)
        formatBulkXML = serverPath & Leer_Ini(Path_Archivo_Ini, "VENTAS_XML_FILE_FORMAT", "I_VENTAS.xml")
        Copiar_Archivo fileName, serverFile
        
        If fBuklVentas(serverFile, formatBulkXML) = 0 Then
            If (i = cantidadTmpI_Ventas) Then
                lblCantidad1 = "Cantidad de Comprobantes a  Procesar : " & i
            End If
        End If
    
    End If
    
    Exit Sub
tratarError:
    


End Sub

Private Sub cmdVerImportados_Click()
    If grillaControl.TextMatrix(4, 2) = "OK" Then
        Screen.MousePointer = vbHourglass
        frmAfipVentasDuplicadas.cargarGrilla Me.txtNroIdVendedor.text
        Screen.MousePointer = vbDefault
        frmAfipVentasDuplicadas.Show vbModal
    End If
End Sub

Private Sub cmpImportar_Click()
Dim datosConcepto() As String
    If cantidadTmpI_Ventas > 0 And cantidadTmpI_Alicuotas Then
        Screen.MousePointer = vbHourglass
        
        datosConcepto = separarCodigoDescripcion(Me.cmbConcepto.text)
        If fImportarAfip(IIf(Me.txtNroIdVendedor.ClipText <> "", Me.txtNroIdVendedor.ClipText, 0), datosConcepto(0)) = 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "Importación Realizada de Forma Exitosa.", vbInformation, "Mensaje al Usuario"
            cmpImportar.Enabled = False
            cmdVerImportados.Enabled = True
            lblCantidad1 = "Cantidad de Comprobantes a  Procesar : 0"
            lblCantidad2 = "Cantidad de Alicuotas a  Procesar : 0"
        End If
    Else
        MsgBox "No Hay Registros para importar", vbInformation, "Mensaje al Usuario"
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()

Dim serverFile As String
Dim Path_Archivo_Ini As String
Dim sql_conceptos As String
    
    Path_Archivo_Ini = App.Path & "\EM_config.ini"
    serverFile = Leer_Ini(Path_Archivo_Ini, "AFIPSERVERFILESHARE", AFIPSERVERFILESHARE)
    
    sql_conceptos = "SELECT CODIGO, ISNULL(DESCRIPCION,'') AS DESCRIPCION, ISNULL(DEFECTO,0) AS DEFECTO FROM dbo.CONCEPTO WHERE ISNULL(VISIBLE,0) = 1 order by CODIGO "
    cargarCombo Me.cmbConcepto, sql_conceptos
    
    lblServerFile.Caption = serverFile
    BorrarTmpI_Ventas
    BorrarTmpI_Alicuotas
   
    With Me.grillaControl
    
        .Cols = 3
        .Rows = 5
        .ColWidth(0) = 300
        .ColWidth(1) = 6275
        .ColWidth(2) = 900
        .TextMatrix(0, 0) = " #"
        .TextMatrix(0, 1) = "Descripcion Validación"
        .TextMatrix(0, 2) = "Estado"
        .TextMatrix(1, 0) = "1"
        .TextMatrix(2, 0) = "2"
        .TextMatrix(3, 0) = "3"
        .TextMatrix(4, 0) = "4"
        .TextMatrix(1, 1) = "Tipos de Comprobantes Habilitados."
        .TextMatrix(2, 1) = "Punto de Venta Habilitado."
        .TextMatrix(3, 1) = "Todos los Comprobantes con Alicuotas."
        .TextMatrix(4, 1) = "No existen comprobantes ya ingresados."
    
    End With
    ' Crear una nueva instancia de la clase
   Set tTip = New ToolTip
     
   ' Establece El tipo ( balloon o normal )
   tTip.Estilo = TTBalloon
   ' Indica el icono a utilizar ( info, Warning , error etc..)
   tTip.Icono = TTIconInfo
   tTip.Delay = 50 ' Tiempo de duración
    
End Sub

Private Sub Label1_Click()

End Sub


Private Sub grillaControl_DblClick()

    If grillaControl.Row = 4 And grillaControl.TextMatrix(4, 2) <> "" Then
        frmAfipVentasDuplicadas.cargarGrilla Me.txtNroIdVendedor.text
        frmAfipVentasDuplicadas.Show vbModal
    End If
End Sub

Private Sub grillaControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  ' para almacenar la fila y columna actual donde está el puntero
  Static f As Long
  Static c As Long
    
  With grillaControl
    ' verifica que el mouse está en una columna o fila
    If .MouseRow = 4 And .MouseCol >= 1 Then
            
          If f <> .MouseRow Or c <> .MouseCol Then
             ' almacena la fila y columna actual
             f = .MouseRow
             c = .MouseCol
               
             ' texto y titulo
             tTip.Titulo = "Ver Detalle"
             tTip.Texto = "Doble Click " + .TextMatrix(.MouseRow, 1) + " " + .TextMatrix(.MouseRow, 2)
                  
             tTip.Crear .hwnd ' crea el Tips
           End If
    Else
        ' Lo destruye
        tTip.Destroy
    End If
  End With
End Sub

Private Sub txtNroIdVendedor_GotFocus()
    With txtNroIdVendedor
        .SelStart = 0
        .SelLength = Len(.text)
    End With
    
    txtRazonSocial.text = ""
    
End Sub

Private Sub txtNroIdVendedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Dim indentificador As Double
        frmBusquedaEmpresa.inicializarFormulario indentificador
        txtNroIdVendedor.Mask = ""
        txtNroIdVendedor.text = indentificador
        'txtNroIdVendedor.Mask = "##-########-#"
        Sendkeys "{tab}"
        Exit Sub
    End If
   
   If Not IsNumeric(Chr(KeyAscii)) Then
       If KeyAscii <> vbKeyBack Then
          KeyAscii = 0
       End If
    End If
    

End Sub

Private Sub txtNroIdVendedor_LostFocus()
On Error Resume Next
    If Len(txtNroIdVendedor.ClipText) > 0 Then
        If Len(txtNroIdVendedor.ClipText) = 11 Then
            Dim ltEmpresa As tEmpresa
            ltEmpresa = recuperarEmpresaPorIdentificador(txtNroIdVendedor.ClipText)
            If ltEmpresa.existe Then
                txtRazonSocial.text = ltEmpresa.razonSocial
                cmbPuntoVenta.Clear
                cargarComboPuntosDeVenta txtNroIdVendedor.ClipText
                Me.cmpImportar.Enabled = False
                Me.cmdVerImportados.Enabled = False
            Else
                With txtNroIdVendedor
                    .SelStart = 0
                    .SelLength = Len(.text)
                    .SetFocus
                End With
               
            End If
        Else
            With txtNroIdVendedor
                .SelStart = 0
                .SelLength = Len(.text)
                .SetFocus
            End With
           
        End If
    End If
End Sub

Private Sub cargarComboPuntosDeVenta(Identificador As Double)
    Dim ltPuntoDeVenta() As tPuntoDeVenta
    Dim lb_Existe As Boolean
    Dim i As Integer
    ltPuntoDeVenta = recuperarPuntosDeVenta(Identificador, lb_Existe, "A")
    
    If lb_Existe Then
        For i = 0 To UBound(ltPuntoDeVenta)
            Me.cmbPuntoVenta.AddItem ltPuntoDeVenta(i).PuntoDeVenta
        Next
        cmbPuntoVenta.ListIndex = 0
    End If
End Sub
