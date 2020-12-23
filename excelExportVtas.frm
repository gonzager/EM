VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form excelExportVtas 
   Caption         =   "Generación de Informes y Archivos De Venta"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   7320
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4980
      TabIndex        =   7
      Top             =   2520
      Width           =   2235
   End
   Begin VB.CommandButton cmdGenerarAlicuotasVentas 
      Caption         =   "Archivo Alicuotas de Vtas"
      Height          =   375
      Left            =   4980
      TabIndex        =   6
      Top             =   2040
      Width           =   2250
   End
   Begin VB.CommandButton cmdGenerarArchivoVentas 
      Caption         =   "Generar Archivo Ventas"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   2040
      Width           =   2250
   End
   Begin VB.Frame frmCriterios 
      Caption         =   "Datos de Generación"
      Height          =   1815
      Left            =   60
      TabIndex        =   8
      Top             =   60
      Width           =   7155
      Begin VB.TextBox txtRazonSocial 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   720
         Width           =   5835
      End
      Begin MSComCtl2.DTPicker dpFechaHasta 
         Height          =   315
         Left            =   5640
         TabIndex        =   3
         Top             =   1200
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Format          =   102694913
         CurrentDate     =   42208
      End
      Begin MSComCtl2.DTPicker dpFechaDesde 
         Height          =   315
         Left            =   2160
         TabIndex        =   2
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   102694913
         CurrentDate     =   42208
      End
      Begin MSMask.MaskEdBox txtNroIdVendedor 
         Height          =   315
         Left            =   2220
         TabIndex        =   0
         Top             =   300
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   13
         Mask            =   "##-########-#"
         PromptChar      =   "_"
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
      Begin VB.Label lblFechaHasta 
         Caption         =   "Fecha Comprobante Hasta:"
         Height          =   255
         Left            =   3660
         TabIndex        =   10
         Top             =   1260
         Width           =   1995
      End
      Begin VB.Label lblFechaDesde 
         Caption         =   "Fecha Comprobante Desde:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1260
         Width           =   1995
      End
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "Sub Diario Ventas"
      Height          =   375
      Left            =   60
      TabIndex        =   4
      Top             =   2040
      Width           =   2250
   End
End
Attribute VB_Name = "excelExportVtas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExportar_Click()
    Exportar_ADO_Excel_SubDiarioVentas Me.dpFechaDesde, Me.dpFechaHasta, Me.txtNroIdVendedor
End Sub



Private Sub cmdGenerarAlicuotasVentas_Click()
    Dim Path_Archivo_Ini As String
    Dim nombreArchivoTXT As String
    Dim ls_sql As String
    
    Path_Archivo_Ini = App.Path & "\EM_config.ini"
    nombreArchivoTXT = "VENTAS_ALICUOTAS_" & Me.txtNroIdVendedor.ClipText & "_" & Format(Me.dpFechaDesde, formatoFechaQuery) & "_" & Format(Me.dpFechaHasta, formatoFechaQuery) & ".txt"
    nombreArchivoTXT = Leer_Ini(Path_Archivo_Ini, "DIR_EXP_VENTAS", App.Path) & "\" & nombreArchivoTXT
    
    ls_sql = "SELECT * FROM V_EXPORTAR_ALICUOTA_VENTAS " + _
             "WHERE FECHA_VENTA between '" & Format(Me.dpFechaDesde, "YYYYMMDD") & "' " + _
             "AND '" & Format(Me.dpFechaHasta, "YYYYMMDD") & "' " + _
             "AND EMPRESA_ID=" & Me.txtNroIdVendedor & " " + _
             "ORDER BY EMPRESA_ID, [TIPO DE COMPROBANTE], [Punto de Venta], [Comprobante Desde] "
    If Exportar_Recordset(nombreArchivoTXT, ls_sql, "", False, 2) Then
        MsgBox "La exportación del archivo cabecera de ventas se ha realizado con Exito", vbInformation, "Exportación de Cabecera de Ventas"
    Else
        MsgBox "Error en la Exportación del Archivo de datos", vbCritical, "Error: Exportación de Archivo de Ventas"
    End If
End Sub

Private Sub cmdGenerarArchivoVentas_Click()
    Dim Path_Archivo_Ini As String
    Dim nombreArchivoTXT As String
    Dim ls_sql As String
    
    Path_Archivo_Ini = App.Path & "\EM_config.ini"
    nombreArchivoTXT = "VENTAS_CBTE_" & Me.txtNroIdVendedor.ClipText & "_" & Format(Me.dpFechaDesde, formatoFechaQuery) & "_" & Format(Me.dpFechaHasta, formatoFechaQuery) & ".txt"
    nombreArchivoTXT = Leer_Ini(Path_Archivo_Ini, "DIR_EXP_VENTAS", App.Path) & "\" & nombreArchivoTXT
    
    ls_sql = "SELECT * FROM V_EXPORTAR_CABECERA_VENTAS " + _
             "WHERE [Fecha de Comprobante] between '" & Format(Me.dpFechaDesde, "YYYYMMDD") & "' " + _
             "AND '" & Format(Me.dpFechaHasta, "YYYYMMDD") & "' " + _
             "AND EMPRESA_ID=" & Me.txtNroIdVendedor & " " & _
             "ORDER BY EMPRESA_ID, [TIPO DE COMPROBANTE], [Punto de Venta], [Comprobante Desde] "
    If Exportar_Recordset(nombreArchivoTXT, ls_sql, "", False, 1) Then
        MsgBox "La exportación del archivo cabecera de ventas se ha realizado con Exito", vbInformation, "Exportación de Cabecera de Ventas"
    Else
        MsgBox "Error en la Exportación del Archivo de datos", vbCritical, "Error: Exportación de Archivo de Ventas"
    End If
    
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.dpFechaDesde.value = Format(fechaActualServer, "Short date")
    Me.dpFechaHasta.value = Format(fechaActualServer, "Short date")
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
    If Len(txtNroIdVendedor.ClipText) > 0 Then
        If Len(txtNroIdVendedor.ClipText) = 11 Then
            Dim ltEmpresa As tEmpresa
            ltEmpresa = recuperarEmpresaPorIdentificador(txtNroIdVendedor.ClipText)
            If ltEmpresa.existe Then
                txtRazonSocial.text = ltEmpresa.razonSocial
                
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
