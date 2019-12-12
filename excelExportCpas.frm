VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form excelExportCpas 
   Caption         =   "Generación de Informes y Archivos de Compras"
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
      Caption         =   "Archivo Alicuotas de Compas"
      Height          =   375
      Left            =   4980
      TabIndex        =   6
      Top             =   2040
      Width           =   2250
   End
   Begin VB.CommandButton cmdGenerarArchivoVentas 
      Caption         =   "Generar Archivo Compras"
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
         Top             =   1380
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Format          =   143917057
         CurrentDate     =   42208
      End
      Begin MSComCtl2.DTPicker dpFechaDesde 
         Height          =   315
         Left            =   2160
         TabIndex        =   2
         Top             =   1380
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   143917057
         CurrentDate     =   42208
      End
      Begin MSMask.MaskEdBox txtNroIdComprador 
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
      Begin VB.Label Label2 
         Caption         =   "Importante: La exportación de compras es por fecha de imputación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   6735
      End
      Begin VB.Label lblIdentificadorComprador 
         Caption         =   " Nro. Identificador Comprador:"
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
         Caption         =   "Fecha Imputación Hasta:"
         Height          =   255
         Left            =   3660
         TabIndex        =   10
         Top             =   1440
         Width           =   1995
      End
      Begin VB.Label lblFechaDesde 
         Caption         =   "Fecha Imputación Desde:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   1995
      End
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "Sub Diario Compras"
      Height          =   375
      Left            =   60
      TabIndex        =   4
      Top             =   2040
      Width           =   2250
   End
End
Attribute VB_Name = "excelExportCpas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExportar_Click()
    Exportar_ADO_Excel_SubDiarioCompras Me.dpFechaDesde, Me.dpFechaHasta, Me.txtNroIdComprador
End Sub



Private Sub cmdGenerarAlicuotasVentas_Click()
   
    Dim Path_Archivo_Ini As String
    Dim nombreArchivoTXT As String
    Dim ls_sql As String

    Path_Archivo_Ini = App.Path & "\EM_config.ini"
    nombreArchivoTXT = "COMPRAS_ALICUOTAS_" & Me.txtNroIdComprador.ClipText & "_" & Format(Me.dpFechaDesde, formatoFechaQuery) & "_" & Format(Me.dpFechaHasta, formatoFechaQuery) & ".txt"
    nombreArchivoTXT = Leer_Ini(Path_Archivo_Ini, "DIR_EXP_COMPRAS", App.Path) & "\" & nombreArchivoTXT

    ls_sql = "SELECT * FROM V_EXPORTAR_ALICUOTA_COMPRAS " + _
             "WHERE FECHA_IMPUTACION between '" & Format(Me.dpFechaDesde, "YYYYMMDD") & "' " + _
             "AND '" & Format(Me.dpFechaHasta, "YYYYMMDD") & "' " + _
             "AND EMPRESA_ID=" & Me.txtNroIdComprador
    If Exportar_Recordset(nombreArchivoTXT, ls_sql, "", False, 2) Then
        MsgBox "La exportación del archivo detalle de compras se ha realizado con Exito", vbInformation, "Exportación de Cabecera de Ventas"
    Else
        MsgBox "Error en la Exportación del Archivo de datos", vbCritical, "Error: Exportación de Archivo de Ventas"
    End If
End Sub

Private Sub cmdGenerarArchivoVentas_Click()

   
    Dim Path_Archivo_Ini As String
    Dim nombreArchivoTXT As String
    Dim ls_sql As String

    Path_Archivo_Ini = App.Path & "\EM_config.ini"
    nombreArchivoTXT = "COMPRAS_CBTE_" & Me.txtNroIdComprador.ClipText & "_" & Format(Me.dpFechaDesde, formatoFechaQuery) & "_" & Format(Me.dpFechaHasta, formatoFechaQuery) & ".txt"
    nombreArchivoTXT = Leer_Ini(Path_Archivo_Ini, "DIR_EXP_COMPRAS", App.Path) & "\" & nombreArchivoTXT

    ls_sql = "SELECT * FROM V_EXPORTAR_CABECERA_COMPRAS " + _
             "WHERE [FECHA_IMPUTACION] between '" & Format(Me.dpFechaDesde, "YYYYMMDD") & "' " + _
             "AND '" & Format(Me.dpFechaHasta, "YYYYMMDD") & "' " + _
             "AND EMPRESA_ID=" & Me.txtNroIdComprador
    If Exportar_Recordset(nombreArchivoTXT, ls_sql, "", False, 2) Then
        MsgBox "La exportación del archivo cabecera de compras se ha realizado con Exito", vbInformation, "Exportación de Cabecera de Ventas"
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

Private Sub txtNroIdComprador_GotFocus()
    With txtNroIdComprador
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
    txtRazonSocial.Text = ""
    
End Sub

Private Sub txtNroIdComprador_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Dim indentificador As Double
        frmBusquedaEmpresa.inicializarFormulario indentificador
        txtNroIdComprador.Mask = ""
        txtNroIdComprador.Text = indentificador
        'txtNroIdComprador.Mask = "##-########-#"
        SendKeys "{tab}"
        Exit Sub
    End If
   
   If Not IsNumeric(Chr(KeyAscii)) Then
       If KeyAscii <> vbKeyBack Then
          KeyAscii = 0
       End If
    End If
    

End Sub

Private Sub txtNroIdComprador_LostFocus()
On Error Resume Next
    If Len(txtNroIdComprador.ClipText) > 0 Then
        If Len(txtNroIdComprador.ClipText) = 11 Then
            Dim ltEmpresa As tEmpresa
            ltEmpresa = recuperarEmpresaPorIdentificador(txtNroIdComprador.ClipText)
            If ltEmpresa.existe Then
                txtRazonSocial.Text = ltEmpresa.razonSocial
                
            Else
                With txtNroIdComprador
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
                End With
               
            End If
        Else
            With txtNroIdComprador
                .SelStart = 0
                .SelLength = Len(.Text)
                .SetFocus
            End With
           
        End If
    End If
End Sub


Private Sub txtNroIdVendedor_Change()

End Sub
