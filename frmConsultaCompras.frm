VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmConsultaCompras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Compras"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   14790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   14790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExport 
      Cancel          =   -1  'True
      Caption         =   "&Excel"
      Height          =   495
      Left            =   10320
      TabIndex        =   17
      Top             =   7080
      Width           =   1995
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   12660
      TabIndex        =   16
      Top             =   7080
      Width           =   1995
   End
   Begin MSFlexGridLib.MSFlexGrid msGrilla 
      Height          =   5355
      Left            =   60
      TabIndex        =   15
      Top             =   1560
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   9446
      _Version        =   393216
   End
   Begin VB.Frame frmCriterios 
      Caption         =   "Criterios de Búsqueda"
      Height          =   1455
      Left            =   60
      TabIndex        =   8
      Top             =   60
      Width           =   14655
      Begin VB.Frame frmFechas 
         Height          =   615
         Left            =   4980
         TabIndex        =   12
         Top             =   600
         Width           =   5475
         Begin MSComCtl2.DTPicker dpFechaDesde 
            Height          =   315
            Left            =   1200
            TabIndex        =   5
            Top             =   180
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Format          =   119472129
            CurrentDate     =   42208
         End
         Begin MSComCtl2.DTPicker dpFechaHasta 
            Height          =   315
            Left            =   3840
            TabIndex        =   6
            Top             =   180
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Format          =   119472129
            CurrentDate     =   42208
         End
         Begin VB.Label lblFechaHasta 
            Caption         =   "Fecha Hasta:"
            Height          =   255
            Left            =   2760
            TabIndex        =   14
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lblFechaDesde 
            Caption         =   "Fecha Desde:"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1035
         End
      End
      Begin VB.OptionButton opFechaImputacion 
         Caption         =   "Por Fecha de Imputación"
         Height          =   375
         Left            =   2580
         TabIndex        =   4
         Top             =   780
         Width           =   2235
      End
      Begin VB.OptionButton opFechaComprobante 
         Caption         =   "Por Fecha de Comprobante"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   780
         Width           =   2595
      End
      Begin VB.TextBox txtRazonSocial 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4980
         TabIndex        =   1
         Top             =   300
         Width           =   5475
      End
      Begin VB.ComboBox cmbTipoComprobante 
         Height          =   315
         Left            =   11880
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   300
         Width           =   2595
      End
      Begin VB.CommandButton cmdRecuperar 
         Caption         =   "&Recuperar Comprobantes de Compra"
         Height          =   435
         Left            =   10620
         TabIndex        =   7
         Top             =   780
         Width           =   3855
      End
      Begin MSMask.MaskEdBox txtNroIdComprador 
         Height          =   315
         Left            =   2280
         TabIndex        =   0
         Top             =   300
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   13
         Mask            =   "##-########-#"
         PromptChar      =   "_"
      End
      Begin VB.Label lblIdentificadorComprador 
         Caption         =   " Nro. Identificador Comprador:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   2115
      End
      Begin VB.Label lblRazonSocial 
         Caption         =   "Razón Social:"
         Height          =   195
         Left            =   3900
         TabIndex        =   10
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label lblTipoComprobante 
         Caption         =   "Tipo Comprobate:"
         Height          =   255
         Left            =   10560
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Menu mnu_operaciones 
      Caption         =   "&Operaciones"
      Begin VB.Menu mnu_BorrarRegistraccion 
         Caption         =   "Borrar Registracion"
      End
   End
End
Attribute VB_Name = "frmConsultaCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const COL_ID = 25
Public operacion As String
Public idCompas As String

Private Sub cmdExport_Click()
Dim ls_sql As String
Dim ls_where As String


    ls_sql = "SELECT * FROM V_SUBDIARIO_COMPRAS "
    ls_where = ""
    
    If Me.opFechaComprobante.value = True Then
        ls_where = "WHERE [F. Comprobante] between '" & Format(Me.dpFechaDesde, formatoFechaQuery) & "' and '" & Format(Me.dpFechaHasta, formatoFechaQuery) & "' "
    Else
        ls_where = "WHERE [F. Imputacion] between '" & Format(Me.dpFechaDesde, formatoFechaQuery) & "' and '" & Format(Me.dpFechaHasta, formatoFechaQuery) & "' "
    End If
    
     
    ls_where = ls_where & IIf(Len(Me.txtNroIdComprador.ClipText) > 0, "AND [ID. Comprador] = " & Me.txtNroIdComprador.ClipText & " ", "")
        
    If Me.cmbTipoComprobante.ListIndex > 0 Then
        Dim ls_tmp() As String
            ls_tmp = Utils.separarCodigoDescripcion(Me.cmbTipoComprobante)
            ls_where = ls_where & "AND TIPO_COMPROBANTE_ID ='" & ls_tmp(0) & "' "
    End If
    
    If Me.opFechaComprobante.value = True Then
        ls_sql = ls_sql & ls_where & "ORDER BY [ID. Comprador], [F. Comprobante], [Total] desc"
    Else
        ls_sql = ls_sql & ls_where & "ORDER BY [ID. Comprador], [F. Imputacion], [Total] desc"
    End If
    UtilsExcel.Exportar_ADO_Excel_RecordSet ls_sql
 
End Sub

Private Sub cmdRecuperar_Click()
    cargarGrilla
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    inicializarFormulario
End Sub

Private Sub mnu_BorrarRegistraccion_Click()
    If msGrilla.Row <= 0 Then
        MsgBox "Debe seleccionar una fila para borrar.", vbCritical, "Error: Al Borrar Registracion"
    ElseIf msGrilla.Row >= 1 And msGrilla.Row <= msGrilla.Rows - 2 Then
        If MsgBox("Seguro desea eliminar la registración que se encuentra seleccionada", vbQuestion + vbYesNoCancel, "Confirmar Borrar Registración") = vbYes Then
            Dim ID As Double
            ID = msGrilla.TextMatrix(msGrilla.Row, COL_ID)
            If borrarCabeceraDetalleCompra(ID) > 0 Then
                ID = msGrilla.Row
                cmdRecuperar_Click
                msGrilla.Row = ID
                msGrilla.RowSel = 1
                msGrilla.Col = 0
                msGrilla.ColSel = 13
                                
            End If
        End If
       
    End If
End Sub

Private Sub msGrilla_DblClick()
    Dim ID As Double
    If msGrilla.Row <= 0 Then
        MsgBox "Debe seleccionar una fila para seleccionar.", vbCritical, "Error: Al seleccionar registro"
    ElseIf msGrilla.Row >= 1 And msGrilla.Row <= msGrilla.Rows - 2 Then
        ID = msGrilla.TextMatrix(msGrilla.Row, COL_ID)
        frmCompras.operacion = "M"
        frmCompras.idCompra = ID
        ID = msGrilla.Row
        frmCompras.Show vbModal
        cmdRecuperar_Click
        msGrilla.Row = ID
        msGrilla.RowSel = 1
        msGrilla.Col = 0
        msGrilla.ColSel = 25
    End If
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
        txtNroIdComprador.PromptInclude = False
        txtNroIdComprador.Mask = "##-########-#"
        txtNroIdComprador.PromptInclude = True
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

Private Sub inicializarFormulario()

    
    Dim sql_tipoComprobante As String
    
    
    Me.dpFechaDesde.value = Format(fechaActualServer, "Short date")
    Me.dpFechaHasta.value = Format(fechaActualServer, "Short date")
     
    sql_tipoComprobante = "SELECT CODIGO, ISNULL(DESCRIPCION,'') AS DESCRIPCION, ISNULL(DEFECTO,0) AS DEFECTO FROM TIPO_COMPROBANTE WHERE ISNULL(VISIBLE,0) = 1 order by CODIGO "
    cargarCombo Me.cmbTipoComprobante, sql_tipoComprobante
    
    Me.cmbTipoComprobante.AddItem "TODAS", 0
    Me.cmbTipoComprobante.ListIndex = 0
   
    Me.opFechaComprobante.value = True
    
End Sub

Private Sub cargarGrilla()
Dim ls_sql As String
Dim ls_where As String
Dim columna As Integer
Dim Ancho_Campo  As Integer
Dim fila As Integer
Dim rst As New ADODB.Recordset

On Error GoTo errores:


    ls_sql = "SELECT * FROM V_SUBDIARIO_COMPRAS "
    ls_where = ""
    
    If Me.opFechaComprobante.value = True Then
        ls_where = "WHERE [F. Comprobante] between '" & Format(Me.dpFechaDesde, formatoFechaQuery) & "' and '" & Format(Me.dpFechaHasta, formatoFechaQuery) & "' "
    Else
        ls_where = "WHERE [F. Imputacion] between '" & Format(Me.dpFechaDesde, formatoFechaQuery) & "' and '" & Format(Me.dpFechaHasta, formatoFechaQuery) & "' "
    End If
    

        
    ls_where = ls_where & IIf(Len(Me.txtNroIdComprador.ClipText) > 0, "AND [ID. Comprador] = " & Me.txtNroIdComprador.ClipText & " ", "")
        
    If Me.cmbTipoComprobante.ListIndex > 0 Then
        Dim ls_tmp() As String
            ls_tmp = Utils.separarCodigoDescripcion(Me.cmbTipoComprobante)
            ls_where = ls_where & "AND TIPO_COMPROBANTE_ID ='" & ls_tmp(0) & "' "
    End If
    
    If Me.opFechaComprobante.value = True Then
        ls_sql = ls_sql & ls_where & "ORDER BY [ID. Comprador], [F. Comprobante], [Total] desc"
    Else
        ls_sql = ls_sql & ls_where & "ORDER BY [ID. Comprador], [F. Imputacion], [Total] desc"
    End If
    
    With rst
        .ActiveConnection = oCon
        .CursorType = adOpenStatic
        .CursorLocation = adUseClient
        .Source = ls_sql
        .Open
    End With
    
    
    With Me.msGrilla
        ' -- Deshabilitar el repintado del control ( Para que la carga sea mas veloz )
        .Redraw = False
        ' -- Seleccionar registros del Grid por Fila
        .SelectionMode = flexSelectionByRow
        ' -- Cantidad de filas inicial
        .Rows = 2
        ' -- Modo de encabezados
        .FixedRows = 1
        .FixedCols = 0
      
        ' -- Cantidad de filas y columnas
        .Rows = 1
        .Cols = rst.Fields.Count
      
        ' -- Redimensionar el Array a la cantidad de campos de la tabla
        ReDim Ancho_Columna(0 To rst.Fields.Count - 1)
      
        ' -- Recorrer los campos del recordset
        For columna = 0 To rst.Fields.Count - 1
            ' -- Añade el título del campo al encabezado de columna
            .TextMatrix(0, columna) = rst.Fields(columna).Name
            ' -- Guardar el ancho del campo en la matriz
            Ancho_Columna(columna) = TextWidth(rst.Fields(columna).Name)
        Next columna
          
        fila = 1
        ' -- Recorrer todos los registros del recordset
        Do While Not rst.EOF
            .Rows = .Rows + 1 ' Añade una nueva fila
            For columna = 0 To rst.Fields.Count - 1
                ' -- Combobar que el valor no es nulo
                If Not IsNull(rst.Fields(columna).value) Then
                    ' -- Agrega el registro en la fila y columna específica
                    .TextMatrix(fila, columna) = rst.Fields(columna).value
                    ' -- Almacena el ancho
                    Ancho_Campo = TextWidth(rst.Fields(columna).value)
                    If Ancho_Columna(columna) < Ancho_Campo Then
                        Ancho_Columna(columna) = Ancho_Campo
                    End If
                End If
            Next
            ' -- Siguiente registro
            rst.MoveNext
            fila = fila + 1 'Incrementa la fila
        Loop
  
          
        ' -- Cierra el recordset y la conexión abierta
        rst.Close

  
        ' -- Establece los ancho de columna de la grilla
        For columna = 0 To Me.msGrilla.Cols - 1
            .ColWidth(columna) = Ancho_Columna(columna) + 240
            If columna = COL_ID Then
                .ColWidth(COL_ID) = 0
            End If
        Next
          
        ' -- Volver a Habilitar el repintado del Grid
        .Redraw = True
    End With
    msGrilla.Rows = msGrilla.Rows + 1
    
    
    Set rst = Nothing
    Exit Sub
    
errores:
    On Error Resume Next
    If oCon.Errors.Count >= 1 Then
        MsgBox oCon.Errors(0).Description, vbCritical, "Nro de Error: " & oCon.Errors(0).Number
        oCon.Errors.Clear
    Else
        MsgBox Err.Number & "-" & Err.Description, vbCritical, "Error: Al Recuperar Datos de Ventas"
        Err.Clear
    End If

    Set rst = Nothing
End Sub

