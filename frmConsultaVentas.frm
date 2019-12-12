VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmConsultaVentas 
   BorderStyle     =   0  'None
   Caption         =   "Consulta de Ventas"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExcel 
      Caption         =   "&Excel"
      Height          =   375
      Left            =   8220
      TabIndex        =   14
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9960
      TabIndex        =   12
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Frame frmCriterios 
      Caption         =   "Criterios de Búsqueda"
      Height          =   1515
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   11535
      Begin VB.CommandButton cmdRecuperar 
         Caption         =   "&Recuperar"
         Height          =   435
         Left            =   9360
         TabIndex        =   11
         Top             =   780
         Width           =   1815
      End
      Begin VB.ComboBox cmbTipoComprobante 
         Height          =   315
         Left            =   6780
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox txtRazonSocial 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4980
         TabIndex        =   1
         Top             =   300
         Width           =   6195
      End
      Begin MSComCtl2.DTPicker dpFechaHasta 
         Height          =   315
         Left            =   3900
         TabIndex        =   2
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   112001025
         CurrentDate     =   42208
      End
      Begin MSComCtl2.DTPicker dpFechaDesde 
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   112001025
         CurrentDate     =   42208
      End
      Begin MSMask.MaskEdBox txtNroIdVendedor 
         Height          =   315
         Left            =   2280
         TabIndex        =   4
         Top             =   300
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   13
         Mask            =   "##-########-#"
         PromptChar      =   "_"
      End
      Begin VB.Label lblTipoComprobante 
         Caption         =   "Tipo Comprobate:"
         Height          =   255
         Left            =   5460
         TabIndex        =   10
         Top             =   900
         Width           =   1575
      End
      Begin VB.Label lblFechaDesde 
         Caption         =   "Fecha Desde:"
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Top             =   900
         Width           =   1035
      End
      Begin VB.Label lblFechaHasta 
         Caption         =   "Fecha Hasta:"
         Height          =   255
         Left            =   2880
         TabIndex        =   7
         Top             =   900
         Width           =   975
      End
      Begin VB.Label lblRazonSocial 
         Caption         =   "Razón Social:"
         Height          =   195
         Left            =   3900
         TabIndex        =   6
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label lblIdentificadorVendedor 
         Caption         =   " Nro. Identificador Vendedor:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2115
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msGrilla 
      Height          =   4395
      Left            =   60
      TabIndex        =   13
      Top             =   1620
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   7752
      _Version        =   393216
   End
   Begin VB.Menu mnu_operaciones 
      Caption         =   "&Operaciones"
      Begin VB.Menu mnu_operacion_borrarRegistracion 
         Caption         =   "&Borrar Registracion"
      End
   End
End
Attribute VB_Name = "frmConsultaVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const COL_ID = 19


Private Sub inicializarFormulario()

    
    Dim sql_tipoComprobante As String
    
    
    Me.dpFechaDesde.value = Format(fechaActualServer, "Short date")
    Me.dpFechaHasta.value = Format(fechaActualServer, "Short date")
     
    sql_tipoComprobante = "SELECT CODIGO, ISNULL(DESCRIPCION,'') AS DESCRIPCION, ISNULL(DEFECTO,0) AS DEFECTO FROM TIPO_COMPROBANTE WHERE ISNULL(VISIBLE,0) = 1 order by CODIGO "
    cargarCombo Me.cmbTipoComprobante, sql_tipoComprobante
    
    Me.cmbTipoComprobante.AddItem "TODAS", 0
    Me.cmbTipoComprobante.ListIndex = 0
   
    
End Sub

Private Sub cargarGrilla()
Dim ls_sql As String
Dim ls_where As String
Dim columna As Integer
Dim Ancho_Campo  As Integer
Dim fila As Integer
Dim rst As New ADODB.Recordset

On Error GoTo errores:


    ls_sql = "SELECT " + _
              "[ID. Vendedor], " + _
              "[Fecha de Venta] as [Fecha], " + _
              "[Total], " + _
              "[Tipo de Comprobante] as [Tipo Cbte], " + _
              "[Comprobante Desde] as [Cbte. Desde], " + _
              "[Comprobante Hasta] as [Cbte. Hasta], " + _
              "[Tipo Documento] as [T. Doc.], " + _
              "[ID. Comprador] as [Comprador], " + _
              "[Razon Social Comprador], " + _
              "[Neto Gravado 21], " + _
              "[Neto Gravado 10.5], " + _
              "[Neto Gravado 27], " + _
              "[IVA 21.0%], " + _
              "[IVA 10.5%], " + _
              "[IVA 27.0%], " + _
              "[exento], " + _
              "[Moneda], " + _
              "[Tipo Cambio], " + _
              "[Anulado], " + _
              "[VENTA_ID] AS ID " + _
            "FROM V_SUBDIARIO_VENTAS "

    
    ls_where = "WHERE [Fecha de Venta] between '" & Format(Me.dpFechaDesde, formatoFechaQuery) & "' and '" & Format(Me.dpFechaHasta, formatoFechaQuery) & "' "
        
    ls_where = ls_where & IIf(Len(Me.txtNroIdVendedor.ClipText) > 0, "AND [ID. Vendedor] = " & Me.txtNroIdVendedor.ClipText & " ", "")
        
    If Me.cmbTipoComprobante.ListIndex > 0 Then
        Dim ls_tmp() As String
            ls_tmp = Utils.separarCodigoDescripcion(Me.cmbTipoComprobante)
            ls_where = ls_where & "AND TIPO_COMPROBANTE_ID ='" & ls_tmp(0) & "' "
    End If
    
    ls_sql = ls_sql & ls_where & "ORDER BY [ID. Vendedor], [Fecha de Venta], [Total] desc"
    
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

Private Sub cmdExcel_Click()
Dim ls_sql As String
Dim ls_where As String




    ls_sql = "SELECT " + _
              "[ID. Vendedor], " + _
              "[Fecha de Venta] as [Fecha], " + _
              "[Total], " + _
              "[Tipo de Comprobante] as [Tipo Cbte], " + _
              "[Comprobante Desde] as [Cbte. Desde], " + _
              "[Comprobante Hasta] as [Cbte. Hasta], " + _
              "[Tipo Documento] as [T. Doc.], " + _
              "[ID. Comprador] as [Comprador], " + _
              "[Razon Social Comprador], " + _
              "[Neto Gravado 21], " + _
              "[Neto Gravado 10.5], " + _
              "[Neto Gravado 27], " + _
              "[IVA 21.0%], " + _
              "[IVA 10.5%], " + _
              "[IVA 27.0%], " + _
              "[exento], " + _
              "[Moneda], " + _
              "[Tipo Cambio], " + _
              "[Anulado], " + _
              "[VENTA_ID] AS ID " + _
            "FROM V_SUBDIARIO_VENTAS "

    
    ls_where = "WHERE [Fecha de Venta] between '" & Format(Me.dpFechaDesde, formatoFechaQuery) & "' and '" & Format(Me.dpFechaHasta, formatoFechaQuery) & "' "
        
    ls_where = ls_where & IIf(Len(Me.txtNroIdVendedor.ClipText) > 0, "AND [ID. Vendedor] = " & Me.txtNroIdVendedor.ClipText & " ", "")
        
    If Me.cmbTipoComprobante.ListIndex > 0 Then
        Dim ls_tmp() As String
            ls_tmp = Utils.separarCodigoDescripcion(Me.cmbTipoComprobante)
            ls_where = ls_where & "AND TIPO_COMPROBANTE_ID ='" & ls_tmp(0) & "' "
    End If
    
    ls_sql = ls_sql & ls_where & "ORDER BY [ID. Vendedor], [Fecha de Venta], [Total] desc"
    
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

Private Sub mnu_operacion_borrarRegistracion_Click()
    If msGrilla.Row <= 0 Then
        MsgBox "Debe seleccionar una fila para borrar.", vbCritical, "Error: Al Borrar Registracion"
    ElseIf msGrilla.Row >= 1 And msGrilla.Row <= msGrilla.Rows - 2 Then
        If MsgBox("Seguro desea eliminar la registración que se encuentra seleccionada", vbQuestion + vbYesNoCancel, "Confirmar Borrar Registración") = vbYes Then
            Dim ID As Double
            ID = msGrilla.TextMatrix(msGrilla.Row, COL_ID)
            If borrarCabeceraDetalle(ID) > 0 Then
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
        frmVentas.operacion = "M"
        frmVentas.idVentas = ID
        ID = msGrilla.Row
        frmVentas.Show vbModal
        cmdRecuperar_Click
        msGrilla.Row = ID
        msGrilla.RowSel = 1
        msGrilla.Col = 0
        msGrilla.ColSel = 15
    End If
End Sub

Private Sub txtNroIdVendedor_GotFocus()

    With txtNroIdVendedor
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

    txtRazonSocial.Text = ""
    
End Sub

Private Sub txtNroIdVendedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Dim indentificador As Double
        frmBusquedaEmpresa.inicializarFormulario indentificador
        txtNroIdVendedor.Mask = ""
        txtNroIdVendedor.Text = indentificador
        txtNroIdVendedor.PromptInclude = False
        txtNroIdVendedor.Mask = "##-########-#"
        txtNroIdVendedor.PromptInclude = True
        SendKeys "{tab}"
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
                txtRazonSocial.Text = ltEmpresa.razonSocial
                
            Else
                With txtNroIdVendedor
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
                End With
               
            End If
        Else
            With txtNroIdVendedor
                .SelStart = 0
                .SelLength = Len(.Text)
                .SetFocus
            End With
           
        End If
    End If
End Sub
