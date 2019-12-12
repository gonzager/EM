VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAfipVentasDuplicadas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comprobantes de Ventas Duplicados"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   11550
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9900
      TabIndex        =   2
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "&Excel"
      Height          =   375
      Left            =   8160
      TabIndex        =   1
      Top             =   4440
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid msGrilla 
      Height          =   4395
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   7752
      _Version        =   393216
   End
   Begin VB.Label lblTotal 
      Caption         =   "Total de Comprobantes: 0"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Menu mnu_operaciones 
      Caption         =   "&Operaciones"
      Begin VB.Menu mnu_borrarRegistracion 
         Caption         =   "Borrar Registración"
      End
      Begin VB.Menu mnu_borrarTodas 
         Caption         =   "Borrar Todas las Registraciones"
      End
   End
End
Attribute VB_Name = "frmAfipVentasDuplicadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const COL_ID = 19
Private l_txtNroIdVendedor As String
Public Sub cargarGrilla(txtNroIdVendedor As String)
Dim ls_sql As String
Dim ls_where As String
Dim columna As Integer
Dim Ancho_Campo  As Integer
Dim fila As Integer
Dim rst As New ADODB.Recordset

On Error GoTo errores:

    l_txtNroIdVendedor = txtNroIdVendedor
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
              "[Exento], " + _
              "[Moneda], " + _
              "[Tipo Cambio], " + _
              "[Anulado], " + _
              "[VENTA_ID] AS ID " + _
            "FROM V_SUBDIARIO_VENTAS VV "

    
    ls_where = "WHERE exists ( " + _
               " SELECT 1 FROM I_VENTAS IV WHERE " + _
               " IV.TIPOCOMPROBANTE = VV.TIPO_COMPROBANTE_ID AND " + _
               " IV.PUNTOVENTA = LEFT(VV.[Comprobante Desde],5) AND " + _
               " RIGHT(IV.NUMEROCOMPROBANTE,8) = RIGHT(VV.[Comprobante Desde],8) ) "
               
                         
        
    ls_where = ls_where & " AND [ID. Vendedor] = " & l_txtNroIdVendedor
        
    
    ls_sql = ls_sql & ls_where & " ORDER BY [Comprobante Desde]"
    
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
    
    lblTotal.Caption = "Total de Comprobantes: " & msGrilla.Rows - 2
    
    Set rst = Nothing
    Exit Sub
    
errores:
    On Error Resume Next
    If oCon.Errors.Count >= 1 Then
        lblTotal.Caption = "Total de Comprobantes: "
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
              "[Exento], " + _
              "[Moneda], " + _
              "[Tipo Cambio], " + _
              "[Anulado], " + _
              "[VENTA_ID] AS ID " + _
            "FROM V_SUBDIARIO_VENTAS VV "

    
    ls_where = "WHERE exists ( " + _
               " SELECT 1 FROM I_VENTAS IV WHERE " + _
               " IV.TIPOCOMPROBANTE = VV.TIPO_COMPROBANTE_ID AND " + _
               " IV.PUNTOVENTA = LEFT(VV.[Comprobante Desde],5) AND " + _
               " RIGHT(IV.NUMEROCOMPROBANTE,8) = RIGHT(VV.[Comprobante Desde],8) ) "
               
                         
        
    ls_where = ls_where & " AND [ID. Vendedor] = " & l_txtNroIdVendedor
        
    
    ls_sql = ls_sql & ls_where & " ORDER BY [Comprobante Desde]"
    
    UtilsExcel.Exportar_ADO_Excel_RecordSet ls_sql
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub mnu_borrarRegistracion_Click()
    If msGrilla.Row <= 0 Then
        MsgBox "Debe seleccionar una fila para borrar.", vbCritical, "Error: Al Borrar Registracion"
    ElseIf msGrilla.Row >= 1 And msGrilla.Row <= msGrilla.Rows - 2 Then
        If MsgBox("Seguro desea eliminar la registración que se encuentra seleccionada", vbQuestion + vbYesNoCancel, "Confirmar Borrar Registración") = vbYes Then
            Dim ID As Double
            ID = msGrilla.TextMatrix(msGrilla.Row, COL_ID)
            If borrarCabeceraDetalle(ID) > 0 Then
                ID = msGrilla.Row
                cargarGrilla l_txtNroIdVendedor
                msGrilla.Row = ID
                msGrilla.RowSel = 1
                msGrilla.Col = 0
                msGrilla.ColSel = 13
                                
            End If
        End If
       
    End If
End Sub

Private Sub mnu_borrarTodas_Click()
    If msGrilla.Row <= 0 Then
        MsgBox "Debe seleccionar una fila para borrar.", vbCritical, "Error: Al Borrar Registracion"
    ElseIf msGrilla.Row >= 1 And msGrilla.Row <= msGrilla.Rows - 2 Then
        If MsgBox("Seguro desea eliminar TODAS las registraciones cargadas", vbQuestion + vbYesNoCancel, "Confirmar Borrar Registración") = vbYes Then
            Screen.MousePointer = vbHourglass
            Dim ID As Double
            Dim i As Integer
            For i = 1 To msGrilla.Rows - 2
                ID = msGrilla.TextMatrix(i, COL_ID)
                borrarCabeceraDetalle ID
            Next i
            cargarGrilla l_txtNroIdVendedor
            Screen.MousePointer = vbDefault
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
        cargarGrilla l_txtNroIdVendedor
        msGrilla.Row = ID
        msGrilla.RowSel = 1
        msGrilla.Col = 0
        msGrilla.ColSel = 15
    End If
End Sub
