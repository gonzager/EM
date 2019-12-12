VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEmpresaABM 
   Caption         =   "Empresas"
   ClientHeight    =   8115
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   11550
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmDatos 
      Caption         =   "Datos de la Empresa"
      Enabled         =   0   'False
      Height          =   2835
      Left            =   60
      TabIndex        =   12
      Top             =   4740
      Width           =   11475
      Begin VB.CheckBox chkMono 
         Caption         =   "Monotibutista"
         Height          =   195
         Left            =   8340
         TabIndex        =   35
         Top             =   780
         Width           =   1335
      End
      Begin VB.Frame frmPuntosVtas 
         Caption         =   "Puntos de Venta"
         Enabled         =   0   'False
         Height          =   1575
         Left            =   120
         TabIndex        =   23
         Top             =   1140
         Width           =   11235
         Begin VB.CommandButton cmdNuevoPV 
            Caption         =   "&Nuevo"
            Height          =   435
            Left            =   4860
            TabIndex        =   32
            Top             =   300
            Width           =   1335
         End
         Begin VB.CommandButton cmdModificarPV 
            Caption         =   "&Modificar"
            Height          =   435
            Left            =   4860
            TabIndex        =   33
            Top             =   840
            Width           =   1335
         End
         Begin VB.Frame frmDatosPV 
            Caption         =   "Datos Puntos de Venta"
            Enabled         =   0   'False
            Height          =   1335
            Left            =   6300
            TabIndex        =   24
            Top             =   120
            Width           =   4815
            Begin VB.CheckBox chkActivoPV 
               Caption         =   "Activo"
               Height          =   255
               Left            =   240
               TabIndex        =   28
               Top             =   960
               Width           =   1335
            End
            Begin VB.CommandButton cmdGrabarPV 
               Caption         =   "G&rabar Punto Venta"
               Height          =   435
               Left            =   2880
               TabIndex        =   29
               Top             =   780
               Width           =   1845
            End
            Begin VB.TextBox txtIDPV 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   3960
               TabIndex        =   25
               Top             =   180
               Width           =   735
            End
            Begin MSMask.MaskEdBox txtPuntoVta 
               Height          =   315
               Left            =   1500
               TabIndex        =   26
               Top             =   540
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   5
               Mask            =   "#####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtNroIdVendedorPV 
               Height          =   315
               Left            =   2340
               TabIndex        =   27
               Top             =   180
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   13
               Mask            =   "##-########-#"
               PromptChar      =   "_"
            End
            Begin VB.Label lblPuntoDeVenta 
               Caption         =   "Punto de Venta"
               Height          =   255
               Left            =   240
               TabIndex        =   31
               Top             =   600
               Width           =   1275
            End
            Begin VB.Label lblIdentificadorVendedorPV 
               Caption         =   " Nro. Identificador Vendedor:"
               Height          =   255
               Left            =   180
               TabIndex        =   30
               Top             =   240
               Width           =   2115
            End
         End
         Begin MSFlexGridLib.MSFlexGrid grillaPuntoVta 
            Height          =   1095
            Left            =   180
            TabIndex        =   34
            Top             =   240
            Width           =   4635
            _ExtentX        =   8176
            _ExtentY        =   1931
            _Version        =   393216
         End
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar Empresa"
         Height          =   435
         Left            =   9720
         TabIndex        =   19
         Top             =   660
         Width           =   1545
      End
      Begin VB.CheckBox chkActivo 
         Caption         =   "Activo"
         Height          =   195
         Left            =   7500
         TabIndex        =   18
         Top             =   780
         Width           =   915
      End
      Begin VB.CheckBox chkComprador 
         Caption         =   "Comprador"
         Height          =   195
         Left            =   6300
         TabIndex        =   17
         Top             =   780
         Width           =   1155
      End
      Begin VB.CheckBox chkVendedor 
         Caption         =   "Vendedor"
         Height          =   195
         Left            =   5220
         TabIndex        =   16
         Top             =   780
         Width           =   1035
      End
      Begin VB.TextBox txtDomicilio 
         Height          =   315
         Left            =   900
         MaxLength       =   255
         TabIndex        =   15
         Top             =   720
         Width           =   4215
      End
      Begin VB.TextBox txtRZ 
         Height          =   315
         Left            =   5040
         MaxLength       =   255
         TabIndex        =   14
         Top             =   240
         Width           =   6195
      End
      Begin MSMask.MaskEdBox txtNroIdVendedor 
         Height          =   315
         Left            =   2340
         TabIndex        =   13
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   13
         Mask            =   "##-########-#"
         PromptChar      =   "_"
      End
      Begin VB.Label lblDomicilio 
         Caption         =   "Domicilio:"
         Height          =   255
         Left            =   180
         TabIndex        =   22
         Top             =   780
         Width           =   735
      End
      Begin VB.Label lblIdentificadorVendedor 
         Caption         =   " Nro. Identificador Vendedor:"
         Height          =   255
         Left            =   180
         TabIndex        =   21
         Top             =   300
         Width           =   2115
      End
      Begin VB.Label lblRZ 
         Caption         =   "Razón Social:"
         Height          =   195
         Left            =   3960
         TabIndex        =   20
         Top             =   300
         Width           =   1035
      End
   End
   Begin VB.Frame frmOperaciones 
      Caption         =   "Operaciones"
      Height          =   915
      Left            =   60
      TabIndex        =   9
      Top             =   3780
      Width           =   11475
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Modificar"
         Height          =   435
         Left            =   2160
         TabIndex        =   11
         Top             =   300
         Width           =   1695
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   435
         Left            =   240
         TabIndex        =   10
         Top             =   300
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   435
      Left            =   9780
      TabIndex        =   8
      Top             =   7620
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid msGrilla 
      Height          =   2595
      Left            =   60
      TabIndex        =   7
      Top             =   1140
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   4577
      _Version        =   393216
      HighLight       =   2
      ScrollBars      =   2
   End
   Begin VB.Frame frmCriterios 
      Caption         =   "Criterios de Búsqueda"
      Height          =   1095
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   11475
      Begin VB.OptionButton opTodos 
         Caption         =   "Comprador"
         Height          =   195
         Index           =   2
         Left            =   7920
         TabIndex        =   3
         Top             =   780
         Width           =   1275
      End
      Begin VB.OptionButton opTodos 
         Caption         =   "Vendedor"
         Height          =   195
         Index           =   1
         Left            =   7920
         TabIndex        =   2
         Top             =   480
         Width           =   1275
      End
      Begin VB.OptionButton opTodos 
         Caption         =   "Todos"
         Height          =   195
         Index           =   0
         Left            =   7920
         TabIndex        =   1
         Top             =   180
         Width           =   1275
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   435
         Left            =   9480
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtRazonSocial 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Top             =   420
         Width           =   6135
      End
      Begin VB.Label lblRazonSocial 
         Caption         =   "Razón Social:"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   480
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmEmpresaABM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const COL_ID = 1
Private Const COL_RAZONSOCIAL = 2
Private Const COL_DOMICILIO = 3
Private Const COL_VENDEDOR = 4
Private Const COL_COMPRADOR = 5
Private Const COL_ACTIVO = 6
Private Const COL_MONO = 7
Private Const COL_PUNTOVENTA = 2
Private Const COL_ACTIVOPV = 3
Private Const COL_IDPV = 4

Private operacion As String
Private operacionPV As String




Public Sub inicializarFormulario()
    inicializarGrilla
    Me.opTodos(0).value = 1
    Me.Show vbModal
  
End Sub


Private Sub cmdBuscar_Click()
    Dim ls_sql As String
    Dim oRec As ADODB.Recordset
    On Error GoTo errores
    Set oRec = New ADODB.Recordset
    Dim tiene As Boolean
    
    ls_sql = "SELECT IDENTIFICADOR, RAZONSOCIAL, ISNULL(DOMICILIO,'') DOMICILIO, " + _
             "CASE WHEN VENDEDOR=1 THEN 'SI' ELSE 'NO' END VENDEDOR, " + _
             "CASE WHEN COMPRADOR=1 THEN 'SI' ELSE 'NO' END COMPRADOR, " + _
             "CASE WHEN ACTIVO=1 THEN 'SI' ELSE 'NO' END ACTIVO, " + _
             "CASE WHEN MONOTRIBUTISTA=1 THEN 'SI' ELSE 'NO' END MONOTRIBUTISTA " + _
             "FROM EMPRESA WHERE RAZONSOCIAL LIKE '%" & Me.txtRazonSocial.Text & "%' " + _
             ""
    If Me.opTodos(1).value = True Then
        ls_sql = ls_sql & " AND VENDEDOR=1"
    ElseIf Me.opTodos(2).value = True Then
        ls_sql = ls_sql & " AND COMPRADOR=1"
    End If
    
    With oRec
        .ActiveConnection = oCon
        .CursorType = adOpenStatic
        .CursorLocation = adUseClient
        .Source = ls_sql
        .Open
    End With
    
    tiene = False
    inicializarGrilla
    Do While Not oRec.EOF
        With msGrilla
            .TextMatrix(.Rows - 1, COL_ID) = oRec!Identificador
            .TextMatrix(.Rows - 1, COL_RAZONSOCIAL) = oRec!razonSocial
            .TextMatrix(.Rows - 1, COL_DOMICILIO) = oRec!domicilio
            .TextMatrix(.Rows - 1, COL_VENDEDOR) = oRec!vendedor
            .TextMatrix(.Rows - 1, COL_COMPRADOR) = oRec!comprador
            .TextMatrix(.Rows - 1, COL_ACTIVO) = oRec!Activo
            .TextMatrix(.Rows - 1, COL_MONO) = oRec!Monotributista
            .Rows = .Rows + 1
        End With
        tiene = True
        oRec.MoveNext
    Loop
    
    msGrilla.RowSel = 1
    msGrilla.ColSel = COL_ID
        
    If tiene Then
        msGrilla.ColSel = COL_ACTIVO
        msGrilla.SetFocus
    End If
    msGrilla_EnterCell

    Set oRec = Nothing
    Exit Sub
errores:
    If oCon.Errors.Count >= 1 Then
        MsgBox oCon.Errors(0).Description, vbCritical, "Nro de Error: " & oCon.Errors(0).Number
        oCon.Errors.Clear
    Else
        MsgBox Err.Number & "-" & Err.Description, vbCritical, "Error: Al insertar cabecera de ventas"
        Err.Clear
    End If

    Set oRec = Nothing

End Sub

Private Sub cmdGrabar_Click()
Dim ltEmpresa As tEmpresa
            
    
    If Me.txtNroIdVendedor.ClipText <> "" Then

        ltEmpresa.Identificador = Me.txtNroIdVendedor.ClipText
        ltEmpresa.razonSocial = Trim(Me.txtRZ.Text)
        ltEmpresa.domicilio = Trim(Me.txtDomicilio.Text)
        ltEmpresa.vendedor = Me.chkVendedor.value
        ltEmpresa.comprador = Me.chkComprador.value
        ltEmpresa.Activo = Me.chkActivo.value
        ltEmpresa.Monotributista = Me.chkMono.value
    
        If ltEmpresa.razonSocial <> "" And ltEmpresa.domicilio <> "" Then
            If operacion = "A" Then
                If Empresa.insertEmpresa(ltEmpresa) > 0 Then
                    MsgBox "La Empresa se genero correctamente", vbInformation, "Alta Empresa OK"
                    cmdBuscar_Click
                Else
                    Me.txtNroIdVendedor.Mask = ""
                    Me.txtNroIdVendedor.PromptInclude = False
                    Me.txtNroIdVendedor.Mask = "##-########-#"
                    Me.txtNroIdVendedor.Text = ""
                    Me.txtNroIdVendedor.PromptInclude = True
                    Me.txtNroIdVendedor.Enabled = True
                    
                    Me.txtRZ.Text = ""
                    Me.txtDomicilio.Text = ""
                    Me.chkVendedor.value = 0
                    Me.chkComprador.value = 0
                    Me.chkActivo.value = 0
                    Me.chkMono.value = 0
                End If
            
            ElseIf operacion = "M" Then
                If Empresa.updateEmpresa(ltEmpresa) > 0 Then
                    MsgBox "Los datos de la empresa se cambiaron correctamente", vbInformation, "Modificacion Empresa OK"
                    cmdBuscar_Click
                End If
            End If
            Me.frmDatos.Enabled = False
            Me.frmPuntosVtas.Enabled = False
            operacion = ""
            operacionPV = ""
        Else
            MsgBox "Deber ingresar los datos de Razón Social y Domicilio", vbCritical, "Error: Carga de Datos"
            Me.txtRazonSocial.SetFocus
        End If
    Else
        Me.txtNroIdVendedor.SetFocus
    End If
    
End Sub

Private Sub cmdGrabarPV_Click()
    Dim ltPuntoVenta As tPuntoDeVenta
    
    ltPuntoVenta.puntoVentaId = CInt(Me.txtIDPV.Text)
    ltPuntoVenta.empresa_id = Trim(Me.txtNroIdVendedorPV.ClipText)
    ltPuntoVenta.PuntoDeVenta = Trim(Me.txtPuntoVta.ClipText)
    ltPuntoVenta.Activo = Me.chkActivoPV.value
    
    If operacionPV = "A" Then
        If Len(ltPuntoVenta.PuntoDeVenta) > 0 Then
            
            If recuperarIDPuntoDeVenta(ltPuntoVenta.empresa_id, ltPuntoVenta.PuntoDeVenta) = 0 Then
                If PuntoDeVenta.insertPuntoDeVenta(ltPuntoVenta) > 0 Then
                    cargarPuntosDeVenta ltPuntoVenta.empresa_id
                    grillaPuntoVta_EnterCell
                    Me.frmDatosPV.Enabled = False
                Else
                    limpiarDatosPV
                
                End If
            Else
                MsgBox "El Punto de Venta " & ltPuntoVenta.PuntoDeVenta & " ya existe para " & ltPuntoVenta.empresa_id, vbCritical, "Error: Punto de Venta Existente"
                Me.txtPuntoVta.SetFocus
            End If
            
        Else
            MsgBox "Debe ingresar ingresar un punto de Venta para continuar", vbCritical, "Error: Punto de Venta"
            Me.txtPuntoVta.SetFocus
        End If
        
    ElseIf operacionPV = "M" Then
        If Len(ltPuntoVenta.PuntoDeVenta) > 0 Then
            If PuntoDeVenta.updatePuntoDeVenta(ltPuntoVenta) > 0 Then
                cargarPuntosDeVenta ltPuntoVenta.empresa_id
                grillaPuntoVta_EnterCell
                Me.frmDatosPV.Enabled = False
            End If
        Else
            MsgBox "Debe ingresar ingresar un punto de Venta para continuar", vbCritical, "Error: Punto de Venta"
            Me.txtPuntoVta.SetFocus
        End If
    End If
    
End Sub

Private Sub cmdModificar_Click()

    If msGrilla.Row <= 0 Then
        MsgBox "Debe seleccionar una fila para seleccionar.", vbCritical, "Error: Al seleccionar registro"
    ElseIf msGrilla.Row >= 1 And msGrilla.Row <= msGrilla.Rows - 2 Then

         operacion = "M"
         habilitarOperacion operacion
         
    End If
End Sub

Private Sub cmdModificarPV_Click()
    operacionPV = "M"
    Me.frmDatosPV.Enabled = True
End Sub

Private Sub cmdNuevo_Click()
    operacion = "A"
    habilitarOperacion operacion
    inciciliarGrillaPV
    limpiarDatosPV
End Sub

Private Sub cmdNuevoPV_Click()
    operacionPV = "A"
    limpiarDatosPV
    
    Me.txtPuntoVta.SetFocus

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub


Private Sub grillaPuntoVta_EnterCell()
Dim ID As Double
    If grillaPuntoVta.Row <= 0 Then
        MsgBox "Debe seleccionar una fila para seleccionar.", vbCritical, "Error: Al seleccionar registro"
    ElseIf grillaPuntoVta.Row >= 1 And grillaPuntoVta.Row <= grillaPuntoVta.Rows - 1 Then
        Me.frmDatosPV.Enabled = False
        cargarDatosPuntoDeVenta grillaPuntoVta.Row

    End If
End Sub


Private Sub msGrilla_EnterCell()
Dim ID As Double
    If msGrilla.Row <= 0 Then
        MsgBox "Debe seleccionar una fila para seleccionar.", vbCritical, "Error: Al seleccionar registro"
    ElseIf msGrilla.Row >= 1 And msGrilla.Row <= msGrilla.Rows - 1 Then
        Me.frmDatos.Enabled = False
        Me.frmPuntosVtas.Enabled = False
        ID = CDbl(IIf(msGrilla.TextMatrix(msGrilla.Row, COL_ID) <> "", msGrilla.TextMatrix(msGrilla.Row, COL_ID), 0))
        cargarRegistroAModificar ID
        cargarPuntosDeVenta ID
        grillaPuntoVta_EnterCell

    End If
End Sub


Private Sub txtDomicilio_GotFocus()
    With txtDomicilio
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtNroIdVendedor_GotFocus()
    With txtNroIdVendedor
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtNroIdVendedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then

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
'On Error Resume Next
    If Len(Me.txtNroIdVendedor.ClipText) > 0 And Me.frmDatos.Enabled Then
        If Len(Me.txtNroIdVendedor.ClipText) = 11 Then
            If Not ValidarCuit(Me.txtNroIdVendedor.ClipText) Then
                MsgBox "El IDENTIFICADOR ingresado no es válido", vbCritical, "Error: Código de Verificación"
                Me.txtNroIdVendedor.SetFocus
                Me.txtRZ.Text = ""
            End If
        Else
            
            MsgBox "La longitud del campo para el numero de identificador no es valida", vbCritical, "Error: Longitud CUIT/CUIL"
            Me.txtNroIdVendedor.SetFocus
            Me.txtRZ.Text = ""
        
        End If
    End If
End Sub

Private Sub txtPuntoVta_GotFocus()
    With txtPuntoVta
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtPuntoVta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{tab}"
    End If
   
   If Not IsNumeric(Chr(KeyAscii)) Then
       If KeyAscii <> vbKeyBack Then
          KeyAscii = 0
       End If
    End If
End Sub

Private Sub txtPuntoVta_LostFocus()
    If Len(txtPuntoVta.ClipText) > 0 Then
        Me.txtPuntoVta.Text = Format(Me.txtPuntoVta.ClipText, "0000#")
    End If
End Sub

Private Sub txtRazonSocial_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        cmdBuscar_Click
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
 
End Sub

Private Sub inicializarGrilla()
    msGrilla.Clear
    msGrilla.Rows = 2
    With msGrilla
        .Cols = 8
        .ColWidth(0) = 300
        .ColWidth(COL_ID) = 1250
        .ColWidth(COL_RAZONSOCIAL) = 3030
        .ColWidth(COL_DOMICILIO) = 2600
        .ColWidth(COL_VENDEDOR) = 1100
        .ColWidth(COL_COMPRADOR) = 1200
        .ColWidth(COL_ACTIVO) = 700
        .ColWidth(COL_MONO) = 1500
     
        .TextMatrix(0, COL_ID) = "IDENTIFICACION"
        .TextMatrix(0, COL_RAZONSOCIAL) = "RAZON SOCIAL"
        .TextMatrix(0, COL_DOMICILIO) = "DOMICILIO"
        .TextMatrix(0, COL_VENDEDOR) = "VENDEDOR"
        .TextMatrix(0, COL_COMPRADOR) = "COMPRADOR"
        .TextMatrix(0, COL_ACTIVO) = "ACTIVO"
        .TextMatrix(0, COL_MONO) = "MONOTRIBUTISTA"
    End With

    msGrilla.ColSel = COL_ACTIVO
    inciciliarGrillaPV
  
    operacion = ""
    operacionPV = ""
End Sub

Private Sub inciciliarGrillaPV()
    grillaPuntoVta.Clear
    grillaPuntoVta.Rows = 2
    With grillaPuntoVta
        .Cols = 5
        .ColWidth(0) = 300
        .ColWidth(COL_ID) = 1500
        .ColWidth(COL_PUNTOVENTA) = 1600
        .ColWidth(COL_ACTIVOPV) = 800
        .ColWidth(COL_IDPV) = 0
        
        .TextMatrix(0, COL_ID) = "IDENTIFICACION"
        .TextMatrix(0, COL_PUNTOVENTA) = "PUNTO DE VENTA"
        .TextMatrix(0, COL_ACTIVOPV) = "ACTIVO"
        .TextMatrix(0, COL_IDPV) = "IDPV"
    End With
    grillaPuntoVta.ColSel = COL_IDPV
End Sub

Private Sub habilitarOperacion(ls_operacion As String)

    Me.frmDatos.Enabled = True
    
    If ls_operacion = "A" Then
        
        Me.txtNroIdVendedor.Mask = ""
        Me.txtNroIdVendedor.PromptInclude = False
        Me.txtNroIdVendedor.Mask = "##-########-#"
        Me.txtNroIdVendedor.Text = ""
        Me.txtNroIdVendedor.PromptInclude = True
        Me.txtNroIdVendedor.Enabled = True
        
        Me.txtRZ.Text = ""
        Me.txtDomicilio.Text = ""
        Me.chkVendedor.value = 0
        Me.chkComprador.value = 0
        Me.chkActivo.value = 1
        Me.chkMono.value = 0
        Me.txtNroIdVendedor.SetFocus
        
    ElseIf ls_operacion = "M" Then
         Me.frmPuntosVtas.Enabled = True
         Me.txtRZ.SetFocus
    End If
    
End Sub

Private Sub cargarRegistroAModificar(ID As Double)
    Dim ls_sql As String
    Dim oRec As ADODB.Recordset
    On Error GoTo errores
    Set oRec = New ADODB.Recordset
    
    
    ls_sql = "SELECT IDENTIFICADOR, RAZONSOCIAL, ISNULL(DOMICILIO,'') DOMICILIO, VENDEDOR, COMPRADOR, ACTIVO, MONOTRIBUTISTA " + _
             "FROM EMPRESA WHERE IDENTIFICADOR = " & ID
             
    With oRec
        .ActiveConnection = oCon
        .CursorType = adOpenStatic
        .CursorLocation = adUseClient
        .Source = ls_sql
        .Open
    End With
        
    If oRec.RecordCount = 1 Then
        Me.txtNroIdVendedor.Mask = ""
        Me.txtNroIdVendedor.PromptInclude = False
        Me.txtNroIdVendedor.Mask = "##-########-#"
        Me.txtNroIdVendedor.Text = oRec!Identificador
        Me.txtNroIdVendedor.PromptInclude = True
        
        Me.txtNroIdVendedor.Enabled = False
        
        Me.txtRZ.Text = oRec!razonSocial
        Me.txtDomicilio.Text = oRec!domicilio
        Me.chkVendedor.value = IIf(oRec!vendedor, 1, 0)
        Me.chkComprador.value = IIf(oRec!comprador, 1, 0)
        Me.chkActivo.value = IIf(oRec!Activo, 1, 0)
        Me.chkMono.value = IIf(oRec!Monotributista, 1, 0)
    Else
        Me.txtNroIdVendedor.Mask = ""
        Me.txtNroIdVendedor.PromptInclude = False
        Me.txtNroIdVendedor.Mask = "##-########-#"
        Me.txtNroIdVendedor.Text = ""
        Me.txtNroIdVendedor.PromptInclude = True
        
        Me.txtNroIdVendedor.Enabled = False
        
        Me.txtRZ.Text = ""
        Me.txtDomicilio.Text = ""
        Me.chkVendedor.value = 0
        Me.chkComprador.value = 0
        Me.chkActivo.value = 0
        Me.chkMono.value = 0
    End If
    
    
    Set oRec = Nothing
    Exit Sub
errores:
    If oCon.Errors.Count >= 1 Then
        MsgBox oCon.Errors(0).Description, vbCritical, "Nro de Error: " & oCon.Errors(0).Number
        oCon.Errors.Clear
    Else
        MsgBox Err.Number & "-" & Err.Description, vbCritical, "Error: Al insertar cabecera de ventas"
        Err.Clear
    End If

    Set oRec = Nothing
End Sub


Private Sub cargarPuntosDeVenta(ID As Double)
    Dim ltPuntoDeVenta() As tPuntoDeVenta
    Dim lb_Existe As Boolean
    Dim i As Integer
    ltPuntoDeVenta = recuperarRegistroPuntoVtas(ID, lb_Existe)
    inciciliarGrillaPV
    If lb_Existe Then
        For i = 0 To UBound(ltPuntoDeVenta)
            With grillaPuntoVta
                .TextMatrix(i + 1, COL_ID) = ltPuntoDeVenta(i).empresa_id
                .TextMatrix(i + 1, COL_PUNTOVENTA) = ltPuntoDeVenta(i).PuntoDeVenta
                .TextMatrix(i + 1, COL_ACTIVOPV) = IIf(ltPuntoDeVenta(i).Activo, "SI", "NO")
                .TextMatrix(i + 1, COL_IDPV) = ltPuntoDeVenta(i).puntoVentaId
                .Rows = .Rows + 1
            End With
        Next
    End If
    grillaPuntoVta.Col = 1
End Sub

Private Sub txtRZ_GotFocus()
    With txtRZ
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub cargarDatosPuntoDeVenta(iRow As Integer)

    Me.txtIDPV.Text = Me.grillaPuntoVta.TextMatrix(iRow, COL_IDPV)
    
    Me.txtNroIdVendedorPV.Mask = ""
    Me.txtNroIdVendedorPV.PromptInclude = False
    Me.txtNroIdVendedorPV.Mask = "##-########-#"
    Me.txtNroIdVendedorPV.Text = Me.grillaPuntoVta.TextMatrix(iRow, COL_ID)
    Me.txtNroIdVendedorPV.PromptInclude = True
    

    Me.txtPuntoVta.Mask = ""
    Me.txtPuntoVta.PromptInclude = False
    Me.txtPuntoVta.Mask = "#####"
    Me.txtPuntoVta.Text = Format(Me.grillaPuntoVta.TextMatrix(iRow, COL_PUNTOVENTA), "0000#")
    Me.txtPuntoVta.PromptInclude = True
    
    Me.chkActivoPV.value = IIf(Me.grillaPuntoVta.TextMatrix(iRow, COL_ACTIVOPV) = "SI", 1, 0)

End Sub

Private Sub limpiarDatosPV()
    Me.frmDatosPV.Enabled = True

    Me.txtIDPV.Text = 0
    
    Me.txtNroIdVendedorPV.Mask = ""
    Me.txtNroIdVendedorPV.PromptInclude = False
    Me.txtNroIdVendedorPV.Mask = "##-########-#"
    Me.txtNroIdVendedorPV.Text = Me.txtNroIdVendedor.ClipText
    Me.txtNroIdVendedorPV.PromptInclude = True
    

    Me.txtPuntoVta.Mask = ""
    Me.txtPuntoVta.PromptInclude = False
    Me.txtPuntoVta.Mask = "#####"
    Me.txtPuntoVta.Text = ""
    Me.txtPuntoVta.PromptInclude = True
    
    Me.chkActivoPV.value = 1
End Sub
