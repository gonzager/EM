VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmConceptoVtasABM 
   Caption         =   "Conceptos de Ventas"
   ClientHeight    =   6930
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   7770
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   5670
      TabIndex        =   12
      Top             =   6360
      Width           =   1995
   End
   Begin VB.Frame frmDatos 
      Caption         =   "Datos Concepto"
      Height          =   1725
      Left            =   0
      TabIndex        =   4
      Top             =   4590
      Width           =   7695
      Begin VB.CommandButton cmdCancela 
         Caption         =   "Cancela"
         Height          =   495
         Left            =   3480
         TabIndex        =   13
         Top             =   1110
         Width           =   1995
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar Concepto"
         Height          =   495
         Left            =   5580
         TabIndex        =   11
         Top             =   1110
         Width           =   1995
      End
      Begin VB.CheckBox chkVisible 
         Caption         =   "Visible"
         Height          =   315
         Left            =   270
         TabIndex        =   9
         Top             =   720
         Width           =   1425
      End
      Begin VB.CheckBox chkDefecto 
         Caption         =   "Defecto"
         Height          =   315
         Left            =   1830
         TabIndex        =   8
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   1830
         MaxLength       =   255
         TabIndex        =   7
         Top             =   330
         Width           =   5775
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   840
         MaxLength       =   3
         TabIndex        =   6
         Top             =   330
         Width           =   855
      End
      Begin VB.Label lblRemember 
         Caption         =   "Recuerde que solo puede exisiter un único registro por Defecto"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3030
         TabIndex        =   10
         Top             =   750
         Width           =   4485
      End
      Begin VB.Label lblCodigo 
         Caption         =   "Código:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame fmrOperaciones 
      Caption         =   "Operaciones"
      Height          =   885
      Left            =   0
      TabIndex        =   1
      Top             =   3720
      Width           =   7695
      Begin VB.CommandButton Modificar 
         Caption         =   "Modificar"
         Height          =   495
         Left            =   2160
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Nuevo"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msGrilla 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   6588
      _Version        =   393216
      HighLight       =   2
      ScrollBars      =   2
   End
End
Attribute VB_Name = "frmConceptoVtasABM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const COL_CODIGO = 1
Private Const COL_DESCRIPCION = 2
Private Const COL_DEFECTO = 3
Private Const COL_VISIBLE = 4
Private operacion As String


Private Sub cmdCancela_Click()
    Me.frmDatos.Enabled = False
    txtCodigo.Enabled = True
    fmrOperaciones.Enabled = True
    msGrilla_EnterCell
End Sub

Private Sub cmdGrabar_Click()
Dim ltConcepto As tConceptos
    
            
    
    If Me.txtCodigo.Text <> "" Then

        ltConcepto.codigo = Me.txtCodigo.Text
        ltConcepto.descripcion = Me.txtDescripcion.Text
        ltConcepto.defecto = Me.chkDefecto
        ltConcepto.visible = Me.chkVisible

    
        If ltConcepto.codigo <> "" Then
            If operacion = "A" Then
                If Concepto.insertConcepto(ltConcepto, "VENTAS") > 0 Then
                    MsgBox "La Concepto se genero correctamente", vbInformation, "Alta Conceptos"
                    fmrOperaciones.Enabled = True
                    cargarGrilla
                Else
                   Me.txtCodigo.SetFocus

                End If
            
            ElseIf operacion = "M" Then
                If Concepto.updateConcepto(ltConcepto, "VENTAS") > 0 Then
                    MsgBox "Los datos del concepto se cambiaron correctamente", vbInformation, "Modificacion Concepto OK"
                    fmrOperaciones.Enabled = True
                    cargarGrilla
                End If
            End If
            Me.frmDatos.Enabled = False
            Me.txtCodigo.Enabled = True

        Else
            MsgBox "Deber ingresar una descripción", vbCritical, "Error: Carga de Datos"
            Me.txtDescripcion.SetFocus
        End If
    Else
        Me.txtCodigo.SetFocus
    End If
End Sub

Private Sub cmdNuevo_Click()
    operacion = "A"
    habilitarOperacion operacion
    fmrOperaciones.Enabled = False
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    
End Sub

Private Sub Form_Load()
    cargarGrilla

End Sub

Private Sub cargarGrilla()
    Dim ls_sql As String
    Dim oRec As ADODB.Recordset
    On Error GoTo errores
    Set oRec = New ADODB.Recordset
    Dim tiene As Boolean
    
    ls_sql = "SELECT CODIGO, DESCRIPCION, DEFECTO, VISIBLE FROM CONCEPTO "
    
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
            .TextMatrix(.Rows - 1, COL_CODIGO) = oRec!codigo
            .TextMatrix(.Rows - 1, COL_DESCRIPCION) = oRec!descripcion
            .TextMatrix(.Rows - 1, COL_DEFECTO) = oRec!defecto
            .TextMatrix(.Rows - 1, COL_VISIBLE) = oRec!visible

            .Rows = .Rows + 1
        End With
        tiene = True
        oRec.MoveNext
    Loop
    
    msGrilla.RowSel = 1
    msGrilla.ColSel = COL_CODIGO
        
    If tiene Then
        msGrilla.ColSel = COL_CODIGO
        'msGrilla.SetFocus
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

Private Sub inicializarGrilla()
    msGrilla.Clear
    msGrilla.Rows = 2
    With msGrilla
        .Cols = 5
        .ColWidth(0) = 300
        .ColWidth(COL_CODIGO) = 1150
        .ColWidth(COL_DESCRIPCION) = 3700
        .ColWidth(COL_DEFECTO) = 1150
        .ColWidth(COL_VISIBLE) = 1150

     
        .TextMatrix(0, COL_CODIGO) = "CODIGO"
        .TextMatrix(0, COL_DESCRIPCION) = "DESCRIPCION"
        .TextMatrix(0, COL_DEFECTO) = "DEFECTO"
        .TextMatrix(0, COL_VISIBLE) = "VISIBLE"

    End With

    msGrilla.ColSel = COL_CODIGO

End Sub

Private Sub Modificar_Click()
    If msGrilla.Row <= 0 Then
        MsgBox "Debe seleccionar una fila para seleccionar.", vbCritical, "Error: Al seleccionar registro"
    ElseIf msGrilla.Row >= 1 And msGrilla.Row <= msGrilla.Rows - 2 Then

         operacion = "M"
         habilitarOperacion operacion
         fmrOperaciones.Enabled = False
         
    End If
End Sub

Private Sub habilitarOperacion(ls_operacion As String)

    Me.frmDatos.Enabled = True
    
    If ls_operacion = "A" Then
        Me.frmDatos.Enabled = True
        Me.txtCodigo.Text = fMaxCodigo("VENTAS")
        Me.txtDescripcion.Text = ""
        Me.txtDescripcion.SetFocus
        Me.chkDefecto.value = 0
        Me.chkVisible.value = 1
        
    ElseIf ls_operacion = "M" Then
         Me.frmDatos.Enabled = True
         Me.txtCodigo.Enabled = False
         Me.txtDescripcion.SelStart = 0
         Me.txtDescripcion.SelLength = Len(txtDescripcion)
         Me.txtDescripcion.SetFocus
    End If
    
End Sub

Private Sub msGrilla_EnterCell()
Dim codigo As String
    If msGrilla.Row <= 0 Then
        MsgBox "Debe seleccionar una fila para seleccionar.", vbCritical, "Error: Al seleccionar registro"
    ElseIf msGrilla.Row >= 1 And msGrilla.Row <= msGrilla.Rows - 1 Then
        Me.frmDatos.Enabled = False
        txtCodigo.Enabled = True
        codigo = IIf(msGrilla.TextMatrix(msGrilla.Row, COL_CODIGO) <> "", msGrilla.TextMatrix(msGrilla.Row, COL_CODIGO), "000")
        cargarRegistroAModificar codigo


    End If
End Sub




Private Sub cargarRegistroAModificar(codigo As String)
    Dim ls_sql As String
    Dim oRec As ADODB.Recordset
    On Error GoTo errores
    Set oRec = New ADODB.Recordset
    
    
    ls_sql = "SELECT CODIGO, DESCRIPCION, DEFECTO, VISIBLE FROM CONCEPTO  WHERE CODIGO = '" & codigo & "'"
             
    With oRec
        .ActiveConnection = oCon
        .CursorType = adOpenStatic
        .CursorLocation = adUseClient
        .Source = ls_sql
        .Open
    End With
        
    If oRec.RecordCount = 1 Then
        Me.txtCodigo = oRec!codigo
        Me.txtDescripcion = oRec!descripcion
        Me.chkVisible.value = IIf(oRec!visible, 1, 0)
        Me.chkDefecto.value = IIf(oRec!defecto, 1, 0)
        
    Else
        Me.txtCodigo = "999"
        Me.txtDescripcion = ""
        Me.chkVisible.value = 0
        Me.chkDefecto.value = 0
    End If
    
    
    Set oRec = Nothing
    Exit Sub
errores:
    If oCon.Errors.Count >= 1 Then
        MsgBox oCon.Errors(0).Description, vbCritical, "Nro de Error: " & oCon.Errors(0).Number
        oCon.Errors.Clear
    Else
        MsgBox Err.Number & "-" & Err.Description, vbCritical, "Error: cargarRegistroAModificar "
        Err.Clear
    End If

    Set oRec = Nothing
End Sub

