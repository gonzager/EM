VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBusquedaEmpresa 
   Caption         =   "Empresas"
   ClientHeight    =   4455
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   11550
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   435
      Left            =   7680
      TabIndex        =   9
      Top             =   3840
      Width           =   1515
   End
   Begin VB.CommandButton cmdSeleccionar 
      Caption         =   "S&eleccionar"
      Height          =   435
      Left            =   9600
      TabIndex        =   8
      Top             =   3840
      Width           =   1515
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
         Width           =   1455
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
Attribute VB_Name = "frmBusquedaEmpresa"
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

Private pl_identificador As Double


Public Sub inicializarFormulario(ByRef l_identificador As Double)
    inicializarGrilla
    Me.opTodos(0).value = 1
    Me.Show vbModal
    l_identificador = pl_identificador
End Sub


Private Sub cmdBuscar_Click()
    Dim ls_Sql As String
    Dim oRec As ADODB.Recordset
    On Error GoTo errores
    Set oRec = New ADODB.Recordset
    Dim tiene As Boolean
    
    ls_Sql = "SELECT IDENTIFICADOR, RAZONSOCIAL, ISNULL(DOMICILIO,'') DOMICILIO, " + _
             "CASE WHEN VENDEDOR=1 THEN 'SI' ELSE 'NO' END VENDEDOR, " + _
             "CASE WHEN COMPRADOR=1 THEN 'SI' ELSE 'NO' END COMPRADOR " + _
             "FROM EMPRESA WHERE RAZONSOCIAL LIKE '%" & Me.txtRazonSocial.Text & "%' AND ACTIVO = 1 "

    If Me.opTodos(1).value = True Then
        ls_Sql = ls_Sql & " AND VENDEDOR=1"
    ElseIf Me.opTodos(2).value = True Then
        ls_Sql = ls_Sql & " AND COMPRADOR=1"
    End If
    
             
    With oRec
        .ActiveConnection = oCon
        .CursorType = adOpenStatic
        .CursorLocation = adUseClient
        .Source = ls_Sql
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
            .Rows = .Rows + 1
        End With
        tiene = True
        oRec.MoveNext
    Loop
    
    If tiene Then
        msGrilla.RowSel = 1
        msGrilla.ColSel = COL_COMPRADOR
        msGrilla.SetFocus
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

Private Sub cmdSalir_Click()
    pl_identificador = 0
    Unload Me
End Sub

Private Sub cmdSeleccionar_Click()
    If msGrilla.Row <= 0 Then
        MsgBox "Debe seleccionar una fila para seleccionar.", vbCritical, "Error: Al seleccionar registro"
    ElseIf msGrilla.Row >= 1 And msGrilla.Row <= msGrilla.Rows - 2 Then
        pl_identificador = msGrilla.TextMatrix(msGrilla.Row, COL_ID)
        Unload Me
    End If
End Sub


Private Sub msGrilla_DblClick()
    If msGrilla.Row <= 0 Then
        MsgBox "Debe seleccionar una fila para seleccionar.", vbCritical, "Error: Al seleccionar registro"
    ElseIf msGrilla.Row >= 1 And msGrilla.Row <= msGrilla.Rows - 2 Then
        pl_identificador = msGrilla.TextMatrix(msGrilla.Row, COL_ID)
        Unload Me
    End If
End Sub

Private Sub msGrilla_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If msGrilla.Row <= 0 Then
            MsgBox "Debe seleccionar una fila para seleccionar.", vbCritical, "Error: Al seleccionar registro"
        ElseIf msGrilla.Row >= 1 And msGrilla.Row <= msGrilla.Rows - 2 Then
            pl_identificador = msGrilla.TextMatrix(msGrilla.Row, COL_ID)
            Unload Me
        End If
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
        .Cols = 6
        .ColWidth(0) = 300
        .ColWidth(COL_ID) = 1500
        .ColWidth(COL_RAZONSOCIAL) = 3550
        .ColWidth(COL_DOMICILIO) = 3400
        .ColWidth(COL_VENDEDOR) = 1120
        .ColWidth(COL_COMPRADOR) = 1150
     
        .TextMatrix(0, COL_ID) = "IDENTIFICACION"
        .TextMatrix(0, COL_RAZONSOCIAL) = "RAZON SOCIAL"
        .TextMatrix(0, COL_DOMICILIO) = "DOMICILIO"
        .TextMatrix(0, COL_VENDEDOR) = "VENDEDOR"
        .TextMatrix(0, COL_COMPRADOR) = "COMPRADOR"
    End With

End Sub
