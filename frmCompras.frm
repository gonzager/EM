VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCompras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compras"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   12495
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   12495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   435
      Left            =   8580
      TabIndex        =   22
      Top             =   7620
      Width           =   1695
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   435
      Left            =   10560
      TabIndex        =   23
      Top             =   7620
      Width           =   1755
   End
   Begin VB.Frame frmTotalGrales 
      Enabled         =   0   'False
      Height          =   1215
      Left            =   6600
      TabIndex        =   55
      Top             =   6300
      Width           =   5835
      Begin VB.TextBox txtOtras 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   120
         TabIndex        =   64
         Top             =   420
         Width           =   1365
      End
      Begin VB.TextBox txtTotalComprobante 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   4320
         TabIndex        =   59
         Top             =   780
         Width           =   1365
      End
      Begin VB.TextBox txtIIBB 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   4305
         TabIndex        =   58
         Top             =   420
         Width           =   1365
      End
      Begin VB.TextBox txtGcias 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2910
         TabIndex        =   57
         Top             =   420
         Width           =   1365
      End
      Begin VB.TextBox txtIVA 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1500
         TabIndex        =   56
         Top             =   420
         Width           =   1365
      End
      Begin VB.Label lblOtras 
         Alignment       =   2  'Center
         Caption         =   "Otras"
         Height          =   195
         Left            =   180
         TabIndex        =   65
         Top             =   180
         Width           =   1335
      End
      Begin VB.Label lblIVA 
         Alignment       =   2  'Center
         Caption         =   "IVA"
         Height          =   195
         Left            =   1560
         TabIndex        =   63
         Top             =   180
         Width           =   1275
      End
      Begin VB.Label lblGanancias 
         Alignment       =   2  'Center
         Caption         =   "Ganancias"
         Height          =   195
         Left            =   3000
         TabIndex        =   62
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label lblIIBB 
         Alignment       =   2  'Center
         Caption         =   "IIBB"
         Height          =   195
         Left            =   4320
         TabIndex        =   61
         Top             =   180
         Width           =   1335
      End
      Begin VB.Label lblTotalComprobate 
         Caption         =   "Total Comprobante"
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
         Height          =   255
         Left            =   2280
         TabIndex        =   60
         Top             =   840
         Width           =   1695
      End
   End
   Begin VB.Frame frmDetallePercepciones 
      Caption         =   "Percepciones"
      Height          =   1755
      Left            =   120
      TabIndex        =   54
      Top             =   6300
      Width           =   6435
      Begin MSFlexGridLib.MSFlexGrid MSPercepcion 
         Height          =   1335
         Left            =   120
         TabIndex        =   25
         Top             =   300
         Width           =   6250
         _ExtentX        =   11033
         _ExtentY        =   2355
         _Version        =   393216
         HighLight       =   2
         ScrollBars      =   2
         SelectionMode   =   1
      End
   End
   Begin VB.Frame frmPercepciones 
      Caption         =   "Carga de Percepciones"
      Height          =   795
      Left            =   60
      TabIndex        =   44
      Top             =   5460
      Width           =   12375
      Begin VB.CommandButton cmdPercepcion 
         Caption         =   "A&gregar Percepcion"
         Height          =   435
         Left            =   10440
         TabIndex        =   21
         Top             =   180
         Width           =   1695
      End
      Begin VB.ComboBox cmbTipoPercepcion 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   240
         Width           =   2535
      End
      Begin VB.ComboBox cmbJurisdiccion 
         Height          =   315
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   240
         Width           =   2535
      End
      Begin MSMask.MaskEdBox txtImportePercepcion 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0,00 ""€"""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   2
         EndProperty
         Height          =   315
         Left            =   9000
         TabIndex        =   20
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label lblImportePercepcion 
         Caption         =   "Importe Percepción:"
         Height          =   195
         Left            =   7500
         TabIndex        =   53
         Top             =   300
         Width           =   1515
      End
      Begin VB.Label lblTipoPercepcion 
         Caption         =   "Percepción:"
         Height          =   195
         Left            =   180
         TabIndex        =   52
         Top             =   300
         Width           =   975
      End
      Begin VB.Label lblJurisdicción 
         Caption         =   "Jurisdicción:"
         Height          =   195
         Left            =   3840
         TabIndex        =   51
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.Frame frmGrillaConceptos 
      Caption         =   "Conceptos"
      Height          =   2370
      Left            =   60
      TabIndex        =   43
      Top             =   3060
      Width           =   12375
      Begin VB.Frame frmTotales 
         Enabled         =   0   'False
         Height          =   555
         Left            =   120
         TabIndex        =   45
         Top             =   1740
         Width           =   12135
         Begin VB.TextBox txtNetoGravado 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   5235
            TabIndex        =   49
            Top             =   180
            Width           =   1610
         End
         Begin VB.TextBox txtTotalIva 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   6885
            TabIndex        =   48
            Top             =   180
            Width           =   1610
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   10185
            TabIndex        =   47
            Top             =   180
            Width           =   1610
         End
         Begin VB.TextBox txtExento 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   8520
            TabIndex        =   46
            Top             =   180
            Width           =   1610
         End
         Begin VB.Label lblTotales 
            Caption         =   "Totales:"
            Height          =   195
            Left            =   4320
            TabIndex        =   50
            Top             =   240
            Width           =   855
         End
      End
      Begin MSFlexGridLib.MSFlexGrid msGrilla 
         Height          =   1500
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   2646
         _Version        =   393216
         Cols            =   6
         HighLight       =   2
         ScrollBars      =   2
         SelectionMode   =   1
      End
   End
   Begin VB.Frame frmAgregarConceptos 
      Caption         =   "Carga de Conceptos"
      Height          =   810
      Left            =   60
      TabIndex        =   39
      Top             =   2220
      Width           =   12375
      Begin VB.ComboBox cmbConcepto 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   300
         Width           =   2535
      End
      Begin VB.ComboBox cmbAlicuota 
         Height          =   315
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   300
         Width           =   1815
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar Concepto"
         Height          =   435
         Left            =   10440
         TabIndex        =   17
         Top             =   240
         Width           =   1695
      End
      Begin MSMask.MaskEdBox txtImporte 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0,00 ""€"""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   2
         EndProperty
         Height          =   315
         Left            =   8880
         TabIndex        =   16
         Top             =   300
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label lblConcepto 
         Caption         =   "Concepto:"
         Height          =   195
         Left            =   180
         TabIndex        =   42
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lblAlicuota 
         Caption         =   "Alicuota I.V.A."
         Height          =   195
         Left            =   3960
         TabIndex        =   41
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label lblImporteConcepto 
         Caption         =   "Importe a Registrar:"
         Height          =   195
         Left            =   7440
         TabIndex        =   40
         Top             =   360
         Width           =   1395
      End
   End
   Begin VB.Frame frmCompras 
      Caption         =   "Compras - Datos del Comprobante"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   12375
      Begin VB.ComboBox cmbCodigoOperacion 
         Height          =   315
         Left            =   10020
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1680
         Width           =   2235
      End
      Begin MSMask.MaskEdBox txtPuntoVta 
         Height          =   315
         Left            =   5580
         TabIndex        =   11
         Top             =   1680
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cmbTipoComprobante 
         Height          =   315
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1680
         Width           =   2535
      End
      Begin VB.ComboBox cmbTipoDocVendedor 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1200
         Width           =   2235
      End
      Begin VB.TextBox txtRazonSocialVendedor 
         Height          =   315
         Left            =   7920
         TabIndex        =   9
         Top             =   1200
         Width           =   4335
      End
      Begin VB.ComboBox cmbMoneda 
         Height          =   315
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   720
         Width           =   2475
      End
      Begin VB.TextBox txtRazonSocial 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5940
         TabIndex        =   2
         Top             =   300
         Width           =   6315
      End
      Begin MSMask.MaskEdBox txtNroIdComprador 
         Height          =   315
         Left            =   2340
         TabIndex        =   1
         Top             =   300
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##-########-#"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker dtpFechaComprobante 
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   119668737
         CurrentDate     =   42201
      End
      Begin MSComCtl2.DTPicker dtpFechaImputacion 
         Height          =   315
         Left            =   4860
         TabIndex        =   4
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   119668737
         CurrentDate     =   42201
      End
      Begin MSMask.MaskEdBox txtTipoCambio 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0,00 ""€"""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   2
         EndProperty
         Height          =   315
         Left            =   11220
         TabIndex        =   6
         Top             =   720
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00000;($#,##0.00000)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtNroIdVendedor 
         Height          =   315
         Left            =   5580
         TabIndex        =   8
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtNroComprobante 
         Height          =   315
         Left            =   7920
         TabIndex        =   12
         Top             =   1680
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "0000000#"
         PromptChar      =   "_"
      End
      Begin VB.Label lblCodigoOperacionVentas 
         Caption         =   "Cod. Oper:"
         Height          =   255
         Left            =   9180
         TabIndex        =   38
         Top             =   1740
         Width           =   855
      End
      Begin VB.Label lblNumeroComprobanteDesde 
         Caption         =   "N. de Comprobante:"
         Height          =   255
         Left            =   6480
         TabIndex        =   37
         Top             =   1740
         Width           =   1455
      End
      Begin VB.Label lblPuntoVta 
         Caption         =   "Punto de Venta:"
         Height          =   255
         Left            =   4200
         TabIndex        =   36
         Top             =   1740
         Width           =   1335
      End
      Begin VB.Label lblTipoComprobante 
         Caption         =   "Tipo Comprobate:"
         Height          =   255
         Left            =   180
         TabIndex        =   35
         Top             =   1740
         Width           =   1575
      End
      Begin VB.Label lblDocumentoVendedor 
         Caption         =   "Tipo Doc. Comprador:"
         Height          =   255
         Left            =   180
         TabIndex        =   34
         Top             =   1260
         Width           =   1635
      End
      Begin VB.Label lblNroIdentificadorVendedor 
         Caption         =   "Nro. Id. Vendedor:"
         Height          =   255
         Left            =   4200
         TabIndex        =   33
         Top             =   1260
         Width           =   1515
      End
      Begin VB.Label lblRazonSocialVendedor 
         Caption         =   "Vendedor:"
         Height          =   195
         Left            =   7140
         TabIndex        =   32
         Top             =   1260
         Width           =   795
      End
      Begin VB.Label lblMoneda 
         Caption         =   "Moneda:"
         Height          =   315
         Left            =   6540
         TabIndex        =   31
         Top             =   780
         Width           =   735
      End
      Begin VB.Label lblTipoCambio 
         Caption         =   "Tipo de Cambio:"
         Height          =   255
         Left            =   9960
         TabIndex        =   30
         Top             =   780
         Width           =   1155
      End
      Begin VB.Label lblFechaImputacion 
         Caption         =   "Fecha Imputación:"
         Height          =   255
         Left            =   3420
         TabIndex        =   29
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label lblFechaComprobante 
         Caption         =   "Fecha Comprobante:"
         Height          =   255
         Left            =   180
         TabIndex        =   28
         Top             =   780
         Width           =   1515
      End
      Begin VB.Label lblIdentificadorComprador 
         Caption         =   " Nro. Identificador Comprador:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   2115
      End
      Begin VB.Label lblRazonSocial 
         Caption         =   "Razón Social Comprador:"
         Height          =   195
         Left            =   3960
         TabIndex        =   26
         Top             =   360
         Width           =   1875
      End
   End
   Begin VB.Menu mnu_edicion 
      Caption         =   "&Edición"
      Begin VB.Menu mnu_edicion_Conceptos 
         Caption         =   "&Conceptos"
         Begin VB.Menu mnu_edicion_borrarConcepto 
            Caption         =   "Borrar Concepto"
         End
      End
      Begin VB.Menu mnu_edicion_Percepciones 
         Caption         =   "&Percepciones"
         Begin VB.Menu mnu_edicion_borrarPercepcion 
            Caption         =   "Borrar Percepción"
         End
      End
   End
End
Attribute VB_Name = "frmCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const COL_CONCEPTO = 1
Private Const COL_ALICUOTA = 2
Private Const COL_NETO_GRAVADO = 3
Private Const COL_IVA = 4
Private Const COL_EXENTO = 5
Private Const COL_TOTAL = 6
Private Const COL_CODIGO_CONCEPTO = 7
Private Const COL_CODIGO_ALICUOTA = 8

Private Const COL_TIPO_PERCEPCION = 1
Private Const COL_JURISDICCION = 2
Private Const COL_PERCEPCION = 3
Private Const COL_ID_TIPO_PERCEPCION = 4
Private Const COL_ID_JURISDICCION = 5



Public operacion As String
Public idCompra As String

 
Private Sub cmbMoneda_Click()
On Error Resume Next
    Dim strtmp() As String
    strtmp = Utils.separarCodigoDescripcion(Me.cmbMoneda)
    If strtmp(0) = tipoMonedaPesos Then
        Me.txtTipoCambio.text = 1
        Me.txtTipoCambio.Enabled = False
    Else
        'Me.txtTipoCambio.Text = 1
        Me.txtTipoCambio.Enabled = True
        Me.txtTipoCambio.SetFocus
    End If
End Sub

Private Sub cmbTipoComprobante_LostFocus()
Dim tipoCalculo As String
Dim ltTipoComprobante As tTipoComprobante
Dim lstpm() As String
Dim ltAlicuota As tAlicuota

Dim alicuotaCero As String
Dim stmp() As String
    stmp = Utils.separarCodigoDescripcion(Me.cmbTipoComprobante)
    ltTipoComprobante.codigo = stmp(0)
    ltTipoComprobante.descripcion = stmp(1)
    tipoCalculo = Utils.recuperarTipoCalculo(ltTipoComprobante)
    alicuotaCero = alicuotaExenta
    If tipoCalculo = tipoCalculoExento Then
        Dim i As Integer
        For i = 0 To Me.cmbAlicuota.ListCount - 1
            Me.cmbAlicuota.ListIndex = i
            lstpm = Utils.separarCodigoDescripcion(Me.cmbAlicuota)
            ltAlicuota.alicuta = lstpm(0)
            If alicuotaCero = ltAlicuota.alicuta Then
                Me.cmbAlicuota.ListIndex = i
                Me.cmbAlicuota.Enabled = False
                Exit For
            End If
        Next

        Me.cmbAlicuota.Enabled = False
    Else
        Dim sql_alicuotas As String
        Me.cmbAlicuota.Enabled = True
        cmbAlicuota.Clear
        sql_alicuotas = "SELECT ALICUOTA, ISNULL(DESCRIPCION,'') AS DESCRIPCION, ISNULL(DEFECTO,0) AS DEFECTO FROM ALICUOTA WHERE ISNULL(VISIBLE,0) = 1"
        cargarCombo Me.cmbAlicuota, sql_alicuotas
        
        
    End If
End Sub

Private Sub cmbTipoDocVendedor_Click()
Dim strtmp() As String
Dim ltDocumento As tDocumento
Dim tmp As String
    strtmp = Utils.separarCodigoDescripcion(Me.cmbTipoDocVendedor)
    ltDocumento.codigo = strtmp(0)
    ltDocumento.descripcion = strtmp(1)
    If ltDocumento.codigo = CODIGOCUIT Or ltDocumento.codigo = CODIGOCUIL Then
        Me.txtNroIdVendedor.Mask = "##-########-#"
    Else
        tmp = Me.txtNroIdVendedor.ClipText
        Me.txtNroIdVendedor.Mask = ""
        Me.txtNroIdVendedor.text = tmp
    End If
End Sub



Private Sub cmbTipoPercepcion_Click()
Dim strtmp() As String
     strtmp = Utils.separarCodigoDescripcion(Me.cmbTipoPercepcion)
     If strtmp(0) <> CODIGOIIBB Then
        Me.cmbJurisdiccion.ListIndex = -1
        Me.cmbJurisdiccion.Enabled = False
     Else
         Me.cmbJurisdiccion.ListIndex = 0
         Me.cmbJurisdiccion.Enabled = True
     End If
End Sub

Private Sub cmdAgregar_Click()

Dim ltCabeceraCompra As tCabecera_Compra

    ltCabeceraCompra = cargarCabeceraCompra
    If cargarCabeceraCompra.existe Then
        If validarDatos Then
            If agregarConcepto Then
                msGrilla.RowSel = msGrilla.Rows - 1
                msGrilla.Rows = msGrilla.Rows + 1
                calcularTotales
                calcularTotalComprobante
                Me.txtImporte.text = ""
                Me.cmbConcepto.SetFocus
                frmCompras.Enabled = False
            End If
        Else
            Me.txtImporte.text = 0
            Me.txtImporte.SetFocus
        End If
    Else
        MsgBox "Datos de la Cabecera Incompletos", vbCritical, "Error: Datos de la Cabecera"
    End If
End Sub

Private Sub cmdGrabar_Click()
    Dim ltCabeceraCompra As tCabecera_Compra
    Dim compraId As Double
    Dim i As Integer
    Dim b As Boolean
    Dim graboDetalle As Boolean
    Dim graboPercepcion As Boolean

    ltCabeceraCompra = cargarCabeceraCompra
    
    If ltCabeceraCompra.existe Then
        If msGrilla.Rows > 2 Then
            If ExisteComprobanteDeCompra(ltCabeceraCompra) = 0 Or Me.idCompra <> "0" Then
        
        
                oCon.BeginTrans
                b = True
                
                If Me.idCompra <> "0" Then
                    b = False
                    If borrarCabeceraDetalleCompra(Me.idCompra) > 0 Then
                        b = True
                    End If
                End If
        
                compraId = insertarCabeceraCompra(ltCabeceraCompra)
                
                If compraId > 0 And b Then
                    
                    For i = 1 To Me.msGrilla.Rows - 2
                        Dim ltDetalleCompra As tDetalle_compra
                        ltDetalleCompra = cargarDetalleCompra(compraId, i)
                        graboDetalle = False
                        If ltDetalleCompra.existe Then
                            If insertarDetalleCompra(ltDetalleCompra) Then
                            graboDetalle = True
                            End If
                        Else
                            Exit For
                        End If
                    Next
                    
                    graboPercepcion = True
                    
                    For i = 1 To Me.MSPercepcion.Rows - 2
                        Dim ltPercepcionCompra As tPercepcion_compra
                        ltPercepcionCompra = cargarPercepcionCompra(compraId, i)
                        graboPercepcion = False
                        If ltPercepcionCompra.existe Then
                            If insertarPercepcionCompra(ltPercepcionCompra) Then
                                graboPercepcion = True
                            Else
                                graboPercepcion = False
                            End If
                        Else
                            Exit For
                        End If
                    Next
                    
                    
                    If graboDetalle = True And graboPercepcion = True Then
                        oCon.CommitTrans
                        MsgBox "La registración se ha realizado en forma correcta.", vbInformation, "Registracion OK"
                        If Me.idCompra <> 0 Then
                            Unload Me
                        Else
                            limpiarFormulario
                            Me.txtNroIdComprador.SetFocus
                        End If
                    Else
                        oCon.RollbackTrans
                    End If
                Else
                    oCon.RollbackTrans
                End If
            Else
                MsgBox "El Comprobante que esta cargando ya se encuentra registrado." + vbCrLf + _
                "Verifique con la siguiente información: " + vbCrLf + _
                "Vendedor: " & ltCabeceraCompra.vendedorId & "-" & ltCabeceraCompra.razonSocialVendedor & vbCrLf & _
                "Tipo de Comprobante: " & ltCabeceraCompra.tipoComprobabte.codigo & "-" & ltCabeceraCompra.tipoComprobabte.descripcion & vbCrLf & _
                "Comprobante Nro: " & ltCabeceraCompra.PuntoDeVenta & "-" & ltCabeceraCompra.nroComprobante _
                , vbCritical, "Error: Comprobante de Compra Duplicado"
            End If
        Else
            MsgBox "Debe ingresar al menos un concepto para poder grabar.", vbCritical, "Error: Detalle de Conceptos"
        End If
    Else
        MsgBox "Datos de la Cabecera Incompletos", vbCritical, "Error: Datos de la Cabecera"
    End If
    
End Sub

Private Sub cmdPercepcion_Click()
Dim ltCabeceraCompra As tCabecera_Compra

    ltCabeceraCompra = cargarCabeceraCompra
    If ltCabeceraCompra.existe Then
        If validarDatosPercepciones Then
            If agregarPercepcion Then
                MSPercepcion.RowSel = MSPercepcion.Rows - 1
                MSPercepcion.Rows = MSPercepcion.Rows + 1
                calcularTotalesPercepciones
                calcularTotalComprobante
                Me.txtImportePercepcion.text = ""
                Me.cmbTipoPercepcion.SetFocus
                frmCompras.Enabled = False
            End If
        Else
            Me.txtImportePercepcion.text = 0
            Me.txtImportePercepcion.SetFocus
        End If
    Else
        MsgBox "Datos de la Cabecera Incompletos", vbCritical, "Error: Datos de la Cabecera"
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub







Private Sub dtpFechaComprobante_Validate(Cancel As Boolean)
    If Me.dtpFechaComprobante > Me.dtpFechaImputacion Then
        MsgBox "La Fecha de Comprobante no puede ser mayor a la fecha de Imputación", vbCritical, "Error: Fecha de Comprobante Invalida"
        Cancel = True
    End If
End Sub



Private Sub dtpFechaImputacion_Validate(Cancel As Boolean)
    If Me.dtpFechaImputacion < Me.dtpFechaComprobante Then
        MsgBox "La Fecha de Imputación no puede ser menor a la fecha de Comprobante", vbCritical, "Error: Fecha de Imputación Invalida"
        Cancel = True
    End If
End Sub

Private Sub Form_Load()
    If operacion = "A" Then
        cargarFormularioAlta
    ElseIf operacion = "M" Then
        cargarFormularioModificacion
    End If
End Sub

Private Sub mnu_edicion_borrarConcepto_Click()
    If msGrilla.Row <= 0 Then
        MsgBox "Debe seleccionar una fila para borrar.", vbCritical, "Error: Al Borrar el Concepto"
    ElseIf msGrilla.Row >= 1 And msGrilla.Row <= msGrilla.Rows - 2 Then
        msGrilla.RemoveItem (msGrilla.Row)
        calcularTotales
        calcularTotalComprobante
        If msGrilla.Rows = 2 Then
            Me.frmCompras.Enabled = True
        End If
    End If
End Sub

Private Sub mnu_edicion_borrarPercepcion_Click()
    If MSPercepcion.Row <= 0 Then
        MsgBox "Debe seleccionar una fila para borrar.", vbCritical, "Error: Al Borrar el Percepción"
    ElseIf MSPercepcion.Row >= 1 And MSPercepcion.Row <= MSPercepcion.Rows - 2 Then
        MSPercepcion.RemoveItem (MSPercepcion.Row)
        calcularTotalesPercepciones
        calcularTotalComprobante
    End If
End Sub

Private Sub msGrilla_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu mnu_edicion_Conceptos
    End If
End Sub

Private Sub MSPercepcion_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu mnu_edicion_Percepciones
    End If
End Sub

Private Sub txtImporte_GotFocus()
    With txtImporte
        .SelStart = 0
        .SelLength = Len(.text)
    End With
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
        cmdAgregar.SetFocus
    End If

   If Not IsNumeric(Chr(KeyAscii)) Then
       If KeyAscii <> vbKeyBack And Chr(KeyAscii) <> separadorDecimal Then
          KeyAscii = 0
        Else
            If InStr(1, txtTipoCambio.text, separadorDecimal) > 0 And KeyAscii <> vbKeyBack And Chr(KeyAscii) <> separadorDecimal Then
                KeyAscii = 0
            End If
       End If
    End If
End Sub

Private Sub txtImportePercepcion_GotFocus()
    With txtImportePercepcion
        .SelStart = 0
        .SelLength = Len(.text)
    End With
End Sub

Private Sub txtImportePercepcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Me.cmdPercepcion.SetFocus
    End If

    If Not IsNumeric(Chr(KeyAscii)) Then
        If KeyAscii <> vbKeyBack And Chr(KeyAscii) <> separadorDecimal Then
           KeyAscii = 0
         Else
             If InStr(1, txtTipoCambio.text, separadorDecimal) > 0 And KeyAscii <> vbKeyBack And Chr(KeyAscii) <> separadorDecimal Then
                 KeyAscii = 0
             End If
        End If
     End If
End Sub

Private Sub txtImportePercepcion_LostFocus()
    With txtImportePercepcion
        .SelStart = 0
        .SelLength = Len(.text)
    End With
End Sub

Private Sub txtNroComprobante_GotFocus()
    With txtNroComprobante
        .SelStart = 0
        .SelLength = Len(.text)
    End With
End Sub

Private Sub txtNroComprobante_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Sendkeys "{tab}"
    End If
   
   If Not IsNumeric(Chr(KeyAscii)) Then
       If KeyAscii <> vbKeyBack Then
          KeyAscii = 0
       End If
    End If
End Sub

Private Sub txtNroComprobante_LostFocus()
    Me.txtNroComprobante.text = Format(Me.txtNroComprobante.ClipText, "0000000#")
End Sub

Sub txtNroIdComprador_GotFocus()
    With txtNroIdComprador
        .SelStart = 0
        .SelLength = Len(.text)
    End With
    
    txtRazonSocial.text = ""
    
End Sub

Private Sub txtNroIdComprador_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
            Dim ltEmpresa As tEmpresa
            ltEmpresa.existe = False
            If txtNroIdComprador.ClipText <> "" Then
                ltEmpresa = recuperarEmpresaPorIdentificador(txtNroIdComprador.ClipText)
            End If
            If Not ltEmpresa.existe Then
                Dim indentificador As Double
                frmBusquedaEmpresa.inicializarFormulario indentificador
                txtNroIdComprador.Mask = ""
                txtNroIdComprador.text = IIf(indentificador <> 0, indentificador, "")
                txtNroIdComprador.PromptInclude = False
                txtNroIdComprador.Mask = "##-########-#"
                txtNroIdComprador.PromptInclude = True
                If indentificador > 0 Then Sendkeys "{tab}"
            Else
                Sendkeys "{tab}"
            End If
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
                txtRazonSocial.text = ltEmpresa.razonSocial
            Else
                With txtNroIdComprador
                    .SelStart = 0
                    .SelLength = Len(.text)
                    .SetFocus
                End With
            End If
        Else
            With txtNroIdComprador
                .SelStart = 0
                .SelLength = Len(.text)
                .SetFocus
            End With

        End If
    End If
End Sub


Private Sub cargarFormularioAlta()
    Dim sql_moneda As String
    Dim sql_tipoDocVendedor As String
    Dim sql_tipoComprobante As String
    Dim sql_codigoOperacionVentas As String
    Dim sql_conceptos As String
    Dim sql_alicuotas As String
    Dim sql_jurisdicciones As String
    Dim sql_percecpiones As String
    Dim strtmp() As String
    
    dtpFechaImputacion.value = Format(fechaActualServer, "Short date")
    dtpFechaComprobante.value = Format(fechaActualServer, "Short date")
   
    
    sql_moneda = "SELECT CODIGO, ISNULL(DESCRIPCION,'') AS DESCRIPCION, ISNULL(DEFECTO,0) AS DEFECTO FROM MONEDA WHERE ISNULL(VISIBLE,0) = 1 order by CODIGO "
    cargarCombo Me.cmbMoneda, sql_moneda
    strtmp = Utils.separarCodigoDescripcion(Me.cmbMoneda)
    If strtmp(0) = tipoMonedaPesos Then
        Me.txtTipoCambio.text = 1
        Me.txtTipoCambio.Enabled = False
    End If
    
    sql_tipoDocVendedor = "SELECT CODIGO, ISNULL(DESCRIPCION,'') AS DESCRIPCION, ISNULL(DEFECTO,0) AS DEFECTO FROM DOCUMENTO WHERE ISNULL(VISIBLE,0) = 1 order by CODIGO "
    cargarCombo Me.cmbTipoDocVendedor, sql_tipoDocVendedor

    sql_tipoComprobante = "SELECT CODIGO, ISNULL(DESCRIPCION,'') AS DESCRIPCION, ISNULL(DEFECTO,0) AS DEFECTO FROM TIPO_COMPROBANTE WHERE ISNULL(VISIBLE,0) = 1 order by CODIGO "
    cargarCombo Me.cmbTipoComprobante, sql_tipoComprobante

    sql_codigoOperacionVentas = "SELECT CODIGO, ISNULL(DESCRIPCION,'') AS DESCRIPCION, ISNULL(DEFECTO,0) AS DEFECTO FROM CODIGO_OPERACION WHERE ISNULL(VISIBLE,0) = 1 order by CODIGO "
    cargarCombo Me.cmbCodigoOperacion, sql_codigoOperacionVentas

    sql_conceptos = "SELECT CODIGO, ISNULL(DESCRIPCION,'') AS DESCRIPCION, ISNULL(DEFECTO,0) AS DEFECTO FROM dbo.CONCEPTOCPA WHERE ISNULL(VISIBLE,0) = 1 order by CODIGO "
    cargarCombo Me.cmbConcepto, sql_conceptos

    sql_alicuotas = "SELECT ALICUOTA, ISNULL(DESCRIPCION,'') AS DESCRIPCION, ISNULL(DEFECTO,0) AS DEFECTO FROM ALICUOTA WHERE ISNULL(VISIBLE,0) = 1"
    cargarCombo Me.cmbAlicuota, sql_alicuotas

    sql_jurisdicciones = "SELECT CODIGO, ISNULL(DESCRIPCION,'') AS DESCRIPCION, ISNULL(DEFECTO,0) AS DEFECTO FROM dbo.JURISDICCION WHERE ISNULL(VISIBLE,0) = 1 order by CODIGO "
    cargarCombo Me.cmbJurisdiccion, sql_jurisdicciones
    
    sql_percecpiones = "SELECT CODIGO, ISNULL(DESCRIPCION,'') AS DESCRIPCION, ISNULL(DEFECTO,0) AS DEFECTO FROM dbo.TIPO_PERCEPCION WHERE ISNULL(VISIBLE,0) = 1 order by CODIGO "
    cargarCombo Me.cmbTipoPercepcion, sql_percecpiones
    
    Me.txtNroIdComprador.Mask = "##-########-#"
    limpiarFormulario
End Sub

Private Sub txtNroIdVendedor_GotFocus()
    With txtNroIdVendedor
        .SelStart = 0
        .SelLength = Len(.text)
    End With
End Sub

Private Sub txtNroIdVendedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Sendkeys "{tab}"
    End If
   
   If Not IsNumeric(Chr(KeyAscii)) Then
       If KeyAscii <> vbKeyBack Then
          KeyAscii = 0
       End If
    End If
End Sub

Private Sub txtNroIdVendedor_LostFocus()
Dim ltDocumento As tDocumento
Dim strtmp() As String


    If Len(txtNroIdVendedor.ClipText) > 0 Then
        strtmp = Utils.separarCodigoDescripcion(Me.cmbTipoDocVendedor)
        ltDocumento.codigo = strtmp(0)
        ltDocumento.descripcion = strtmp(1)
        If ltDocumento.codigo = CODIGOCUIT Or ltDocumento.codigo = CODIGOCUIL Then
            If Len(Me.txtNroIdVendedor.ClipText) = 11 Then
                If ValidarCuit(Me.txtNroIdVendedor.ClipText) Then
                    Me.txtRazonSocialVendedor.text = "INGRESE LA RAZON SOCIAL DEL VENDEDOR"
                    Me.txtRazonSocialVendedor.SetFocus
                Else
                    MsgBox "El CUIT/CUIL ingresado no es válido", vbCritical, "Error: Código de Verificación"
                    Me.txtNroIdVendedor.SetFocus
                    Me.txtRazonSocialVendedor.text = ""
                End If
            Else
                MsgBox "La longitud del campo para el numero de CUIT/CUIL no es valida", vbCritical, "Error: Longitud CUIT/CUIL"
                Me.txtNroIdVendedor.SetFocus
                Me.txtRazonSocialVendedor.text = ""
            End If
        Else
            Me.txtRazonSocialVendedor.text = "INGRESE LA RAZON SOCIAL DEL VENDEDOR"
            Me.txtRazonSocialVendedor.SetFocus
        End If
    End If

End Sub



Private Sub txtPorcentual_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
        Me.cmdPercepcion.SetFocus
    End If

   If Not IsNumeric(Chr(KeyAscii)) Then
       If KeyAscii <> vbKeyBack And Chr(KeyAscii) <> separadorDecimal Then
          KeyAscii = 0
        Else
            If InStr(1, txtTipoCambio.text, separadorDecimal) > 0 And KeyAscii <> vbKeyBack Then
                KeyAscii = 0
            End If
       End If
    End If
End Sub


Private Sub txtPuntoVta_GotFocus()
    With txtPuntoVta
        .SelStart = 0
        .SelLength = Len(.text)
    End With
End Sub

Private Sub txtPuntoVta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Sendkeys "{tab}"
    End If
   
   If Not IsNumeric(Chr(KeyAscii)) Then
       If KeyAscii <> vbKeyBack Then
          KeyAscii = 0
       End If
    End If
End Sub

Private Sub txtPuntoVta_LostFocus()
    Me.txtPuntoVta.text = Format(Me.txtPuntoVta.ClipText, "0000#")
End Sub

Private Sub txtRazonSocialVendedor_GotFocus()
    With txtRazonSocialVendedor
        .SelStart = 0
        .SelLength = Len(.text)
    End With
End Sub

Private Sub txtRazonSocialVendedor_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtTipoCambio_GotFocus()
    With txtTipoCambio
        .SelStart = 0
        .SelLength = Len(.text)
    End With
End Sub

Private Sub txtTipoCambio_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
        Sendkeys "{tab}"
    End If

   If Not IsNumeric(Chr(KeyAscii)) Then
       If KeyAscii <> vbKeyBack And Chr(KeyAscii) <> separadorDecimal Then
          KeyAscii = 0
       Else
            If InStr(1, txtTipoCambio.text, separadorDecimal) > 0 And KeyAscii <> vbKeyBack Then
                KeyAscii = 0
            End If
       End If

    End If
End Sub

Private Sub limpiarFormulario()
   
    inicializarGrilla
    inicializarGrillaPercepcion
    
    txtNroIdVendedor.Mask = ""
    txtNroIdVendedor.text = ""
    txtNroIdVendedor.Mask = "##-########-#"
    txtRazonSocialVendedor.text = ""
    txtNroComprobante.text = ""
    Me.txtTotalComprobante.text = Format(0, formatoMoneda)
    Me.frmCompras.Enabled = True
End Sub


Private Sub inicializarGrilla()
    msGrilla.Clear
    msGrilla.Rows = 2
    With msGrilla
        .Cols = 9
        .ColWidth(0) = 300
        .ColWidth(COL_CONCEPTO) = 3500
        
        .ColWidth(COL_ALICUOTA) = 1370
        .ColAlignment(COL_ALICUOTA) = flexAlignRightCenter
        
        .ColWidth(COL_NETO_GRAVADO) = 1650
        .ColAlignment(COL_NETO_GRAVADO) = flexAlignRightCenter

        .ColWidth(COL_IVA) = 1650
        .ColAlignment(COL_IVA) = flexAlignRightCenter
        
        .ColWidth(COL_EXENTO) = 1650
        .ColAlignment(COL_EXENTO) = flexAlignRightCenter
        
        .ColWidth(COL_TOTAL) = 1650
        .ColAlignment(COL_TOTAL) = flexAlignRightCenter
        
        .ColWidth(COL_CODIGO_CONCEPTO) = 0
        .ColWidth(COL_CODIGO_ALICUOTA) = 0
        
        .TextMatrix(0, COL_CONCEPTO) = "CONCEPTO"
        .TextMatrix(0, COL_ALICUOTA) = "ALICUOTA"
        .TextMatrix(0, COL_NETO_GRAVADO) = "NETO GRAVADO"
        .TextMatrix(0, COL_EXENTO) = "EXENTO"
        .TextMatrix(0, COL_IVA) = "I.V.A"
        .TextMatrix(0, COL_TOTAL) = "TOTAL"
        .TextMatrix(0, COL_CODIGO_CONCEPTO) = "CODIGO CONCEPTO"
        .TextMatrix(0, COL_CODIGO_ALICUOTA) = "CODIGO ALICUOTA"
    End With
    
    Me.txtNetoGravado.text = Format(0, formatoMoneda)
    Me.txtTotalIva.text = Format(0, formatoMoneda)
    Me.txtTotal.text = Format(0, formatoMoneda)
    Me.txtExento.text = Format(0, formatoMoneda)
    
End Sub


Private Sub inicializarGrillaPercepcion()
    Me.MSPercepcion.Clear
    Me.MSPercepcion.Rows = 2
    With MSPercepcion
        .Cols = 6
        .ColWidth(0) = 300
        
        .ColWidth(COL_TIPO_PERCEPCION) = 1700
        .ColWidth(COL_JURISDICCION) = 2200
        .ColAlignment(COL_PERCEPCION) = flexAlignRightCenter
        .ColWidth(COL_PERCEPCION) = 1700
        .ColWidth(COL_ID_TIPO_PERCEPCION) = 0
        .ColWidth(COL_ID_JURISDICCION) = 0
        
        .TextMatrix(0, COL_TIPO_PERCEPCION) = "TIPO"
        .TextMatrix(0, COL_JURISDICCION) = "JURISDICCION"
        .TextMatrix(0, COL_PERCEPCION) = "PERCEPCION"
        .TextMatrix(0, COL_ID_TIPO_PERCEPCION) = "CODIGO TIPO PERCEPCION"
        .TextMatrix(0, COL_ID_JURISDICCION) = "CODIGO JURISDICCION"
        
    End With
    
    Me.txtOtras.text = Format(0, formatoMoneda)
    Me.txtIVA.text = Format(0, formatoMoneda)
    Me.txtGcias.text = Format(0, formatoMoneda)
    Me.txtIIBB.text = Format(0, formatoMoneda)
    
End Sub



Private Function agregarConcepto() As Boolean
Dim datosConcepto() As String
Dim datosAlicuota() As String
Dim importeIVA() As Currency
Dim ltTipoComprobante As tTipoComprobante
Dim stmp() As String
Dim tipoCalculo As String

agregarConcepto = False
On Error GoTo errores
    
    datosConcepto = separarCodigoDescripcion(Me.cmbConcepto.text)
    datosAlicuota = separarCodigoDescripcion(Me.cmbAlicuota.text)
        
    msGrilla.TextMatrix(msGrilla.Rows - 1, COL_CONCEPTO) = datosConcepto(1)
    msGrilla.TextMatrix(msGrilla.Rows - 1, COL_CODIGO_CONCEPTO) = datosConcepto(0)
    
    msGrilla.TextMatrix(msGrilla.Rows - 1, COL_ALICUOTA) = datosAlicuota(1) & " %"
    msGrilla.TextMatrix(msGrilla.Rows - 1, COL_CODIGO_ALICUOTA) = datosAlicuota(0)
    
    stmp = Utils.separarCodigoDescripcion(Me.cmbTipoComprobante)
    ltTipoComprobante.codigo = stmp(0)
    ltTipoComprobante.descripcion = stmp(1)
    
    tipoCalculo = Utils.recuperarTipoCalculo(ltTipoComprobante)
    If tipoCalculo = tipoCalculoDirecto Then


        importeIVA = calculoIVATipoDirecto(CCur(Replace(Me.txtImporte.text, "$", "")), CCur(datosAlicuota(1)))

        If datosAlicuota(0) <> alicuotaExenta Then
            msGrilla.TextMatrix(msGrilla.Rows - 1, COL_NETO_GRAVADO) = Format(Me.txtImporte.text, formatoMoneda)
            msGrilla.TextMatrix(msGrilla.Rows - 1, COL_IVA) = Format((importeIVA(0)), formatoMoneda)
            msGrilla.TextMatrix(msGrilla.Rows - 1, COL_TOTAL) = Format((importeIVA(1)), formatoMoneda)
            msGrilla.TextMatrix(msGrilla.Rows - 1, COL_EXENTO) = Format(CCur(0), formatoMoneda)
        Else
            msGrilla.TextMatrix(msGrilla.Rows - 1, COL_NETO_GRAVADO) = Format(CCur(0), formatoMoneda)
            msGrilla.TextMatrix(msGrilla.Rows - 1, COL_IVA) = Format(CCur(importeIVA(0)), formatoMoneda)
            msGrilla.TextMatrix(msGrilla.Rows - 1, COL_TOTAL) = Format(CCur(importeIVA(1)), formatoMoneda)
            msGrilla.TextMatrix(msGrilla.Rows - 1, COL_EXENTO) = Format(CCur(importeIVA(1)), formatoMoneda)
        End If
        
    ElseIf tipoCalculo = tipoCalculoInverso Then
        
        importeIVA = calculoIVATipoInverso(CCur(Me.txtImporte.text), CCur(datosAlicuota(1)))
        
        If datosAlicuota(0) <> alicuotaExenta Then
            msGrilla.TextMatrix(msGrilla.Rows - 1, COL_TOTAL) = Format(Me.txtImporte.text, formatoMoneda)
            msGrilla.TextMatrix(msGrilla.Rows - 1, COL_IVA) = Format((importeIVA(0)), formatoMoneda)
            msGrilla.TextMatrix(msGrilla.Rows - 1, COL_NETO_GRAVADO) = Format((importeIVA(1)), formatoMoneda)
            msGrilla.TextMatrix(msGrilla.Rows - 1, COL_EXENTO) = Format(CCur(0), formatoMoneda)
        Else
            msGrilla.TextMatrix(msGrilla.Rows - 1, COL_NETO_GRAVADO) = Format(CCur(0), formatoMoneda)
            msGrilla.TextMatrix(msGrilla.Rows - 1, COL_IVA) = Format(CCur(importeIVA(0)), formatoMoneda)
            msGrilla.TextMatrix(msGrilla.Rows - 1, COL_TOTAL) = Format(CCur(importeIVA(1)), formatoMoneda)
            msGrilla.TextMatrix(msGrilla.Rows - 1, COL_EXENTO) = Format(CCur(importeIVA(1)), formatoMoneda)
        End If
    
    ElseIf tipoCalculo = tipoCalculoExento Then
                
'        importeIVA = calculoIVATipoDirecto(CCur(Me.txtImporte.Text), CCur(0))
'
'        If datosAlicuota(0) <> alicuotaExenta Then
'            msGrilla.TextMatrix(msGrilla.Rows - 1, COL_NETO_GRAVADO) = Format(0, formatoMoneda)
'            msGrilla.TextMatrix(msGrilla.Rows - 1, COL_IVA) = Format((importeIVA(0)), formatoMoneda)
'            msGrilla.TextMatrix(msGrilla.Rows - 1, COL_TOTAL) = Format((importeIVA(1)), formatoMoneda)
'            msGrilla.TextMatrix(msGrilla.Rows - 1, COL_EXENTO) = Format(CCur(Me.txtImporte.Text), formatoMoneda)
'        Else
'            msGrilla.TextMatrix(msGrilla.Rows - 1, COL_NETO_GRAVADO) = Format(CCur(0), formatoMoneda)
'            msGrilla.TextMatrix(msGrilla.Rows - 1, COL_IVA) = Format(CCur(importeIVA(0)), formatoMoneda)
'            msGrilla.TextMatrix(msGrilla.Rows - 1, COL_TOTAL) = Format(CCur(importeIVA(1)), formatoMoneda)
'            msGrilla.TextMatrix(msGrilla.Rows - 1, COL_EXENTO) = Format(CCur(importeIVA(1)), formatoMoneda)
'        End If
'


        'CAMBIO ESTE PEDAZO DE CODIGO POR EL QUE ESTA COMENTADO ARRIBA
        msGrilla.TextMatrix(msGrilla.Rows - 1, COL_NETO_GRAVADO) = Format(0, formatoMoneda)
        msGrilla.TextMatrix(msGrilla.Rows - 1, COL_IVA) = Format(CCur(0), formatoMoneda)
        msGrilla.TextMatrix(msGrilla.Rows - 1, COL_TOTAL) = Format(CCur(Me.txtImporte.text), formatoMoneda)
        msGrilla.TextMatrix(msGrilla.Rows - 1, COL_EXENTO) = Format(CCur(0), formatoMoneda)
    End If
        
    agregarConcepto = True
    Exit Function
errores:
    
        MsgBox Err.Description, Err.Number
        Err.Clear
    
    agregarConcepto = False
End Function

Private Sub calcularTotales()
On Error Resume Next
Dim i As Integer
Dim totalGravado As Currency
Dim totalExento As Currency
Dim totalIVA As Currency
Dim total As Currency

    Me.txtNetoGravado.text = ""
    Me.txtTotalIva.text = ""
    Me.txtTotal.text = ""
    Me.txtExento.text = ""

    totalGravado = 0
    totalIVA = 0
    total = 0
    totalExento = 0
    For i = 1 To Me.msGrilla.Rows - 2
        totalGravado = totalGravado + CCur(Replace(msGrilla.TextMatrix(i, COL_NETO_GRAVADO), "$", ""))
        totalIVA = totalIVA + CCur(Replace(msGrilla.TextMatrix(i, COL_IVA), "$", ""))
        total = total + CCur(Replace(msGrilla.TextMatrix(i, COL_TOTAL), "$", ""))
        totalExento = totalExento + CCur(Replace(msGrilla.TextMatrix(i, COL_EXENTO), "$", ""))
    Next
    
    Me.txtNetoGravado.text = Format(totalGravado, formatoMoneda)
    Me.txtTotalIva.text = Format(totalIVA, formatoMoneda)
    Me.txtTotal.text = Format(total, formatoMoneda)
    Me.txtExento.text = Format(totalExento, formatoMoneda)

    
End Sub

Private Function validarDatos() As Boolean
    validarDatos = True
    If Len(Me.txtImporte) = 0 Or Me.txtImporte.ClipText = "0" Then
        MsgBox "Debe ingresar un importe para agregar el concepto.", vbCritical, "Error: Importe inválido"
        validarDatos = False
    End If
End Function


Private Function agregarPercepcion() As Boolean
Dim datosPercecpion() As String
Dim datosJurisdiccion() As String
Dim importeIVA() As Currency

Dim stmp() As String


agregarPercepcion = False
On Error GoTo errores
    
    datosPercecpion = separarCodigoDescripcion(Me.cmbTipoPercepcion.text)
    datosJurisdiccion = separarCodigoDescripcion(Me.cmbJurisdiccion.text)
        
    MSPercepcion.TextMatrix(MSPercepcion.Rows - 1, COL_TIPO_PERCEPCION) = datosPercecpion(1)
    MSPercepcion.TextMatrix(MSPercepcion.Rows - 1, COL_ID_TIPO_PERCEPCION) = datosPercecpion(0)
    
    MSPercepcion.TextMatrix(MSPercepcion.Rows - 1, COL_JURISDICCION) = datosJurisdiccion(1)
    MSPercepcion.TextMatrix(MSPercepcion.Rows - 1, COL_ID_JURISDICCION) = datosJurisdiccion(0)

    MSPercepcion.TextMatrix(MSPercepcion.Rows - 1, COL_PERCEPCION) = Format(Me.txtImportePercepcion, formatoMoneda)
   
    
    agregarPercepcion = True
    Exit Function
errores:
    
        MsgBox Err.Description, Err.Number
        Err.Clear
End Function


Private Function cargarCabeceraCompra() As tCabecera_Compra
On Error GoTo errores

Dim tCabeceraCompras As tCabecera_Compra
Dim ltEmpresa As tEmpresa
Dim ltTipoComprobante As tTipoComprobante
Dim ltMonedas As tMoneda
Dim ltDocumento As tDocumento
Dim ltCodigoOperacion As tCodigoOperacion
Dim strtmp() As String


    tCabeceraCompras.existe = False

    'Cargo la empresa
    ltEmpresa.existe = False
    ltEmpresa.Identificador = IIf(Me.txtNroIdComprador.ClipText <> "", Me.txtNroIdComprador.ClipText, 0)
    ltEmpresa.razonSocial = Trim(Me.txtRazonSocial)
    If ltEmpresa.Identificador <> 0 Then
        ltEmpresa.existe = True
    Else
        MsgBox "Debe ingresar una empresa para continuar", vbCritical, "Error: Ingresar Identificador"
        Exit Function
    End If
    ' Fin Carga de la empresa
    
    'Cargo el Tipo de Comprobante
    ltTipoComprobante.existe = False
    strtmp = separarCodigoDescripcion(Me.cmbTipoComprobante)
    ltTipoComprobante.codigo = strtmp(0)
    ltTipoComprobante.descripcion = strtmp(1)
    If ltTipoComprobante.codigo <> "" Then ltTipoComprobante.existe = True
    'Fin Carga Tipo de Comprobante
    
    'Cargo la Moneda
    ltMonedas.existe = False
    strtmp = separarCodigoDescripcion(Me.cmbMoneda)
    ltMonedas.codigo = strtmp(0)
    ltMonedas.descripcion = strtmp(1)
    If ltMonedas.codigo <> "" Then ltMonedas.existe = True
    'Fin Carga de la Moneda
    

    'Cargo el tipo de documento del Vendedor
    ltDocumento.existe = False
    strtmp = separarCodigoDescripcion(Me.cmbTipoDocVendedor)
    ltDocumento.codigo = strtmp(0)
    ltDocumento.descripcion = strtmp(1)
    If ltDocumento.codigo <> "" Then ltDocumento.existe = True
    'Fin de carga del documento de Vendedor
    
    'Cargo el tipo de operacion de ventas
    ltCodigoOperacion.existe = False
    strtmp = separarCodigoDescripcion(Me.cmbCodigoOperacion)
    ltCodigoOperacion.codigo = strtmp(0)
    ltCodigoOperacion.descripcion = strtmp(1)
    If ltCodigoOperacion.codigo <> "" Then ltCodigoOperacion.existe = True
    'Fin de carga del tipo de operacion de ventas
   

    tCabeceraCompras.Empresa = ltEmpresa
    tCabeceraCompras.tipoComprobabte = ltTipoComprobante
    tCabeceraCompras.moneda = ltMonedas
    
    tCabeceraCompras.tipoDocumento = ltDocumento
    tCabeceraCompras.codigoOperacion = ltCodigoOperacion
    
    tCabeceraCompras.fechaImputacion = Me.dtpFechaImputacion.value
    tCabeceraCompras.fechaCompra = Me.dtpFechaComprobante.value
    
    tCabeceraCompras.vendedorId = Me.txtNroIdVendedor.ClipText
    tCabeceraCompras.razonSocialVendedor = Trim(Me.txtRazonSocialVendedor.text)
    tCabeceraCompras.nroComprobante = Me.txtNroComprobante.ClipText
    
    tCabeceraCompras.tipoCambio = Me.txtTipoCambio.text
    tCabeceraCompras.PuntoDeVenta = Me.txtPuntoVta.ClipText
    tCabeceraCompras.totalOperacion = CCur(Replace(Me.txtTotalComprobante, "$", ""))
    
    tCabeceraCompras.existe = tCabeceraCompras.Empresa.existe And _
                            tCabeceraCompras.codigoOperacion.existe And _
                            tCabeceraCompras.moneda.existe And _
                            tCabeceraCompras.tipoComprobabte.existe And _
                            tCabeceraCompras.tipoDocumento.existe And _
                            (Len(tCabeceraCompras.vendedorId) > 0) And _
                            (Len(tCabeceraCompras.razonSocialVendedor) > 0) And _
                            (Len(tCabeceraCompras.nroComprobante) > 0) And _
                            (Len(tCabeceraCompras.PuntoDeVenta) > 0)
                            
    cargarCabeceraCompra = tCabeceraCompras
    Exit Function

errores:
    MsgBox Err.Description, Err.Number
    Err.Clear
    'tCabeceraCompras.existe = False
    cargarCabeceraCompra = tCabeceraCompras
End Function


Private Function cargarDetalleCompra(idCompra As Double, iRow As Integer) As tDetalle_compra
On Error GoTo errores
Dim tDetalleCompras As tDetalle_compra
Dim ltAlicuota As tAlicuota
Dim ltConcepto As tConcepto

    tDetalleCompras.existe = False
   
    ltConcepto.codigo = Me.msGrilla.TextMatrix(iRow, COL_CODIGO_CONCEPTO)
    ltConcepto.descripcion = Me.msGrilla.TextMatrix(iRow, COL_CONCEPTO)
    If ltConcepto.codigo <> "" Then ltConcepto.existe = True
    
    
    ltAlicuota.alicuta = Me.msGrilla.TextMatrix(iRow, COL_CODIGO_ALICUOTA)
    ltAlicuota.descripcion = Me.msGrilla.TextMatrix(iRow, COL_ALICUOTA)
    If ltAlicuota.alicuta <> "" Then ltAlicuota.existe = True
    
    
    tDetalleCompras.compra_id = idCompra
    tDetalleCompras.Concepto = ltConcepto
    tDetalleCompras.alicuota = ltAlicuota
    tDetalleCompras.neto_gravado = CCur(Replace(Me.msGrilla.TextMatrix(iRow, COL_NETO_GRAVADO), "$", ""))
    tDetalleCompras.iva = CCur(Replace(Me.msGrilla.TextMatrix(iRow, COL_IVA), "$", ""))
    tDetalleCompras.exento = CCur(Replace(Me.msGrilla.TextMatrix(iRow, COL_EXENTO), "$", ""))
    tDetalleCompras.total = CCur(Replace(Me.msGrilla.TextMatrix(iRow, COL_TOTAL), "$", ""))
    tDetalleCompras.existe = True
    cargarDetalleCompra = tDetalleCompras
    Exit Function

errores:
    MsgBox Err.Description, Err.Number
    Err.Clear
    cargarDetalleCompra = tDetalleCompras

End Function

Private Function cargarPercepcionCompra(idCompra As Double, iRow As Integer) As tPercepcion_compra
On Error GoTo errores
Dim tPercepcionCompras As tPercepcion_compra
Dim ltTipoPercepcion As tTipoPercepcion
Dim ltJurisdiccion As tJurisdiccion

    tPercepcionCompras.existe = False
   
    ltTipoPercepcion.percipcion = Me.MSPercepcion.TextMatrix(iRow, COL_ID_TIPO_PERCEPCION)
    ltTipoPercepcion.descripcion = Me.MSPercepcion.TextMatrix(iRow, COL_TIPO_PERCEPCION)
    If ltTipoPercepcion.percipcion <> "" Then ltTipoPercepcion.existe = True
    
    ltJurisdiccion.jurisdiccion = "00"
    ltJurisdiccion.descripcion = "NO APLICA"
    ltJurisdiccion.exite = True
    If Me.MSPercepcion.TextMatrix(iRow, COL_ID_JURISDICCION) <> "" Then
        ltJurisdiccion.jurisdiccion = Me.MSPercepcion.TextMatrix(iRow, COL_ID_JURISDICCION)
        ltJurisdiccion.descripcion = Me.MSPercepcion.TextMatrix(iRow, COL_JURISDICCION)
    End If
    
    
    tPercepcionCompras.compra_id = idCompra
    tPercepcionCompras.tuTipoPercepcion = ltTipoPercepcion
    tPercepcionCompras.tuJurisdiccion = ltJurisdiccion
    tPercepcionCompras.totalPercepcion = CCur(Replace(Me.MSPercepcion.TextMatrix(iRow, COL_PERCEPCION), "$", ""))

    tPercepcionCompras.existe = True
    cargarPercepcionCompra = tPercepcionCompras
    Exit Function

errores:
    MsgBox Err.Description, Err.Number
    Err.Clear
    cargarPercepcionCompra = tPercepcionCompras

End Function
Private Sub calcularTotalesPercepciones()
On Error Resume Next
Dim i As Integer
Dim percepcion As Currency
Dim totalIIBB As Currency
Dim totalGcias As Currency
Dim totalIVA As Currency
Dim totalOtras As Currency


    Me.txtIIBB.text = ""
    Me.txtIVA.text = ""
    Me.txtGcias.text = ""
    Me.txtOtras.text = ""
    
    totalIIBB = 0
    totalIVA = 0
    totalGcias = 0
    totalOtras = 0
    percepcion = 0
    
    For i = 1 To Me.MSPercepcion.Rows - 2
        
        percepcion = CCur(Replace(Me.MSPercepcion.TextMatrix(i, COL_PERCEPCION), "$", ""))
        Select Case Me.MSPercepcion.TextMatrix(i, COL_ID_TIPO_PERCEPCION)
            Case CODIGOIIBB
                totalIIBB = totalIIBB + percepcion
            Case CODIGOIVA
                totalIVA = totalIVA + percepcion
            Case CODIGOGCIAS
                totalGcias = totalGcias + percepcion
            Case Else
                totalOtras = totalOtras + percepcion
            
        End Select
       
    Next
    
    Me.txtIIBB.text = Format(totalIIBB, formatoMoneda)
    Me.txtIVA.text = Format(totalIVA, formatoMoneda)
    Me.txtGcias.text = Format(totalGcias, formatoMoneda)
    Me.txtOtras.text = Format(totalOtras, formatoMoneda)
    
End Sub


Private Sub calcularTotalComprobante()
On Error Resume Next
Dim totalComprobante As Currency
Dim totalConceptos As Currency
Dim totalPercepciones As Currency

    totalConceptos = CCur(Replace(txtTotal.text, "$", ""))
    totalPercepciones = CCur(Replace(txtOtras.text, "$", "")) + CCur(Replace(txtIVA.text, "$", "")) + CCur(Replace(txtGcias.text, "$", "")) + CCur(Replace(txtIIBB.text, "$", ""))

    totalComprobante = totalConceptos + totalPercepciones
    Me.txtTotalComprobante.text = Format(totalComprobante, formatoMoneda)
End Sub

Private Function validarDatosPercepciones() As Boolean
    validarDatosPercepciones = True
    If Len(Me.txtImportePercepcion) = 0 Or Me.txtImportePercepcion.ClipText = "0" Then
        MsgBox "Debe ingresar un importe para agregar la percepción", vbCritical, "Error: Importe inválido"
        validarDatosPercepciones = False
    End If
End Function

Private Sub cargarFormularioModificacion()
On Error GoTo errores

    Dim sql_cabecera As String
    Dim sql_moneda As String
    Dim sql_tipoDocVendedor As String
    Dim sql_tipoComprobante As String
    Dim sql_codigoOperacionCompras As String
    Dim sql_conceptos As String
    Dim sql_alicuotas As String
    Dim sql_jurisdicciones As String
    Dim sql_percecpiones As String
    
    
    Dim oRec As New ADODB.Recordset
    
    sql_cabecera = "SELECT C.* FROM CABECERA_COMPRA C WHERE C.COMPRA_ID=" & Me.idCompra
    
   
    With oRec
        .ActiveConnection = oCon
        .CursorType = adOpenStatic
        .CursorLocation = adUseClient
        .Source = sql_cabecera
        .Open
    End With
    
    If oRec.RecordCount = 1 Then
        Me.txtNroIdComprador.Mask = ""
        Me.txtNroIdComprador.text = oRec!empresa_id
        Me.txtNroIdComprador.PromptInclude = False
        Me.txtNroIdComprador.Mask = "##-########-#"
        Me.txtNroIdComprador.Enabled = False
        txtNroIdComprador_LostFocus
        
        Me.dtpFechaImputacion.value = Format(oRec!FECHA_IMPUTACION, "short date")
        Me.dtpFechaComprobante.value = Format(oRec!FECHA_COMPRA, "short date")
        Me.txtPuntoVta.text = oRec!PUNTO_VENTA
        Me.txtNroComprobante.text = oRec!NRO_COMPROBANTE
        Me.txtNroIdVendedor.PromptInclude = False
        Me.txtNroIdVendedor.text = oRec!VENDEDOR_ID
        Me.txtRazonSocialVendedor.text = oRec!RAZON_SOCIAL_VENDEDOR

        sql_moneda = "SELECT CODIGO, ISNULL(DESCRIPCION,'') AS DESCRIPCION, ISNULL(DEFECTO,0) AS DEFECTO FROM MONEDA WHERE ISNULL(VISIBLE,0) = 1 order by CODIGO "
        cargarCombo Me.cmbMoneda, sql_moneda
        setLisIndexCombo Me.cmbMoneda, oRec!TIPO_MONEDA_ID
        Me.txtTipoCambio.text = oRec!TIPO_DE_CAMBIO
        Me.cmbMoneda.Enabled = False

        sql_tipoDocVendedor = "SELECT CODIGO, ISNULL(DESCRIPCION,'') AS DESCRIPCION, ISNULL(DEFECTO,0) AS DEFECTO FROM DOCUMENTO WHERE ISNULL(VISIBLE,0) = 1 order by CODIGO "
        cargarCombo Me.cmbTipoDocVendedor, sql_tipoDocVendedor
        setLisIndexCombo Me.cmbTipoDocVendedor, oRec!TIPO_DOCUMENTO_ID

        sql_tipoComprobante = "SELECT CODIGO, ISNULL(DESCRIPCION,'') AS DESCRIPCION, ISNULL(DEFECTO,0) AS DEFECTO FROM TIPO_COMPROBANTE WHERE ISNULL(VISIBLE,0) = 1 order by CODIGO "
        cargarCombo Me.cmbTipoComprobante, sql_tipoComprobante
        setLisIndexCombo Me.cmbTipoComprobante, oRec!TIPO_COMPROBANTE_ID
        Me.cmbTipoComprobante.Enabled = False

        sql_codigoOperacionCompras = "SELECT CODIGO, ISNULL(DESCRIPCION,'') AS DESCRIPCION, ISNULL(DEFECTO,0) AS DEFECTO FROM CODIGO_OPERACION WHERE ISNULL(VISIBLE,0) = 1 order by CODIGO "
        cargarCombo Me.cmbCodigoOperacion, sql_codigoOperacionCompras
        setLisIndexCombo Me.cmbCodigoOperacion, oRec!CODIGO_OPERACION_ID
        'Me.cmbCodigoOperacion.Enabled = False

        sql_conceptos = "SELECT CODIGO, ISNULL(DESCRIPCION,'') AS DESCRIPCION, ISNULL(DEFECTO,0) AS DEFECTO FROM dbo.CONCEPTOCPA WHERE ISNULL(VISIBLE,0) = 1 order by CODIGO "
        cargarCombo Me.cmbConcepto, sql_conceptos

        sql_alicuotas = "SELECT ALICUOTA, ISNULL(DESCRIPCION,'') AS DESCRIPCION, ISNULL(DEFECTO,0) AS DEFECTO FROM ALICUOTA WHERE ISNULL(VISIBLE,0) = 1"
        cargarCombo Me.cmbAlicuota, sql_alicuotas
    
        sql_jurisdicciones = "SELECT CODIGO, ISNULL(DESCRIPCION,'') AS DESCRIPCION, ISNULL(DEFECTO,0) AS DEFECTO FROM dbo.JURISDICCION WHERE ISNULL(VISIBLE,0) = 1 order by CODIGO "
        cargarCombo Me.cmbJurisdiccion, sql_jurisdicciones
    
        sql_percecpiones = "SELECT CODIGO, ISNULL(DESCRIPCION,'') AS DESCRIPCION, ISNULL(DEFECTO,0) AS DEFECTO FROM dbo.TIPO_PERCEPCION WHERE ISNULL(VISIBLE,0) = 1 order by CODIGO "
        cargarCombo Me.cmbTipoPercepcion, sql_percecpiones
    
    End If
    inicializarGrilla
    inicializarGrillaPercepcion
    
    
    cargarDetalle
    cargarPercepcion
    calcularTotalComprobante
    
    Exit Sub
errores:
    On Error Resume Next
    If oCon.Errors.Count >= 1 Then
        MsgBox oCon.Errors(0).Description, vbCritical, "Nro de Error: " & oCon.Errors(0).Number
        oCon.Errors.Clear
    Else
        MsgBox Err.Number & "-" & Err.Description, vbCritical, "Error: Al Cargar modificar Comprobante"
        Err.Clear
    End If

    Set oRec = Nothing

End Sub


Private Sub cargarDetalle()


    Dim sql_detalle As String
 
    Dim oRec As New ADODB.Recordset
    
    sql_detalle = "SELECT " + _
                  "C.DESCRIPCION AS CONCEPTO, " + _
                  "A.DESCRIPCION AS ALICUOTA, " + _
                  "DC.NETO_GRAVADO, " + _
                  "DC.EXENTO, " + _
                  "DC.IVA, " + _
                  "DC.TOTAL, " + _
                  "DC.CONCEPTO_ID, " + _
                  "DC.ALICUOTA_ID " + _
                  "FROM DETALLE_COMPRA DC INNER JOIN CONCEPTOCPA C ON DC.CONCEPTO_ID = C.CODIGO " + _
                  "INNER JOIN ALICUOTA A ON DC.ALICUOTA_ID = A.ALICUOTA " + _
                  "WHERE DC.CABECERA_COMPRA_ID =" + Me.idCompra
    
    With oRec
        .ActiveConnection = oCon
        .CursorType = adOpenStatic
        .CursorLocation = adUseClient
        .Source = sql_detalle
        .Open
    End With
    
    Dim i As Integer
    Dim j As Integer
    For i = 1 To oRec.RecordCount
        With msGrilla
            .TextMatrix(i, COL_CONCEPTO) = oRec!Concepto
            .TextMatrix(i, COL_CODIGO_CONCEPTO) = oRec!concepto_id
            .TextMatrix(i, COL_ALICUOTA) = oRec!alicuota & " %"
            .TextMatrix(i, COL_CODIGO_ALICUOTA) = oRec!ALICUOTA_ID
            .TextMatrix(i, COL_NETO_GRAVADO) = Format(oRec!neto_gravado, formatoMoneda)
            .TextMatrix(i, COL_IVA) = Format(oRec!iva, formatoMoneda)
            .TextMatrix(i, COL_TOTAL) = Format(oRec!total, formatoMoneda)
            .TextMatrix(i, COL_EXENTO) = Format(oRec!exento, formatoMoneda)
            
        End With
        oRec.MoveNext
        msGrilla.Rows = msGrilla.Rows + 1
    Next i
    
    calcularTotales

End Sub


Private Sub cargarPercepcion()


    Dim sql_detalle As String
 
    Dim oRec As New ADODB.Recordset
    
    sql_detalle = "SELECT " + _
                  "T.DESCRIPCION AS PERCEPCION, " + _
                  "J.DESCRIPCION AS JURISDICCION, " + _
                  "P.IMPORTE_PERCEPCION, " + _
                  "P.TIPO_PERCEPCION_ID, " + _
                  "P.JURISDICCION_ID " + _
                  "FROM " + _
                  "PERCEPCION_COMPRA P INNER JOIN TIPO_PERCEPCION T ON P.TIPO_PERCEPCION_ID = T.CODIGO " + _
                  "INNER JOIN JURISDICCION J ON P.JURISDICCION_ID = J.CODIGO " + _
                  "WHERE P.CABECERA_COMPRA_ID=" & Me.idCompra
    
    With oRec
        .ActiveConnection = oCon
        .CursorType = adOpenStatic
        .CursorLocation = adUseClient
        .Source = sql_detalle
        .Open
    End With
    
'Private Const COL_TIPO_PERCEPCION = 1
'Private Const COL_JURISDICCION = 2
'Private Const COL_PERCEPCION = 3
'Private Const COL_ID_TIPO_PERCEPCION = 4
'Private Const COL_ID_JURISDICCION = 5

    Dim i As Integer
    For i = 1 To oRec.RecordCount
        With MSPercepcion
            .TextMatrix(i, COL_TIPO_PERCEPCION) = oRec!percepcion
            .TextMatrix(i, COL_JURISDICCION) = ""
            If oRec!JURISDICCION_ID <> "00" Then .TextMatrix(i, COL_JURISDICCION) = oRec!jurisdiccion
            .TextMatrix(i, COL_PERCEPCION) = Format(oRec!IMPORTE_PERCEPCION, formatoMoneda)
            .TextMatrix(i, COL_ID_TIPO_PERCEPCION) = oRec!TIPO_PERCEPCION_ID
            .TextMatrix(i, COL_ID_JURISDICCION) = oRec!JURISDICCION_ID

            
        End With
        oRec.MoveNext
        MSPercepcion.Rows = MSPercepcion.Rows + 1
    Next i
    
    calcularTotalesPercepciones

End Sub

