VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmVentas 
   BorderStyle     =   0  'None
   Caption         =   "Ventas"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   12495
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   12495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9060
      TabIndex        =   37
      Top             =   7140
      Width           =   1335
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10740
      TabIndex        =   36
      Top             =   7140
      Width           =   1335
   End
   Begin VB.Frame frmGrillaConceptos 
      Caption         =   "Conceptos"
      Height          =   3750
      Left            =   60
      TabIndex        =   32
      Top             =   3300
      Width           =   12375
      Begin VB.Frame frmTotales 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   120
         TabIndex        =   40
         Top             =   3060
         Width           =   12135
         Begin VB.TextBox txtNetoGravado 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   5235
            TabIndex        =   44
            Top             =   180
            Width           =   1610
         End
         Begin VB.TextBox txtTotalIva 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   6885
            TabIndex        =   43
            Top             =   180
            Width           =   1610
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   10185
            TabIndex        =   42
            Top             =   180
            Width           =   1610
         End
         Begin VB.TextBox txtExento 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   8520
            TabIndex        =   41
            Top             =   180
            Width           =   1610
         End
         Begin VB.Label lblTotales 
            Caption         =   "Totales:"
            Height          =   195
            Left            =   4320
            TabIndex        =   45
            Top             =   240
            Width           =   855
         End
      End
      Begin MSFlexGridLib.MSFlexGrid msGrilla 
         Height          =   2940
         Left            =   120
         TabIndex        =   33
         Top             =   180
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   5186
         _Version        =   393216
         Cols            =   6
         HighLight       =   2
         ScrollBars      =   2
         SelectionMode   =   1
      End
   End
   Begin VB.Frame frmAgregarConceptos 
      Caption         =   "Carga de Conceptos"
      Height          =   815
      Left            =   60
      TabIndex        =   25
      Top             =   2460
      Width           =   12375
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar Concepto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10440
         TabIndex        =   17
         Top             =   180
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
      Begin VB.ComboBox cmbAlicuota 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   300
         Width           =   1815
      End
      Begin VB.ComboBox cmbConcepto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   300
         Width           =   2535
      End
      Begin VB.Label lblImporteConcepto 
         Caption         =   "Importe a Registrar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7440
         TabIndex        =   31
         Top             =   360
         Width           =   1395
      End
      Begin VB.Label lblAlicuota 
         Caption         =   "Alicuota I.V.A."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3960
         TabIndex        =   30
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label lblConcepto 
         Caption         =   "Concepto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   29
         Top             =   360
         Width           =   795
      End
   End
   Begin VB.Frame frmDatosComprobante 
      Caption         =   "Ventas - Datos del Comprobante"
      Height          =   2280
      Left            =   60
      TabIndex        =   18
      Top             =   60
      Width           =   12375
      Begin VB.CheckBox chkLote 
         Caption         =   "Habilita Lote"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6360
         TabIndex        =   5
         Top             =   780
         Width           =   1275
      End
      Begin MSMask.MaskEdBox txtNroComprobante 
         Height          =   315
         Left            =   9300
         TabIndex        =   6
         Top             =   720
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
      Begin VB.TextBox txtRazonSocialComprador 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8220
         TabIndex        =   10
         Top             =   1260
         Width           =   3915
      End
      Begin MSMask.MaskEdBox txtNroIdComprador 
         Height          =   315
         Left            =   5640
         TabIndex        =   9
         Top             =   1260
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
      Begin MSMask.MaskEdBox txtNroIdVendedor 
         Height          =   315
         Left            =   2340
         TabIndex        =   0
         Top             =   240
         Width           =   1515
         _ExtentX        =   2672
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
      Begin VB.ComboBox cmbCodigoOperacion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7500
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1740
         Width           =   3075
      End
      Begin VB.ComboBox cmbPuntoVenta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5220
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   915
      End
      Begin VB.TextBox txtRazonSocial 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5100
         TabIndex        =   1
         Top             =   240
         Width           =   4755
      End
      Begin MSComCtl2.DTPicker dtpFechaComprobante 
         Height          =   315
         Left            =   10680
         TabIndex        =   2
         Top             =   240
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
         Format          =   53018625
         CurrentDate     =   42201
      End
      Begin VB.ComboBox cmbTipoComprobante 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   2715
      End
      Begin VB.ComboBox cmbTipoDocComprador 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1260
         Width           =   2235
      End
      Begin VB.ComboBox cmbMoneda 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1740
         Width           =   2475
      End
      Begin MSMask.MaskEdBox txtNroComprobanteHasta 
         Height          =   315
         Left            =   11040
         TabIndex        =   7
         Top             =   720
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
         Left            =   4800
         TabIndex        =   12
         Top             =   1740
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
      Begin VB.Label lblTipoCambio 
         Caption         =   "Tipo de Cambio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3540
         TabIndex        =   39
         Top             =   1800
         Width           =   1155
      End
      Begin VB.Label lblNroComprobanteHasta 
         Caption         =   "Hasta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10500
         TabIndex        =   38
         Top             =   780
         Width           =   555
      End
      Begin VB.Label lblRazonSocialComprador 
         Caption         =   "Comprador:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7320
         TabIndex        =   35
         Top             =   1320
         Width           =   915
      End
      Begin VB.Label lblNroIdentificadorComprador 
         Caption         =   "Nro. Id. Comprador:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   34
         Top             =   1320
         Width           =   1515
      End
      Begin VB.Label lblCodigoOperacionVentas 
         Caption         =   "Código Operación:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6060
         TabIndex        =   28
         Top             =   1800
         Width           =   1395
      End
      Begin VB.Label lblNumeroComprobanteDesde 
         Caption         =   "Comprobante Desde:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7740
         TabIndex        =   27
         Top             =   780
         Width           =   1635
      End
      Begin VB.Label lblPuntoVenta 
         Caption         =   "Punto Venta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4140
         TabIndex        =   26
         Top             =   780
         Width           =   1215
      End
      Begin VB.Label lblRazonSocial 
         Caption         =   "Razón Social:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4020
         TabIndex        =   24
         Top             =   300
         Width           =   1035
      End
      Begin VB.Label lblIdentificadorVendedor 
         Caption         =   " Nro. Identificador Vendedor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   300
         Width           =   2115
      End
      Begin VB.Label lblFecha 
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10080
         TabIndex        =   22
         Top             =   300
         Width           =   735
      End
      Begin VB.Label lblTipoComprobante 
         Caption         =   "T. Comprobate:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   21
         Top             =   780
         Width           =   1575
      End
      Begin VB.Label lblDocumentoComprador 
         Caption         =   "Tipo Doc. Comprador:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   20
         Top             =   1320
         Width           =   1635
      End
      Begin VB.Label lblMoneda 
         Caption         =   "Moneda:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   19
         Top             =   1800
         Width           =   735
      End
   End
   Begin VB.Menu mnu_Edicion 
      Caption         =   "&Edición"
      Begin VB.Menu mnu_Edicion_Conceptos 
         Caption         =   "&Concepto"
         Begin VB.Menu mnu_Edicion_BorrarConcepto 
            Caption         =   "Borrar Concepto"
         End
      End
   End
End
Attribute VB_Name = "frmVentas"
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
Public operacion As String
Public idVentas As String




Private Sub chkLote_Click()
    Me.txtNroComprobanteHasta.text = Format(Me.txtNroComprobante.ClipText, "000000#")
    If chkLote.value = 1 Then
    
        Me.txtNroComprobanteHasta.Enabled = True
    Else
        Me.txtNroComprobanteHasta.Enabled = False
        
    End If
End Sub




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




Private Sub cmbTipoDocComprador_Click()
Dim strtmp() As String
Dim ltDocumento As tDocumento
Dim tmp As String
    strtmp = Utils.separarCodigoDescripcion(Me.cmbTipoDocComprador)
    ltDocumento.codigo = strtmp(0)
    ltDocumento.descripcion = strtmp(1)
    If ltDocumento.codigo = CODIGOCUIT Or ltDocumento.codigo = CODIGOCUIL Then
        Me.txtNroIdComprador.Mask = "##-########-#"
    Else
        tmp = Me.txtNroIdComprador.ClipText
        Me.txtNroIdComprador.Mask = ""
        Me.txtNroIdComprador.text = tmp
    End If
End Sub

Private Sub cmdAgregar_Click()
    Dim ltdCabeceraVenta As tCabecera_Venta
    ltdCabeceraVenta = cargarCabeceraVentas
    If ltdCabeceraVenta.existe Then
        If validarDatos Then
            If agregarConcepto Then
                msGrilla.RowSel = msGrilla.Rows - 1
                msGrilla.Rows = msGrilla.Rows + 1
                calcularTotales
                Me.txtImporte.text = ""
                Me.cmbConcepto.SetFocus
                frmDatosComprobante.Enabled = False
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
    Dim ltCabeceraVenta As tCabecera_Venta
    Dim ventaId As Double
    Dim i As Integer
    Dim b As Boolean
    Dim graboDetalle As Boolean

    ltCabeceraVenta = cargarCabeceraVentas

    If ltCabeceraVenta.existe Then
        If msGrilla.Rows > 2 Then
                oCon.BeginTrans
                b = True
                If Me.idVentas <> 0 Then
                    b = False
                    If borrarCabeceraDetalle(Me.idVentas) > 0 Then
                        b = True
                    End If
                End If
                
                ventaId = insertarCabeceraVenta(ltCabeceraVenta)
                If ventaId > 0 And b Then
                    
                    For i = 1 To Me.msGrilla.Rows - 2
                        Dim ltDetalleVenta As tDetalle_Venta
                        ltDetalleVenta = cargarDetalleVentas(ventaId, i)
                        graboDetalle = False
                        If ltDetalleVenta.existe Then
                            If insertarDetalleVenta(ltDetalleVenta) Then
                                graboDetalle = True
                            End If
                        Else
                            Exit For
                        End If
                    Next
                    If graboDetalle = True Then
                        oCon.CommitTrans
                        MsgBox "La registración se ha realizado en forma correcta.", vbInformation, "Registracion OK"
                        If Me.idVentas <> 0 Then
                            Unload Me
                        Else
                        
                            limpiarFormulario
                            Me.txtNroIdVendedor.SetFocus
                        End If
                    Else
                        oCon.RollbackTrans
                    End If
                Else
                    oCon.RollbackTrans
                End If
    
        Else
            MsgBox "Debe ingresar al menos un concepto para poder grabar.", vbCritical, "Error: Detalle de Conceptos"
        End If
    Else
            MsgBox "Datos de la Cabecera Incompletos", vbCritical, "Error: Datos de la Cabecera"
        End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
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
        If msGrilla.Rows = 2 Then
            Me.frmDatosComprobante.Enabled = True
        End If
    End If

End Sub

Private Sub msGrilla_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu mnu_edicion_Conceptos
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
Dim ltPuntoDeVenta As tPuntoDeVenta
Dim ltTipoComprobante As tTipoComprobante
Dim strtmp() As String
    If Len(Me.txtNroComprobante.ClipText) > 0 And Len(Me.txtNroIdVendedor.ClipText) > 0 Then
        strtmp = Utils.separarCodigoDescripcion(Me.cmbTipoComprobante)
        ltTipoComprobante.codigo = strtmp(0)
        ltTipoComprobante.descripcion = strtmp(1)
        ltPuntoDeVenta.empresa_id = Me.txtNroIdVendedor.ClipText
        ltPuntoDeVenta.PuntoDeVenta = Me.cmbPuntoVenta.text
        
        ltPuntoDeVenta.puntoVentaId = recuperarIDPuntoDeVenta(ltPuntoDeVenta.empresa_id, ltPuntoDeVenta.PuntoDeVenta)
        If ltPuntoDeVenta.puntoVentaId > 0 Then
            If existeRegistradoNroComprobanteEmpresa(ltPuntoDeVenta, Format(Me.txtNroComprobante, "0000000#"), ltTipoComprobante.codigo) Then
                MsgBox "El número de comprobante " & Me.cmbPuntoVenta.text & "-" & Format(Me.txtNroComprobante, "0000000#") & " ya se encuentra registrado para el Vendedor " & Format(Me.txtNroIdVendedor.ClipText, "##-########-#") & " con el tipo de comprobante " & ltTipoComprobante.descripcion, vbCritical, "Error: Ya existe el Comprobante"
                Me.txtNroComprobante.SetFocus
            Else
                'If Len(txtNroComprobanteHasta.ClipText) = 0 And Me.chkLote.value = 1 Then
                    Me.txtNroComprobanteHasta.text = Format(Me.txtNroComprobante.ClipText, "000000#")
                'End If
            End If
        End If
    End If
    
End Sub

Private Sub txtNroComprobanteHasta_GotFocus()
    With txtNroComprobanteHasta
        .SelStart = 0
        .SelLength = Len(.text)
    End With
End Sub



Private Sub txtNroIdComprador_GotFocus()
    With txtNroIdComprador
        .SelStart = 0
        .SelLength = Len(.text)
    End With
End Sub

Private Sub txtNroIdComprador_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Sendkeys "{tab}"
    End If
   
   If Not IsNumeric(Chr(KeyAscii)) Then
       If KeyAscii <> vbKeyBack Then
          KeyAscii = 0
       End If
    End If
End Sub

Private Sub txtNroIdComprador_LostFocus()
Dim ltTipoComprobante As tTipoComprobante
Dim ltDocumento As tDocumento
Dim strtmp() As String


        strtmp = Utils.separarCodigoDescripcion(Me.cmbTipoComprobante)
        ltTipoComprobante.codigo = strtmp(0)
        ltTipoComprobante.descripcion = strtmp(1)
        Me.txtRazonSocialComprador.Enabled = True
        'If ltTipoComprobante.CODIGO = TIPOCOMROBANTEB And IIf(txtNroIdComprador.ClipText = "", 0, txtNroIdComprador.ClipText) = 0 Then
        If IIf(txtNroIdComprador.ClipText = "", 0, txtNroIdComprador.ClipText) = 0 Then
            txtNroIdComprador.Mask = ""
            txtNroIdComprador.PromptInclude = False
            txtNroIdComprador.text = "00000000000"
            txtNroIdComprador.Mask = "##-########-#"
            txtNroIdComprador.PromptInclude = True
            Me.txtRazonSocialComprador.text = "CONSUMIDOR FINAL"
            Me.txtRazonSocialComprador.Enabled = False
        Else
            If Len(txtNroIdComprador.ClipText) > 0 Then
                strtmp = Utils.separarCodigoDescripcion(Me.cmbTipoDocComprador)
                ltDocumento.codigo = strtmp(0)
                ltDocumento.descripcion = strtmp(1)
                If ltDocumento.codigo = CODIGOCUIT Or ltDocumento.codigo = CODIGOCUIL Then
                    If Len(Me.txtNroIdComprador.ClipText) = 11 Then
                        If ValidarCuit(Me.txtNroIdComprador.ClipText) Then
                            If Me.txtRazonSocialComprador.text = "CONSUMIDOR FINAL" Then Me.txtRazonSocialComprador.SetFocus
                        Else
                            MsgBox "El CUIT/CUIL ingresado no es válido", vbCritical, "Error: Código de Verificación"
                            Me.txtNroIdComprador.SetFocus
                            Me.txtRazonSocialComprador.text = ""
                        End If
                    Else
                        MsgBox "La longitud del campo para el numero de CUIT/CUIL no es valida", vbCritical, "Error: Longitud CUIT/CUIL"
                        Me.txtNroIdComprador.SetFocus
                        Me.txtRazonSocialComprador.text = ""
                    End If
                Else
                    If Me.txtRazonSocialComprador.text = "CONSUMIDOR FINAL" Then Me.txtRazonSocialComprador.SetFocus
                End If
            End If
        End If
End Sub

Private Sub txtNroIdVendedor_GotFocus()
    
    With txtNroIdVendedor
        .SelStart = 0
        .SelLength = Len(.text)
    End With
    
    txtRazonSocial.text = ""
    Me.cmbPuntoVenta.Clear
End Sub

Private Sub txtNroIdVendedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
            Dim ltEmpresa As tEmpresa
            ltEmpresa.existe = False
            If txtNroIdVendedor.ClipText <> "" Then
                ltEmpresa = recuperarEmpresaPorIdentificador(txtNroIdVendedor.ClipText)
            End If
            If Not ltEmpresa.existe Then
                Dim indentificador As Double
                frmBusquedaEmpresa.inicializarFormulario indentificador
                txtNroIdVendedor.Mask = ""
                txtNroIdVendedor.text = IIf(indentificador <> 0, indentificador, "")
                txtNroIdVendedor.PromptInclude = False
                txtNroIdVendedor.Mask = "##-########-#"
                txtNroIdVendedor.PromptInclude = True
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

Private Sub txtNroIdVendedor_LostFocus()
    On Error Resume Next
    If Len(txtNroIdVendedor.ClipText) > 0 Then
        If Len(txtNroIdVendedor.ClipText) = 11 Then
            Dim ltEmpresa As tEmpresa
            ltEmpresa = recuperarEmpresaPorIdentificador(txtNroIdVendedor.ClipText)
            If ltEmpresa.existe Then
                txtRazonSocial.text = ltEmpresa.razonSocial
                cargarComboPuntosDeVenta ltEmpresa.Identificador
                cargarTiposComprobantesPermitidos ltEmpresa.Monotributista
                cmbTipoComprobante_LostFocus
            Else
                With txtNroIdVendedor
                    .SelStart = 0
                    .SelLength = Len(.text)
                    .SetFocus
                End With
                Me.cmbPuntoVenta.Clear
            End If
        Else
           
            With txtNroIdVendedor
                .SelStart = 0
                .SelLength = Len(.text)
                .SetFocus
            End With
            
            Me.cmbPuntoVenta.Clear
        End If
    End If
End Sub


Private Sub cargarTiposComprobantesPermitidos(lb_monotributista As Boolean)

Dim sql_tipoComprobante As String
Dim ls_where As String

    sql_tipoComprobante = "SELECT CODIGO, ISNULL(DESCRIPCION,'') AS DESCRIPCION, ISNULL(DEFECTO,0) AS DEFECTO FROM TIPO_COMPROBANTE "
    ls_where = "WHERE ISNULL(VISIBLE,0) = 1 "
    
    If lb_monotributista Then
        'solo puede cargar comprobantes C Y DE EXPORTACION
        ls_where = ls_where + " AND  CODIGO IN ('011','019') "
    Else
        'No puede cargar comprobantes tipo C
        ls_where = ls_where + " AND  CODIGO NOT IN ('011') "
    End If
    
    sql_tipoComprobante = sql_tipoComprobante + ls_where + " ORDER BY CODIGO "
    cargarCombo Me.cmbTipoComprobante, sql_tipoComprobante

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
                
        importeIVA = calculoIVATipoDirecto(CCur(Me.txtImporte.text), CCur(0))
        
        If datosAlicuota(0) <> alicuotaExenta Then
            msGrilla.TextMatrix(msGrilla.Rows - 1, COL_NETO_GRAVADO) = Format(0, formatoMoneda)
            msGrilla.TextMatrix(msGrilla.Rows - 1, COL_IVA) = Format((importeIVA(0)), formatoMoneda)
            msGrilla.TextMatrix(msGrilla.Rows - 1, COL_TOTAL) = Format((importeIVA(1)), formatoMoneda)
            msGrilla.TextMatrix(msGrilla.Rows - 1, COL_EXENTO) = Format(CCur(Me.txtImporte.text), formatoMoneda)
        Else
            msGrilla.TextMatrix(msGrilla.Rows - 1, COL_NETO_GRAVADO) = Format(CCur(0), formatoMoneda)
            msGrilla.TextMatrix(msGrilla.Rows - 1, COL_IVA) = Format(CCur(importeIVA(0)), formatoMoneda)
            msGrilla.TextMatrix(msGrilla.Rows - 1, COL_TOTAL) = Format(CCur(importeIVA(1)), formatoMoneda)
            msGrilla.TextMatrix(msGrilla.Rows - 1, COL_EXENTO) = Format(CCur(importeIVA(1)), formatoMoneda)
        End If

'        'CAMBIO ESTE PEDAZO DE CODIGO POR EL QUE ESTA COMENTADO ARRIBA
'        msGrilla.TextMatrix(msGrilla.Rows - 1, COL_NETO_GRAVADO) = Format(0, formatoMoneda)
'        msGrilla.TextMatrix(msGrilla.Rows - 1, COL_IVA) = Format(CCur(0), formatoMoneda)
'        msGrilla.TextMatrix(msGrilla.Rows - 1, COL_TOTAL) = Format(CCur(Me.txtImporte.Text), formatoMoneda)
'        msGrilla.TextMatrix(msGrilla.Rows - 1, COL_EXENTO) = Format(CCur(0), formatoMoneda)

    End If
        
    agregarConcepto = True
    Exit Function
errores:
    
        MsgBox Err.Description, Err.Number
        Err.Clear
    
    agregarConcepto = False
End Function


Private Function validarDatos() As Boolean
Dim datosConcepto() As String
    validarDatos = True
    If Len(Me.txtImporte) = 0 Or Me.txtImporte.ClipText = "0" Then
        datosConcepto = separarCodigoDescripcion(Me.cmbConcepto.text)
        If datosConcepto(0) <> COMPROBANTEANULADO Then
            MsgBox "Debe ingresar un importe para agregar el concepto.", vbCritical, "Error: Importe inválido"
            validarDatos = False
        End If
    End If
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



Private Function cargarCabeceraVentas() As tCabecera_Venta
On Error GoTo errores

Dim tCabeceraVentas As tCabecera_Venta
Dim ltEmpresa As tEmpresa
Dim ltTipoComprobante As tTipoComprobante
Dim ltMonedas As tMoneda
Dim ltPuntoDeVenta As tPuntoDeVenta
Dim ltDocumento As tDocumento
Dim ltCodigoOperacion As tCodigoOperacion
Dim strtmp() As String


    tCabeceraVentas.existe = False

    'Cargo la empresa
    ltEmpresa.existe = False
    ltEmpresa.Identificador = IIf(Me.txtNroIdVendedor.ClipText <> "", Me.txtNroIdVendedor.ClipText, 0)
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
    
    'Cargo el Punto de Venta
    ltPuntoDeVenta.existe = False
    ltPuntoDeVenta.puntoVentaId = recuperarIDPuntoDeVenta(ltEmpresa.Identificador, Me.cmbPuntoVenta)
    If ltPuntoDeVenta.puntoVentaId > 0 Then
        ltPuntoDeVenta.existe = True
        ltPuntoDeVenta.empresa_id = ltEmpresa.Identificador
        ltPuntoDeVenta.PuntoDeVenta = Me.cmbPuntoVenta.text
    End If
    'Fin carga del punto de Venta
    
    'Cargo el tipo de documento del Comprador
    ltDocumento.existe = False
    strtmp = separarCodigoDescripcion(Me.cmbTipoDocComprador)
    ltDocumento.codigo = strtmp(0)
    ltDocumento.descripcion = strtmp(1)
    If ltDocumento.codigo <> "" Then ltDocumento.existe = True
    'Fin de carga del documento de Comprador
    
    'Cargo el tipo de operacion de ventas
    ltCodigoOperacion.existe = False
    strtmp = separarCodigoDescripcion(Me.cmbCodigoOperacion)
    ltCodigoOperacion.codigo = strtmp(0)
    ltCodigoOperacion.descripcion = strtmp(1)
    If ltCodigoOperacion.codigo <> "" Then ltCodigoOperacion.existe = True
    'Fin de carga del tipo de operacion de ventas
   
    tCabeceraVentas.Empresa = ltEmpresa
    tCabeceraVentas.tipoComprobabte = ltTipoComprobante
    tCabeceraVentas.moneda = ltMonedas
    tCabeceraVentas.PuntoDeVenta = ltPuntoDeVenta
    tCabeceraVentas.tipoDocumento = ltDocumento
    tCabeceraVentas.codigoOperacionVenta = ltCodigoOperacion
    
    tCabeceraVentas.fechaVenta = Me.dtpFechaComprobante.value
    tCabeceraVentas.compradorId = Me.txtNroIdComprador.ClipText
    tCabeceraVentas.razonSocialComprador = Trim(Me.txtRazonSocialComprador.text)
    tCabeceraVentas.nroComprobanteDesde = Me.txtNroComprobante.ClipText
    tCabeceraVentas.nroComprobanteHasta = Me.txtNroComprobanteHasta.ClipText
    tCabeceraVentas.tipoCambio = Me.txtTipoCambio.text
    
    tCabeceraVentas.existe = tCabeceraVentas.Empresa.existe And _
                            tCabeceraVentas.codigoOperacionVenta.existe And _
                            tCabeceraVentas.moneda.existe And _
                            tCabeceraVentas.PuntoDeVenta.existe And _
                            tCabeceraVentas.tipoComprobabte.existe And _
                            tCabeceraVentas.tipoDocumento.existe And _
                            (Len(tCabeceraVentas.compradorId) > 0) And _
                            (Len(tCabeceraVentas.razonSocialComprador) > 0) And _
                            (Len(tCabeceraVentas.nroComprobanteDesde) > 0) And _
                            (Len(tCabeceraVentas.nroComprobanteHasta) > 0)
                            
    cargarCabeceraVentas = tCabeceraVentas
    Exit Function

errores:
    MsgBox Err.Description, Err.Number
    Err.Clear
    'tCabeceraVentas.existe = False
    cargarCabeceraVentas = tCabeceraVentas
End Function


Private Function cargarDetalleVentas(idVenta As Double, iRow As Integer) As tDetalle_Venta
On Error GoTo errores
Dim tDetalleVentas As tDetalle_Venta
Dim ltAlicuota As tAlicuota
Dim ltConcepto As tConcepto

    tDetalleVentas.existe = False
   
    ltConcepto.codigo = Me.msGrilla.TextMatrix(iRow, COL_CODIGO_CONCEPTO)
    ltConcepto.descripcion = Me.msGrilla.TextMatrix(iRow, COL_CONCEPTO)
    If ltConcepto.codigo <> "" Then ltConcepto.existe = True
    
    
    ltAlicuota.alicuta = Me.msGrilla.TextMatrix(iRow, COL_CODIGO_ALICUOTA)
    ltAlicuota.descripcion = Me.msGrilla.TextMatrix(iRow, COL_ALICUOTA)
    If ltAlicuota.alicuta <> "" Then ltAlicuota.existe = True
    
    
    tDetalleVentas.venta_id = idVenta
    tDetalleVentas.Concepto = ltConcepto
    tDetalleVentas.alicuota = ltAlicuota
    tDetalleVentas.neto_gravado = CCur(Replace(Me.msGrilla.TextMatrix(iRow, COL_NETO_GRAVADO), "$", ""))
    tDetalleVentas.iva = CCur(Replace(Me.msGrilla.TextMatrix(iRow, COL_IVA), "$", ""))
    tDetalleVentas.exento = CCur(Replace(Me.msGrilla.TextMatrix(iRow, COL_EXENTO), "$", ""))
    tDetalleVentas.total = CCur(Replace(Me.msGrilla.TextMatrix(iRow, COL_TOTAL), "$", ""))
    tDetalleVentas.existe = True
    cargarDetalleVentas = tDetalleVentas
    Exit Function

errores:
    MsgBox Err.Description, Err.Number
    Err.Clear
    cargarDetalleVentas = tDetalleVentas

End Function

Private Sub cargarFormularioModificacion()
On Error GoTo errores

    Dim sql_cabecera As String
    Dim sql_moneda As String
    Dim sql_tipoDocComprador As String
    Dim sql_tipoComprobante As String
    Dim sql_codigoOperacionVentas As String
    Dim sql_conceptos As String
    Dim sql_alicuotas As String
    
    Dim oRec As New ADODB.Recordset
    
    sql_cabecera = "SELECT C.*, PV.PUNTO_VENTA FROM CABECERA_VENTA C INNER JOIN PUNTO_VENTA PV ON C.PUNTO_VENTA_ID = PV.PUNTO_VENTA_ID WHERE C.VENTA_ID=" & Me.idVentas
    
   
    With oRec
        .ActiveConnection = oCon
        .CursorType = adOpenStatic
        .CursorLocation = adUseClient
        .Source = sql_cabecera
        .Open
    End With
    
    If oRec.RecordCount = 1 Then
        Me.txtNroIdVendedor.Mask = ""
        Me.txtNroIdVendedor.text = oRec!empresa_id
        Me.txtNroIdVendedor.PromptInclude = False
        Me.txtNroIdVendedor.Mask = "##-########-#"
        Me.txtNroIdVendedor.Enabled = False
        txtNroIdVendedor_LostFocus
        setLisIndexCombo Me.cmbPuntoVenta, oRec!PUNTO_VENTA
        
        Me.dtpFechaComprobante.value = Format(oRec!FECHA_VENTA, "short date")
        Me.txtNroComprobante.text = oRec!NRO_COMPROBANTE_DESDE
        Me.txtNroComprobanteHasta.text = oRec!NRO_COMPROBANTE_HASTA
        'Me.txtNroIdComprador.Mask = ""
        Me.txtNroIdComprador.PromptInclude = False
        Me.txtNroIdComprador.text = oRec!COMPRADOR_ID
        
        Me.txtRazonSocialComprador.text = oRec!RAZON_SOCIAL_COMPRADOR
        
        
        sql_moneda = "SELECT CODIGO, ISNULL(DESCRIPCION,'') AS DESCRIPCION, ISNULL(DEFECTO,0) AS DEFECTO FROM MONEDA WHERE ISNULL(VISIBLE,0) = 1 order by CODIGO "
        cargarCombo Me.cmbMoneda, sql_moneda
        setLisIndexCombo Me.cmbMoneda, oRec!TIPO_MONEDA_ID
        Me.txtTipoCambio.text = oRec!TIPO_DE_CAMBIO
        Me.cmbMoneda.Enabled = False
        
        sql_tipoDocComprador = "SELECT CODIGO, ISNULL(DESCRIPCION,'') AS DESCRIPCION, ISNULL(DEFECTO,0) AS DEFECTO FROM DOCUMENTO WHERE ISNULL(VISIBLE,0) = 1 order by CODIGO "
        cargarCombo Me.cmbTipoDocComprador, sql_tipoDocComprador
        setLisIndexCombo Me.cmbTipoDocComprador, oRec!TIPO_DOCUMENTO_ID
        
        sql_tipoComprobante = "SELECT CODIGO, ISNULL(DESCRIPCION,'') AS DESCRIPCION, ISNULL(DEFECTO,0) AS DEFECTO FROM TIPO_COMPROBANTE WHERE ISNULL(VISIBLE,0) = 1 order by CODIGO "
        cargarCombo Me.cmbTipoComprobante, sql_tipoComprobante
        setLisIndexCombo Me.cmbTipoComprobante, oRec!TIPO_COMPROBANTE_ID
        Me.cmbTipoComprobante.Enabled = False
        
        sql_codigoOperacionVentas = "SELECT CODIGO, ISNULL(DESCRIPCION,'') AS DESCRIPCION, ISNULL(DEFECTO,0) AS DEFECTO FROM CODIGO_OPERACION WHERE ISNULL(VISIBLE,0) = 1 order by CODIGO "
        cargarCombo Me.cmbCodigoOperacion, sql_codigoOperacionVentas
        setLisIndexCombo Me.cmbCodigoOperacion, oRec!CODIGO_OPERACION_ID
        'Me.cmbCodigoOperacion.Enabled = False
        
        sql_conceptos = "SELECT CODIGO, ISNULL(DESCRIPCION,'') AS DESCRIPCION, ISNULL(DEFECTO,0) AS DEFECTO FROM dbo.CONCEPTO WHERE ISNULL(VISIBLE,0) = 1 order by CODIGO "
        cargarCombo Me.cmbConcepto, sql_conceptos

        sql_alicuotas = "SELECT ALICUOTA, ISNULL(DESCRIPCION,'') AS DESCRIPCION, ISNULL(DEFECTO,0) AS DEFECTO FROM ALICUOTA WHERE ISNULL(VISIBLE,0) = 1"
        cargarCombo Me.cmbAlicuota, sql_alicuotas
    
    End If
    
    inicializarGrilla
    cargarDetalle

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

Private Sub cargarFormularioAlta()
    Dim sql_moneda As String
    Dim sql_tipoDocComprador As String
    Dim sql_tipoComprobante As String
    Dim sql_codigoOperacionVentas As String
    Dim sql_conceptos As String
    Dim sql_alicuotas As String
    Dim strtmp() As String
    
    
    dtpFechaComprobante.value = Format(fechaActualServer, "Short date")
   
    
    sql_moneda = "SELECT CODIGO, ISNULL(DESCRIPCION,'') AS DESCRIPCION, ISNULL(DEFECTO,0) AS DEFECTO FROM MONEDA WHERE ISNULL(VISIBLE,0) = 1 order by CODIGO "
    cargarCombo Me.cmbMoneda, sql_moneda
    strtmp = Utils.separarCodigoDescripcion(Me.cmbMoneda)
    If strtmp(0) = tipoMonedaPesos Then
        Me.txtTipoCambio.text = 1
        Me.txtTipoCambio.Enabled = False
    End If
    
    sql_tipoDocComprador = "SELECT CODIGO, ISNULL(DESCRIPCION,'') AS DESCRIPCION, ISNULL(DEFECTO,0) AS DEFECTO FROM DOCUMENTO WHERE ISNULL(VISIBLE,0) = 1 order by CODIGO "
    cargarCombo Me.cmbTipoDocComprador, sql_tipoDocComprador
    
    sql_tipoComprobante = "SELECT CODIGO, ISNULL(DESCRIPCION,'') AS DESCRIPCION, ISNULL(DEFECTO,0) AS DEFECTO FROM TIPO_COMPROBANTE WHERE ISNULL(VISIBLE,0) = 1 order by CODIGO "
    cargarCombo Me.cmbTipoComprobante, sql_tipoComprobante
    
    sql_codigoOperacionVentas = "SELECT CODIGO, ISNULL(DESCRIPCION,'') AS DESCRIPCION, ISNULL(DEFECTO,0) AS DEFECTO FROM CODIGO_OPERACION WHERE ISNULL(VISIBLE,0) = 1 order by CODIGO "
    cargarCombo Me.cmbCodigoOperacion, sql_codigoOperacionVentas
    
    sql_conceptos = "SELECT CODIGO, ISNULL(DESCRIPCION,'') AS DESCRIPCION, ISNULL(DEFECTO,0) AS DEFECTO FROM dbo.CONCEPTO WHERE ISNULL(VISIBLE,0) = 1 order by CODIGO "
    cargarCombo Me.cmbConcepto, sql_conceptos
    
    sql_alicuotas = "SELECT ALICUOTA, ISNULL(DESCRIPCION,'') AS DESCRIPCION, ISNULL(DEFECTO,0) AS DEFECTO FROM ALICUOTA WHERE ISNULL(VISIBLE,0) = 1"
    cargarCombo Me.cmbAlicuota, sql_alicuotas
    
    Me.txtNroIdVendedor.Mask = "##-########-#"
    limpiarFormulario
End Sub

Private Sub limpiarFormulario()
   
    inicializarGrilla
    txtNroIdComprador.Mask = ""
    txtNroIdComprador.text = ""
    txtNroIdComprador.Mask = "##-########-#"
    txtRazonSocialComprador.text = ""
    
    chkLote.value = False

    txtNroComprobante.text = ""
    txtNroComprobanteHasta.text = ""
    txtNroComprobanteHasta.Enabled = False
    Me.frmDatosComprobante.Enabled = True
    
    
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

Private Sub cargarDetalle()


    Dim sql_detalle As String
 
    Dim oRec As New ADODB.Recordset
    
    sql_detalle = "SELECT " + _
                  "C.DESCRIPCION AS CONCEPTO, " + _
                  "A.DESCRIPCION AS ALICUOTA, " + _
                  "DV.NETO_GRAVADO, " + _
                  "DV.EXENTO, " + _
                  "DV.IVA, " + _
                  "DV.TOTAL, " + _
                  "DV.CONCEPTO_ID, " + _
                  "dv.ALICUOTA_ID " + _
                  "FROM DETALLE_VENTA DV INNER JOIN CONCEPTO C ON DV.CONCEPTO_ID = C.CODIGO " + _
                  "INNER JOIN ALICUOTA A ON DV.ALICUOTA_ID = A.ALICUOTA " + _
                  "WHERE DV.CABECERA_VENTA_ID =" + Me.idVentas
    
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

Private Sub txtRazonSocialComprador_GotFocus()
    With txtRazonSocialComprador
        .SelStart = 0
        .SelLength = Len(.text)
    End With
End Sub

Private Sub txtRazonSocialComprador_KeyPress(KeyAscii As Integer)
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
        cmbConcepto.SetFocus
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
