VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmPrincipal 
   BackColor       =   &H8000000C&
   Caption         =   "EM - Ventas y Compras"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   1170
   ClientWidth     =   4560
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   1
      Top             =   10080
      Width           =   20250
      _ExtentX        =   35719
      _ExtentY        =   979
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            TextSave        =   "12/12/2019"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "MAYÚS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            TextSave        =   "NÚM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            TextSave        =   "INS"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4680
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":0452
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":076C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":0BBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":1010
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20250
      _ExtentX        =   35719
      _ExtentY        =   1588
      ButtonWidth     =   2778
      ButtonHeight    =   1429
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ventas"
            Object.ToolTipText     =   "Registraciones de Ventas"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Compras"
            Object.ToolTipText     =   "Registraciones de Compras"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consulta Ventas"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consulta Compras"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Key             =   "Sale del Sistema"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnu_registraciones 
      Caption         =   "&Registraciones"
      Begin VB.Menu mnu_registraciones_ventas 
         Caption         =   "&Ventas"
      End
      Begin VB.Menu mnu_registraciones_compras 
         Caption         =   "&Compras"
      End
   End
   Begin VB.Menu mnu_consultas 
      Caption         =   "&Consultas"
      Begin VB.Menu mnu_consultas_ResgistracionesVtas 
         Caption         =   "&Registraciones de &Ventas"
      End
      Begin VB.Menu mnu_consultas_registracionesCras 
         Caption         =   "&Registraciones de &Compras"
      End
   End
   Begin VB.Menu mnu_amb 
      Caption         =   "&ABM"
      Begin VB.Menu mnu_abm_empresas 
         Caption         =   "&Empresas"
      End
      Begin VB.Menu mnu_abm_conceptos 
         Caption         =   "&Conceptos"
         Begin VB.Menu mnuConceptoComras 
            Caption         =   "Compras"
         End
         Begin VB.Menu mnuConceptoVentas 
            Caption         =   "Ventas"
         End
      End
   End
   Begin VB.Menu mnu_exportacion 
      Caption         =   "&Exportación"
      Begin VB.Menu mnu_exportacion_ventas 
         Caption         =   "V&entas"
      End
      Begin VB.Menu mnu_exportaciones_compras 
         Caption         =   "C&ompras"
      End
   End
   Begin VB.Menu mnuImportacines 
      Caption         =   "&Importaciones"
      Begin VB.Menu mnuImportarASOPROFARMA 
         Caption         =   "ASOPROFARMA"
      End
      Begin VB.Menu mnuVentasAfip 
         Caption         =   "Ventas de AFIP"
      End
   End
   Begin VB.Menu mnu_salir 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("¿Seguro desea salir del Sistema?", vbQuestion + vbYesNo, "Salir") = vbYes Then
        End
    Else
        Cancel = True
    End If
End Sub

Private Sub mnu_abm_empresas_Click()
    frmEmpresaABM.inicializarFormulario
End Sub

Private Sub mnu_consultas_registracionesCras_Click()
    frmConsultaCompras.Show vbModal
End Sub

Private Sub mnu_consultas_ResgistracionesVtas_Click()
    frmConsultaVentas.Show vbModal
End Sub

Private Sub mnu_exportacion_ventas_Click()
    excelExportVtas.Show vbModal
End Sub

Private Sub mnu_informes_ABM_Empresas_Click()

    excelExportVtas.Show vbModal
End Sub

Private Sub mnu_informes_ventas_subdiario_Click()
    excelExportVtas.Show vbModal
End Sub

Private Sub mnu_exportaciones_compras_Click()
    excelExportCpas.Show vbModal
End Sub

Private Sub mnu_registraciones_compras_Click()
    frmCompras.operacion = "A"
    frmCompras.idCompra = 0
    frmCompras.Show vbModal
End Sub

Private Sub mnu_registraciones_ventas_Click()
    frmVentas.operacion = "A"
    frmVentas.idVentas = 0
    frmVentas.Show vbModal
End Sub

Private Sub mnu_salir_Click()
    If MsgBox("¿Seguro desea salir del Sistema?", vbQuestion + vbYesNo, "Salir") = vbYes Then
        End
    End If
End Sub

Private Sub mnuConceptoComras_Click()
    frmConceptoCpasABM.Show vbModal
End Sub

Private Sub mnuConceptoVentas_Click()
    frmConceptoVtasABM.Show vbModal
End Sub

Private Sub mnuImportarASOPROFARMA_Click()
    frmAsoProFarma.Show vbModal
End Sub

Private Sub mnuVentasAfip_Click()
    frmImporVentasAfip.Show vbModal
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button
    Case "Ventas"
        frmVentas.operacion = "A"
        frmVentas.idVentas = 0
        frmVentas.Show vbModal
    Case "Compras"
        frmCompras.operacion = "A"
        frmCompras.idCompra = 0
        frmCompras.Show vbModal
    Case "Consulta Ventas"
        frmConsultaVentas.Show vbModal
    Case "Consulta Compras"
        frmConsultaCompras.Show vbModal
    Case "Salir"
        If MsgBox("¿Seguro desea salir del Sistema?", vbQuestion + vbYesNo, "Salir") = vbYes Then
        End
    End If
    End Select
End Sub
