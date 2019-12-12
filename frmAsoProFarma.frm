VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAsoProFarma 
   Caption         =   "Importar Archivo de ASOPROFARMA"
   ClientHeight    =   2610
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   7650
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5850
      TabIndex        =   5
      Top             =   2130
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frmDatos 
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7605
      Begin VB.CommandButton cmpImportar 
         Caption         =   "Importar Datos"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   1470
         Width           =   2415
      End
      Begin VB.CommandButton cmdArchivo 
         Caption         =   "Copiar y Leer Archivo"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   510
         Width           =   2415
      End
      Begin VB.Label lblCantidad 
         Caption         =   "Cantidad de Comprobantes a  Procesar : 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1110
         Width           =   4935
      End
      Begin VB.Label lblServerFile 
         Caption         =   "lblServerFile"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   150
         TabIndex        =   2
         Top             =   180
         Width           =   7125
      End
   End
End
Attribute VB_Name = "frmAsoProFarma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdArchivo_Click()
Dim Linea As String
Dim i As Integer
Dim errTam As Boolean
Dim fileName As String
Dim serverFile As String
Dim serverPath As String
Dim Path_Archivo_Ini As String
Dim formatBulkXML  As String

On Error GoTo tratarError
    
    errTram = False
    'Titulo del CommonDialog
    CommonDialog1.DialogTitle = "Seleccione el archivo a Importar"
    
    'Extension del CommonDialog. Archivos txt
    CommonDialog1.Filter = "Archivos tipo csv|*.csv"

    'Abrimos el CommonDialog
    CommonDialog1.ShowOpen

    If CommonDialog1.fileName = "" Then
        'salimos de la rutina ya que no se ha seleccionado ningún archivo
        Exit Sub
    Else
        i = 0
        fileName = CommonDialog1.fileName
        'Abrimos el archivo para leerlo, pasándole la ruta con la propiedad FileName del Commondialog
        Open fileName For Input As #1
  
        While Not EOF(1) And Not errTam
            'Leemos la línea
            Line Input #1, Linea
            If Len(Linea) <> 246 Then errTam = True
            i = i + 1
            
        Wend
  
        'Cerramos el archivo abierto anteriormente
        Close #1
        
        If errTam Then
            MsgBox "Error de formato en la linea " & i & " del Archivo: " & vbcrcl & fileName, vbCritical, "Error en el Archivo"
            Exit Sub
        End If
        
        Path_Archivo_Ini = App.Path & "\EM_config.ini"
        serverPath = Leer_Ini(Path_Archivo_Ini, "SERVERFILESHARE", SERVERFILESHARE) + "\"
        serverFile = serverPath & Obtener_Nombre_Archivo(fileName)
        formatBulkXML = serverPath & Leer_Ini(Path_Archivo_Ini, "ASOPROFARMA_XML_FILE_FORMAT", "asoprofarma.xml")
        Copiar_Archivo fileName, serverFile
        
        If fBuklAsoProFarma(serverFile, formatBulkXML) = 0 Then
            If (i = cantidadTmpAsoprofarma) Then
                lblCantidad = "Cantidad de Comprobantes a  Procesar : " & i
                cmpImportar.Enabled = True
            End If
        End If
    
    End If
    
    Exit Sub
tratarError:
    


End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmpImportar_Click()

    If cantidadTmpAsoprofarma > 0 Then
        Screen.MousePointer = vbHourglass
        If fImportarAsoProForma() = 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "Importación Realizada de Forma Exitosa.", vbInformation, "Mensaje al Usuario"
            cmpImportar.Enabled = False
            lblCantidad = "Cantidad de Comprobantes a  Procesar : 0"
        End If
    Else
        MsgBox "No Hay Registros para importar", vbInformation, "Mensaje al Usuario"
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()

Dim serverFile As String
Dim Path_Archivo_Ini As String
    
    Path_Archivo_Ini = App.Path & "\EM_config.ini"
    serverFile = Leer_Ini(Path_Archivo_Ini, "SERVERFILESHARE", SERVERFILESHARE)
    
    lblServerFile.Caption = serverFile
    
End Sub

Private Sub Label1_Click()

End Sub

