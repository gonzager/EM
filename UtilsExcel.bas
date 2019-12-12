Attribute VB_Name = "UtilsExcel"
Option Explicit
Public Const COL_FECHACOMPRA = 1
Public Const LET_FECHACOMPRA = "A"
Public Const COL_TIPO_COMPROBANTE = 2
Public Const LET_TIPO_COMPROBABTE = "B"
Public Const COL_NROCOMPROBANTE = 3
Public Const LET_NROCOMPROBANTE = "C"
Public Const COL_TIPO_DOC_VENDEDOR = 4
Public Const LET_TIPO_DOC_VENDEDOR = "D"
Public Const COL_IDVENDEDOR = 5
Public Const LET_IDVENDEDOR = "E"
Public Const COL_RAZONSOCIALVENDEDOR = 6
Public Const LET_RAZONSOCIALVENDEDOR = "F"

Public Const COL_TOTALOPERACION = 7
Public Const LET_TOTALOPERACION = "G"
Public Const COL_NETOGRAVADO21 = 8
Public Const LET_NETOGRAVADO21 = "H"
Public Const COL_NETOGRVADO105 = 9
Public Const LET_NETOGRVADO105 = "I"
Public Const COL_NETOGRAVADO27 = 10
Public Const LET_NETOGRAVADO27 = "J"
Public Const COL_IVA21 = 11
Public Const LET_IVA21 = "K"
Public Const COL_IVA105 = 12
Public Const LET_IVA105 = "L"
Public Const COL_IVA27 = 13
Public Const LET_IVA27 = "M"
Public Const COL_EXENTO = 14
Public Const LET_EXENTO = "N"
Public Const COL_PER_GCIAS = 15
Public Const LET_PER_GCIAS = "O"
Public Const COL_PER_IVA = 16
Public Const LET_PER_IVA = "P"
Public Const COL_PER_IIBB_CABA = 17
Public Const LET_PER_IIBB_CABA = "Q"
Public Const COL_PER_IIBB_BSAS = 18
Public Const LET_PER_IIBB_BSAS = "R"
Public Const COL_PER_IIBB_OTRA = 19
Public Const LET_PER_IIBB_OTRA = "S"
Public Const COL_PER_OTRA_PERC = 20
Public Const LET_PER_OTRA_PERC = "T"

Public Const FIL_COMINEZADATO = 9


Public Function Exportar_ADO_Excel_SubDiarioVentas(fechaDesde As Date, fechaHasta As Date, Identificador As String) As Boolean
      
    On Error GoTo errSub
      
   
    Dim rec         As New ADODB.Recordset
    
    Dim Excel       As Object
    Dim Libro       As Object
    Dim Hoja        As Object
    Dim arrData     As Variant
    Dim iRec        As Long
    Dim iCol        As Integer
    Dim iRow        As Integer
    Dim sOutputPathXLS As String
    Dim Path_Archivo_Ini As String
    'Me.Enabled = False
      
    Screen.MousePointer = vbHourglass
    
'    Path_Archivo_Ini = App.Path & "\EM_config.ini"
'    sOutputPathXLS = Leer_Ini(Path_Archivo_Ini, "archivoExportadoVentas", App.Path & "Exportado.xls")

      
    ' -- Crear los objetos para utilizar el Excel
    Set Excel = CreateObject("Excel.Application")
    Set Libro = Excel.Workbooks.Add
    ' -- Hacer referencia a la hoja
    Set Hoja = Libro.Worksheets(1)
    Excel.Visible = True: Excel.UserControl = True
    

    With rec
        .ActiveConnection = oCon
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
    End With
    
    rec.Source = "SELECT " + _
                    "[Fecha de Venta], " + _
                    "[Tipo de Comprobante], " + _
                    "[Comprobante Desde], " + _
                    "[Comprobante Hasta], " + _
                    "[Tipo Documento], " + _
                    "[ID. Comprador], " + _
                    "[Razon Social Comprador], " + _
                    "[Total], " + _
                    "[Neto Gravado 21], " + _
                    "[Neto Gravado 10.5], " + _
                    "[Neto Gravado 27], " + _
                    "[iva 21.0%], " + _
                    "[iva 10.5%], " + _
                    "[iva 27.0%], " + _
                    "[exento], " + _
                    "[ANULADO] as [Observaciones] " + _
                  "FROM V_SUBDIARIO_VENTAS " + _
                  "WHERE [Fecha de Venta] between '" & Format(fechaDesde, formatoFechaQuery) & "' and '" & Format(fechaHasta, formatoFechaQuery) & "' " + _
                  "AND [ID. Vendedor] = " & Identificador & " " + _
                  "ORDER BY [Fecha de Venta], [Comprobante Desde]"
    rec.Open
    
    
    iCol = rec.Fields.Count
    For iCol = 1 To rec.Fields.Count
        Hoja.cells(8, iCol).value = rec.Fields(iCol - 1).Name
    Next
      
    If Val(Mid(Excel.Version, 1, InStr(1, Excel.Version, ".") - 1)) > 8 Then
        Hoja.cells(9, 1).CopyFromRecordset rec
    Else
  
        arrData = rec.GetRows
  
        iRec = UBound(arrData, 2) + 1
          
        For iCol = 0 To rec.Fields.Count - 1
            For iRow = 0 To iRec - 1
  
                If IsDate(arrData(iCol, iRow)) Then
                    arrData(iCol, iRow) = Format(arrData(iCol, iRow))
  
                ElseIf IsArray(arrData(iCol, iRow)) Then
                    arrData(iCol, iRow) = "Array Field"
                End If
            Next iRow
        Next iCol
              
        ' -- Traspasa los datos a la hoja de Excel
        Hoja.cells(8, 1).Resize(iRec, rec.Fields.Count).value = GetData(arrData)
    End If
  
    Hoja.Range(Hoja.cells(FIL_COMINEZADATO, 8), Hoja.cells(rec.RecordCount + FIL_COMINEZADATO, rec.Fields.Count)).NumberFormat = "#,##0.00"


  
    'Totales por cada columna
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO, 8).Formula = "=SUM(H9:H" & rec.RecordCount + FIL_COMINEZADATO - 1 & ")"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO, 9).Formula = "=SUM(I9:I" & rec.RecordCount + FIL_COMINEZADATO - 1 & ")"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO, 10).Formula = "=SUM(J9:J" & rec.RecordCount + FIL_COMINEZADATO - 1 & ")"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO, 11).Formula = "=SUM(K9:K" & rec.RecordCount + FIL_COMINEZADATO - 1 & ")"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO, 12).Formula = "=SUM(L9:L" & rec.RecordCount + FIL_COMINEZADATO - 1 & ")"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO, 13).Formula = "=SUM(M9:M" & rec.RecordCount + FIL_COMINEZADATO - 1 & ")"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO, 14).Formula = "=SUM(N9:N" & rec.RecordCount + FIL_COMINEZADATO - 1 & ")"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO, 15).Formula = "=SUM(O9:O" & rec.RecordCount + FIL_COMINEZADATO - 1 & ")"
  
    'Totales de abajo
    Hoja.cells(rec.RecordCount + 11, 1) = "Total"
    Hoja.cells(rec.RecordCount + 11, 2).Formula = "=SUM(H9:H" & rec.RecordCount + FIL_COMINEZADATO - 1 & ")"
    
    Hoja.cells(rec.RecordCount + 12, 1) = "Total Exentos"
    Hoja.cells(rec.RecordCount + 12, 2).Formula = "=SUM(O9:O" & rec.RecordCount + FIL_COMINEZADATO - 1 & ")"
    
    Hoja.cells(rec.RecordCount + 13, 1) = "Total Neto Gravado"
    Hoja.cells(rec.RecordCount + 13, 2).Formula = "=SUM(I" & rec.RecordCount + FIL_COMINEZADATO & ":K" & rec.RecordCount + FIL_COMINEZADATO & ")"
    
    Hoja.cells(rec.RecordCount + 14, 1) = "Total IVA 21"
    Hoja.cells(rec.RecordCount + 14, 2).Formula = "=SUM(L9:L" & rec.RecordCount + FIL_COMINEZADATO - 1 & ")"
    Hoja.cells(rec.RecordCount + 15, 1) = "Total IVA 10.5"
    Hoja.cells(rec.RecordCount + 15, 2).Formula = "=SUM(M9:M" & rec.RecordCount + FIL_COMINEZADATO - 1 & ")"
    Hoja.cells(rec.RecordCount + 16, 1) = "Total IVA 27"
    Hoja.cells(rec.RecordCount + 16, 2).Formula = "=SUM(N9:N" & rec.RecordCount + FIL_COMINEZADATO - 1 & ")"
  

    
    Hoja.Range(Hoja.cells(rec.RecordCount + 11, 2), Hoja.cells(rec.RecordCount + 16, 2)).NumberFormat = "#,##0.00"
  
    Hoja.Range(Hoja.cells(rec.RecordCount + 11, 1), Hoja.cells(rec.RecordCount + 16, 2)).Select
    With Excel.Selection.Font
        .Name = "Arial"
        .FontStyle = "Negrita"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
    End With
    With Excel.Selection.Interior
        .ColorIndex = 15
    End With
    
    '8 son los renglones de arriba y 8 los renglones de abajo
    Hoja.Range(Hoja.cells(1, 1), Hoja.cells(rec.RecordCount + 8 + 8, rec.Fields.Count)).Columns.AutoFit
    
    Dim RowDesdeTipoComprobante As Integer
    RowDesdeTipoComprobante = IIf(rec.RecordCount > 0, rec.RecordCount + 8 + 8, 16)
   
    If rec.State = adStateOpen Then rec.Close
    
    'AGRUPACION POR TIPO DE COMPROBANTE
    With rec
        .ActiveConnection = oCon
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .Source = "SELECT [Tipo de Comprobante], " + _
                         "Sum([Neto Gravado 21] + CASE WHEN TC.CALCULO='I' THEN [IVA 21.0%] ELSE 0 END)  as [Neto Gravado 21], " + _
                         "Sum ([Neto Gravado 10.5] + CASE WHEN TC.CALCULO='I' THEN [IVA 10.5%] ELSE 0 END) [Neto Gravado 10.5], " + _
                         "Sum([Neto Gravado 27] + CASE WHEN TC.CALCULO='I' THEN [IVA 27.0%] ELSE 0 END) as [Neto Gravado 27] , " + _
                         "Sum([Neto Gravado 21] + [Neto Gravado 10.5] + [Neto Gravado 27] + CASE WHEN TC.CALCULO='I' THEN [IVA 21.0%] + [IVA 10.5%] + [IVA 10.5%] ELSE 0 END ) as [Total Neto Gravado] " + _
                  "FROM V_SUBDIARIO_VENTAS inner join TIPO_COMPROBANTE TC on " + _
                  "V_SUBDIARIO_VENTAS.TIPO_COMPROBANTE_ID = TC.CODIGO " + _
                  "WHERE [Fecha de Venta] between '" & Format(fechaDesde, formatoFechaQuery) & "' and '" & Format(fechaHasta, formatoFechaQuery) & "' " + _
                  "AND [ID. Vendedor] = " & Identificador & " " + _
                  "GROUP BY [Tipo de Comprobante] " + _
                  "ORDER BY [Tipo de Comprobante]"
    End With
    rec.Open
    Dim iRowTipoComprobante As Integer
    iRowTipoComprobante = 0
    Do While Not rec.EOF
        Hoja.cells(RowDesdeTipoComprobante + 2 + iRowTipoComprobante, 1) = "Tipo Comprobante"
        Hoja.cells(RowDesdeTipoComprobante + 2 + iRowTipoComprobante, 2) = rec![Tipo de Comprobante]
        Hoja.cells(RowDesdeTipoComprobante + 3 + iRowTipoComprobante, 1) = "Neto Gravado 21"
        Hoja.cells(RowDesdeTipoComprobante + 3 + iRowTipoComprobante, 2) = rec![Neto Gravado 21]
        Hoja.cells(RowDesdeTipoComprobante + 4 + iRowTipoComprobante, 1) = "Neto Gravado 10.5"
        Hoja.cells(RowDesdeTipoComprobante + 4 + iRowTipoComprobante, 2) = rec![Neto Gravado 10.5]
        Hoja.cells(RowDesdeTipoComprobante + 5 + iRowTipoComprobante, 1) = "Neto Gravado 27"
        Hoja.cells(RowDesdeTipoComprobante + 5 + iRowTipoComprobante, 2) = rec![Neto Gravado 27]
        Hoja.cells(RowDesdeTipoComprobante + 6 + iRowTipoComprobante, 1) = "Total Neto Gravado"
        Hoja.cells(RowDesdeTipoComprobante + 6 + iRowTipoComprobante, 2) = rec![Total Neto Gravado]
        iRowTipoComprobante = iRowTipoComprobante + 6
        rec.MoveNext
    Loop
    
    Hoja.Range(Hoja.cells(RowDesdeTipoComprobante + 2, 2), Hoja.cells(RowDesdeTipoComprobante + iRowTipoComprobante, 2)).NumberFormat = "#,##0.00"
  
    Hoja.Range(Hoja.cells(RowDesdeTipoComprobante + 2, 1), Hoja.cells(RowDesdeTipoComprobante + iRowTipoComprobante, 2)).Select
    With Excel.Selection.Font
        .Name = "Arial"
        .FontStyle = "Negrita"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
    End With
    With Excel.Selection.Interior
        .ColorIndex = 15
    End With
    
    If rec.State = adStateOpen Then rec.Close
    
    
    

   With rec
        .ActiveConnection = oCon
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .Source = "select IDENTIFICADOR, RAZONSOCIAL, DOMICILIO from EMPRESA WHERE IDENTIFICADOR =" & Identificador
    End With
    
    
    rec.Open
    iRow = 1
    If rec.RecordCount = 1 Then
        Hoja.cells(iRow, 1).value = "CUIT/CUIL"
        Hoja.cells(iRow, 2).value = rec!Identificador
        Hoja.cells(iRow + 1, 1).value = "RAZON SOCIAL"
        Hoja.cells(iRow + 1, 2).value = rec!razonSocial
        Hoja.cells(iRow + 2, 1).value = "DOMICILIO"
        Hoja.cells(iRow + 2, 2).value = rec!domicilio
    End If
    
    Hoja.cells(iRow + 4, 1).value = "Sub Diario de Ventas"
    Hoja.cells(iRow + 4, 3).value = "Fecha Desde:"
    Hoja.cells(iRow + 4, 4).value = "'" & Format(fechaDesde, "dd") & "/" & Format(fechaDesde, "mm") & "/" & Format(fechaDesde, "yyyy")
    Hoja.cells(iRow + 4, 5).value = "Fecha Hasta:"
    Hoja.cells(iRow + 4, 6).value = "'" & Format(fechaHasta, "dd") & "/" & Format(fechaHasta, "mm") & "/" & Format(fechaHasta, "yyyy")
    
    If rec.State = adStateOpen Then rec.Close
    

    Set rec = Nothing
    
    ' -- guardar el libro
    'Libro.saveAs sOutputPathXLS
    'Libro.Close
    ' -- Elimina las referencias Xls
    Set Hoja = Nothing
    Set Libro = Nothing
    'Excel.quit
    Set Excel = Nothing
      
    Exportar_ADO_Excel_SubDiarioVentas = True
    Screen.MousePointer = vbDefault
    Exit Function
errSub:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical, "Error"
    Exportar_ADO_Excel_SubDiarioVentas = False
    If rec.State = adStateOpen Then rec.Close
    Set rec = Nothing
    
End Function
  
Private Function GetData(vValue As Variant) As Variant
    Dim X As Long, Y As Long, xMax As Long, yMax As Long, T As Variant
      
    xMax = UBound(vValue, 2): yMax = UBound(vValue, 1)
      
    ReDim T(xMax, yMax)
    For X = 0 To xMax
        For Y = 0 To yMax
            T(X, Y) = vValue(Y, X)
        Next Y
    Next X
      
    GetData = T
End Function



Public Function Exportar_ADO_Excel_SubDiarioCompras(fechaDesde As Date, fechaHasta As Date, Identificador As String) As Boolean
      
    On Error GoTo errSub
      
   
    Dim rec         As New ADODB.Recordset
    Dim Excel       As Object
    Dim Libro       As Object
    Dim Hoja        As Object
    Dim arrData     As Variant
    Dim iRec        As Long
    Dim iCol        As Integer
    Dim iRow        As Integer
    Dim sOutputPathXLS As String
    Dim Path_Archivo_Ini As String
    Dim HastaRow    As Integer
      
    Screen.MousePointer = vbHourglass

    ' -- Crear los objetos para utilizar el Excel
    Set Excel = CreateObject("Excel.Application")
    Set Libro = Excel.Workbooks.Add
    ' -- Hacer referencia a la hoja
    Set Hoja = Libro.Worksheets(1)
    Excel.Visible = True: Excel.UserControl = True
    
    With rec
        .ActiveConnection = oCon
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        
    End With
    
    
    rec.Source = "SELECT " + _
                    "[F. Comprobante] [Fecha de Comprobante      ], " + _
                    "[T. Comprobante], " + _
                    "[Comprobante], " + _
                    "[T. Documento], " + _
                    "[ID. Vendedor], " + _
                    "[Razon Social Vendedor], " + _
                    "[Total] as [Total Cbte], " + _
                    "[Neto Gravado 21.00%] as [N Gravado 21%], " + _
                    "[Neto Gravado 10.5%] as [N Gravado 10.5%], " + _
                    "[Neto Gravado 27.0%] as [N Gravado 27%], " + _
                    "[IVA 21.0%], " + _
                    "[IVA 10.5%], " + _
                    "[IVA 27.0%], " + _
                    "[Exento]," + _
                    "[Perc. Ganancias], " + _
                    "[Perc. I.V.A], " + _
                    "[Perc. IIBB CABA], " + _
                    "[Perc IIBB Pcia Bs As], " + _
                    "[IIBB Otra J], " + _
                    "[Otras Perc], [Concepto] " + _
                  "FROM V_SUBDIARIO_COMPRAS " + _
                  "WHERE [F. Imputacion] between '" & Format(fechaDesde, formatoFechaQuery) & "' and '" & Format(fechaHasta, formatoFechaQuery) & "' " + _
                  "AND [ID. Comprador] = " & Identificador & " " + _
                  "ORDER BY  [F. Imputacion], [Comprobante]"
    rec.Open
    
    
    iCol = rec.Fields.Count
    For iCol = 1 To rec.Fields.Count
        Hoja.cells(8, iCol).value = rec.Fields(iCol - 1).Name
    Next
      
    If Val(Mid(Excel.Version, 1, InStr(1, Excel.Version, ".") - 1)) > 8 Then
        Hoja.cells(9, 1).CopyFromRecordset rec
    Else
  
        arrData = rec.GetRows
  
        iRec = UBound(arrData, 2) + 1
          
        For iCol = 0 To rec.Fields.Count - 1
            For iRow = 0 To iRec - 1
  
                If IsDate(arrData(iCol, iRow)) Then
                    arrData(iCol, iRow) = Format(arrData(iCol, iRow))
  
                ElseIf IsArray(arrData(iCol, iRow)) Then
                    arrData(iCol, iRow) = "Array Field"
                End If
            Next iRow
        Next iCol
              
        ' -- Traspasa los datos a la hoja de Excel
        Hoja.cells(8, 1).Resize(iRec, rec.Fields.Count).value = GetData(arrData)
    End If
  
    Hoja.Range(Hoja.cells(FIL_COMINEZADATO, COL_TOTALOPERACION), Hoja.cells(rec.RecordCount + FIL_COMINEZADATO, rec.Fields.Count)).NumberFormat = "#,##0.00"

    Hoja.Range(Hoja.cells(1, 1), Hoja.cells(rec.RecordCount + FIL_COMINEZADATO - 1, rec.Fields.Count)).Columns.AutoFit


    'Totales por cada columna
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO, COL_TOTALOPERACION).Formula = "=SUM(" & LET_TOTALOPERACION & FIL_COMINEZADATO & ":" & LET_TOTALOPERACION & rec.RecordCount + FIL_COMINEZADATO - 1 & ")"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO, COL_NETOGRAVADO21).Formula = "=SUM(" & LET_NETOGRAVADO21 & FIL_COMINEZADATO & ":" & LET_NETOGRAVADO21 & rec.RecordCount + FIL_COMINEZADATO - 1 & ")"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO, COL_NETOGRVADO105).Formula = "=SUM(" & LET_NETOGRVADO105 & FIL_COMINEZADATO & ":" & LET_NETOGRVADO105 & rec.RecordCount + FIL_COMINEZADATO - 1 & ")"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO, COL_NETOGRAVADO27).Formula = "=SUM(" & LET_NETOGRAVADO27 & FIL_COMINEZADATO & ":" & LET_NETOGRAVADO27 & rec.RecordCount + FIL_COMINEZADATO - 1 & ")"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO, COL_IVA21).Formula = "=SUM(" & LET_IVA21 & FIL_COMINEZADATO & ":" & LET_IVA21 & rec.RecordCount + FIL_COMINEZADATO - 1 & ")"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO, COL_IVA105).Formula = "=SUM(" & LET_IVA105 & FIL_COMINEZADATO & ":" & LET_IVA105 & rec.RecordCount + FIL_COMINEZADATO - 1 & ")"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO, COL_IVA27).Formula = "=SUM(" & LET_IVA27 & FIL_COMINEZADATO & ":" & LET_IVA27 & rec.RecordCount + FIL_COMINEZADATO - 1 & ")"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO, COL_EXENTO).Formula = "=SUM(" & LET_EXENTO & FIL_COMINEZADATO & ":" & LET_EXENTO & rec.RecordCount + FIL_COMINEZADATO - 1 & ")"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO, COL_PER_GCIAS).Formula = "=SUM(" & LET_PER_GCIAS & FIL_COMINEZADATO & ":" & LET_PER_GCIAS & rec.RecordCount + FIL_COMINEZADATO - 1 & ")"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO, COL_PER_IVA).Formula = "=SUM(" & LET_PER_IVA & FIL_COMINEZADATO & ":" & LET_PER_IVA & rec.RecordCount + FIL_COMINEZADATO - 1 & ")"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO, COL_PER_IIBB_CABA).Formula = "=SUM(" & LET_PER_IIBB_CABA & FIL_COMINEZADATO & ":" & LET_PER_IIBB_CABA & rec.RecordCount + FIL_COMINEZADATO - 1 & ")"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO, COL_PER_IIBB_BSAS).Formula = "=SUM(" & LET_PER_IIBB_BSAS & FIL_COMINEZADATO & ":" & LET_PER_IIBB_BSAS & rec.RecordCount + FIL_COMINEZADATO - 1 & ")"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO, COL_PER_IIBB_OTRA).Formula = "=SUM(" & LET_PER_IIBB_OTRA & FIL_COMINEZADATO & ":" & LET_PER_IIBB_OTRA & rec.RecordCount + FIL_COMINEZADATO - 1 & ")"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO, COL_PER_OTRA_PERC).Formula = "=SUM(" & LET_PER_OTRA_PERC & FIL_COMINEZADATO & ":" & LET_PER_OTRA_PERC & rec.RecordCount + FIL_COMINEZADATO - 1 & ")"
    

    'Totales de abajo
    'Hoja.cells(rec.RecordCount + FIL_COMINEZADATO + 2, 1) = "Total"
    'Hoja.cells(rec.RecordCount + FIL_COMINEZADATO + 2, 2).Formula = "=SUM(" & LET_TOTALOPERACION & FIL_COMINEZADATO & ":" & LET_TOTALOPERACION & rec.RecordCount + FIL_COMINEZADATO - 1 & ")"

    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO + 2, 1) = "Total IVA 10.5"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO + 2, 2).Formula = "=SUM(" & LET_IVA105 & FIL_COMINEZADATO & ":" & LET_IVA105 & rec.RecordCount + FIL_COMINEZADATO - 1 & ")"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO + 3, 1) = "Total IVA 21"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO + 3, 2).Formula = "=SUM(" & LET_IVA21 & FIL_COMINEZADATO & ":" & LET_IVA21 & rec.RecordCount + FIL_COMINEZADATO - 1 & ")"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO + 4, 1) = "Total IVA 27"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO + 4, 2).Formula = "=SUM(" & LET_IVA27 & FIL_COMINEZADATO & ":" & LET_IVA27 & rec.RecordCount + FIL_COMINEZADATO - 1 & ")"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO + 5, 1) = "Total Crédito Fiscal"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO + 5, 2).Formula = "=SUM(B" & rec.RecordCount + FIL_COMINEZADATO + 2 & ":" & "B" & rec.RecordCount + FIL_COMINEZADATO + 4 & ")"
    
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO + 7, 1) = "Percepciones IVA"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO + 7, 2).Formula = "=SUM(" & LET_PER_IVA & FIL_COMINEZADATO & ":" & LET_PER_IVA & rec.RecordCount + FIL_COMINEZADATO - 1 & ")"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO + 8, 1) = "Percepciones Gcias"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO + 8, 2).Formula = "=SUM(" & LET_PER_GCIAS & FIL_COMINEZADATO & ":" & LET_PER_GCIAS & rec.RecordCount + FIL_COMINEZADATO - 1 & ")"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO + 10, 1) = "Percep. IIBB CABA"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO + 10, 2).Formula = "=SUM(" & LET_PER_IIBB_CABA & FIL_COMINEZADATO & ":" & LET_PER_IIBB_CABA & rec.RecordCount + FIL_COMINEZADATO - 1 & ")"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO + 11, 1) = "Percep. IIBB Bs.As"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO + 11, 2).Formula = "=SUM(" & LET_PER_IIBB_BSAS & FIL_COMINEZADATO & ":" & LET_PER_IIBB_BSAS & rec.RecordCount + FIL_COMINEZADATO - 1 & ")"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO + 12, 1) = "Percep. IIBB Otras"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO + 12, 2).Formula = "=SUM(" & LET_PER_IIBB_OTRA & FIL_COMINEZADATO & ":" & LET_PER_IIBB_OTRA & rec.RecordCount + FIL_COMINEZADATO - 1 & ")"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO + 13, 1) = "Total Percep. IIBB"
    Hoja.cells(rec.RecordCount + FIL_COMINEZADATO + 13, 2).Formula = "=SUM(B" & rec.RecordCount + FIL_COMINEZADATO + 10 & ":" & "B" & rec.RecordCount + FIL_COMINEZADATO + 12 & ")"
        

    Hoja.Range(Hoja.cells(rec.RecordCount + FIL_COMINEZADATO + 5, 2), Hoja.cells(rec.RecordCount + FIL_COMINEZADATO + 5, 2)).Select
    'xlEdgeTop = 8
    'xlContinuous = 1
    'xlThin = 2
    With Excel.Selection.Borders(8)
        .LineStyle = 1
        .ColorIndex = 0
        '.TintAndShade = 0
        .Weight = 2
    End With

    Hoja.Range(Hoja.cells(rec.RecordCount + FIL_COMINEZADATO + 13, 2), Hoja.cells(rec.RecordCount + FIL_COMINEZADATO + 13, 2)).Select
    'xlEdgeTop = 8
    'xlContinuous = 1
    'xlThin = 2
    With Excel.Selection.Borders(8)
        .LineStyle = 1
        .ColorIndex = 0
        '.TintAndShade = 0
        .Weight = 2
    End With


    Hoja.Range(Hoja.cells(rec.RecordCount + FIL_COMINEZADATO + 2, 2), Hoja.cells(rec.RecordCount + FIL_COMINEZADATO + 13, 2)).NumberFormat = "#,##0.00"
    
    
    Hoja.Range(Hoja.cells(rec.RecordCount + FIL_COMINEZADATO + 2, 1), Hoja.cells(rec.RecordCount + FIL_COMINEZADATO + 13, 2)).Select
    With Excel.Selection.Font
        .Name = "Arial"
        .FontStyle = "Negrita"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
    End With
    With Excel.Selection.Interior
        .ColorIndex = 15
    End With
    
    'ASIENTO
    HastaRow = rec.RecordCount + FIL_COMINEZADATO + 13 + 2
    If rec.State = adStateOpen Then rec.Close
    
    With rec
        .ActiveConnection = oCon
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .Source = " select CONCEPTO , " + _
                  " SUM (DEBE) AS DEBE, " + _
                  " SUM(HABER) As HABER " + _
                  " From dbo.V_ASIENTOS_COMPRAS " + _
                  " Where " + _
                  " IDENTIFICADOR =" & Identificador + _
                  " AND FECHA_IMPUTACION between '" & Format(fechaDesde, formatoFechaQuery) & "' and '" & Format(fechaHasta, formatoFechaQuery) & "' " + _
                  " GROUP BY CONCEPTO, ORDEN " + _
                  " ORDER BY ORDEN, CONCEPTO "
    End With
    
    rec.Open
    
    iCol = rec.Fields.Count
    For iCol = 1 To rec.Fields.Count
        Hoja.cells(HastaRow, iCol).value = rec.Fields(iCol - 1).Name
    Next
      
    If Val(Mid(Excel.Version, 1, InStr(1, Excel.Version, ".") - 1)) > 8 Then
        Hoja.cells(HastaRow + 1, 1).CopyFromRecordset rec
    End If
    
    Hoja.cells(rec.RecordCount + HastaRow + 1, 2).Formula = "=SUM(B" & HastaRow + 1 & ":B" & rec.RecordCount + HastaRow & ")"
    Hoja.cells(rec.RecordCount + HastaRow + 1, 3).Formula = "=SUM(C" & HastaRow + 1 & ":C" & rec.RecordCount + HastaRow & ")"
    
    
    'ASIENTO
    
    
    
    
    ' -- Cierra el recordset y la base de datos y los objetos ADO
    If rec.State = adStateOpen Then rec.Close
    
    With rec
        .ActiveConnection = oCon
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .Source = "select IDENTIFICADOR, RAZONSOCIAL, DOMICILIO from EMPRESA WHERE IDENTIFICADOR =" & Identificador
    End With
    
    rec.Open
    iRow = 1
    If rec.RecordCount = 1 Then
        Hoja.cells(iRow, 1).value = "CUIT/CUIL"
        Hoja.cells(iRow, 2).value = rec!Identificador
        Hoja.cells(iRow + 1, 1).value = "RAZON SOCIAL"
        Hoja.cells(iRow + 1, 2).value = rec!razonSocial
        Hoja.cells(iRow + 2, 1).value = "DOMICILIO"
        Hoja.cells(iRow + 2, 2).value = rec!domicilio
    End If
    
    Hoja.cells(iRow + 4, 1).value = "Sub Diario de Compras"
    Hoja.cells(iRow + 4, 3).value = "Fecha Desde:"
    Hoja.cells(iRow + 4, 4).value = "'" & Format(fechaDesde, "dd") & "/" & Format(fechaDesde, "mm") & "/" & Format(fechaDesde, "yyyy")
    Hoja.cells(iRow + 4, 5).value = "Fecha Hasta:"
    Hoja.cells(iRow + 4, 6).value = "'" & Format(fechaHasta, "dd") & "/" & Format(fechaHasta, "mm") & "/" & Format(fechaHasta, "yyyy")
    
    If rec.State = adStateOpen Then rec.Close
    

    Set rec = Nothing
    
    ' -- guardar el libro
    'Libro.saveAs sOutputPathXLS
    'Libro.Close
    ' -- Elimina las referencias Xls
    Set Hoja = Nothing
    Set Libro = Nothing
    'Excel.quit
    Set Excel = Nothing
      
    Exportar_ADO_Excel_SubDiarioCompras = True
    Screen.MousePointer = vbDefault
    Exit Function
errSub:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical, "Error"
    Exportar_ADO_Excel_SubDiarioCompras = False
    If rec.State = adStateOpen Then rec.Close
    Set rec = Nothing
    
End Function




Public Function Exportar_ADO_Excel_RecordSet(ls_sql As String) As Boolean
      
    On Error GoTo errSub
      
   
    Dim rec         As New ADODB.Recordset
    Dim Excel       As Object
    Dim Libro       As Object
    Dim Hoja        As Object
    Dim arrData     As Variant
    Dim iRec        As Long
    Dim iCol        As Integer
    Dim iRow        As Integer
    Dim sOutputPathXLS As String
    Dim Path_Archivo_Ini As String
      
    Screen.MousePointer = vbHourglass

    ' -- Crear los objetos para utilizar el Excel
    Set Excel = CreateObject("Excel.Application")
    Set Libro = Excel.Workbooks.Add
    ' -- Hacer referencia a la hoja
    Set Hoja = Libro.Worksheets(1)
    Excel.Visible = True: Excel.UserControl = True
    
    With rec
        .ActiveConnection = oCon
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .Source = ls_sql
    End With
    
    rec.Open
    
    
    iCol = rec.Fields.Count
    For iCol = 1 To rec.Fields.Count
        Hoja.cells(1, iCol).value = rec.Fields(iCol - 1).Name
    Next
      
    If Val(Mid(Excel.Version, 1, InStr(1, Excel.Version, ".") - 1)) > 8 Then
        Hoja.cells(2, 1).CopyFromRecordset rec
    Else
  
        arrData = rec.GetRows
  
        iRec = UBound(arrData, 2) + 1
          
        For iCol = 0 To rec.Fields.Count - 1
            For iRow = 0 To iRec - 1
  
                If IsDate(arrData(iCol, iRow)) Then
                    arrData(iCol, iRow) = Format(arrData(iCol, iRow))
  
                ElseIf IsArray(arrData(iCol, iRow)) Then
                    arrData(iCol, iRow) = "Array Field"
                End If
            Next iRow
        Next iCol
              
        ' -- Traspasa los datos a la hoja de Excel
        Hoja.cells(1, 1).Resize(iRec, rec.Fields.Count).value = GetData(arrData)
    End If
  
  

    
    'Hoja.Range(Hoja.cells(1, rec.Fields.Count), Hoja.cells(rec.RecordCount, rec.Fields.Count)).NumberFormat = "#,##0.00"

    Hoja.Range(Hoja.cells(1, 1), Hoja.cells(rec.RecordCount, rec.Fields.Count)).Columns.AutoFit

    ' -- Cierra el recordset y la base de datos y los objetos ADO
    If rec.State = adStateOpen Then rec.Close
    Set rec = Nothing

    Set Hoja = Nothing
    Set Libro = Nothing
    'Excel.quit
    Set Excel = Nothing
      
    Exportar_ADO_Excel_RecordSet = True
    Screen.MousePointer = vbDefault
    Exit Function
errSub:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical, "Error"
    Exportar_ADO_Excel_RecordSet = False
    If rec.State = adStateOpen Then rec.Close
    Set rec = Nothing
    
End Function
