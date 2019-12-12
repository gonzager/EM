Attribute VB_Name = "Constantes"
Option Explicit
Public Const TIPOCOMROBANTEB As String = "006"
Public Const COMPROBANTEANULADO As String = "999"
Public Const formatoNroComprobante As String = "0000000#"
Public Const formatoMoneda As String = "$ #,##0.00"
Public Const alicuotaExenta As String = "0003"
Public Const tipoCalculoDirecto As String = "D"
Public Const tipoCalculoInverso As String = "I"
Public Const tipoCalculoExento As String = "E"
Public Const APPLICATION As String = "EM"
Public Const tipoMonedaPesos As String = "PES"
Public Const formatoFechaQuery As String = "YYYY-MM-DD"
Public Const CODIGOCUIT As String = "80"
Public Const CODIGOCUIL As String = "86"
Public Const CODIGOIIBB As String = "01"
Public Const CODIGOIVA As String = "02"
Public Const CODIGOGCIAS As String = "03"



    
Public separadorDecimal As String
Public oCon As ADODB.Connection
Dim Path_Archivo_Ini As String




Public Type tTipoComprobante
    CODIGO As String
    descripcion As String
    existe As Boolean
End Type

Public Type tMoneda
    CODIGO As String
    descripcion As String
    existe As Boolean
End Type


Public Type tDocumento
    CODIGO As String
    descripcion As String
    existe As Boolean
End Type

Public Type tCodigoOperacion
    CODIGO As String
    descripcion As String
    existe As Boolean
End Type


Public Type tConcepto
    CODIGO As String
    descripcion As String
    existe As Boolean
End Type

Public Type tAlicuota
    alicuta As String
    descripcion As String
    existe As Boolean
End Type


Public Type tTipoPercepcion
    percipcion As String
    descripcion As String
    existe As Boolean
End Type

Public Type tJurisdiccion
    jurisdiccion As String
    descripcion As String
    exite As Boolean
End Type



Public Type tExportVentas

'N° | Nombre del Campo                                                  | Desde | Hasta | Tipo de dato | Tam | Observaciones
'01 | Fecha de comprobante                                              | 001   | 008 | Numérico     | 8  | Fto: AAAAMMDD
'02 | Tipo de comprobante                                               | 009   | 011 | Numérico     | 3  | Según tabla Comprobantes
'03 | Punto de venta                                                    | 012   | 016 | Numérico     | 5  |
'04 | Número de comprobante                                             | 017   | 036 | Numérico     | 20 |
'05 | Número de comprobante hasta                                       | 037   | 056 | Numérico     | 20 |
'06 | Código de documento del comprador                                 | 057   | 058 | Numérico     | 2  | Según tabla Documentos
'07 | Número de identificación del comprador                            | 059   | 078 | Alfanumérico | 20 |
'08 | Apellido y nombre del comprador                                   | 079   | 108 | Alfanumérico | 30 |
'09 | Importe total de la operación                                     | 109   | 123 | Numérico     | 15 | 13 enteros 2 decimales sin punto decimal
'10 | Importe total de conceptos que no integran el precio neto gravado | 124   | 138 | Numérico     | 15 | 13 enteros 2 decimales sin punto decimal
'11 | Percepción a no categorizados                                     | 139   | 153 | Numérico     | 15 | 13 enteros 2 decimales sin punto decimal
'12 | Importe operaciones exentas                                       | 154   | 168 | Numérico     | 15 | 13 enteros 2 decimales sin punto decimal
'13 | Importe de percepciones o pagos a cuenta de impuestos nacionales  | 169   | 183 | Numérico     | 15 | 13 enteros 2 decimales sin punto decimal
'14 | Importe de percepciones de ingresos brutos                        | 184   | 198 | Numérico     | 15 | 13 enteros 2 decimales sin punto decimal
'15 | Importe de percepciones impuestos municipales                     | 199   | 213 | Numérico     | 15 | 13 enteros 2 decimales sin punto decimal
'16 | Importe impuestos internos                                        | 214   | 228 | Numérico     | 15 | 13 enteros 2 decimales sin punto decimal
'17 | Código de Moneda                                                  | 229   | 231 | Alfanumérico | 3  | Según tabla Monedas
'18 | Tipo de cambio                                                    | 232   | 241 | Numérico     | 10 | 4 enteros 6 decimales sin punto decimal
'19 | Cantidad de alícuotas de IVA                                      | 242   | 242 | Numérico     | 1  |
'20 | Código de operación                                               | 243   | 243 | Alfanumérico | 1  | Según tabla Codigo_Operación, de No Corresponder v
'21 | Otros Tributos                                                    | 244   | 258 | Numérico     | 15 |
'22 | Fecha de vencimiento de pago                                      | 259   | 266 | Numérico     | 8  | Fto: AAAAMMDD


    fechaDeComprobante As String
    tipoDeComprobante As String
    PuntoDeVenta As String
    numeroDeComprobante As String
    numeroDeComprobanteHasta As String
    codigoDocumentoComprador As String
    numeroIdentificadorComprador As String
    apellidoYNombreComprador As String
    importeTotalOperacion As String
    importeNoIntegranNetoGravado As String
    percepcionNoCategorizados As String
    importeOperacionesExentas As String
    importePercepcionesPagosACuenta As String
    importePercepcionesIIBB As String
    importePercepcionesMunicipales As String
    importeImpuestosInternos As String
    codigoDeMoneda As String
    tipoDeCambio As String
    cantidadDeAlicuotas As String
    codigoDeOperacion As String
    otrosTributos As String
    fechaVtoPago As String
    
End Type
