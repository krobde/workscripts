Attribute VB_Name = "IMPRIMIR_A_PDFS"
'*******************************************************************************
'* mdlPDFCreator
'* código para la impresión de un informe en la impresora virtual PDFCreator
'* ESH 18/09/11 19:54
'*******************************************************************************

Option Compare Database
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Const MaxTime As Long = 10        ' segundos
Public Const SleepTime As Long = 250     ' milisegundos
Public Const SW_NORMAL As Long = 1&


'*******************************************************************************
'* ImprimirInformePDFCreator
'* código para imprimir un informe en PDF por medio de PDFCreator
'* deberá incluir una referencia a la librería PDFCreator, que evidentemente
'* forma parte de la instalación de PDFCreator
'* Argumentos: strInforme       => nombre del informe
'*             strRutaArchivo   => ruta del archivo pdf destino
'*             strNombreArchivo => (opcional) nombre del archivo destino
'*             bytFormato       => (opcional) formato del archivo destino
'*             strExtension     => (opcional) extensión del archivo destino
'*             blnAbrir         => (opcional) abrir el archivo recien creado
'* uso: ImprimirInformePDFCreator
'* ESH 18/09/11 21:58
'* http://www.mvp-access.es/emilio/
'* Si utilizas este código, respeta la autoría y los créditos
'*******************************************************************************

Public Sub ImprimirInformePDFCreator(strInforme As String, strRutaArchivo As String, strSubo As Long, Optional strNombreArchivo As String, Optional bytFormato As Byte, Optional strExtension As String, Optional blnAbrir As Boolean)
Dim strImpresora As String
Dim strArchivo As String
Dim i As Long
Dim PDFPrinter As PDFCreator.clsPDFCreator

On Error GoTo ImprimirInformePDFCreator_TratamientoErrores

DoCmd.Hourglass True
' si no se pasa nombre de archivo, utilizaré el mismo que el del informe
If strNombreArchivo = vbNullString Then strNombreArchivo = strInforme
' si ya existe el archivo confirmo si continuamos y en su caso lo elimino
If Not Dir$(strRutaArchivo & "\" & strNombreArchivo & "." & strExtension) = vbNullString Then
   If MsgBox("El archivo """ & strRutaArchivo & "\" & strNombreArchivo & ".pdf" & """ ya existe ¿Sobreescribir?", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
      Kill strRutaArchivo & "\" & strNombreArchivo & "." & strExtension
   Else
      GoTo ImprimirInformePDFCreator_Salir
   End If
End If
' configuro PDFCreator
Set PDFPrinter = New clsPDFCreator
With PDFPrinter
   .cStart "/NoProcessingAtStartup"
   .cOption("UseAutosave") = 1
   .cOption("UseAutosaveDirectory") = 1
   .cOption("AutosaveDirectory") = strRutaArchivo      ' ruta del archivo
   .cOption("AutosaveFilename") = strNombreArchivo     ' nombre de archivo
   ' AutosaveFormat: valores posibles: 0 = PDF, 1 = PNG, 2 = JPEG, 3 = BMP, 4 = PCX, 5 = TIFF, 6 = PS, 7 = EPS, 8 = TXT, 9 = PDF/A-1b, 10 = PDF/X, 11 = PSD, 12 = PCL, 13 = RAW
   .cOption("AutosaveFormat") = bytFormato
   ' guardo la impresora por defecto
   strImpresora = .cDefaultPrinter
   .cDefaultPrinter = "PDFCreator"
   .cClearCache
   ' imprimo el archivo
   'DoCmd.OpenReport strInforme, acViewNormal, , "ID_SUBO=" & strSubo
   
   
   
   DoCmd.OpenReport strInforme, acViewNormal, , "c_regul.INE = " & strSubo
   
   
   
   .cPrinterStop = False

   ' espero a que acabe de imprimir
   Do While (.cOutputFilename = vbNullString) And (i < (MaxTime * 500 / SleepTime))
      i = i + 1
      Sleep 100
   Loop
   ' recupero el nombre del archivo de salida
   strArchivo = .cOutputFilename
   ' vuelvo a aplicar la impresora por defecto inicial
   .cDefaultPrinter = strImpresora
   Sleep 800
   .cClose
End With
' le doy un tiempo para descargar PDFCreator de la memoria
Sleep 300

' si la variable strArchivo está vacía, fuerzo un error
If strArchivo = vbNullString Then
   Err.Raise 513, "Creación de PDF", "¡Se ha producido un error!"
End If
' si asi se ha solicitado abro el pdf
If blnAbrir Then ShellExecute 0&, "open", strArchivo, 0&, vbNullString, SW_NORMAL


ImprimirInformePDFCreator_Salir:
   On Error GoTo 0
   DoCmd.Hourglass False
   Exit Sub
   
ImprimirInformePDFCreator_TratamientoErrores:
   MsgBox "Error " & Err & " en proc.: ImprimirInformePDFCreator de Módulo: mdlPDFCreator (" & Err.Description & ")", vbCritical + vbOKOnly, "ATENCION"
   Resume ImprimirInformePDFCreator_Salir
Resume Next
End Sub        ' ImprimirInformePDFCreator

Public Sub arrea()


Dim rs As DAO.Recordset
Set rs = CurrentDb.OpenRecordset("SELECT INE, MUNICIPIOS FROM c_regul_")

   Do While Not rs.EOF
      ''expression.OpenReport(ReportName, View, FilterName,
      ''      WhereCondition, WindowMode) -- 2010 has OpenArgs
      ImprimirInformePDFCreator "Informe1", "d:\prueba\", rs!INE, "Entrega_Cuenta_2017_" & rs!INE & "_" & rs!MUNICIPIOS, 0, "pdf", False
      'DoCmd.OpenReport "Invoices",<..>,,"CustomerID=" & rs!CustomerID
      ''OutputTo or other relevant code

      rs.MoveNext
   Loop
End Sub
