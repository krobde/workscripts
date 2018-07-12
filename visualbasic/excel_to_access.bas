Attribute VB_Name = "excel_to_access"
Option Compare Database

 Private Sub xlsAdd_Click()
 
Dim rec As DAO.Recordset
Dim xls As Object
Dim xlsSht As Object
Dim xlsSht2 As Object
Dim xlsWrkBk As Object
Dim xlsPath As String
Dim xlsPath2 As String
Dim xlsFile As String
Dim fullXlsFile As String
Dim fullFile As String
Dim fullFile2 As String
 
Dim Msg, Style, title, Response
  Msg = "Importing is Done, Files are imported!"    ' Define message.
  Style = vbOKOnly
  title = "Import Mesage"
 
    xlsPath = "C:\BASE\TRABAJOS\ESTHER\CIERRE 2010\MUNICIPIOS 2010\"    ' Set the xls path for new files.
    xlsPath2 = "C:\BASE\TRABAJOS\ESTHER\CIERRE 2010\MUNICIPIOS 2010\hechos\"    ' Set the 2nd xls path to store imported files.
    xlsFile = Dir(xlsPath & "*.xls", vbNormal)     ' Retrieve the first entry.
 
    Do While xlsFile <> ""    ' Start the loop.
        ' Ignore the current directory and the encompassing directory.
        fullXlsFile = xlsPath & xlsFile
        fullFile = xlsPath & xlsFile
        fullFile2 = xlsPath2 & xlsFile
        If Right(fullXlsFile, 4) = ".xls" Then 'import it
        DoCmd.SetWarnings False
        Set xls = CreateObject("Excel.Application")
        Set xlsWrkBk = GetObject(fullXlsFile)
        Set xlsSht = xlsWrkBk.Worksheets(1)
        'Set xlsSht2 = xlsWrkBk.Worksheets(2)
 
        'Open 1st table
        Set rec = CurrentDb.OpenRecordset("Tabla_importacion")
        rec.AddNew
        
        rec.Fields("municipio") = xlsFile
        rec.Fields("cargo_b_r") = Nz(StrConv(xlsSht.cells(5, "B"), vbProperCase), "dastos null")
        rec.Fields("cargo_b_l") = Nz(StrConv(xlsSht.cells(6, "B"), vbProperCase), "dastos null")
        rec.Fields("minoraciones_r") = Nz(StrConv(xlsSht.cells(7, "B"), vbProperCase), "dastos null")
        rec.Fields("minoraciones_l") = Nz(StrConv(xlsSht.cells(8, "B"), vbProperCase), "dastos null")
        rec.Fields("cargo_liq_r") = Nz(StrConv(xlsSht.cells(9, "B"), vbProperCase), "dastos null")
        rec.Fields("cargo_liq_l") = Nz(StrConv(xlsSht.cells(10, "B"), vbProperCase), "dastos null")
        rec.Fields("bonif_x_dom") = Nz(StrConv(xlsSht.cells(11, "B"), vbProperCase), "dastos null")
        rec.Fields("cargo_liq_ajus_r") = Nz(StrConv(xlsSht.cells(12, "B"), vbProperCase), "dastos null")
        rec.Fields("cargo_liq_ajus_l") = Nz(StrConv(xlsSht.cells(13, "B"), vbProperCase), "dastos null")
        rec.Fields("recaudacion_r") = Nz(StrConv(xlsSht.cells(16, "B"), vbProperCase), "dastos null")
        rec.Fields("recaudacion_l") = Nz(StrConv(xlsSht.cells(17, "B"), vbProperCase), "dastos null")
        rec.Fields("pendiente_r") = Nz(StrConv(xlsSht.cells(18, "B"), vbProperCase), "dastos null")
        rec.Fields("pendiente_l") = Nz(StrConv(xlsSht.cells(19, "B"), vbProperCase), "dastos null")
        
        rec.Update
 
        'How do I open the second table here to continue exportind the rest of the data?
 
        DoCmd.SetWarnings False
        End If
 
        'Closing excel
        xlsWrkBk.Application.Quit
 
    'Moving the imported Excel file
    Name fullFile As fullFile2
    xlsFile = Dir()
 
    Loop
Response = MsgBox(Msg, Style, title)
End Sub
