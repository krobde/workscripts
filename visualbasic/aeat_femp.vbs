Private Sub Comando0_Click()

    Dim strCtipo0, strCtipo1, strCtipo3, strCtipo4, strCtipo5  As String    'consultas sql
    Dim strFicheroAEAT As String                                            'fichero formato AEAT
    Dim strTemp0, strTemp1, strTemp3, strTemp4, strTemp5 As String          'Temporales para formatear el fichero aeat
    Dim strtmp, strfulltmp As String
    Dim identificador As Recordset
    Dim qd0, qd1, qd3, qd4, qd5 As dao.QueryDef
    Dim n As String
    
        
    
    
    strFicheroAEAT = "c:\FicheroAEAT.txt" 'Fichero formato AEAT
    
    strTemp0 = "c:\temp0.txt" 'temporal registro tipo 0
    strTemp1 = "c:\temp1.txt" 'temporal registro tipo 1
    strTemp3 = "c:\temp3.txt" 'temporal registro tipo 3
    strTemp4 = "c:\temp4.txt" 'temporal registro tipo 4
    strTemp5 = "c:\temp5.txt" 'temporal registro tipo 5
    
    
    Set identificador = CurrentDb.OpenRecordset("SELECT * FROM TIPO3")
        
    strCtipo0 = "SELECT TIPO0.TIPO, TIPO0.TIPO_OF, TIPO0.COD_OF, TIPO0.ANIO_NENVIO, TIPO0.FECHA_ENV, TIPO0.NUM_LIQUIS, TIPO0.IMPORTE_PTE, TIPO0.N_RESPONSABLES, TIPO0.IMPORTE_RESP, TIPO0.PERIODO, TIPO0.IND_MONEDA, TIPO0.LIBRE FROM TIPO0"
    
    Set qd0 = CurrentDb.CreateQueryDef("query0", strCtipo0)
    
    DoCmd.TransferText acExportFixed, "TIPO0", "query0", strTemp0 ' generamos el archivo temporal tipo 0
    
    Open strFicheroAEAT For Append As #9   'rellenamos el fichero final con la cabecera tipo 0
    Open strTemp0 For Input As #10
    
    Line Input #10, strtmp
    strtmp = strtmp & vbNewLine
    Print #9, strtmp
    
    Close #9
    Close #10                               'Cerramos los ficheros
    
    CurrentDb.QueryDefs.Delete ("query0")
    
    
    
    'For Each n In identificador.Fields("ID_DEUDA").Value
    Do While identificador.EOF = False
    n = identificador.Fields("ID_DEUDA").Value
    
    strCtipo1 = "SELECT TIPO1.TIPO, TIPO1.ANIO_NENVIO, TIPO1.ORDEN_REG, TIPO1.ID_DEUDA, TIPO1.DNI, TIPO1.APEYNOM, TIPO1.SIGLAF, TIPO1.NOMBREVIAF, TIPO1.NUMEROF, TIPO1.LETRAF, TIPO1.ESCALERAF, TIPO1.PISOF, TIPO1.PUERTAF, TIPO1.COD_DOMI_AEATF, TIPO1.COD_PROVF, TIPO1.COD_MUNIF, TIPO1.CPF, TIPO1.COD_CONC_PRESUP, TIPO1.TIPO_LIQ_DEUDA, TIPO1.FECHA_LIQUIDACION, TIPO1.FORM_COBRO_DEUDA, TIPO1.ID_LIQ_SANCION, TIPO1.ID_RECURRIDA, TIPO1.EJERCICIO_DEUDA, TIPO1.DESCR_OBJ_TRIB, TIPO1.COD_MUNI, TIPO1.CPT, TIPO1.IMPORTE_PRIN, TIPO1.IMPORTE_REC, TIPO1.IMPORTE_TOTAL, TIPO1.IMPORTE_ING_F_PLAZO, TIPO1.FECHA_ING_F_PLAZO, TIPO1.FECHA_NOT_VOL, TIPO1.TIPO_NOTI, TIPO1.FECHA_VENC_VOL, TIPO1.FECHA_CERT_PROV_APREM, TIPO1.REF_ORG_EMISOR FROM TIPO1 where TIPO1.ID_DEUDA =" & n & ";"
    strCtipo3 = "SELECT TIPO3.TIPO, TIPO3.ANIO_NENVIO, TIPO3.ID_DEUDA, TIPO3.ORDEN_REG, TIPO3.BLOQUE1, TIPO3.BLOQUE2, TIPO3.BLOQUE3, TIPO3.BLOQUE4, TIPO3.BLOQUE5, TIPO3.BLOQUE6, TIPO3.BLOQUE7, TIPO3.BLOQUE8 FROM TIPO3 where TIPO3.ID_DEUDA =" & n & ";"
    strCtipo4 = "SELECT TIPO4.TIPO, TIPO4.ANIO_NENVIO, TIPO4.ORDEN_REG, TIPO4.ID_DEUDA, TIPO4.INF_ADICIONAL, TIPO4.COD_ORG_ACTO, TIPO4.COD_TIP_RECURSO, TIPO4.PERIOD_PRESC, TIPO4.FECHA_ACT_INT_PRESC, TIPO4.LIBRE FROM TIPO4 where TIPO4.ID_DEUDA =" & n & ";"
    strCtipo5 = "SELECT TIPO5.TIPO, TIPO5.ANIO_NENVIO, TIPO5.ORDEN_REG, TIPO5.ID_DEUDA, TIPO5.SIGLA, TIPO5.NOMBREVIA, TIPO5.NUMERO, TIPO5.LETRA, TIPO5.ESCALERA, TIPO5.PISO, TIPO5.PUERTA, TIPO5.COD_DOMI_AEAT, TIPO5.COD_PROV, TIPO5.COD_MUNI, TIPO5.CP, TIPO5.FECHA_NOTAPREM, TIPO5.TIPO_NOTI, TIPO5.FECHA_FIN_APREM, TIPO5.FECHA_INI_EMBARGO, TIPO5.IMPORTE_EMBARGO, TIPO5.LIBRE FROM TIPO5 where TIPO5.ID_DEUDA =" & n & ";"
    
    Set qd1 = CurrentDb.CreateQueryDef("query1", strCtipo1)
    Set qd3 = CurrentDb.CreateQueryDef("query3", strCtipo3)
    Set qd4 = CurrentDb.CreateQueryDef("query4", strCtipo4)
    Set qd5 = CurrentDb.CreateQueryDef("query5", strCtipo5)
    
    
    DoCmd.TransferText acExportFixed, "TIPO1", "query1", strTemp1
    DoCmd.TransferText acExportFixed, "TIPO3", "query3", strTemp3
    DoCmd.TransferText acExportFixed, "TIPO4", "query4", strTemp4
    DoCmd.TransferText acExportFixed, "TIPO5", "query5", strTemp5
    
    
    CurrentDb.QueryDefs.Delete ("query1")
    CurrentDb.QueryDefs.Delete ("query3")
    CurrentDb.QueryDefs.Delete ("query4")
    CurrentDb.QueryDefs.Delete ("query5")
    
    'Output Transaction Header EE2 es la configuracion de exportación.
    'DoCmd.TransferText acExportFixed, "EE2", strFirstQuery, strFicheroAEAT
    'OutPut Transactions
    'DoCmd.TransferText acExportFixed, "Tipo1", strSecondQuery, strFicheroAEAT
    'OutPut Transaction Trailer
    'DoCmd.TransferText acExportFixed, "EE2", strThirdQuery, strFicheroAEAT
    
    Open strFicheroAEAT For Append As #9 'Abrimos fichero AEAT
    Open strTemp1 For Input As #11
    Open strTemp3 For Input As #13
    Open strTemp4 For Input As #14
    Open strTemp5 For Input As #15
    
    
    'Open strResult For Output As #4 'Open Results File
    
    Do While Not EOF(11)
    Line Input #11, strtmp
    strfulltmp = strfulltmp & strtmp & vbNewLine 'inserta tipo 1
    Loop
    Do While Not EOF(13)
    Line Input #13, strtmp
    strfulltmp = strfulltmp & strtmp & vbNewLine 'inserta tipo 3
    Loop
    Do While Not EOF(14)
    Line Input #14, strtmp
    strfulltmp = strfulltmp & strtmp & vbNewLine 'inserta tipo 4
    Loop
    Do While Not EOF(15)
    Line Input #15, strtmp
    strfulltmp = strfulltmp & strtmp & vbNewLine 'inserta tipo 5
    Loop
    
    Print #9, strfulltmp 'Make Merged Results File
    
    Close #9 'Close Files
    Close #15
    Close #14
    Close #13
    Close #11
    
    identificador.MoveNext
    
    Loop
    'Next
    
    identificador.Close

    
End Sub
