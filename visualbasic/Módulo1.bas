Attribute VB_Name = "Módulo1"
Option Compare Database
Public Function formato()


'Dim rst As DAO.Recordset

Dim LineData As String


'Dim CODE1 As Integer ' Holder for Account Number in text file
'Dim COOMPANIE As String ' Holder for the Cusip Number or Ticker Symbol
'Dim TYPEOP As String ' Holder for the number of shares held in account
'Dim MOUNANT As String ' Holder for Long or Short Position
'Dim TYPEMOU As String ' Holder for value of security
'Dim DATE1 As String ' Holder for Net Asset Value
'Dim PRO As String
 Dim sql0 As String

'Registro tipo 0

Dim tipoRegistro As String
Dim codDelegacion As Integer
Dim codOfiLiq As String
Dim nifOfiLiq As String
Dim fIniPerLiq As String
Dim fFinPerLiq As String
Dim fSoporte As String
Dim moneda As String
Dim blancos0 As String

'registro tipo 1
Dim secuencialT1 As Integer
Dim claveLiqSIR As String
Dim refExtLiq As String
Dim antClavLiq As String
Dim importe As Double
Dim flIngEje As String
Dim nifDeudor As String
Dim apeynom As String
Dim fNotAprem As String
Dim fAceptDeuda As String
Dim indDeuda As String
Dim blancos1 As String

'registro tipo 2
Dim secuencialT2 As Integer
Dim indPendInicio As String
Dim importePendIni As Double
Dim importeRehabil As Double
Dim importeReactiva As Double
Dim impTotPerLiq As Double
Dim fContUltIng As String
Dim fIngUltIng As String
Dim impAnulado As Double
Dim impCancelIncobr As Double
Dim impCancelOT As Double
Dim impCancelPP As Double
Dim impTotIng As Double
Dim fCancelAnul As String
Dim fCancelIncob As String
Dim fCancelOT As String
Dim fCancelPP As String
Dim impPendFinLiq As Double
Dim impIntIngPerant As Double
Dim impTotIngInt As Double
Dim fContUltII As String
Dim fIngUltII As String
Dim impTotCostas As Double
Dim impDisminucion As Double

'Registro tipo 3
Dim secuencialT3 As Integer
Dim indTipIng As Integer
Dim impIngreso As Double
Dim fIngreso As String
Dim fAplicContIng As String
Dim blancos3 As String

'Registro tipo 5 y 7
Dim totDeudas As Double
Dim tImPenIniLiq As Double
Dim tImPenFinLiq As Double
Dim tImCarPerLiq As Double
Dim tImpReact As Double
Dim tImpRehab As Double
Dim tIngConcDeu As Double
Dim tIngConcInt As Double
Dim tImpCancAnu As Double
Dim tImpCancPI As Double
Dim tImpCancOT As Double
Dim tImpCancPP As Double
Dim tImpIngPerliq As Double
Dim tImpCosPerliq As Double
Dim tImpDismi As Double
Dim numeroTipo1 As Integer
Dim numeroTipo2 As Integer
Dim numeroTipo3 As Integer
Dim numTotTip01235 As Integer
Dim numTotTip012357 As Integer
Dim blancos57 As String

'Registro 8 y 9
'Coste del servicio

Dim altas As Double
Dim importe1 As Double
Dim importe2 As Double
Dim costeTotal As Double
Dim blancos8 As String
'Dim codOfiLiq As String
Dim blancos9 As String







' Open the text file
'Set rst = CurrentDb.OpenRecordset("Table1", dbOpenDynaset)
Open "C:\AEAT\suministrosaapp\EJECUTIVA\IFSUFIR02_502A0000.TXT" For Input As #1
Do Until EOF(1)

Line Input #1, LineData
tipoRegistro = Left(LineData, 1)
Select Case tipoRegistro

Case "0"

codDelegacion = Mid(LineData, 2, 2)
codOfiLiq = Mid(LineData, 4, 6)
nifOfiLiq = Mid(LineData, 10, 9)
fIniPerLiq = Mid(LineData, 19, 8)
fFinPerLiq = Mid(LineData, 27, 8)
fSoporte = Mid(LineData, 35, 8)
moneda = Mid(LineData, 43, 1)
'blancos0 = Mid(LineData, 44, 1)

If ifTableExists("TIPO0") = False Then

sql0 = "SELECT '" & codDelegacion & "' as codDelegacion, '" & codOfiLiq & "' as codOfiLiq, '" & nifOfiLiq & "' as nifOfiLiq, '" & fIniPerLiq & "' as fIniPerLiq, '" & fFinPerLiq & "' as fFinPerLiq, '" & fSoporte & "' as fSoporte, '" & moneda & "' as moneda INTO TIPO0"

Else

sql0 = "INSERT INTO TIPO0 VALUES ('" & codDelegacion & "', '" & codOfiLiq & "', '" & nifOfiLiq & "', '" & fIniPerLiq & "', '" & fFinPerLiq & "', '" & fSoporte & "', '" & moneda & "')"

End If

CurrentDb.Execute sql0

Case "1"
secuencialT1 = Mid(LineData, 2, 7)
claveLiqSIR = Mid(LineData, 9, 17)
refExtLiq = Mid(LineData, 26, 20)
antClavLiq = Mid(LineData, 46, 20)
importe = Mid(LineData, 66, 11)
flIngEje = Mid(LineData, 77, 8)
nifDeudor = Mid(LineData, 85, 9)
apeynom = Mid(LineData, 94, 40)
fNotAprem = Mid(LineData, 134, 8)
fAceptDeuda = Mid(LineData, 142, 8)
indDeuda = Mid(LineData, 150, 1)

If ifTableExists("TIPO1") = False Then
sql0 = "SELECT '" & secuencialT1 & "' as secuencialT1, '" & claveLiqSIR & "' as claveLiqSIR, '" & refExtLiq & "' as refExtLiq, '" & antClavLiq & "' as antClavLiq, '" & importe & "' as importe, '" & flIngEje & "' as flIngEje, '" & nifDeudor & "' as nifDeudor, '" & apeynom & "' as apeynom, '" & fNotAprem & "' as fNotAprem, '" & fAceptDeuda & "' as fAceptDeuda, '" & indDeuda & "' as indDeuda INTO TIPO1"
Else
sql0 = "INSERT INTO TIPO1 VALUES ('" & secuencialT1 & "', '" & claveLiqSIR & "', '" & refExtLiq & "', '" & antClavLiq & "', '" & importe & "', '" & flIngEje & "', '" & nifDeudor & "', '" & apeynom & "', '" & fNotAprem & "', '" & fAceptDeuda & "', '" & indDeuda & "')"
End If

CurrentDb.Execute sql0

Case "2"


secuencialT2 = Mid(LineData, 2, 7)
claveLiqSIR = Mid(LineData, 9, 17)
indPendInicio = Mid(LineData, 26, 1)
importePendIni = Mid(LineData, 27, 11)
importeRehabil = Mid(LineData, 38, 11)
importeReactiva = Mid(LineData, 49, 11)
impTotPerLiq = Mid(LineData, 60, 11)
fContUltIng = Mid(LineData, 71, 8)
fIngUltIng = Mid(LineData, 79, 8)
impAnulado = Mid(LineData, 87, 11)
impCancelIncobr = Mid(LineData, 98, 11)
impCancelOT = Mid(LineData, 109, 11)
impCancelPP = Mid(LineData, 120, 11)
impTotIng = Mid(LineData, 131, 11)
fCancelAnul = Mid(LineData, 142, 8)
fCancelIncob = Mid(LineData, 150, 8)
fCancelOT = Mid(LineData, 158, 8)
fCancelPP = Mid(LineData, 166, 8)
impPendFinLiq = Mid(LineData, 174, 11)
impIntIngPerant = Mid(LineData, 185, 11)
impTotIngInt = Mid(LineData, 196, 11)
fContUltII = Mid(LineData, 207, 8)
fIngUltII = Mid(LineData, 215, 8)
impTotCostas = Mid(LineData, 223, 11)
impDisminucion = Mid(LineData, 234, 11)

If ifTableExists("TIPO2") = False Then

sql0 = "SELECT '" & secuencialT2 & "' as secuencialT2, '" & claveLiqSIR & "' as claveLiqSIR, '" & indPendInicio & "' as indPendInicio, '" & importePendIni & "' as importePendIni, '" & importeRehabil & "' as importeRehabil, '" & importeReactiva & "' as importeReactiva, '" & impTotPerLiq & "' as impTotPerLiq, '" & fContUltIng & "' as fContUltIng, '" & fIngUltIng & "' as fIngUltIng, '" & impAnulado & "' as impAnulado, '" & impCancelIncobr & "' as impCancelIncobr, '" & impCancelOT & "' as impCancelOT, '" & impCancelPP & "' as impCancelPP, '" & impTotIng & "' as impTotIng, '" & fCancelAnul & "' as fCancelAnul, '" & fCancelIncob & "' as fCancelIncob, '" & fCancelOT & "' as fCancelOT, '" & fCancelPP & "' as fCancelPP, '" & impPendFinLiq & "' as impPendFinLiq, '" & impIntIngPerant & "' as impIntIngPerant, '" & impTotIngInt & "' as impTotIngInt, '" & fContUltII & "' as fContUltII, '" & fIngUltII & "' as fIngUltII, '" & impTotCostas & "' as impTotCostas, '" & impDisminucion & "' as impDisminucion INTO TIPO2"

Else

sql0 = "INSERT INTO TIPO2 VALUES ('" & secuencialT2 & "', '" & claveLiqSIR & "', '" & indPendInicio & "', '" & importePendIni & "', '" & importeRehabil & "', '" & importeReactiva & "', '" & impTotPerLiq & "', '" & fContUltIng & "', '" & fIngUltIng & "', '" & impAnulado & "', '" & impCancelIncobr & "', '" & impCancelOT & "', '" & impCancelPP & "', '" & impTotIng & "', '" & fCancelAnul & "', '" & fCancelIncob & "', '" & fCancelOT & "', '" & fCancelPP & "', '" & impPendFinLiq & "', '" & impIntIngPerant & "', '" & impTotIngInt & "', '" & fContUltII & "', '" & fIngUltII & "', '" & impTotCostas & "', '" & impDisminucion & "')"

End If

CurrentDb.Execute sql0




Case "3"

secuencialT3 = Mid(LineData, 0, 0)
claveLiqSIR = Mid(LineData, 0, 0)
indTipIng = Mid(LineData, 0, 0)
impIngreso = Mid(LineData, 0, 0)
fIngreso = Mid(LineData, 0, 0)
fAplicContIng = Mid(LineData, 0, 0)

If ifTableExists("TIPO3") = False Then

sql0 = "SELECT '" & secuencialT3 & "' as secuencialT3, '" & claveLiqSIR & "' as claveLiqSIR, '" & indTipIng & "' as indTipIng, '" & impIngreso & "' as impIngreso, '" & fIngreso & "' as fIngreso, '" & fAplicContIng & "' as fAplicContIng INTO TIPO3"

Else

sql0 = "INSERT INTO TIPO3 VALUES ('" & secuencialT3 & "', '" & claveLiqSIR & "', '" & indTipIng & "', '" & impIngreso & "', '" & fIngreso & "', '" & fAplicContIng & "')"

End If

CurrentDb.Execute sql0


Case "5"

totDeudas = Mid(LineData, 2, 13)
tImPenIniLiq = Mid(LineData, 15, 13)
tImPenFinLiq = Mid(LineData, 28, 13)
tImCarPerLiq = Mid(LineData, 41, 13)
tImpReact = Mid(LineData, 54, 13)
tImpRehab = Mid(LineData, 67, 13)
tIngConcDeu = Mid(LineData, 80, 13)
tIngConcInt = Mid(LineData, 93, 13)
tImpCancAnu = Mid(LineData, 106, 13)
tImpCancPI = Mid(LineData, 119, 13)
tImpCancOT = Mid(LineData, 132, 13)
tImpCancPP = Mid(LineData, 145, 13)
tImpIngPerliq = Mid(LineData, 158, 13)
tImpCosPerliq = Mid(LineData, 171, 13)
tImpDismi = Mid(LineData, 184, 13)
numeroTipo1 = Mid(LineData, 197, 7)
numeroTipo2 = Mid(LineData, 204, 7)
numeroTipo3 = Mid(LineData, 211, 7)
numTotTip01235 = Mid(LineData, 218, 8)


If ifTableExists("TIPO5") = False Then

sql0 = "SELECT '" & totDeudas & "' as totDeudas, '" & tImPenIniLiq & "' as tImPenIniLiq, '" & tImPenFinLiq & "' as tImPenFinLiq, '" & tImCarPerLiq & "' as tImCarPerLiq, '" & tImpReact & "' as tImpReact, '" & tImpRehab & "' as tImpRehab, '" & tIngConcDeu & "' as tIngConcDeu, '" & tIngConcInt & "' as tIngConcInt, '" & tImpCancAnu & "' as tImpCancAnu, '" & tImpCancPI & "' as tImpCancPI, '" & tImpCancOT & "' as tImpCancOT, '" & tImpCancPP & "' as tImpCancPP, '" & tImpIngPerliq & "' as tImpIngPerliq, '" & tImpCosPerliq & "' as tImpCosPerliq, '" & tImpDismi & "' as tImpDismi, '" & numeroTipo1 & "' as numeroTipo1, '" & numeroTipo2 & "' as numeroTipo2, '" & numeroTipo3 & "' as numeroTipo3, '" & numTotTip01235 & "' as numTotTip01235 INTO TIPO5"

Else

sql0 = "INSERT INTO TIPO5 VALUES ('" & totDeudas & "' , '" & tImPenIniLiq & "' , '" & tImPenFinLiq & "' , '" & tImCarPerLiq & "' , '" & tImpReact & "' , '" & tImpRehab & "' , '" & tIngConcDeu & "' , '" & tIngConcInt & "' , '" & tImpCancAnu & "' , '" & tImpCancPI & "' , '" & tImpCancOT & "' , '" & tImpCancPP & "' , '" & tImpIngPerliq & "' , '" & tImpCosPerliq & "' , '" & tImpDismi & "' , '" & numeroTipo1 & "' , '" & numeroTipo2 & "' , '" & numeroTipo3 & "' , '" & numTotTip01235 & "')"

End If

CurrentDb.Execute sql0

Case "7"


totDeudas = Mid(LineData, 2, 13)
tImPenIniLiq = Mid(LineData, 15, 13)
tImPenFinLiq = Mid(LineData, 28, 13)
tImCarPerLiq = Mid(LineData, 41, 13)
tImpReact = Mid(LineData, 54, 13)
tImpRehab = Mid(LineData, 67, 13)
tIngConcDeu = Mid(LineData, 80, 13)
tIngConcInt = Mid(LineData, 93, 13)
tImpCancAnu = Mid(LineData, 106, 13)
tImpCancPI = Mid(LineData, 119, 13)
tImpCancOT = Mid(LineData, 132, 13)
tImpCancPP = Mid(LineData, 145, 13)
tImpIngPerliq = Mid(LineData, 158, 13)
tImpCosPerliq = Mid(LineData, 171, 13)
tImpDismi = Mid(LineData, 184, 13)
numeroTipo1 = Mid(LineData, 197, 7)
numeroTipo2 = Mid(LineData, 204, 7)
numeroTipo3 = Mid(LineData, 211, 7)
numTotTip012357 = Mid(LineData, 218, 8)


If ifTableExists("TIPO7") = False Then

sql0 = "SELECT '" & totDeudas & "' as totDeudas, '" & tImPenIniLiq & "' as tImPenIniLiq, '" & tImPenFinLiq & "' as tImPenFinLiq, '" & tImCarPerLiq & "' as tImCarPerLiq, '" & tImpReact & "' as tImpReact, '" & tImpRehab & "' as tImpRehab, '" & tIngConcDeu & "' as tIngConcDeu, '" & tIngConcInt & "' as tIngConcInt, '" & tImpCancAnu & "' as tImpCancAnu, '" & tImpCancPI & "' as tImpCancPI, '" & tImpCancOT & "' as tImpCancOT, '" & tImpCancPP & "' as tImpCancPP, '" & tImpIngPerliq & "' as tImpIngPerliq, '" & tImpCosPerliq & "' as tImpCosPerliq, '" & tImpDismi & "' as tImpDismi, '" & numeroTipo1 & "' as numeroTipo1, '" & numeroTipo2 & "' as numeroTipo2, '" & numeroTipo3 & "' as numeroTipo3, '" & numTotTip01235 & "' as numTotTip012357 INTO TIPO7"

Else

sql0 = "INSERT INTO TIPO7 VALUES ('" & totDeudas & "' , '" & tImPenIniLiq & "' , '" & tImPenFinLiq & "' , '" & tImCarPerLiq & "' , '" & tImpReact & "' , '" & tImpRehab & "' , '" & tIngConcDeu & "' , '" & tIngConcInt & "' , '" & tImpCancAnu & "' , '" & tImpCancPI & "' , '" & tImpCancOT & "' , '" & tImpCancPP & "' , '" & tImpIngPerliq & "' , '" & tImpCosPerliq & "' , '" & tImpDismi & "' , '" & numeroTipo1 & "' , '" & numeroTipo2 & "' , '" & numeroTipo3 & "' , '" & numTotTip012357 & "')"

End If

CurrentDb.Execute sql0


Case "8"

altas = Mid(LineData, 2, 13)
importe1 = Mid(LineData, 15, 13)
importe2 = Mid(LineData, 28, 13)
costeTotal = Mid(LineData, 41, 13)

If ifTableExists("TIPO8") = False Then

sql0 = "SELECT '" & altas & "' as altas, '" & importe1 & "' as importe1, '" & importe2 & "' as importe2, '" & costeTotal & "' as costeTotal INTO TIPO8"

Else

sql0 = "INSERT INTO TIPO8 VALUES ('" & altas & "', '" & importe1 & "', '" & importe2 & "', '" & costeTotal & "')"

End If

CurrentDb.Execute sql0


Case "9"
altas = Mid(LineData, 2, 13)
importe1 = Mid(LineData, 15, 13)
importe2 = Mid(LineData, 28, 13)
costeTotal = Mid(LineData, 41, 13)
claveLiqSIR = Mid(LineData, 54, 17)
codOfiLiq = Mid(LineData, 71, 6)

sql0 = "SELECT '" & altas & "' as altas, '" & importe1 & "' as importe1, '" & importe2 & "' as importe2, '" & costeTotal & "' as costeTotal, '" & claveLiqSIR & "' as claveLiqSIR, '" & codOfiLiq & "' as codOfiLiq INTO TIPO9"

Else

sql0 = "INSERT INTO TIPO9 VALUES ('" & altas & "', '" & importe1 & "', '" & importe2 & "', '" & costeTotal & "', '" & claveLiqSIR & "', '" & codOfiLiq & "')"

End If

CurrentDb.Execute sql0

Case Else

End Select

'CODE1 = Left(LineData, 2)
'COMPANIE = Mid(LineData, 3, 1)
'TYPEOP = Mid(LineData, 4, 4)
'MOUNTANT = Mid(LineData, 8, 4)
'TYPEMOU = Mid(LineData, 12, 4)
'DATE1 = Mid(LineData, 16, 4)
'PRO = Mid(LineData, 20, 4)
'With rst

'.AddNew

'rst.Fields(1) = CODE1
'rst.Fields(2) = COMPANIE
'rst.Fields(3) = TYPEOP
'rst.Fields(4) = MOUNTANT
'rst.Fields(5) = TYPEOP
'rst.Fields(6) = DATE1
'rst.Fields(7) = PRO
'.Update
'End With
Loop
' Close the data file.
Close #1
End Function

Public Function ifTableExists(tblName As String) As Boolean

ifTableExists = False
If DCount("[Name]", "MSysObjects", "[Name] = '" & tblName & "'") = 1 Then
ifTableExists = True
End If

End Function
