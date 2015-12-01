Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
'para paypal
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const conSwNormal = 1
'Para el programa
Public TOT As Variant
Sub iniciar_pen()
Pen.Show
End Sub
Sub donar() 'Envia a la pagina de paypal
ShellExecute hwnd, "open", _
"https://www.paypal.com/cgi-bin/webscr?cmd=_xclick&business=sergio%2eovalle%40lynxsig%2ecom&item_name=Pensilvania&item_number=01&no_shipping=1&cn=Favor%20de%20dejar%20un%20Comentario&tax=0&currency_code=USD&lc=GT&bn=PP%2dDonationsBF&charset=UTF%2d8https://www.paypal.com/cgi-bin/webscr?cmd=_xclick&business=sergio%2eovalle%40lynxsig%2ecom&item_name=Pensilvania&item_number=01&no_shipping=1&cn=Favor%20de%20dejar%20un%20Comentario&tax=0&currency_code=USD&lc=GT&bn=PP%2dDonationsBF&charset=UTF%2d8" _
, vbNullString, vbNullString, conSwNormal
End Sub
Sub PENSIL()
    
    'COORDENADAS PARCIALES LATITUD
    Range("E1").Value = "Y PAR"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=RC3*COS(RADIANS(RC4))"
    Range("E2").Select
    Selection.AutoFill Destination:=Range("E2:E" & TOT + 1)
    Range("E2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    'CALCULO DE ERROR
    Range("C" & TOT + 6).Value = "E Lat"
    Range("D" & TOT + 6).Value = "=ABS(SUM(C[1]))"
    Range("D" & TOT + 6).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    'CALCULO DEL FACTOR DE CORRECCION
    Range("C" & TOT + 8).Value = "F C Lat"
    Range("D" & TOT + 8).Value = _
    "=R[-2]C/(ABS(SUMIF(C[1],""<0""))+ABS(SUMIF(C[1],"">0"")))"
    Range("D" & TOT + 8).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

    'COLUMNA DE CORRECCION
    Range("F1").Value = "CORR"
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "=ABS(RC[-1]*R" & TOT + 8 & "C4)"
    Range("F2").Select
    Selection.AutoFill Destination:=Range("F2:F" & TOT + 1)
    Range("F2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    'Y CORREGIDA
    Range("I1").Value = "Y COR"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = _
    "=IF(SUMIF(C[-4],"">0"")>ABS(SUMIF(C[-4],""<0"")),RC[-4]-RC[-3],-(RC[-4]+RC[-3]))"
    Range("I2").Select
    Selection.AutoFill Destination:=Range("I2:I" & TOT + 1)
    Range("I2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

    'Y TOTAL
    Range("K1").Value = "Y TOT"
    Range("K2").Value = "=RC[-2]"
    Range("K3").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C+RC[-2]"
    Range("K3").Select
    Selection.AutoFill Destination:=Range("K3:K" & TOT + 1)
    Range("K2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False


    'COORDENADAS PARCIALES LONGITUD
    Range("G1").Value = "X PAR"
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "=RC3*SIN(RADIANS(RC4))"
    Range("G2").Select
    Selection.AutoFill Destination:=Range("G2:G" & TOT + 1)
    Range("G2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    'CALCULO DEL ERROR
    Range("C" & TOT + 7).Value = "E Lon"
    Range("D" & TOT + 7).Value = "=ABS(SUM(C[3]))"
    Range("D" & TOT + 7).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    'CALCULO DEL FACTOR DE CORRECCION
    Range("C" & TOT + 9).Value = "F C Lon"
    Range("D" & TOT + 9).Value = _
    "=R[-2]C/(ABS(SUMIF(C[3],""<0""))+ABS(SUMIF(C[3],"">0"")))"
    Range("D" & TOT + 9).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    'CALCULO DE LA CORRECCION
    Range("H1").Value = "CORR"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "=ABS(RC[-1]*R" & TOT + 9 & "C4)"
    Range("H2").Select
    Selection.AutoFill Destination:=Range("H2:H" & TOT + 1)
    Range("H2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
        
    'X CORREGIDA
    Range("J1").Value = "X COR"
    Range("J2").Select
    ActiveCell.FormulaR1C1 = _
    "=IF(SUMIF(C[-3],"">0"")>ABS(SUMIF(C[-3],""<0"")),RC[-3]-RC[-2],-(RC[-3]+RC[-2]))"
    Range("J2").Select
    Selection.AutoFill Destination:=Range("J2:J" & TOT + 1)
    Range("J2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
        
    'X TOTAL
    Range("L1").Value = "X TOT"
    Range("L2").Value = "=RC[-2]"
    Range("L3").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C+RC[-2]"
    Range("L3").Select
    Selection.AutoFill Destination:=Range("L3:L" & TOT + 1)
    Range("L2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

    'ERROR LINEAL DE CIERRE
    Range("C" & TOT + 10).Value = "E L C"
    Range("D" & TOT + 10).Value = _
    "=SQRT(R[-4]C^2+R[-3]C^2)"
    Range("D" & TOT + 10).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    'DISTANCIA
    Range("C" & TOT + 5).Value = "DISTANCIA"
    Range("D" & TOT + 5).Value = _
    "=SUM(C[-1])"
    Range("D" & TOT + 5).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    'PRESICIÓN
    Range("C" & TOT + 11).Value = "PRESICIÓN"
    Range("D" & TOT + 11).Value = _
    "=R[-1]C/R[-6]C"
    Range("D" & TOT + 11).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

    'CALCULO XY
    Range("M2").Value = _
    "=IF(R[1]C[-1]="""",RC[-2]*R2C12,RC[-2]*R[1]C[-1])"
    Range("M2").Select
    Selection.AutoFill Destination:=Range("M2:M" & TOT + 1)
    Range("M2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("M1").Value = Empty
    
    'CALCULO YX
    Range("N2").Value = _
    "=IF(R[1]C[-3]="""",RC[-2]*R2C11,RC[-2]*R[1]C[-3])"
    Range("N2").Select
    Selection.AutoFill Destination:=Range("N2:N" & TOT + 1)
    Range("N2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("N1").Value = Empty
    
    'ÁREA
    Range("C" & TOT + 4).Value = "ÁREA"
    Range("D" & TOT + 4).Value = _
    "=(SUM(C[9])-SUM(C[10]))/2"
    Range("D" & TOT + 4).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("M:N").Value = Empty
End Sub
