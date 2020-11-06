Sub Diario()
Sheets("hoja1").Select
J = WorksheetFunction.CountA(Range("A9:A300"))
Dim fila As String
fila = 9
For W = 1 To J
If Left(Range("I" & fila), 7) = "Enviado" Then
Sheets("Hoja1").Range("a" & fila & ":I" & fila).Select
Selection.Copy
Sheets("DIARIO").Select
B = WorksheetFunction.CountA(Range("a1:a300"))
Range("a" & B + 1).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Sheets("hoja1").Select
Range("a9" & ":a" & fila).Clear
Range("f9" & ":f" & fila).Clear
Range("g9" & ":g" & fila).Clear
Range("h9" & ":h" & fila).Clear
Range("i9" & ":i" & fila).Clear
fila = fila + 1
End If
Next

Range("a9").Select

End Sub

