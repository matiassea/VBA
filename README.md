# VBA
Codigo VBA

## Distiantas formas de Paste

https://powerspreadsheets.com/excel-vba-copy-paste/

## Buscar intitucion segun Businees Unit
```
Worksheets("Hoja2").Activate
Dim lastRow As Long
Set ws = Worksheets("Hoja1")
lastRow = ws.Range("A" & Rows.Count).End(xlUp).Row 'conteo de columna
For i = 2 To lastRow
        If Cells(i, 2).Value = "Instituto Profesional AIEP S.A." Then
        Cells(i, 1) = "CHL04"
        ElseIf Cells(i, 2).Value = "UNAB" Then
        Cells(i, 1) = "CHL01"
        ElseIf Cells(i, 2).Value = "Universidad Privada del Norte" Then
        Cells(i, 1) = "PER03"
        ElseIf Cells(i, 2).Value = "Univ. De Viña del Mar Chile OP" Then
        Cells(i, 1) = "CHL32"
        ElseIf Cells(i, 2).Value = "Universidad Perú Ciencias Aplicadas" Then
        Cells(i, 1) = "PER05"
        ElseIf Cells(i, 2).Value = "UDLA Chile" Then
        Cells(i, 1) = "CHL02"
        ElseIf Cells(i, 2).Value = "Cibertec" Then
        Cells(i, 1) = "PER06"
        ElseIf Cells(i, 2).Value = "IEDE Chile" Then
        Cells(i, 1) = "CHL05"
        ElseIf Cells(i, 2).Value = "Inmobiliaria Educ SPA (IESA)" Then
        Cells(i, 1) = "CHL18"
        ElseIf Cells(i, 2).Value = "Laureate Chile II SPA" Then
        Cells(i, 1) = "CHL25"
        ElseIf Cells(i, 2).Value = "Servicios Andinos" Then
        Cells(i, 1) = "CHL28"
        ElseIf Cells(i, 2).Value = "Immob Inversiones SanGenarosDos" Then
        Cells(i, 1) = "CHL31"
        ElseIf Cells(i, 2).Value = "Servicios Profesionales Andrés Bello" Then
        Cells(i, 1) = "CHL08"
        End If
Next i

Dim lastRow As Long
Set ws = Worksheets("Hoja1")
lastRow = ws.Range("A" & Rows.Count).End(xlUp).Row 'conteo de columna
For i = 2 To lastRow
        If Cells(i, 2) = "CHL04" Then
        Cells(i, 46).Value = "Instituto Profesional AIEP S.A."
        ElseIf Cells(i, 2) = "CHL01" Then
        Cells(i, 46).Value = "UNAB"
        ElseIf Cells(i, 2) = "PER03" Then
        Cells(i, 46).Value = "Universidad Privada del Norte"
        ElseIf Cells(i, 2) = "CHL32" Then
        Cells(i, 46).Value = "Univ. De Viña del Mar Chile OP"
        ElseIf Cells(i, 2) = "PER05" Then
        Cells(i, 46).Value = "Universidad Perú Ciencias Aplicadas"
        ElseIf Cells(i, 2) = "CHL02" Then
        Cells(i, 46).Value = "UDLA Chile"
        ElseIf Cells(i, 2) = "PER06" Then
        Cells(i, 46).Value = "Cibertec"
        ElseIf Cells(i, 2) = "CHL05" Then
        Cells(i, 46).Value = "IEDE Chile"
        ElseIf Cells(i, 2) = "CHL18" Then
        Cells(i, 46).Value = "Inmobiliaria Educ SPA (IESA)"
        ElseIf Cells(i, 2) = "CHL25" Then
        Cells(i, 46).Value = "Laureate Chile II SPA"
        ElseIf Cells(i, 2) = "CHL28" Then
        Cells(i, 46).Value = "Servicios Andinos"
        ElseIf Cells(i, 2) = "CHL31" Then
        Cells(i, 46).Value = "Immob Inversiones SanGenarosDos"
        ElseIf Cells(i, 2) = "CHL08" Then
        Cells(i, 46).Value = "Servicios Profesionales Andrés Bello"
        End If
Next i
```

## Concatenas String de distintas columnas

```
Dim lastRow As Long
Set ws = Worksheets("Hoja1")
lastRow = ws.Range("A" & Rows.Count).End(xlUp).Row 'conteo de columna
For i = 2 To lastRow
    Cells(i, 6).Value = Cells(i, 1) & Cells(i, 3)
Next i
```


## Conteo de lineas de la PO o PR
```
Dim ws As Worksheet
Dim lastRow As Long, x As Long
Dim items As Object

Application.ScreenUpdating = False
  
Set ws = Worksheets("sheet1")
Cells(1, 47).Value = "Cantidad de lineas" 'Donde va el titulo
            Cells(1, 47).Font.Bold = True 'Donde va el titulo en negrita
    
lastRow = ws.Range("A" & Rows.Count).End(xlUp).Row 'conteo de columna
    
    Set items = CreateObject("Scripting.Dictionary")
    For x = 2 To lastRow
        If Not items.Exists(ws.Cells(x, 46).Value) Then 'columna de conteo columna 1
            items.Add ws.Cells(x, 46).Value, 1 'columna de conteo columna 1
            ws.Cells(x, 47).Value = items(ws.Cells(x, 46).Value) 'columna donde deja = columna de conteo columna 1
        Else
            items(ws.Cells(x, 46).Value) = items(ws.Cells(x, 46).Value) + 1 'columna de conteo columna 1 = columna de conteo columna 1 + 1
            ws.Cells(x, 47).Value = items(ws.Cells(x, 46).Value) 'columna donde deja = columna de conteo columna 1
        End If
    Next x
End Sub
```

## Clasificacion del tipo de compras Catalogo, Contrato o Sourcing
```
Dim lastRow As Long
Set ws = Worksheets("Hoja1")

lastRow = ws.Range("A" & Rows.Count).End(xlUp).Row 'conteo de columna
For c = 2 To lastRow
   Select Case True
   '23 = Id Articulo
   '28 = Itm Id Vndr
    
    Case Cells(c, 16).Value = "" And Cells(c, 17).Value = ""
            Cells(c, 47) = "Sourcing"
            Cells(c, 47).Font.Bold = True 'negrita
            Cells(c, 47).Font.ColorIndex = 4 'pintar en verde
   
    Case InStr(Cells(c, 16).Value, "CNTRT") > 0 Or InStr(Cells(c, 17).Value, "CNTRT") > 0
            Cells(c, 47) = "Contrato"
            Cells(c, 47).Font.Bold = True 'negrita
            Cells(c, 47).Font.ColorIndex = 17 'pintar en azulado
            
    Case InStr(Cells(c, 16).Value, "PER") > 0 Or InStr(Cells(c, 17).Value, "PER") > 0
            Cells(c, 47) = "Contrato"
            Cells(c, 47).Font.Bold = True 'negrita
            Cells(c, 47).Font.ColorIndex = 17 'pintar en azulado
    
    Case Cells(c, 16).Value <> "" And Cells(c, 17).Value = ""
            Cells(c, 47) = "Catalogo"
            Cells(c, 47).Font.Bold = True 'negrita
            Cells(c, 47).Font.ColorIndex = 33 'pintar en calipto
    
    Case Cells(c, 16).Value = "" And Cells(c, 17).Value <> ""
            Cells(c, 47) = "Catalogo"
            Cells(c, 47).Font.Bold = True 'negrita
            Cells(c, 47).Font.ColorIndex = 33 'pintar en calipto
            
    Case Cells(c, 16).Value <> "" And Cells(c, 17).Value <> ""
            Cells(c, 47) = "Catalogo"
            Cells(c, 47).Font.Bold = True 'negrita
            Cells(c, 47).Font.ColorIndex = 33 'pintar en calipto
    End Select
Next c
```

## Trasnformacion a Pesos o Dolares

```
'Dim d As Integer
'Worksheets("sheet1").Activate
For d = 2 To 52844
    If Cells(d, 22) = "CLP" Then
    Cells(d, 48) = Cells(d, 29) / 649
    Else
    Cells(d, 48) = Cells(d, 29) / 3.3
    End If
Next d
End Sub
```

## Definicion de Threshold
```
Sub desglose()
'Dim d As Integer
'Worksheets("sheet1").Activate
For d = 2 To 13222
    If Cells(d, 4) <= 100 And 0 <= Cells(d, 4) Then
    Cells(d, 5) = "< 100"
    ElseIf Cells(d, 4) <= 200 And 100 < Cells(d, 4) Then
    Cells(d, 5) = "< 200"
    ElseIf Cells(d, 4) <= 300 And 200 < Cells(d, 4) Then
    Cells(d, 5) = "< 300"
    ElseIf Cells(d, 4) <= 400 And 300 < Cells(d, 4) Then
    Cells(d, 5) = "< 400"
    ElseIf Cells(d, 4) <= 500 And 400 < Cells(d, 4) Then
    Cells(d, 5) = "< 500"
    ElseIf Cells(d, 4) > 500 Then
    Cells(d, 5) = "> 500"
    End If
Next d
End Sub
```


```
Sub mayor_a()

Application.ScreenUpdating = False 'Apagar el parpadeo de pantalla, Evita los movimientos de pantalla que se producen al seleccionar celdas, hojas y libros
Application.Calculation = xlCalculationManual 'Apagar los cálculos automáticos, Evita que se recalcule todo cada vez que se pegan o modifican datos
Application.EnableEvents = False 'Apagar los eventos automáticos, Evita que se disparen macros de evento si las hubiere
ActiveSheet.DisplayPageBreaks = False 'Apagar visualización de saltos de página, Sirve para evitar algunos problemas de compatibilidad entre macros Excel 2003 vs. 2007/2010
Cells(1, 24).Value = "Catalogos-Contratos"
            Cells(1, 24).Font.Bold = True 'negrita
For c = 2 To 257
   Select Case True
    
    Case InStr(Cells(c, 24).Value, "CNTR") > 0
        Cells(c, 49) = "Contrato"
        
    Case Len(Cells(c, 24)) > 0
        Cells(c, 49) = "Catalogo"
        
    Case Else
        Cells(c, 49) = "Sourcing"
    End Select
Next c
End Sub
```

```
Sub countPO2()

Dim ws As Worksheet
Dim lastRow As Long, x As Long
Dim items As Object

Application.ScreenUpdating = False
  
Set ws = Worksheets("Hoja1")
Cells(1, 48).Value = "Cantidad de lineas" 'Donde va el titulo
            Cells(1, 48).Font.Bold = True 'Donde va el titulo en negrita
    
lastRow = ws.Range("A" & Rows.Count).End(xlUp).Row 'conteo de columna
    
    Set items = CreateObject("Scripting.Dictionary")
    For x = 2 To lastRow
        If Not items.Exists(ws.Cells(x, 47).Value) Then 'columna de conteo columna 1
            items.Add ws.Cells(x, 47).Value, 1 'columna de conteo columna 1
            ws.Cells(x, 48).Value = items(ws.Cells(x, 47).Value) 'columna donde deja = columna de conteo columna 1
        Else
            items(ws.Cells(x, 47).Value) = items(ws.Cells(x, 47).Value) + 1 'columna de conteo columna 1 = columna de conteo columna 1 + 1
            ws.Cells(x, 48).Value = items(ws.Cells(x, 47).Value) 'columna donde deja = columna de conteo columna 1
        End If
    Next x
End Sub
```
-----------------------------------------------------------------------------------------------------------------------------------
``` 
Sub PO_ID()
Application.ScreenUpdating = False 'Apagar el parpadeo de pantalla, Evita los movimientos de pantalla que se producen al seleccionar celdas, hojas y libros
Application.Calculation = xlCalculationManual 'Apagar los cálculos automáticos, Evita que se recalcule todo cada vez que se pegan o modifican datos
Application.EnableEvents = False 'Apagar los eventos automáticos, Evita que se disparen macros de evento si las hubiere
ActiveSheet.DisplayPageBreaks = False 'Apagar visualización de saltos de página, Sirve para evitar algunos problemas de compatibilidad entre macros Excel 2003 vs. 2007/2010


Cells(1, 47).Value = "Concatenado"
            Cells(1, 47).Font.Bold = True 'negrita

For c = 2 To 257
        
        Cells(c, 47) = Cells(c, 46) & Cells(c, 38)
            
Next c

End Sub
```
-----------------------------------------------------------------------------------------------------------------------------------
``` 
Sub instituciones()

Worksheets("Hoja1").Activate

Cells(1, 46).Value = "PO_ID"
            Cells(1, 46).Font.Bold = True 'negrita

For i = 2 To 257
        If Cells(i, 44).Value = "Instituto Profesional AIEP S.A." Then
        Cells(i, 46) = "CHL04"
        ElseIf Cells(i, 44).Value = "UNAB" Then
        Cells(i, 46) = "CHL01"
        ElseIf Cells(i, 44).Value = "Universidad Privada del Norte" Then
        Cells(i, 46) = "PER03"
        ElseIf Cells(i, 44).Value = "Univ. De Viña del Mar Chile OP" Then
        Cells(i, 46) = "CHL32"
        ElseIf Cells(i, 44).Value = "Universidad Perú Ciencias Aplicadas" Then
        Cells(i, 46) = "PER05"
        ElseIf Cells(i, 44).Value = "UDLA Chile" Then
        Cells(i, 46) = "CHL02"
        ElseIf Cells(i, 44).Value = "Cibertec" Then
        Cells(i, 46) = "PER06"
        ElseIf Cells(i, 44).Value = "IEDE Chile" Then
        Cells(i, 46) = "CHL05"
        ElseIf Cells(i, 44).Value = "Inmobiliaria Educ SPA (IESA)" Then
        Cells(i, 46) = "CHL18"
        ElseIf Cells(i, 44).Value = "Laureate Chile II SPA" Then
        Cells(i, 46) = "CHL25"
        ElseIf Cells(i, 44).Value = "Servicios Andinos" Then
        Cells(i, 46) = "CHL28"
        ElseIf Cells(i, 44).Value = "Immob Inversiones SanGenarosDos" Then
        Cells(i, 46) = "CHL31"
        ElseIf Cells(i, 44).Value = "Servicios Profesionales Andrés Bello" Then
        Cells(i, 46) = "CHL08"
        End If
Next i

End Sub
 ```
-----------------------------------------------------------------------------------------------------------------------------------
```
'https://wellsr.com/vba/excel/vba-variable-scope/
Public Function transformar3(fecha As String)
'esta formula tira primero el mes y despues el dia
'https://exceltotal.com/cadenas-de-texto-en-vba/
'InStr(fecha, "/") + 2
On Error Resume Next
If InStr(fecha, "/") > 0 Then
transformar3 = Format(Mid(fecha, InStr(fecha, "/") + 1, 2) & "-" & Mid(fecha, InStr(fecha, "/") - 2, 2) & "-" & Mid(fecha, InStr(fecha, "/") + 4, 4), "dd-mm-yyyy")
Else
transformar3 = Format(Mid(fecha, InStr(fecha, "-") + 1, InStr(fecha, "-") - 1) & "-" & Mid(fecha, 1, InStr(fecha, "-") - 1) & "-" & Mid(fecha, InStr(fecha, "-") + 4, 4), "dd-mm-yyyy")
End If
'MsgBox "Termine de calcular las fechas en matriz" & Ultimate_Column & "x" & Ultimate_Row & " ,ahora falta colocar el mes y generar tabla dinamica"
End Function
```
-----------------------------------------------------------------------------------------------------------------------------------
```
Sub FECHA()
Application.ScreenUpdating = False 'Apagar el parpadeo de pantalla, Evita los movimientos de pantalla que se producen al seleccionar celdas, hojas y libros
Application.Calculation = xlCalculationManual 'Apagar los cálculos automáticos, Evita que se recalcule todo cada vez que se pegan o modifican datos
Application.EnableEvents = False 'Apagar los eventos automáticos, Evita que se disparen macros de evento si las hubiere
ActiveSheet.DisplayPageBreaks = False 'Apagar visualización de saltos de página, Sirve para evitar algunos problemas de compatibilidad entre macros Excel 2003 vs. 2007/2010
Worksheets("Hoja1").Activate

Cells(1, 41).Value = "Fecha"
            Cells(1, 41).Font.Bold = True 'negrita

For c = 2 To 1163
        Cells(c, 41) = "=transformar3(" & "AG" & c & ")"
Next c

End Sub
```
