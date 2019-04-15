Attribute VB_Name = "Módulo2"
Sub PO_ID()
Dim lastRow As Long

Application.ScreenUpdating = False 'Apagar el parpadeo de pantalla, Evita los movimientos de pantalla que se producen al seleccionar celdas, hojas y libros
Application.Calculation = xlCalculationManual 'Apagar los cálculos automáticos, Evita que se recalcule todo cada vez que se pegan o modifican datos
Application.EnableEvents = False 'Apagar los eventos automáticos, Evita que se disparen macros de evento si las hubiere
ActiveSheet.DisplayPageBreaks = False 'Apagar visualización de saltos de página, Sirve para evitar algunos problemas de compatibilidad entre macros Excel 2003 vs. 2007/2010


Set ws = Worksheets("Hoja1")
lastRow = ws.Range("A" & Rows.Count).End(xlUp).Row 'conteo de columna

Cells(1, 47).Value = "Concatenado"
            Cells(1, 47).Font.Bold = True 'negrita

For c = 2 To lastRow
        
        Cells(c, 47) = Cells(c, 46) & Cells(c, 38)
            
Next c

End Sub

