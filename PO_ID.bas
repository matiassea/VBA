Attribute VB_Name = "M�dulo1"
Sub PO_ID()

Application.ScreenUpdating = False 'Apagar el parpadeo de pantalla, Evita los movimientos de pantalla que se producen al seleccionar celdas, hojas y libros
Application.Calculation = xlCalculationManual 'Apagar los c�lculos autom�ticos, Evita que se recalcule todo cada vez que se pegan o modifican datos
Application.EnableEvents = False 'Apagar los eventos autom�ticos, Evita que se disparen macros de evento si las hubiere
ActiveSheet.DisplayPageBreaks = False 'Apagar visualizaci�n de saltos de p�gina, Sirve para evitar algunos problemas de compatibilidad entre macros Excel 2003 vs. 2007/2010


Cells(1, 47).Value = "Concatenado"
            Cells(1, 47).Font.Bold = True 'negrita

For c = 2 To 257
        
        Cells(c, 47) = Cells(c, 46) & Cells(c, 38)
            
Next c

End Sub

