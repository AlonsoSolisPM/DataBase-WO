Sub EliminarDuplicadosMasViejos1()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim ticket1 As String, ticket2 As String
    Dim fecha1 As Date, fecha2 As Date
    Dim categoria1 As String, categoria2 As String
    Dim categoriaConservada As String
    
    ' Define la hoja de trabajo
    Set ws = ThisWorkbook.Sheets("Todas las tiendas") ' Cambia "Hoja1" al nombre de tu hoja

    ' Encuentra la última fila con datos
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    ' Recorre cada fila
    For i = 2 To lastRow
        ticket1 = ws.Cells(i, 5).Value ' Columna Ticket
        fecha1 = ws.Cells(i, 11).Value ' Columna Modificado
        categoria1 = Trim(ws.Cells(i, 14).Value) ' Columna Categoría
        
        ' Compara con todas las filas siguientes
        For j = i + 1 To lastRow
            ticket2 = ws.Cells(j, 5).Value
            fecha2 = ws.Cells(j, 11).Value
            categoria2 = Trim(ws.Cells(j, 14).Value)
            
            ' Si se encuentra un duplicado
            If ticket1 = ticket2 Then
                ' Compara las fechas
                If fecha1 < fecha2 Then
                    ' Si la primera fecha es más antigua, elimina la fila i
                    ws.Rows(i).Delete
                    i = i - 1
                    Exit For
                ElseIf fecha2 < fecha1 Then
                    ' Si la segunda fecha es más antigua, elimina la fila j
                    ws.Rows(j).Delete
                    j = j - 1
                Else
                    ' Si las fechas son iguales, conserva la categoría de la fila con la fecha más reciente
                    If categoria1 = "" And categoria2 <> "" Then
                        ws.Cells(i, 14).Value = categoria2
                        ws.Rows(i).Delete
                        i = i - 1
                    ElseIf categoria2 = "" And categoria1 <> "" Then
                        ws.Cells(j, 14).Value = categoria1
                        ws.Rows(j).Delete
                        j = j - 1
                    Else
                        ' Si ambas categorías están vacías o llenas, elimina la fila j
                        ws.Rows(j).Delete
                        j = j - 1
                    End If
                End If
                ' Actualiza la última fila
                lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
                Exit For
            End If
        Next j
    Next i
    
    MsgBox "Proceso completado."
End Sub


Sub EliminarDuplicadosMasViejos2()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim ticket1 As String, ticket2 As String
    Dim fecha1 As Date, fecha2 As Date
    Dim categoria1 As String, categoria2 As String
    
    ' Define la hoja de trabajo
    Set ws = ThisWorkbook.Sheets("Todas las tiendas") ' Cambia "Todas las tiendas" al nombre de tu hoja

    ' Encuentra la última fila con datos
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    ' Recorre cada fila
    For i = 2 To lastRow
        ticket1 = ws.Cells(i, 5).Value ' Columna Ticket
        fecha1 = ws.Cells(i, 11).Value ' Columna Modificado
        categoria1 = Trim(ws.Cells(i, 14).Value) ' Columna Categoría
        
        ' Compara con todas las filas siguientes
        For j = i + 1 To lastRow
            ticket2 = ws.Cells(j, 5).Value
            fecha2 = ws.Cells(j, 11).Value
            categoria2 = Trim(ws.Cells(j, 14).Value)
            
            ' Si se encuentra un duplicado
            If ticket1 = ticket2 Then
                ' Compara las fechas
                If fecha1 < fecha2 Then
                    ' Si la primera fecha es más antigua, conserva la segunda
                    ' Si la categoría de la fila más nueva está vacía, rellena con la más antigua si está disponible
                    If categoria2 = "" And categoria1 <> "" Then
                        ws.Cells(j, 14).Value = categoria1
                    End If
                    ' Elimina la fila más antigua
                    ws.Rows(i).Delete
                    i = i - 1
                    Exit For
                ElseIf fecha2 < fecha1 Then
                    ' Si la segunda fecha es más antigua, conserva la primera
                    ' Si la categoría de la fila más reciente está vacía, rellena con la más antigua si está disponible
                    If categoria1 = "" And categoria2 <> "" Then
                        ws.Cells(i, 14).Value = categoria2
                    End If
                    ' Elimina la fila más antigua
                    ws.Rows(j).Delete
                    j = j - 1
                Else
                    ' Si las fechas son iguales, conserva una y gestiona las categorías
                    If categoria1 = "" And categoria2 <> "" Then
                        ws.Cells(i, 14).Value = categoria2
                    ElseIf categoria2 = "" And categoria1 <> "" Then
                        ws.Cells(j, 14).Value = categoria1
                    End If
                    ' Elimina una de las filas duplicadas
                    ws.Rows(j).Delete
                    j = j - 1
                End If
                ' Actualiza la última fila
                lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
                Exit For
            End If
        Next j
    Next i
    
    MsgBox "Proceso completado."
End Sub

