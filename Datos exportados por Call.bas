' Variable global para acumular mensajes de estado
Dim logMensajes As String

Sub EjecutarProcesamientoKACE()
    ' Inicializar el log
    logMensajes = "Resultados del proceso:" & vbCrLf

    ' Ejecutar cada paso del procesamiento
    Call ImportarDatosDesdeDescargas
    Call EliminarColumnaPrioridad
    Call ReemplazoFechas
    Call ReemplazoTick
    Call ReemplazoEstado
    Call ReemplazoRemitente
    Call AjustarFormatoColumnas


    ' Mostrar el resultado final al usuario
    MsgBox logMensajes, vbInformation, "Proceso Completado"
End Sub

Sub ImportarDatosDesdeDescargas()
    Dim wbDestino As Workbook
    Dim wsDestino As Worksheet
    Dim wbOrigen As Workbook
    Dim rutaDescargas As String
    Dim archivoOrigen As String
    Dim rutaArchivo As String

    ' Ruta de Descargas
    rutaDescargas = "C:\Users\jsolis\Downloads\"
    archivoOrigen = "export_list.xlsx"
    rutaArchivo = rutaDescargas & archivoOrigen

    ' Verificar si el archivo existe
    If Dir(rutaArchivo) = "" Then
        logMensajes = logMensajes & "- Archivo de origen no encontrado: " & archivoOrigen & vbCrLf
        Exit Sub
    End If

    ' Desactivar actualización de pantalla
    Application.ScreenUpdating = False

    ' Abrir el archivo de origen en segundo plano
    Set wbOrigen = Workbooks.Open(rutaArchivo, ReadOnly:=True)

    ' Hoja de destino
    Set wbDestino = ThisWorkbook
    Set wsDestino = wbDestino.Sheets("Importados")

    ' Limpiar datos previos en la hoja de destino
    wsDestino.Cells.Clear

    ' Copiar datos desde la primera hoja del archivo origen
    wbOrigen.Sheets(1).UsedRange.Copy wsDestino.Range("A1")

    ' Cerrar archivo de origen
    wbOrigen.Close False

    ' Reactivar actualización de pantalla
    Application.ScreenUpdating = True

    ' Registrar éxito
    logMensajes = logMensajes & "- Datos importados correctamente desde: " & archivoOrigen & vbCrLf
End Sub

Sub EliminarColumnaPrioridad()
    Dim wsExport As Worksheet
    Dim rng As Range
    Dim colPrioridad As Long

    ' Hoja de datos importados
    Set wsExport = ThisWorkbook.Sheets("Importados")

    ' Buscar y eliminar la columna "Prioridad"
    Set rng = wsExport.Rows(1).Find(What:="Prioridad", LookAt:=xlWhole)
    If Not rng Is Nothing Then
        colPrioridad = rng.Column
        wsExport.Columns(colPrioridad).Delete
        logMensajes = logMensajes & "- Columna 'Prioridad' eliminada." & vbCrLf
    Else
        logMensajes = logMensajes & "- Columna 'Prioridad' no encontrada." & vbCrLf
    End If
End Sub

Sub ReemplazoFechas()
    Dim wsExport As Worksheet
    Dim lastRow As Long
    Dim colCreado As Long, colVencimiento As Long, colModificado As Long
    Dim i As Long
    Dim celda As Range
    Dim fecha As Date
    
    ' Hoja de datos importados
    Set wsExport = ThisWorkbook.Sheets("Importados")
    
    ' Encontrar la última fila con datos
    lastRow = wsExport.Cells(wsExport.Rows.Count, "A").End(xlUp).Row
    
    ' Buscar las columnas para Creado, Vencimiento y Modificado
    colCreado = BuscarColumna(wsExport, "Creado")
    colVencimiento = BuscarColumna(wsExport, "Vencimiento")
    colModificado = BuscarColumna(wsExport, "Modificado")
    
    ' Procesar la columna "Creado"
    If colCreado > 0 Then
        For i = 2 To lastRow
            Set celda = wsExport.Cells(i, colCreado)
            If IsDate(celda.Value) Then
                ' Asegurar formato consistente
                celda.Value = Format(celda.Value, "dd/mm/yyyy")
            End If
        Next i
    End If
    
    ' Procesar la columna "Vencimiento"
    If colVencimiento > 0 Then
        For i = 2 To lastRow
            Set celda = wsExport.Cells(i, colVencimiento)
            If IsDate(celda.Value) Then
                ' Asegurar formato consistente
                celda.Value = Format(celda.Value, "dd/mm/yyyy")
            End If
        Next i
    End If
    
    ' Procesar la columna "Modificado"
    If colModificado > 0 Then
        For i = 2 To lastRow
            Set celda = wsExport.Cells(i, colModificado)
            If IsDate(celda.Value) Then
                ' Asegurar formato consistente
                celda.Value = Format(celda.Value, "dd/mm/yyyy")
            End If
        Next i
    End If
End Sub


Sub ReemplazoTick()
    Call ReemplazoColumna("Número", "TICK:", "")
    logMensajes = logMensajes & "- 'TICK:' reemplazado correctamente en la columna 'Número'." & vbCrLf
End Sub

Sub ReemplazoEstado()
    Call ReemplazoColumna("Estado", "Nuevo", "Abierto")
    logMensajes = logMensajes & "- Estado 'Nuevo' reemplazado por 'Abierto'." & vbCrLf
End Sub

Sub ReemplazoRemitente()
    Dim wsExport As Worksheet
    Dim colRemitente As Long

    ' Hoja de datos importados
    Set wsExport = ThisWorkbook.Sheets("Importados")

    ' Buscar columna "Remitente"
    colRemitente = BuscarColumna(wsExport, "Remitente")
    If colRemitente > 0 Then
        With wsExport
            .Range(.Cells(2, colRemitente), .Cells(.Rows.Count, colRemitente).End(xlUp)).Replace What:="Recepción", Replacement:="Tienda", LookAt:=xlPart
            .Range(.Cells(2, colRemitente), .Cells(.Rows.Count, colRemitente).End(xlUp)).Replace What:="Receptor", Replacement:="Tienda", LookAt:=xlPart
            .Range(.Cells(2, colRemitente), .Cells(.Rows.Count, colRemitente).End(xlUp)).Replace What:="Tesorería", Replacement:="Tienda", LookAt:=xlPart
        End With
        logMensajes = logMensajes & "- Reemplazos en columna 'Remitente' completados." & vbCrLf
    Else
        logMensajes = logMensajes & "- Columna 'Remitente' no encontrada." & vbCrLf
    End If
End Sub

' ---- Funciones Reutilizables ----

Sub ReemplazoColumna(columnaNombre As String, buscar As String, reemplazar As String)
    Dim wsExport As Worksheet
    Dim lastRow As Long
    Dim colIndex As Long

    ' Hoja de datos importados
    Set wsExport = ThisWorkbook.Sheets("Importados")

    ' Buscar la columna por su encabezado
    colIndex = BuscarColumna(wsExport, columnaNombre)
    If colIndex > 0 Then
        With wsExport
            lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
            .Range(.Cells(2, colIndex), .Cells(lastRow, colIndex)).Replace What:=buscar, Replacement:=reemplazar, LookAt:=xlPart
        End With
    Else
        logMensajes = logMensajes & "- Columna '" & columnaNombre & "' no encontrada." & vbCrLf
    End If
End Sub

Function BuscarColumna(ws As Worksheet, encabezado As String) As Long
    Dim rng As Range
    Set rng = ws.Rows(1).Find(What:=encabezado, LookAt:=xlWhole)
    If Not rng Is Nothing Then
        BuscarColumna = rng.Column
    Else
        BuscarColumna = 0
    End If
End Function

Sub AjustarFormatoColumnas()
    Dim wsExport As Worksheet
    Dim lastRow As Long
    Dim colCreado As Long, colVencimiento As Long, colModificado As Long
    Dim colNumero As Long, colTitulo As Long, colEstado As Long, colRemitente As Long, colPropietario As Long
    Dim rng As Range
    
    ' Establecer la hoja de trabajo
    Set wsExport = ThisWorkbook.Sheets("Importados")
    
    ' Encontrar la última fila con datos
    lastRow = wsExport.Cells(wsExport.Rows.Count, "A").End(xlUp).Row
    
    ' Encontrar las columnas por sus encabezados
    colCreado = BuscarColumna(wsExport, "Creado")
    colVencimiento = BuscarColumna(wsExport, "Vencimiento")
    colModificado = BuscarColumna(wsExport, "Modificado")
    colNumero = BuscarColumna(wsExport, "Número")
    colTitulo = BuscarColumna(wsExport, "Título")
    colEstado = BuscarColumna(wsExport, "Estado")
    colRemitente = BuscarColumna(wsExport, "Remitente")
    colPropietario = BuscarColumna(wsExport, "Propietario")
    
    ' Ajustar alineación y formato para todas las columnas
    With wsExport
        ' Centrar datos para todas las columnas excepto "Título"
        Set rng = .Range("A2", .Cells(lastRow, .Cells(1, .Columns.Count).End(xlToLeft).Column))
        rng.HorizontalAlignment = xlCenter
        
        ' Excluir "Título" de la alineación centrada
        If colTitulo > 0 Then
            .Columns(colTitulo).HorizontalAlignment = xlLeft
        End If
        
        ' Formatear columnas de fechas
        If colCreado > 0 Then .Columns(colCreado).NumberFormat = "dd/mm/yyyy"
        If colVencimiento > 0 Then .Columns(colVencimiento).NumberFormat = "dd/mm/yyyy"
        If colModificado > 0 Then .Columns(colModificado).NumberFormat = "dd/mm/yyyy"
        
        ' Formatear "Número" como número
        If colNumero > 0 Then .Columns(colNumero).NumberFormat = "0"
        
        ' Formatear "Título", "Estado", "Remitente" y "Propietario" como general
        If colTitulo > 0 Then .Columns(colTitulo).NumberFormat = "General"
        If colEstado > 0 Then .Columns(colEstado).NumberFormat = "General"
        If colRemitente > 0 Then .Columns(colRemitente).NumberFormat = "General"
        If colPropietario > 0 Then .Columns(colPropietario).NumberFormat = "General"
        
        ' Ajustar el tamaño de las columnas
        .Columns.AutoFit
        
        ' Excluir "Título" del ajuste de tamaño
        If colTitulo > 0 Then
            .Columns(colTitulo).ColumnWidth = 30 ' Define un ancho fijo para "Título"
        End If
    End With
End Sub

