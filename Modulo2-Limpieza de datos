Sub ImportarDatosDesdeDescargas()
    Dim wbDestino As Workbook
    Dim wsDestino As Worksheet
    Dim wbOrigen As Workbook
    Dim rutaDescargas As String
    Dim archivoOrigen As String
    Dim rutaArchivo As String

    ' Establece la ruta de la carpeta de Descargas
    rutaDescargas = "C:\Users\jsolis\Downloads\"
    
    ' Nombre del archivo que quieres copiar (export_list.xlsx)
    archivoOrigen = "export_list.xlsx"
    
    ' Crear la ruta completa del archivo
    rutaArchivo = rutaDescargas & archivoOrigen
    
    ' Verificar si el archivo existe
    If Dir(rutaArchivo) = "" Then
        MsgBox "El archivo " & archivoOrigen & " no se encuentra en la carpeta de Descargas.", vbExclamation
        Exit Sub
    End If
    
    ' Abrir el archivo de origen
    Set wbOrigen = Workbooks.Open(rutaArchivo)
    
    ' Establecer la hoja de destino en el libro actual
    Set wbDestino = ThisWorkbook ' Suponiendo que este es "1.Todas las Tiendas - Registro"
    Set wsDestino = wbDestino.Sheets("Importados") ' Hoja de destino donde pegarás los datos
    
    ' Limpiar cualquier dato previo en la hoja Worksheet antes de pegar los nuevos datos
    wsDestino.Cells.Clear
    
    ' Copiar el contenido de la primera hoja del archivo de origen
    wbOrigen.Sheets(1).UsedRange.Copy wsDestino.Range("A1")
    
    ' Cerrar el archivo de origen
    wbOrigen.Close False
    
    MsgBox "Datos importados correctamente desde " & archivoOrigen & "."
End Sub


Sub EliminarColumnaPrioridad()
    Dim wsExport As Worksheet
    Dim rng As Range
    Dim colPrioridad As Long
    
    ' Establecer la hoja exportada
    Set wsExport = ThisWorkbook.Sheets("Importados")
    
    ' Encontrar la columna con el encabezado "Prioridad" y eliminarla
    Set rng = wsExport.Rows(1).Find(What:="Prioridad", LookAt:=xlWhole)
    If Not rng Is Nothing Then
        colPrioridad = rng.Column
        wsExport.Columns(colPrioridad).Delete
    End If
    
    MsgBox "Columna 'Prioridad' eliminada."
End Sub

Sub ReemplazoFechas()
    Dim wsExport As Worksheet
    Dim lastRow As Long
    Dim rng As Range
    Dim colCreado As Long, colVencimiento As Long, colModificado As Long
    
    ' Establecer la hoja exportada
    Set wsExport = ThisWorkbook.Sheets("Importados")
    
    ' Encontrar la última fila con datos
    lastRow = wsExport.Cells(wsExport.Rows.Count, "A").End(xlUp).Row
    
    ' Buscar las columnas para Creado, Vencimiento y Modificado
    Set rng = wsExport.Rows(1).Find(What:="Creado", LookAt:=xlWhole)
    If Not rng Is Nothing Then colCreado = rng.Column
    
    Set rng = wsExport.Rows(1).Find(What:="Vencimiento", LookAt:=xlWhole)
    If Not rng Is Nothing Then colVencimiento = rng.Column
    
    Set rng = wsExport.Rows(1).Find(What:="Modificado", LookAt:=xlWhole)
    If Not rng Is Nothing Then colModificado = rng.Column
    
    ' Reemplazo en las columnas de fecha
    If colCreado > 0 Then wsExport.Range(wsExport.Cells(2, colCreado), wsExport.Cells(lastRow, colCreado)).Replace What:="/", Replacement:="b", LookAt:=xlPart
    If colVencimiento > 0 Then wsExport.Range(wsExport.Cells(2, colVencimiento), wsExport.Cells(lastRow, colVencimiento)).Replace What:="/", Replacement:="b", LookAt:=xlPart
    If colModificado > 0 Then wsExport.Range(wsExport.Cells(2, colModificado), wsExport.Cells(lastRow, colModificado)).Replace What:="/", Replacement:="b", LookAt:=xlPart
    
    If colCreado > 0 Then wsExport.Range(wsExport.Cells(2, colCreado), wsExport.Cells(lastRow, colCreado)).Replace What:="b", Replacement:="/", LookAt:=xlPart
    If colVencimiento > 0 Then wsExport.Range(wsExport.Cells(2, colVencimiento), wsExport.Cells(lastRow, colVencimiento)).Replace What:="b", Replacement:="/", LookAt:=xlPart
    If colModificado > 0 Then wsExport.Range(wsExport.Cells(2, colModificado), wsExport.Cells(lastRow, colModificado)).Replace What:="b", Replacement:="/", LookAt:=xlPart
    
    MsgBox "Reemplazo en las columnas de fecha completado."
End Sub

Sub ReemplazoTick()
    Dim wsExport As Worksheet
    Dim lastRow As Long
    Dim rng As Range
    Dim colNumero As Long
    
    ' Establecer la hoja exportada
    Set wsExport = ThisWorkbook.Sheets("Importados")
    
    ' Encontrar la última fila con datos
    lastRow = wsExport.Cells(wsExport.Rows.Count, "A").End(xlUp).Row
    
    ' Buscar la columna "Número"
    Set rng = wsExport.Rows(1).Find(What:="Número", LookAt:=xlWhole)
    If Not rng Is Nothing Then
        colNumero = rng.Column
        wsExport.Range(wsExport.Cells(2, colNumero), wsExport.Cells(lastRow, colNumero)).Replace What:="TICK:", Replacement:="", LookAt:=xlPart
    End If
    
    MsgBox "Reemplazo de 'TICK:' en la columna Número completado."
End Sub

Sub ReemplazoEstado()
    Dim wsExport As Worksheet
    Dim lastRow As Long
    Dim rng As Range
    Dim colEstado As Long
    
    ' Establecer la hoja exportada
    Set wsExport = ThisWorkbook.Sheets("Importados")
    
    ' Encontrar la última fila con datos
    lastRow = wsExport.Cells(wsExport.Rows.Count, "A").End(xlUp).Row
    
    ' Buscar la columna "Estado"
    Set rng = wsExport.Rows(1).Find(What:="Estado", LookAt:=xlWhole)
    If Not rng Is Nothing Then
        colEstado = rng.Column
        wsExport.Range(wsExport.Cells(2, colEstado), wsExport.Cells(lastRow, colEstado)).Replace What:="Nuevo", Replacement:="Abierto", LookAt:=xlPart
    End If
    
    MsgBox "Reemplazo de 'Nuevo' por 'Abierto' en la columna Estado completado."
End Sub

Sub ReemplazoRemitente()
    Dim wsExport As Worksheet
    Dim lastRow As Long
    Dim rng As Range
    Dim colRemitente As Long
    
    ' Establecer la hoja exportada
    Set wsExport = ThisWorkbook.Sheets("Importados")
    
    ' Encontrar la última fila con datos
    lastRow = wsExport.Cells(wsExport.Rows.Count, "A").End(xlUp).Row
    
    ' Buscar la columna "Remitente"
    Set rng = wsExport.Rows(1).Find(What:="Remitente", LookAt:=xlWhole)
    If Not rng Is Nothing Then
        colRemitente = rng.Column
        wsExport.Range(wsExport.Cells(2, colRemitente), wsExport.Cells(lastRow, colRemitente)).Replace What:="Recepción", Replacement:="Tienda", LookAt:=xlPart
        wsExport.Range(wsExport.Cells(2, colRemitente), wsExport.Cells(lastRow, colRemitente)).Replace What:="Receptor", Replacement:="Tienda", LookAt:=xlPart
        wsExport.Range(wsExport.Cells(2, colRemitente), wsExport.Cells(lastRow, colRemitente)).Replace What:="Tesorería", Replacement:="Tienda", LookAt:=xlPart
    End If
    
    MsgBox "Reemplazos en la columna Remitente completados."
End Sub

