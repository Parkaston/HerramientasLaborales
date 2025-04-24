' ------------------------------------------------------------------------------
' Script: ActualizarPlanFormativo
' Autor: Guillermo Luna Alvarez
' Descripción: Este macro combina múltiples documentos Word según una hoja Excel,
'              pegando contenidos en un documento base y actualizando fechas específicas.
'
' Este archivo no contiene datos personales. Todos los valores provienen
' de archivos externos (Excel/Word), y deben ser anonimizados si se usan
' para pruebas públicas. No subir archivos adjuntos reales a repositorios públicos.
'
' Requiere:
' - Word con controles de contenido ("FechaInicio", "FechaFin")
' - Carpeta "Plantillas" con documentos .docx nombrados como se indica en la hoja Excel
' - Carpeta "Archivos de salida" para guardar el documento final
' ------------------------------------------------------------------------------



Sub ActualizarPlanFormativo()
    Dim ws As Worksheet
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim plantillaDoc As Object
    Dim lastRow As Long
    Dim nombreArchivo As String
    Dim fechaInicio As String
    Dim fechaFin As String
    Dim wordFilePath As String
    Dim plantillaPath As String
    Dim salidaPath As String
    Dim nombreArchivoSalida As String
    Dim i As Long
    Dim cc As Object
    Dim rng As Object
    Dim nuevaSeccion As Object

    ' Configurar hoja de cálculo y aplicación de Word
    Set ws = ThisWorkbook.Sheets("CALCULO") '
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = True
    
    plantillaPath = ThisWorkbook.Path & "\Plantillas\"
    salidaPath = ThisWorkbook.Path & "\Archivos de salida\"

    ' Verificar si el archivo base existe
    If Dir(ThisWorkbook.Path & "\PlanFormativoBase.docx") = "" Then
        MsgBox "El archivo PlanFormativoBase.docx no se encontró.", vbExclamation
        Exit Sub
    End If

    ' Crear la carpeta de salida si no existe
    If Dir(salidaPath, vbDirectory) = "" Then
        MkDir salidaPath
    End If

    ' Abrir el documento base
    Set plantillaDoc = wordApp.Documents.Open(ThisWorkbook.Path & "\PlanFormativoBase.docx")

    ' Obtener la última fila con datos en la hoja
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Iterar sobre cada fila de la hoja
    For i = 2 To lastRow
        ' Obtener el nombre del archivo y las fechas
        nombreArchivo = ws.Cells(i, 1).Value
        fechaInicio = ws.Cells(i, 3).Value
        fechaFin = ws.Cells(i, 4).Value

        ' Construir la ruta del archivo de Word correspondiente
        wordFilePath = plantillaPath & nombreArchivo & ".docx"

        ' Verificar si el archivo de Word existe
        If Dir(wordFilePath) <> "" Then
            ' Abrir el documento de Word de la plantilla
            Set wordDoc = wordApp.Documents.Open(wordFilePath)

            ' Mover el cursor al final del documento base antes de pegar
            Set rng = plantillaDoc.Content
            rng.Collapse Direction:=0 ' wdCollapseEnd
            rng.InsertAfter vbCrLf ' Insertar un salto de línea antes de pegar el contenido
            rng.Collapse Direction:=0 ' wdCollapseEnd nuevamente

            ' Copiar y pegar el contenido de la plantilla en el documento base
            wordDoc.Content.Copy
            rng.Paste

            ' Reemplazar el contenido del control de contenido "FechaInicio" y "FechaFin" en la sección recién pegada
            For Each cc In plantillaDoc.ContentControls
                If cc.Range.InRange(rng) Then ' Asegurarse de que el control de contenido esté en la sección recién pegada
                    If cc.Title = "FechaInicio" Then
                        cc.Range.Text = fechaInicio
                    ElseIf cc.Title = "FechaFin" Then
                        cc.Range.Text = fechaFin
                    End If
                End If
            Next cc

            ' Cerrar el documento de Word de la plantilla
            wordDoc.Close False
        Else
            MsgBox "El archivo para " & nombreArchivo & " no se encontró.", vbExclamation
        End If
    Next i

    ' Preguntar al usuario por el nombre del archivo de salida
    nombreArchivoSalida = InputBox("Introduce el nombre para el archivo de salida:", "Guardar como", "PlanFormativoPersonalizado")

    ' Guardar el documento base con el nuevo nombre en la carpeta de salida
    If nombreArchivoSalida <> "" Then
        plantillaDoc.SaveAs2 salidaPath & nombreArchivoSalida & ".docx"
    End If

    ' Cerrar el documento base y la aplicación de Word
    plantillaDoc.Close False
    wordApp.Quit
    Set plantillaDoc = Nothing
    Set wordApp = Nothing

    MsgBox "Proceso completado."
End Sub
