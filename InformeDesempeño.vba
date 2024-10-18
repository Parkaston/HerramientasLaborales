Sub RellenarPlantillaWordConMarcador()
    Dim ws As Worksheet
    Dim wsObjetivos As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    Dim codigo As String
    Dim nombre As String
    Dim objetivos As String
    Dim rutaArchivoExcel As String
    Dim rutaCarpeta As String
    Dim nombreArchivo As String
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim marcador As Object
    Dim marcadorConseguido As Object
    Dim savePath As String
    Dim plantillaPath As String
    Dim tabla As Object
    Dim tablaConseguido As Object
    Dim filaTabla As Integer
    Dim filaTablaConseguido As Integer

    ' Establecer la hoja de trabajo "CALCULO" y la hoja "Objetivos"
    Set ws = ThisWorkbook.Sheets("CALCULO")
    Set wsObjetivos = ThisWorkbook.Sheets("Objetivos") ' Hoja donde guardas los objetivos

    ' Encontrar la última fila con datos en la columna A (Código de curso)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Obtener la ruta donde está ubicado el archivo Excel
    rutaArchivoExcel = ThisWorkbook.Path

    ' Ruta de la plantilla "Plantilla 0.docx" ubicada en la misma carpeta que el Excel
    plantillaPath = rutaArchivoExcel & "\Plantilla 0.docx"

    ' Crear una instancia de Word
    On Error Resume Next
    Set wordApp = CreateObject("Word.Application")
    On Error GoTo 0

    If wordApp Is Nothing Then
        MsgBox "No se pudo iniciar Microsoft Word."
        Exit Sub
    End If

    ' Abrir el documento de Word existente ("Plantilla 0.docx")
    Set wordDoc = wordApp.Documents.Open(plantillaPath)

    ' Hacer visible Word
    wordApp.Visible = True

    ' Llamar a la función para reemplazar controles de contenido por datos
    Call ReemplazarControlesContenido(wordDoc, ws)

    ' Tabla en el marcador "UnidadesCompetencia"
    On Error Resume Next
    Set marcador = wordDoc.Bookmarks("UnidadesCompetencia").Range
    On Error GoTo 0

    If Not marcador Is Nothing Then
        ' Obtener la tabla después del marcador
        Set tabla = marcador.Tables(1)
        
        ' Iterar sobre cada fila para obtener el código, nombre y objetivos del curso (empezando en la fila 2)
        For currentRow = 2 To lastRow ' Comienza en la fila 2 para evitar los títulos
            codigo = ws.Cells(currentRow, 1).Value
            nombre = ws.Cells(currentRow, 2).Value
            objetivos = ObtenerObjetivosGenerales(codigo, wsObjetivos)

            ' Saltar filas vacías
            If codigo <> "" And nombre <> "" Then
                ' Añadir una nueva fila a la tabla
                filaTabla = tabla.Rows.Count
                tabla.Rows.Add ' Agrega una fila nueva

                ' Escribir el código y nombre del curso en la primera columna de la nueva fila
                tabla.Cell(filaTabla + 1, 1).Range.Text = codigo & " - " & nombre

                ' Escribir los objetivos en la segunda columna de la nueva fila y deshabilitar negrita
                With tabla.Cell(filaTabla + 1, 2).Range
                    .Text = objetivos
                    .Font.Bold = False ' Desactivar negrita
                End With
            End If
        Next currentRow
    Else
        MsgBox "El marcador 'UnidadesCompetencia' no se encontró en el documento."
        wordDoc.Close False
        wordApp.Quit
        Exit Sub
    End If
    
' Tabla en el marcador "UnidadesCompetenciaConseguido"
On Error Resume Next
Set marcadorConseguido = wordDoc.Bookmarks("UnidadesCompetenciaConseguido").Range
On Error GoTo 0

' Verificar si el marcador existe y contiene una tabla
If Not marcadorConseguido Is Nothing Then
    ' Asegurarse de que hay una tabla en el marcador
    If marcadorConseguido.Tables.Count > 0 Then
        Set tablaConseguido = marcadorConseguido.Tables(1)

        ' Iterar sobre cada fila para obtener el código, nombre y objetivos del curso
        For currentRow = 2 To lastRow ' Comienza en la fila 2 para evitar los títulos
            codigo = ws.Cells(currentRow, 1).Value
            nombre = ws.Cells(currentRow, 2).Value
            objetivos = ObtenerObjetivosGenerales(codigo, wsObjetivos)

            ' Saltar filas vacías
            If codigo <> "" And nombre <> "" Then
                ' Añadir una nueva fila a la tabla
                tablaConseguido.Rows.Add ' Agrega una fila nueva
                filaTablaConseguido = tablaConseguido.Rows.Count ' Contar las filas para obtener la última agregada

                ' Insertar el contenido en la fila recién añadida usando InsertAfter y formato visible
                ' Primera columna: Código y Nombre del curso
                With tablaConseguido.Cell(filaTablaConseguido, 1).Range
                    .Text = codigo & " - " & nombre
                    .Font.Color = wdColorBlack ' Asegurar que el texto sea visible (negro)
                    .Font.Size = 11 ' Tamaño de fuente estándar
                End With

                ' Segunda columna: Objetivos generales, sin negrita
                With tablaConseguido.Cell(filaTablaConseguido, 2).Range
                    .Text = objetivos
                    .Font.Color = wdColorBlack ' Asegurar que el texto sea visible (negro)
                    .Font.Size = 11 ' Tamaño de fuente estándar
                    .Font.Bold = False ' Quitar negrita
                End With

                ' Tercera columna: Vacía
                With tablaConseguido.Cell(filaTablaConseguido, 3).Range
                    .Text = ""
                    .Font.Color = wdColorBlack ' Asegurar que no haya formato oculto
                    .Font.Size = 11 ' Tamaño de fuente estándar
                End With
            End If
        Next currentRow

        ' Forzar el refresco de campos visibles en el documento
        wordDoc.Fields.Update
        wordDoc.Application.ScreenRefresh
        
    Else
        MsgBox "No se encontró una tabla en el marcador 'UnidadesCompetenciaConseguido'."
    End If
Else
    MsgBox "El marcador 'UnidadesCompetenciaConseguido' no se encontró en el documento."
    wordDoc.Close False
    wordApp.Quit
    Exit Sub
End If




    ' Pedirle al usuario un nombre para el archivo con un emoji en el mensaje
    nombreArchivo = InputBox("Ingrese el nombre del archivo (sin extensión): ??")

    ' Crear la carpeta "Archivos de salida" si no existe
    rutaCarpeta = rutaArchivoExcel & "\Archivos de salida"
    If Dir(rutaCarpeta, vbDirectory) = "" Then
        MkDir rutaCarpeta
    End If

    ' Definir la ruta completa donde se guardará el archivo
    savePath = rutaCarpeta & "\" & nombreArchivo & ".docx"

    ' Guardar el documento de Word
    wordDoc.SaveAs2 Filename:=savePath, FileFormat:=16 ' wdFormatDocumentDefault (Word 2010 y superior)

    ' Cerrar el documento de Word
    wordDoc.Close

    ' Cerrar la aplicación de Word
    wordApp.Quit

    ' Informar al usuario que el archivo se ha guardado en: " & savePath"
    MsgBox "El archivo se ha guardado en: " & savePath

End Sub

' Función para reemplazar los controles de contenido en el documento de Word
Sub ReemplazarControlesContenido(doc As Object, ws As Worksheet)
    Dim cc As Object
    Dim mapeoCC As Object
    Set mapeoCC = CreateObject("Scripting.Dictionary")

    ' Añadir el mapeo de los controles de contenido a las celdas de Excel
    mapeoCC.Add "ApellidoAlumno", ws.Range("K28").Value
    mapeoCC.Add "NombreAlumno", ws.Range("K12").Value
    mapeoCC.Add "NacimientoAlumno", ws.Range("K14").Value
    mapeoCC.Add "DniAlumno", ws.Range("K13").Value
    mapeoCC.Add "TelefonoAlumno", ws.Range("K29").Value
    mapeoCC.Add "DireccionAlumno", ws.Range("K30").Value
    mapeoCC.Add "PoblacionAlumno", ws.Range("K31").Value
    mapeoCC.Add "ProvinciaAlumno", ws.Range("K32").Value
    mapeoCC.Add "CPAlumno", ws.Range("K33").Value
    mapeoCC.Add "FamiliaProfesional", ws.Range("K34").Value
    mapeoCC.Add "Tutor", ws.Range("F2").Value
    mapeoCC.Add "NombreEmpresa", ws.Range("K1").Value
    mapeoCC.Add "DireccionEmpresa", ws.Range("K35").Value
    mapeoCC.Add "PoblacionEmpresa", ws.Range("K36").Value
    mapeoCC.Add "ProvinciaEmpresa", ws.Range("K37").Value
    mapeoCC.Add "CPEmpresa", ws.Range("K38").Value
    mapeoCC.Add "TelefonoEmpresa", ws.Range("K7").Value
    mapeoCC.Add "EmailEmpresa", ws.Range("K6").Value
    mapeoCC.Add "TutorEmpresa", ws.Range("K8").Value
    mapeoCC.Add "TelefonoTutorEmpresa", ws.Range("K39").Value

    ' Iterar sobre los controles de contenido del documento de Word
    For Each cc In doc.ContentControls
        If mapeoCC.Exists(cc.Title) Then
            ' Reemplazar el texto del control de contenido por el valor de la celda correspondiente
            cc.Range.Text = mapeoCC(cc.Title)
        End If
    Next cc
End Sub

' Función para obtener los objetivos generales según el código del curso
Function ObtenerObjetivosGenerales(codigoCurso As String, wsObjetivos As Worksheet) As String
    Dim lastRowObjetivos As Long
    Dim currentRowObjetivos As Long
    Dim codigoObjetivo As String
    Dim objetivos As String

    ' Encontrar la última fila con datos en la hoja "Objetivos"
    lastRowObjetivos = wsObjetivos.Cells(wsObjetivos.Rows.Count, 1).End(xlUp).Row

    ' Iterar sobre la hoja de objetivos para encontrar el código y devolver los objetivos
    For currentRowObjetivos = 2 To lastRowObjetivos ' Asume que los datos empiezan en la fila 2
        codigoObjetivo = wsObjetivos.Cells(currentRowObjetivos, 1).Value ' Código de la materia en la columna A
        If codigoObjetivo = codigoCurso Then
            ' Objetivos generales en la columna B
            objetivos = wsObjetivos.Cells(currentRowObjetivos, 2).Value
            ObtenerObjetivosGenerales = objetivos
            Exit Function
        End If
    Next currentRowObjetivos

    ' Si no encuentra los objetivos, devuelve un mensaje de error
    ObtenerObjetivosGenerales = "No se encontraron objetivos para este curso"
End Function