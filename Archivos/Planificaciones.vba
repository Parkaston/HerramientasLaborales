Sub InsertarDatosDesdeExcelEnWord()
    
    ' Declaración de variables
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim tbl As Object
    Dim xlSheet As Worksheet
    Dim i As Integer, lastRow As Integer
    Dim codigo As String, denominacion As String, fechaInicio As String, fechaFin As String
    Dim nombreArchivo As String
    
    ' Ruta fija del archivo Word
    Dim pathWord As String
    pathWord = "C:\Users\ContratoFor\Desktop\Pruebas pdf\Pruebas pdf\Archivos\Planibase.docx" ' Cambia esta ruta a la ubicación del archivo Word fijo

    ' Carpeta de salida predeterminada
    Dim outputFolder As String
    outputFolder = "C:\Users\ContratoFor\Desktop\Pruebas pdf\Pruebas pdf\Archivos\Archivos de salida\" ' Cambia esta ruta a la carpeta de salida deseada

    ' Verifica si la carpeta de salida existe; si no, la crea
    If Dir(outputFolder, vbDirectory) = "" Then
        MkDir outputFolder
    End If

    ' Obtener la hoja activa
    Set xlSheet = ThisWorkbook.Sheets(1) ' Ajusta el índice o nombre de la hoja si es necesario

    ' Iniciar la aplicación de Word
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True

    ' Abrir el archivo Word fijo
    Set wdDoc = wdApp.Documents.Open(pathWord)

    ' Seleccionar la tabla en Word (asume que es la primera tabla)
    Set tbl = wdDoc.Tables(1)

    ' Obtener la última fila con datos en Excel
    lastRow = xlSheet.Cells(xlSheet.Rows.Count, "A").End(xlUp).Row

    ' Recorrer las filas en Excel y añadirlas a la tabla en Word
    For i = 2 To lastRow ' Asume que la fila 1 es el encabezado
        codigo = xlSheet.Cells(i, 1).Value
        denominacion = xlSheet.Cells(i, 2).Value
        fechaInicio = xlSheet.Cells(i, 3).Value
        fechaFin = xlSheet.Cells(i, 4).Value
        horas = xlSheet.Cells(i, 5).Value
        ' Añadir una nueva fila a la tabla en Word
        With tbl.Rows.Add
            .Cells(1).Range.Text = codigo
            .Cells(2).Range.Text = denominacion & vbCrLf & "(" & horas & " horas)"
            .Cells(5).Range.Text = "8000000705" & vbCrLf & "(Teleformación)" & vbCrLf & "Grupo cfcom 2.0, s.l"
            .Cells(6).Range.Text = fechaInicio & "  A  " & fechaFin & vbCrLf & "(Teleformación)"
            .Cells(7).Range.Text = "NO TIENE SESIONES PRESENCIALES"
        End With
    Next i

    For Each cc In wdDoc.ContentControls
        Select Case cc.Title
            Case "NombreAlumno"
                cc.Range.Text = xlSheet.Cells(12, 11).Value ' Dato en K12
                cc.Range.Text = xlSheet.Cells(13, 11).Value ' Dato en K13
        End Select
    Next cc
                
                
                
                
    ' Solicitar al usuario el nombre del archivo para guardar
    nombreArchivo = InputBox("Ingrese el nombre con el que desea guardar el archivo Word:", "Guardar como")

    ' Guardar el documento con el nombre proporcionado en la carpeta de salida predeterminada
    If nombreArchivo <> "" Then
        wdDoc.SaveAs2 outputFolder & nombreArchivo & ".docx"
    Else
        MsgBox "No se ha guardado el archivo. Nombre no proporcionado.", vbExclamation
    End If

    ' Cerrar el documento de Word
    wdDoc.Close False
    wdApp.Quit

    ' Liberar objetos
    Set xlSheet = Nothing
    Set tbl = Nothing
    Set wdDoc = Nothing
    Set wdApp = Nothing

End Sub