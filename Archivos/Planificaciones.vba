' -----------------------------------------------------------
' Script: InsertarDatosDesdeExcelEnWord
' Autor: Guillermo Luna Alvarez
' Descripción: Este macro automatiza la inserción de datos desde un archivo Excel a un documento Word.
' NOTA: Este código no contiene datos personales. Cualquier dato que se procese en tiempo de ejecución
' proviene de archivos externos y debe ser anonimizado antes de ser compartido.
' -----------------------------------------------------------




Sub InsertarDatosDesdeExcelEnWord()

    ' Declaración de variables
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim tbl As Object
    Dim xlSheet As Worksheet
    Dim i As Integer, lastRow As Integer
    Dim codigo As String, denominacion As String, fechaInicio As String, fechaFin As String, horas As String
    Dim nombreArchivo As String
    Dim sumaMaterias As String
    Dim cc As Object
    
    ' Ruta fija del archivo Word
    Dim pathWord As String
    pathWord = "C://RUTA DEL ARCHIVO//" ' Cambia esta ruta a la ubicación del archivo Word fijo

    ' Carpeta de salida predeterminada
    Dim outputFolder As String
    outputFolder = "//RUTA DE SALIDA DESEADA//" ' Cambia esta ruta a la carpeta de salida deseada

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

    ' Inicializar la cadena para concatenar los códigos de las materias
    sumaMaterias = ""

    ' Obtener la última fila con datos en Excel
    lastRow = xlSheet.Cells(xlSheet.Rows.Count, "A").End(xlUp).Row

    ' Recorrer las filas en Excel y añadirlas a la tabla en Word
    For i = 2 To lastRow ' Asume que la fila 1 es el encabezado
        codigo = xlSheet.Cells(i, 1).Value
        denominacion = xlSheet.Cells(i, 2).Value
        fechaInicio = xlSheet.Cells(i, 3).Value
        fechaFin = xlSheet.Cells(i, 4).Value
        horas = xlSheet.Cells(i, 5).Value
        
        ' Concatenar el código de la materia a la cadena de sumaMaterias
        If sumaMaterias = "" Then
            sumaMaterias = codigo
        Else
            sumaMaterias = sumaMaterias & "+" & codigo
        End If
        
        ' Añadir una nueva fila a la tabla en Word
        With tbl.Rows.Add
            .Cells(1).Range.Text = codigo
            .Cells(2).Range.Text = denominacion & vbCrLf & "(" & horas & " horas)"
            .Cells(5).Range.Text = "" & vbCrLf & "(Teleformación)" & vbCrLf & ""
            .Cells(6).Range.Text = fechaInicio & "  A  " & fechaFin & vbCrLf & "(Teleformación)"
            .Cells(7).Range.Text = "NO TIENE SESIONES PRESENCIALES"
        End With
    Next i

    ' Reemplazar los valores en los controles de contenido
    For Each cc In wdDoc.ContentControls
        Select Case cc.Title
            Case "NombreAlumno"
                cc.Range.Text = xlSheet.Cells(12, 11).Value ' Dato en K12
            Case "DniAlumno"
                cc.Range.Text = xlSheet.Cells(13, 11).Value ' Dato en K13
            Case "SumaDeMaterias"
                cc.Range.Text = sumaMaterias ' Asignar la cadena concatenada a SumaDeMaterias
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
