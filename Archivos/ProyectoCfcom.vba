Sub RellenarFormularioYCrearCuadros()
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim ws As Worksheet
    Dim cc As Object
    Dim nombreArchivo As String
    Dim rutaArchivo As String
    Dim fila As Long
    Dim ultimaFila As Long
    Dim rango As Object
    Dim tabla As Object

    ' Verificar si la hoja de trabajo existe
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CALCULO")
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "La hoja de trabajo 'CÁLCULO' no existe.", vbCritical
        Exit Sub
    End If

    ' Crear una instancia de Word y abrir el documento para rellenar el formulario
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True ' Opcional, para ver Word mientras se ejecuta el script
    Set wdDoc = wdApp.Documents.Open("C:\Users\ContratoFor\Desktop\Pruebas pdf\Pruebas pdf\Archivos\Formulariollenar.docx")  ' Ruta del archivo que queremos modificar

    ' Rellenar campos de formulario utilizando nombres únicos
    For Each cc In wdDoc.ContentControls
        Select Case cc.Title
            Case "NombreCampo"
                cc.Range.Text = ws.Cells(2, 1).Value ' Dato en A2
            Case "ApellidoCampo"
                cc.Range.Text = ws.Cells(2, 2).Value ' Dato en B2
            Case "Fecha1Campo"
                cc.Range.Text = ws.Cells(2, 3).Value ' Dato en C2
            Case "Fecha2Campo"
                cc.Range.Text = ws.Cells(2, 4).Value ' Dato en D2
        End Select
    Next cc

 ' Generar cuadros para cada fila de Excel en el mismo documento
' Encontrar la última fila con datos en la primera columna
ultimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

' Mover el cursor al final del documento
wdDoc.Content.InsertParagraphAfter
wdDoc.Content.Paragraphs.Last.Range.Select

' Recorrer cada fila con datos
For fila = 2 To ultimaFila
    ' Insertar una nueva tabla de una celda para el cuadro
    Set rango = wdDoc.Content.Paragraphs.Last.Range
    Set tabla = wdDoc.Tables.Add(Range:=rango, NumRows:=1, NumColumns:=1)
    tabla.Borders.Enable = True

    ' Insertar el contenido del cuadro en la celda de la tabla con placeholders
  
With tabla.Cell(1, 1).Range
    .Text = "DATOS DEL CENTRO DE FORMACIÓN" & vbCrLf & _
            "Formación a impartir: Código: [CÓDIGO] Denominación: [DENOMINACIÓN]" & vbCrLf & _
            ChrW(&H2610) & " Centro Sistema Educativo. Código de centro autorizado:" & vbCrLf & _
            ChrW(&H2610) & " Centro Acreditado. Código de centro en Registro Estatal de centros de formación: 8000000705" & vbCrLf & _
            ChrW(&H2610) & " Si la formación se imparte mediante teleformación, en su caso, especificar código/s del/os Centros Presenciales vinculados:" & vbCrLf & vbCrLf & _
            "Nombre Centro: [NOMBRE CENTRO] CIF/NIF/NIE: [CIF/NIF/NIE]" & vbCrLf & _
            "URL (Entidades de teleformación)" & vbCrLf & _
            "Dirección: [DIRECCIÓN] CP: [CP] Municipio: [MUNICIPIO]" & vbCrLf & _
            "Provincia: [PROVINCIA] Teléfono: [TELÉFONO] Correo electrónico: [CORREO ELECTRÓNICO]" & vbCrLf & _
            "D./Dña. [NOMBRE] en concepto de [CONCEPTO] NIF/NIE: [NIF/NIE]" & vbCrLf & _
            "Tutor/a del centro – D./Dña. [TUTOR/A] NIF/NIE: [NIF/NIE TUTOR/A]"
End With

    ' Mover el cursor fuera de la tabla y añadir un salto de página
    If fila < ultimaFila Then
        Set rango = wdDoc.Content
        rango.Collapse Direction:=0 ' wdCollapseEnd
        rango.InsertBreak Type:=7 ' wdPageBreak
    End If
Next fila


    ' Solicitar al usuario el nombre del archivo de salida
    nombreArchivo = InputBox("Ingrese el nombre del archivo de salida (sin extensión):", "Guardar como")
    
    ' Asegurarse de que el usuario no haya dejado el nombre en blanco
    If nombreArchivo = "" Then
        MsgBox "No se ingresó un nombre de archivo. El archivo no se guardará.", vbExclamation
    Else
        ' Definir la ruta completa para guardar el archivo
        rutaArchivo = "C:\Users\ContratoFor\Desktop\Pruebas pdf\Pruebas pdf\Archivos\Archivos de salida\" & nombreArchivo & ".docx"
        
        ' Guardar el documento con el nombre proporcionado por el usuario
        wdDoc.SaveAs rutaArchivo
        MsgBox "Archivo guardado en: " & rutaArchivo, vbInformation
    End If
    
    ' Cerrar el documento y Word
    wdDoc.Close False
    wdApp.Quit
    
    ' Liberar objetos
    Set wdDoc = Nothing
    Set wdApp = Nothing
End Sub