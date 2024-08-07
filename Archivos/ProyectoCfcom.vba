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
    
' Generar cuadros para cada fila de Excel en el mismo documento
 ' Encontrar la última fila con datos en la primera columna
    ultimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Mover el cursor al marcador en la cuarta página
    wdDoc.Bookmarks("CuartaPagina").Select

    ' Insertar el título antes de generar los cuadros
    wdApp.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    wdApp.Selection.Font.Bold = True
    wdApp.Selection.TypeText Text:="4.- CENTROS IMPARTIDORES DE LA ACTIVIDAD FORMATIVA"
    wdApp.Selection.TypeParagraph
    wdApp.Selection.Font.Bold = False

    ' Recorrer cada fila con datos
    For fila = 2 To ultimaFila
        ' Insertar una nueva tabla de una celda para el cuadro
        Set rango = wdApp.Selection.Range
        Set tabla = wdDoc.Tables.Add(Range:=rango, NumRows:=1, NumColumns:=1)
        tabla.Borders.Enable = True

        ' Insertar el contenido del cuadro en la celda de la tabla con placeholders
        With tabla.Cell(1, 1).Range
            .Text = "DATOS DEL CENTRO DE FORMACIÓN" & vbCrLf & _
                    "Formación a impartir: Código:         SIN GENERAR             Denominación: COMPETENCIAS DIGITALES BÁSICAS" & vbCrLf & vbCrLf & _
                    ChrW(&H2610) & " Centro Sistema Educativo. Código de centro autorizado:" & vbCrLf & _
                    ChrW(&H2610) & " Centro Acreditado. Código de centro en Registro Estatal de centros de formación: 8000000705" & vbCrLf & _
                    ChrW(&H2610) & " Si la formación se imparte mediante teleformación, en su caso, especificar código/s del/os Centros Presenciales vinculados:" & vbCrLf & vbCrLf & _
                    "Nombre Centro: Grupo CFCOM 2.0, S.L.              CIF/NIF/NIE: B98551401" & vbCrLf & _
                    "URL (Entidades de teleformación)" & vbCrLf & _
                    "Dirección: Calle Chiva, 20, B                    CP:        46018                      Municipio: VALENCIA" & vbCrLf & _
                    "Provincia:   VALENCIA       Teléfono         962067573            Correo electrónico INFO@CONTRATO-FORMACION.COM" & vbCrLf & _
                    "D./Dña. JOSE VICENTE ROIG           en concepto de                GERENTE               NIF/NIE         44869822L" & vbCrLf & _
                    "Tutor/a del centro – D./Dña. ROCÍO LÓPEZ ROMERO                 NIF/NIE  26752178G"
        End With

        ' Mover el cursor fuera de la tabla y añadir un párrafo después de cada cuadro
        wdApp.Selection.Collapse Direction:=0 ' wdCollapseEnd
        wdApp.Selection.TypeParagraph
    Next fila

' Insertar un salto de página después del último cuadro
Set rango = wdDoc.Content
rango.Collapse Direction:=0 ' wdCollapseEnd
rango.InsertBreak Type:=7 ' wdPageBreak


    ' Rellenar campos de formulario utilizando nombres únicos
    For Each cc In wdDoc.ContentControls
         Select Case cc.Title
        Case "NombreEmpresa"
                cc.Range.Text = ws.Cells(1, 11).Value ' Dato en K1
            Case "CifEmpresa"
                cc.Range.Text = ws.Cells(2, 11).Value ' Dato en K2
            Case "NombreJefe"
                cc.Range.Text = ws.Cells(3, 11).Value ' Dato en K3
            Case "CargoJefe"
                cc.Range.Text = ws.Cells(4, 11).Value ' Dato en K4
            Case "DniJefe"
                cc.Range.Text = ws.Cells(5, 11).Value ' Dato en K5
            Case "MailEmpresa"
                cc.Range.Text = ws.Cells(6, 11).Value ' Dato en K6
            Case "TelefonoEmpresa"
                cc.Range.Text = ws.Cells(7, 11).Value ' Dato en K7
            Case "TutorEmpresa"
                cc.Range.Text = ws.Cells(8, 11).Value ' Dato en K8
            Case "DniTutor"
                cc.Range.Text = ws.Cells(9, 11).Value ' Dato en K9
            Case "HorasSemanales"
                cc.Range.Text = ws.Cells(10, 11).Value ' Dato en K10
            Case "ConvenioAplicable"
                cc.Range.Text = ws.Cells(11, 11).Value ' Dato en K11
            Case "NombreTrabajador"
                cc.Range.Text = ws.Cells(12, 11).Value ' Dato en K12
            Case "DniTrabajador"
                cc.Range.Text = ws.Cells(13, 11).Value ' Dato en K13
            Case "FechaNacimientoTrabajador"
                cc.Range.Text = ws.Cells(14, 11).Value ' Dato en K14
            Case "FechaInicioContrato"
                cc.Range.Text = ws.Cells(15, 11).Value ' Dato en K15
            Case "FechaFinContrato"
                cc.Range.Text = ws.Cells(16, 11).Value ' Dato en K16
            Case "OcupacionOPuesto"
                cc.Range.Text = ws.Cells(17, 11).Value ' Dato en K17
            Case "CNO"
                cc.Range.Text = ws.Cells(18, 11).Value ' Dato en K18
            Case "ProvinciaPuesto"
                cc.Range.Text = ws.Cells(19, 11).Value ' Dato en K19
            Case "HorasContratoAñoUno"
                cc.Range.Text = ws.Cells(20, 11).Value ' Dato en K20
            Case "HorasContratoAñoDos"
                cc.Range.Text = ws.Cells(21, 11).Value ' Dato en K21
            Case "HorasItinerario"
                cc.Range.Text = ws.Cells(22, 11).Value ' Dato en K22
            Case "DiasLaboral"
                cc.Range.Text = ws.Cells(23, 11).Value ' Dato en K23
            Case "HorarioLaboral"
                cc.Range.Text = ws.Cells(24, 11).Value ' Dato en K24
            Case "HorarioFormacion"
                cc.Range.Text = ws.Cells(25, 11).Value ' Dato en K25
            Case "DireccionCentroTrabajo"
                cc.Range.Text = ws.Cells(26, 11).Value ' Dato en K26
        End Select
    Next cc

 


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
