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
    Dim pos As Object

    ' Variables para almacenar datos de las celdas
    Dim codigo As String
    Dim denominacion As String
    Dim horas As String
    Dim modalidad As String
    Dim codCentro As String
    Dim columna6 As String
    Dim denominacionCentro As String
    Dim tutor As String
    Dim nif As String

    ' Verificar si la hoja de trabajo existe
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CALCULO")
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "La hoja de trabajo 'CALCULO' no existe.", vbCritical
        Exit Sub
    End If

    ' Crear una instancia de Word y abrir el documento para rellenar el formulario
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True ' Opcional, para ver Word mientras se ejecuta el script
    Set wdDoc = wdApp.Documents.Open("/////SetearRuta//////")  ' Ruta del archivo que queremos modificar

    ' Encontrar la última fila con datos en la primera columna
    ultimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' ***** PARTE 1: TABLA DE MATERIAS EN LA HOJA 3 *****

    ' Mover el cursor al marcador en la tercera página
    wdDoc.Bookmarks("TerceraPagina").Select

    ' Insertar el título antes de generar la tabla
    With wdApp.Selection
        .ParagraphFormat.Alignment = 0 ' wdAlignParagraphLeft
        .Font.Bold = True
        .TypeText Text:="2.A Itinerario de especialidades formativas del Catálogo de Especialidades Formativas del Sistema Nacional de Empleo"
        .TypeParagraph
    End With

    ' Inicializar la posición del cursor después del título
    Set pos = wdApp.Selection.Range
    pos.Collapse Direction:=0 ' wdCollapseEnd
    pos.InsertParagraphAfter
    pos.Collapse Direction:=0 ' wdCollapseEnd

    ' Crear una nueva tabla con 5 columnas y (ultimaFila - 1) filas (porque la primera fila de datos es la fila 2)
    Set tabla = wdDoc.Tables.Add(Range:=pos, NumRows:=ultimaFila - 1 + 1, NumColumns:=5) ' +1 para la fila del encabezado
    tabla.Borders.Enable = True
    ' Cambiar el tamaño de fuente de toda la tabla a 10 puntos
    tabla.Range.Font.Size = 9

    ' Insertar encabezados en la primera fila
    With tabla.Rows(1)
        .Cells(1).Range.Text = "Código"
        .Cells(2).Range.Text = "Denominación"
        .Cells(3).Range.Text = "Nº Horas"
        .Cells(4).Range.Text = "Modalidad"
        .Cells(5).Range.Text = "Cod. Centro Inscrito Reg.E."
        ' Asegurarse de que los encabezados no estén en negrita
    .Range.Font.Bold = False
    End With

    ' Recorrer cada fila con datos y rellenar la tabla
    For fila = 2 To ultimaFila
        ' Obtener los datos de la fila actual
        codigo = ws.Cells(fila, 1).Value
        denominacion = ws.Cells(fila, 9).Value
        horas = ws.Cells(fila, 5).Value
        modalidad = ws.Cells(fila, 7).Value
        codCentro = ws.Cells(fila, 8).Value

        ' Rellenar la tabla con los datos
        With tabla.Rows(fila - 1 + 1) ' -1 porque la primera fila es el encabezado, +1 porque la primera fila de datos es la fila 2
            .Cells(1).Range.Text = codigo
            .Cells(2).Range.Text = denominacion
            .Cells(3).Range.Text = horas
            .Cells(4).Range.Text = modalidad
            .Cells(5).Range.Text = codCentro
        ' Asegurarse de que las celdas no estén en negrita
        .Range.Font.Bold = False
        End With
    Next fila

    ' ***** PARTE 2: CUADROS PARA LA HOJA 4 *****

    ' Mover el cursor al marcador en la cuarta página
    wdDoc.Bookmarks("CuartaPagina").Select

    ' Insertar el título antes de generar los cuadros
    With wdApp.Selection
        .ParagraphFormat.Alignment = 0 ' wdAlignParagraphLeft
        .Font.Bold = True
        .TypeText Text:="4.- CENTROS IMPARTIDORES DE LA ACTIVIDAD FORMATIVA"
        .TypeParagraph
    End With

    ' Inicializar la posición del cursor después del título
    Set pos = wdApp.Selection.Range
    pos.Collapse Direction:=0 ' wdCollapseEnd
    pos.InsertParagraphAfter
    pos.Collapse Direction:=0 ' wdCollapseEnd

    ' Recorrer cada fila con datos
    For fila = 2 To ultimaFila
        ' Obtener los datos de la fila actual
        codigoCentro = ws.Cells(fila, 1).Value
        denominacionCentro = ws.Cells(fila, 9).Value
        tutor = ws.Cells(fila, 6).Value
        nif = ws.Cells(9, 11).Value

        ' Insertar una nueva tabla de una celda para el cuadro
        Set tabla = wdDoc.Tables.Add(Range:=pos, NumRows:=1, NumColumns:=1)
        tabla.Borders.Enable = True
        tabla.Range.Font.Size = 9
        ' Insertar el contenido del cuadro en la celda de la tabla con los datos de Excel
      With tabla.Cell(1, 1).Range
            .Text = "DATOS DEL CENTRO DE FORMACIÓN" & vbCrLf & vbCrLf & _
                "Formación a impartir: Código: " & codigoCentro & " Denominación: " & denominacionCentro & vbCrLf & _
                 ChrW(&H2610) & " Centro Sistema Educativo. Código de centro autorizado: " & vbCrLf & _
                 ChrW(&H2611) & " " & vbCrLf & _
                 ChrW(&H2610) & " Si la formación se imparte mediante teleformación, en su caso, especificar código/s del/os Centros Presenciales vinculados: " & vbCrLf & vbCrLf & _
                 "Nombre Centro:             CIF/NIF/NIE: " & vbCrLf & _
                 "URL (Entidades de teleformación): " & vbCrLf & _
                 "Dirección:                     CP:                            Municipio: 
                 "Provincia:   VALENCIA       Teléfono                 Correo electrónico " & vbCrLf & _
                 "D./Dña.   en concepto de                            NIF/NIE        
                  "Tutor/a del centro – D./Dña. " & tutor & "                 NIF/NIE  " & nif
            .Font.Bold = False
        End With

        ' Mover el cursor fuera de la tabla y añadir un párrafo después de cada cuadro
        pos.Collapse Direction:=0 ' wdCollapseEnd
        pos.InsertParagraphAfter
        pos.Collapse Direction:=0 ' wdCollapseEnd
        ' Añadir un salto de párrafo para separar cada cuadro
        pos.InsertBreak Type:=7 ' wdPageBreak
        pos.InsertBreak Type:=7 ' wdPageBreak
        pos.Collapse Direction:=0 ' wdCollapseEnd
    Next fila

    ' ***** PARTE 3: NUEVA TABLA EN SEGUNDOCUADRO *****

    ' Mover el cursor al marcador en el segundo cuadro
    wdDoc.Bookmarks("segundocuadro").Select

    ' Insertar el título antes de generar la tabla
    With wdApp.Selection
        .ParagraphFormat.Alignment = 0 ' wdAlignParagraphLeft
        .Font.Bold = True
        .TypeText Text:="Actividad Formativa"
        .TypeParagraph
    End With

    ' Inicializar la posición del cursor después del título
    Set pos = wdApp.Selection.Range
    pos.Collapse Direction:=0 ' wdCollapseEnd
    pos.InsertParagraphAfter
    pos.Collapse Direction:=0 ' wdCollapseEnd

    ' Crear una nueva tabla con 6 columnas y (ultimaFila - 1) filas (porque la primera fila de datos es la fila 2)
    Set tabla = wdDoc.Tables.Add(Range:=pos, NumRows:=ultimaFila - 1 + 1, NumColumns:=6) ' +1 para la fila del encabezado
    tabla.Borders.Enable = True
    tabla.Range.Font.Size = 9
    ' Insertar encabezados en la primera fila
    With tabla.Rows(1)
        .Cells(1).Range.Text = "Código"
        .Cells(2).Range.Text = "Fecha de inicio"
        .Cells(3).Range.Text = "Fecha de fin"
        .Cells(4).Range.Text = "Horas semanales de Actividad formativa"
        .Cells(5).Range.Text = "Dias de la semana"
        .Cells(6).Range.Text = "Horario"
        .Range.Font.Bold = False
    End With

    ' Recorrer cada fila con datos y rellenar la tabla
    For fila = 2 To ultimaFila
        ' Obtener los datos de la fila actual
        codigo = ws.Cells(fila, 1).Value
        denominacion = ws.Cells(fila, 9).Value
        horas = ws.Cells(fila, 5).Value
        modalidad = ws.Cells(fila, 7).Value
        codCentro = ws.Cells(fila, 8).Value
        columna6 = ws.Cells(25, 11).Value

        ' Rellenar la tabla con los datos
        With tabla.Rows(fila - 1 + 1) ' -1 porque la primera fila es el encabezado, +1 porque la primera fila de datos es la fila 2
            .Cells(1).Range.Text = codigo
            .Cells(2).Range.Text = denominacion
            .Cells(3).Range.Text = horas
            .Cells(4).Range.Text = modalidad
            .Cells(5).Range.Text = codCentro
            .Cells(6).Range.Text = columna6
            .Range.Font.Bold = False
        End With
    Next fila

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
            Case "Horas"
                cc.Range.Text = ws.Cells(10, 11).Value ' Dato en K10
            Case "Convenio"
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
        rutaArchivo = "//Insertar ruta de salida de archivo///"
        
        ' Guardar el documento con el nombre proporcionado por el usuario
        wdDoc.SaveAs rutaArchivo
        MsgBox "Archivo guardado en: " & rutaArchivo, vbInformation
    End If


    ' Liberar objetos
    Set wdDoc = Nothing
    Set wdApp = Nothing
End Sub


