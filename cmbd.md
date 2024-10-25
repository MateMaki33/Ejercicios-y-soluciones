# MACROS PARA EXCEL

## MACRO PARA EL MODULO 

- Recuerda que debes habilitar la pestaña programación en archivo/opciones/personalizar cinta de opciones y activar la pestaña 
- Haz click en la pestaña y accede a VBA o Visual Basic
- Ahora debes incluir un módulo en VBA
- Incluye este código en el módulo y guarda como XSLM. 
- Ejecuta la macro para generar la estructura básica de ejemplo:

```
Sub CrearDatasetConteosYGrafico()
    Dim wsDatos As Worksheet
    Dim wsConteos As Worksheet
    Dim wsGrafico As Worksheet
    Dim chartObj As ChartObject
    Dim lastRow As Long
    Dim diagCount As Object
    Dim key As Variant
    Dim rowCount As Long
    Dim i As Long
    Dim tbl As ListObject

    ' Eliminar todas las hojas excepto la hoja activa
    Dim ws As Worksheet
    Application.DisplayAlerts = False ' Desactivar advertencias

    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> ThisWorkbook.ActiveSheet.Name Then
            ws.Delete
        End If
    Next ws

    Application.DisplayAlerts = True ' Reactivar advertencias

    ' Crear nueva hoja "Dataset CMBD"
    On Error Resume Next ' Ignorar error si ya existe una hoja con ese nombre
    Set wsDatos = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsDatos.Name = "Dataset CMBD"
    On Error GoTo 0 ' Volver a activar el manejo de errores

    ' Crear un conjunto de datos de ejemplo
    With wsDatos
        .Cells.Clear
        .Cells(1, 1).Value = "ID Paciente"
        .Cells(1, 2).Value = "Sexo"
        .Cells(1, 3).Value = "Fecha de Nacimiento"
        .Cells(1, 4).Value = "Residencia"
        .Cells(1, 5).Value = "Fuente de Financiación"
        .Cells(1, 6).Value = "Fecha de Ingreso"
        .Cells(1, 7).Value = "Diagnóstico Principal"
        .Cells(1, 8).Value = "Procedimientos"

        Dim sexos As Variant
        Dim residencias As Variant
        Dim fuentes As Variant
        Dim diagnosticos As Variant
        Dim procedimientos As Variant
        
        sexos = Array("Masculino", "Femenino")
        residencias = Array("Madrid", "Barcelona", "Valencia", "Sevilla")
        fuentes = Array("Seguro", "Privado", "Público", "Ninguno")
        diagnosticos = Array("Infección Respiratoria", "Diabetes", "Hipertensión", "Fractura", "Cáncer", "Covid-19", "Asma", "Apendicitis")
        procedimientos = Array("Cirugía Mayor", "Consulta Externa", "Hospitalización", "Urgencias")

        For i = 1 To 10
            .Cells(i + 1, 1).Value = "P" & Format(i, "000")
            .Cells(i + 1, 2).Value = sexos(Application.WorksheetFunction.RandBetween(0, UBound(sexos)))
            .Cells(i + 1, 3).Value = Date - Application.WorksheetFunction.RandBetween(20 * 365, 70 * 365)
            .Cells(i + 1, 4).Value = residencias(Application.WorksheetFunction.RandBetween(0, UBound(residencias)))
            .Cells(i + 1, 5).Value = fuentes(Application.WorksheetFunction.RandBetween(0, UBound(fuentes)))
            .Cells(i + 1, 6).Value = Date - Application.WorksheetFunction.RandBetween(1, 30)
            .Cells(i + 1, 7).Value = diagnosticos(Application.WorksheetFunction.RandBetween(0, UBound(diagnosticos)))
            .Cells(i + 1, 8).Value = procedimientos(Application.WorksheetFunction.RandBetween(0, UBound(procedimientos)))
        Next i
    End With

    ' Crear nueva hoja para los conteos de diagnósticos
    On Error Resume Next
    Set wsConteos = ThisWorkbook.Sheets("Conteos Diagnósticos")
    If wsConteos Is Nothing Then
        Set wsConteos = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsConteos.Name = "Conteos Diagnósticos"
    Else
        wsConteos.Cells.Clear ' Limpiar la hoja existente si ya existe
    End If
    On Error GoTo 0

    ' Inicializar el conteo de diagnósticos
    Call ActualizarConteos(wsDatos, wsConteos)

    ' Crear nueva hoja para el gráfico
    On Error Resume Next
    Set wsGrafico = ThisWorkbook.Sheets("Gráfico Dinámico")
    If wsGrafico Is Nothing Then
        Set wsGrafico = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsGrafico.Name = "Gráfico Dinámico"
    Else
        wsGrafico.Cells.Clear ' Limpiar la hoja existente si ya existe
    End If
    On Error GoTo 0

    ' Crear gráfico
    Call CrearGrafico(wsConteos, wsGrafico)

    MsgBox "Se ha creado el dataset, los conteos y el gráfico dinámico correctamente.", vbInformation
End Sub

Sub ActualizarConteos(wsDatos As Worksheet, wsConteos As Worksheet)
    Dim diagCount As Object
    Dim key As Variant
    Dim rowCount As Long
    Dim lastRow As Long
    Dim i As Long

    Set diagCount = CreateObject("Scripting.Dictionary")
    lastRow = wsDatos.Cells(wsDatos.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRow
        key = wsDatos.Cells(i, 7).Value
        If diagCount.Exists(key) Then
            diagCount(key) = diagCount(key) + 1
        Else
            diagCount.Add key, 1
        End If
    Next i

    ' Escribir los conteos en la hoja
    wsConteos.Cells.Clear
    wsConteos.Cells(1, 1).Value = "Diagnóstico Principal"
    wsConteos.Cells(1, 2).Value = "Cantidad"
    rowCount = 2
    For Each key In diagCount.Keys
        wsConteos.Cells(rowCount, 1).Value = key
        wsConteos.Cells(rowCount, 2).Value = diagCount(key)
        rowCount = rowCount + 1
    Next key

    ' Crear una tabla de Excel
    Dim tbl As ListObject
    Set tbl = wsConteos.ListObjects.Add(xlSrcRange, wsConteos.Range("A1:B" & rowCount - 1), , xlYes)
    tbl.Name = "TablaConteos"
End Sub

Sub CrearGrafico(wsConteos As Worksheet, wsGrafico As Worksheet)
    Dim chartObj As ChartObject

    ' Crear gráfico
    Set chartObj = wsGrafico.ChartObjects.Add(Left:=50, Width:=600, Top:=50, Height:=400)

    ' Configurar el gráfico
    With chartObj.Chart
        .ChartType = xlPie ' Gráfico de pastel
        .HasTitle = True
        .ChartTitle.Text = "Distribución de Diagnósticos"
        .SetSourceData Source:=wsConteos.ListObjects("TablaConteos").Range
        .SeriesCollection(1).XValues = wsConteos.Range("TablaConteos[Diagnóstico Principal]")
        .SeriesCollection(1).Values = wsConteos.Range("TablaConteos[Cantidad]")
    End With
End Sub


```

## MACRO PARA ACTUALIZACIONES AUTOMATICAS DE LOS DATOS

- Una vez creada la estructura, selecciona la hoja donde está la tabla con los datos, dentro de visual basic.
- Haz doble click y se abrirá otra ventana para macro.
- Inserta la macro:

```
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Me.Range("A2:H1000")) Is Nothing Then
        Dim wsConteos As Worksheet
        Dim wsGrafico As Worksheet
        
        ' Verifica que la hoja de conteos existe
        On Error Resume Next
        Set wsConteos = ThisWorkbook.Sheets("Conteos Diagnósticos")
        Set wsGrafico = ThisWorkbook.Sheets("Gráfico Dinámico")
        On Error GoTo 0

        If Not wsConteos Is Nothing Then
            ' Llama a la función de actualización
            Call ActualizarConteos(Me, wsConteos)
            
            ' Actualiza el gráfico
            Call CrearGrafico(wsConteos, wsGrafico)
        End If
    End If
End Sub

```


