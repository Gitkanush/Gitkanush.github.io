' Macro para realizar la actualización de los datos desde SAP optimizada
' para una respuesta más rápida del archivo y SAP.
' Esta macro es usada para el archivo Generador de precios.
' @author Antonio Alfredo Ramírez Ramírez
' @since 18/01/2020
' @version 1.2.5 15/06/2020

Sub ActualizarDatosGenerador()
    Dim fila As ListRow
    Dim respaldo As ListObject
    Dim principal As ListObject
    Dim cont As Integer
    
    Set respaldo = Datos.ListObjects("tRespaldo")
    Set principal = Generador.ListObjects("tGenerador")
    FastWB True
    MsgBox "Se actualizarán los datos desde SAP" & vbNewLine & "Por favor espere, esto puede tardar", vbInformation
    respaldo.DataBodyRange.ClearContents
    ' respaldamos los datos capturados manualmente en la hoja del generador
    With principal
        .ListColumns("Clave").DataBodyRange.Copy
        respaldo.ListColumns("Clave").DataBodyRange.PasteSpecial xlPasteValues, xlPasteSpecialOperationNone
        .ListColumns("Venta atípica").DataBodyRange.Copy
        respaldo.ListColumns("Venta atípica").DataBodyRange.PasteSpecial xlPasteValues, xlPasteSpecialOperationNone
        .ListColumns("Pedido realizado").DataBodyRange.Copy
        respaldo.ListColumns("Pedido realizado").DataBodyRange.PasteSpecial xlPasteValues, xlPasteSpecialOperationNone
        .ListColumns("Número de OC").DataBodyRange.Copy
        respaldo.ListColumns("Número de OC").DataBodyRange.PasteSpecial xlPasteValues, xlPasteSpecialOperationNone
        .ListColumns("Excepción de compra").DataBodyRange.Copy
        respaldo.ListColumns("Excepción de compra").DataBodyRange.PasteSpecial xlPasteValues, xlPasteSpecialOperationNone
        .ListColumns("Proveedor").DataBodyRange.Copy
        respaldo.ListColumns("Proveedor").DataBodyRange.PasteSpecial xlPasteValues, xlPasteSpecialOperationNone
        ' borramos las claves de productos de la hoja generador
        .ListColumns("Clave").DataBodyRange.ClearContents
    End With
    
    ' copiamos las claves de la tabla datos al generador
    With Datos.ListObjects("tDatos")
        cont = 1
        For Each fila In .ListRows
            If .DataBodyRange(fila.Index, .ListColumns("Costo").Index).Value <> 0 Then
                If .DataBodyRange(fila.Index, .ListColumns("Tipo").Index).Value <> "D" Then
                    principal.DataBodyRange(cont, principal.ListColumns("Clave").Index) = .DataBodyRange(fila.Index, .ListColumns("Clave").Index)
                    cont = cont + 1
                End If
            End If
        Next fila
    End With
    
    ' obtenemos los datos respaldados según las nuevas claves cargadas
    With principal
        cont = 1
        For Each fila In .ListRows
            If .DataBodyRange(fila.Index, .ListColumns("Clave").Index) = respaldo.DataBodyRange(cont, respaldo.ListColumns("Clave").Index) Then
                .DataBodyRange(fila.Index, .ListColumns("Venta atípica").Index) = respaldo.DataBodyRange(cont, respaldo.ListColumns("Venta atípica").Index)
                .DataBodyRange(fila.Index, .ListColumns("Pedido realizado").Index) = respaldo.DataBodyRange(cont, respaldo.ListColumns("Pedido realizado").Index)
                .DataBodyRange(fila.Index, .ListColumns("Número de OC").Index) = respaldo.DataBodyRange(cont, respaldo.ListColumns("Número de OC").Index)
                .DataBodyRange(fila.Index, .ListColumns("Excepción de compra").Index) = respaldo.DataBodyRange(cont, respaldo.ListColumns("Excepción de compra").Index)
                .DataBodyRange(fila.Index, .ListColumns("Proveedor").Index) = respaldo.DataBodyRange(cont, respaldo.ListColumns("Proveedor").Index)
                cont = cont + 1
            End If
        Next fila
    End With
    MsgBox "Proceso terminado" & vbNewLine & "Puede continuar", vbExclamation
    XlResetSettings
End Sub

' Macro para crear una hoja de respaldo de datos de pedidos'
' usada para registrar los pedidos hechos en un día y poder'
' realizar comparaciones de compras para AMASA'
' @author Antonio Alfredo Ramírez Ramírez'
' @since 01feb2020'
' @version 1.1.0 29/06/2020

Sub RespaldoDiario()
    Dim fila As ListRow
    Dim origen As ListObject
    Dim ws As Worksheet
    Dim cont As Integer
    Dim resp As Integer
    Dim columnas(5) As String
    Dim i As Integer
    FastWB True
    Generador.Unprotect
    Set origen = Generador.ListObjects("tGenerador")
    resp = MsgBox("Desea generar una hoja nueva con el reporte (Si)" & vbNewLine & "O cargar en columnas nuevas al final de la tabla (No)", vbYesNo)
    If resp = vbYes Then
        ' Se crea una nueva hoja después de la última'
        MsgBox "Se creará una hoja nueva con el nombre:" & vbNewLine & WorksheetFunction.Text(Now, "ddmmyy-hhmm")
        Set ws = ThisWorkbook.Sheets.Add(After:=Agenda)
        ws.Name = WorksheetFunction.Text(Now, "ddmmyy-hhmm")
        ws.Tab.Color = RGB(102, 204, 102)
    
        ' insertamos los datos de las celdas con datos sobre pedidos realizados
        ws.Cells(1, 1) = "RESPALDO DE PEDIDOS REALIZADOS EL " & UCase(FormatDateTime(Now, vbLongDate))
        origen.HeaderRowRange.Copy
        ws.Range("A2").PasteSpecial xlPasteValues, xlPasteSpecialOperationNone
        cont = 3
        With origen
            For Each fila In .ListRows
                If .DataBodyRange(fila.Index, .ListColumns("Pedido realizado").Index).Value > 0 Then
                    For i = 1 To .ListColumns.Count
                        ws.Cells(cont, i) = .DataBodyRange(fila.Index, i)
                    Next i
                    .DataBodyRange(fila.Index, .ListColumns("Pedido realizado").Index).ClearContents
                    .DataBodyRange(fila.Index, .ListColumns("Número de OC").Index).ClearContents
                    .DataBodyRange(fila.Index, .ListColumns("Excepción de compra").Index).ClearContents
                    .DataBodyRange(fila.Index, .ListColumns("Proveedor").Index).ClearContents
                    cont = cont + 1
                End If
            Next fila
        End With
        ws.Protect
    Else
        MsgBox "Se crearán columnas nuevas con el registro de los pedidos", vbInformation
        ' copiamos los datos de la tabla en columnas nuevas
        columnas(1) = "Proveedor"
        columnas(2) = "Excepción de compra"
        columnas(3) = "Número de OC"
        columnas(4) = "Pedido realizado"
        columnas(5) = "Clave"
        With origen
            ' insertamos las nuevas columnas
            For i = 1 To 5
                Generador.Columns(.ListColumns.Count + 2).EntireColumn.Insert
                Generador.Cells(7, .ListColumns.Count + 2) = columnas(i)
                Generador.Cells(7, .ListColumns.Count + 2).Interior.ColorIndex = 35
                Generador.Cells(7, .ListColumns.Count + 2).ColumnWidth = 12
                With Generador.Cells(7, .ListColumns.Count + 2)
                    .Borders(xlEdgeBottom).ColorIndex = 1
                    .Borders(xlEdgeTop).ColorIndex = 1
                    .Borders(xlEdgeLeft).ColorIndex = 1
                    .Borders(xlEdgeRight).ColorIndex = 1
                End With
            Next i
            Generador.Cells(6, .ListColumns.Count + 2) = "Pedidos realizados el " & WorksheetFunction.Text(Now, "dd/mm/yy")
            For Each fila In .ListRows
                Generador.Cells(fila.Index + 7, .ListColumns.Count + 2) = .DataBodyRange(fila.Index, .ListColumns("Pedido realizado").Index)
                Generador.Cells(fila.Index + 7, .ListColumns.Count + 3) = .DataBodyRange(fila.Index, .ListColumns("Número de OC").Index)
                Generador.Cells(fila.Index + 7, .ListColumns.Count + 4) = .DataBodyRange(fila.Index, .ListColumns("Excepción de compra").Index)
                Generador.Cells(fila.Index + 7, .ListColumns.Count + 5) = .DataBodyRange(fila.Index, .ListColumns("Proveedor").Index)
                Generador.Cells(fila.Index + 7, .ListColumns.Count + 6) = .DataBodyRange(fila.Index, .ListColumns("Clave").Index)
            Next fila
            Generador.Range(Cells(6, .ListColumns.Count + 2), Cells(.ListRows.Count + 1, .ListColumns.Count + 6)).Select
            With Selection
                .Borders(xlEdgeBottom).ColorIndex = 1
                .Borders(xlEdgeTop).ColorIndex = 1
                .Borders(xlEdgeLeft).ColorIndex = 1
                .Borders(xlEdgeRight).ColorIndex = 1
                .Borders(xlInsideVertical).ColorIndex = 1
                .Borders(xlInsideHorizontal).ColorIndex = 1
                .Interior.ColorIndex = 19
            End With
        End With
    End If
    Generador.Select
    XlResetSettings
    Generador.Protect AllowFiltering:=True
End Sub

' Función privada del libro de Excel que se ejecuta al abrir el libro
' Modificada para actualziar los datos y cargar la lista de grupos para filtro
' @author Antonio Alfredo Ramírez Ramírez
' @since 20/06/2020
' @version 1.0.0 20/05/2020
' @version 1.0.1 27/06/2020

Private Sub Workbook_Open()
    ActiveWorkbook.RefreshAll
    Application.EnableEvents = False
    Dim fila As ListRow
    
    With Generador
        .Unprotect
        .cbGrupo.Clear
        .cbGrupo.AddItem "Ninguno"
        With Separador.ListObjects("tSeparadores")
            ' cargamos los datos de los grupos sin repeticiones
            For Each fila In .ListRows
                If fila.Index = 1 Then
                    Generador.cbGrupo.AddItem .DataBodyRange(fila.Index, .ListColumns("Grupo").Index).Value
                ElseIf .DataBodyRange(fila.Index, .ListColumns("Grupo").Index) <> .DataBodyRange(fila.Index - 1, .ListColumns("Grupo").Index) Then
                    Generador.cbGrupo.AddItem .DataBodyRange(fila.Index, .ListColumns("Grupo").Index).Value
                End If
            Next fila
        End With
        .ListObjects("tGenerador").Range.AutoFilter
        .Protect AllowFiltering:=True
    End With
    Application.EnableEvents = True
    Generador.Select
End Sub

' Macro usada para respaldar los pedidos completos del mes que termina realizados en el generador,
' de compras, siempre y cuando los pedidos se hayan reportado en la misma hoja al generar
' columnas dentro del generador de compras
' @author Antonio Alfredo Ramírez Ramírez
' @since 01/08/2020
' @version 1.0.0 01/08/2020

Sub RespaldoMensual()
    Dim hoja As Worksheet
    Dim nombre As String
    Dim inicio As Integer
    Dim final As Integer
    FastWS
    Generador.Unprotect
    ' copiamos la hoja del generador completa
    nombre = InputBox("Ingrese el nombre de la hoja para respaldo mensual", "Nombre de respaldo")
    Generador.Copy After:=Agenda
    ActiveSheet.Name = nombre
    Cells.Copy
    Cells.PasteSpecial xlPasteValues, xlPasteSpecialOperationNone
    ActiveSheet.Protect nombre
    
    ' borramos las columnas que hayan sido agregadas por el reporte diario, excepto del último día
    inicio = Generador.ListObjects("tGenerador").Range.Columns.Count + 7
    final = Generador.ListObjects("tGenerador").Range.Columns.Count + 107
    Generador.Select
    Generador.Range(Cells(, inicio), Cells(, final)).EntireColumn.Delete
    Sheets(nombre).Select
    ActiveSheet.Protect (nombre)
    ActualizarDatosGenerador
    Generador.Select
    XlResetSettings
End Sub