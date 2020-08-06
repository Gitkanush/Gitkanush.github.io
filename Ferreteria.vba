' Macro para crear la lista de precios actualizada en la pestaña correspondiente,
' los encabezados para la tabla de precios ya están capturados y establecidos,
' cada uno de los campos se rellenará con la información pertinente.
' Esta macro toma como base la relacionada con el respaldo de compras diario'
' @author Antonio Alfredo Ramírez Ramírez'
' @since 01feb2020'
' @version 1.0.0 01feb2020
' @version 1.0.1 05feb2020
' @version 1.1.0 28feb2020
' @version 1.1.1 11abr2020
' @version 1.1.2 19may2020
' @version 1.1.3 15jun2020

Sub PreciosFerreteria()
    Dim fila As ListRow
    Dim rango As Integer
    Dim tDeterminador As ListObject
    Set tDeterminador = Determinador.ListObjects("tDeterminador")
    FastWB True
    Determinador.Unprotect
    Precios.Unprotect
    
    ' revisamos si se han actualizado los tipos de productos recientemente
    resp = MsgBox("Se han actualizado los tipos de productos.", vbYesNo)
    If resp = vbNo Then
        MsgBox "Es necesario revisar y/o" & vbNewLine & "Actualizar el tipo de producto", vbCritical
        CargarTipo
        Grupo.Select
    Else
        ' borramos los datos de los precios anteriores
        MsgBox "Se cargarán los nuevos precios en la lista" & vbNewLine & "Por favor, espere", vbInformation
        Precios.ListObjects("tPrecios").DataBodyRange.ClearContents
        Precios.lblTitulo.Caption = "LISTA ACTUALIZADA DE PRECIOS AL " & UCase(FormatDateTime(Now, vbLongDate))
        Precios.ListObjects("tPrecios").DataBodyRange.Interior.ColorIndex = 0
        
        ' cargamos los precios nuevos en la lista de precios y en analizador de precios
        i = 1
        For Each fila In tDeterminador.ListRows
            With Precios.ListObjects("tPrecios")
                If tDeterminador.DataBodyRange(fila.Index, tDeterminador.ListColumns("Precio Piso").Index).Value > 0# Then
                    ' se cargan los precios en la hoja para la lista de precios
                    .DataBodyRange(i, 1) = tDeterminador.DataBodyRange(fila.Index, tDeterminador.ListColumns("Clave").Index)
                    .DataBodyRange(i, 2) = tDeterminador.DataBodyRange(fila.Index, tDeterminador.ListColumns("Descripción").Index)
                    .DataBodyRange(i, 3) = tDeterminador.DataBodyRange(fila.Index, tDeterminador.ListColumns("UDM").Index)
                    .DataBodyRange(i, 4) = tDeterminador.DataBodyRange(fila.Index, tDeterminador.ListColumns("Precio Auto Constructor").Index)
                    .DataBodyRange(i, 5) = tDeterminador.DataBodyRange(fila.Index, tDeterminador.ListColumns("Precio Profesional").Index)
                    .DataBodyRange(i, 6) = tDeterminador.DataBodyRange(fila.Index, tDeterminador.ListColumns("Precio Reventa").Index)
                    .DataBodyRange(i, 7) = tDeterminador.DataBodyRange(fila.Index, tDeterminador.ListColumns("Precio Piso").Index)
                    .DataBodyRange(i, 8) = tDeterminador.DataBodyRange(fila.Index, tDeterminador.ListColumns("Precio Sucursal").Index)
                    i = i + 1
                End If
            End With
        Next fila
        CargarPrecios
        
        ' actualizamos los precios en SAP
        resp = MsgBox("¿Desea cargar los nuevos precios en SAP?", vbYesNo)
        If resp = vbYes Then
            ActualizarPrecioSAP
        Else
            Precios.Select
        End If
        
        ' generamos el archivo PDF con la nueva lista de precios sin el precio sucursal
        resp = MsgBox("¿Desea generar la lista de precios en PDF?", vbYesNo)
        If resp = vbYes Then
            With Precios
                .ListObjects("tPrecios").DataBodyRange.Interior.ColorIndex = 0

                ' ordenamos los datos por la descrición en orden alfabetico
                .ListObjects("tPrecios").Sort.SortFields.Add Key:=Range("tPrecios[[#All],[Descripción]]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                With .ListObjects("tPrecios").Sort
                    .Header = xlYes
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With
                rango = .ListObjects("tPrecios").ListRows.Count
                .ExportAsFixedFormat xlTypePDF, Filename:="\\dc3\Ferreteria\Listas de Precios PDF-XLSX-Ferreteria\DEO-COM-ListaDePrecios-" & Day(Date) & "-" & Month(Date) & "-" & Year(Date) & ".pdf", Quality:=xlQualityStandard, OpenAfterPublish:=True
            End With
        Else
            Precios.Select
        End If
    End If
    XlResetSettings
    Determinador.Protect AllowFiltering:=True
    Precios.Protect AllowFiltering:=True
End Sub

' Macro para realizar la actualización de los datos desde SAP optimizada
' para una respuesta más rápida del archivo y SAP.
' Esta macro es usada para el archivo determinador de precios.
' @author Antonio Alfredo Ramírez Ramírez
' @since 18 ene 2020
' @version 1.0.0 18ene2020
' @version 1.1.0 05feb2020
' @version 1.2.0 15feb2020
' @version 1.2.1 29feb2020
' @version 1.2.2 28mar2020
' @version 1.2.3 18abr2020
' @version 1.2.4 16may2020
' @version 1.2.5 15jun2020

Sub ActualizarDatosFerreteria()
    FastWB True
    Dim fila As ListRow
    Dim tDeterminador As ListObject
    Dim i As Integer, dato As Integer
    Set tDeterminador = Determinador.ListObjects("tDeterminador")
    
    Determinador.Unprotect
    MsgBox "Se actualizará la lista de productos." & vbNewLine & "Por favor, espere", vbInformation
    
    'actualizamos los datos de la hoja principal de datos y respaldamos datos del determinador
    With tDeterminador
        .ListColumns("Clave").DataBodyRange.Copy
        Datos.ListObjects("tRespaldo").ListColumns("Clave").DataBodyRange.PasteSpecial xlPasteValues, xlPasteSpecialOperationNone
        .ListColumns("Ajuste por %").DataBodyRange.Copy
        Datos.ListObjects("tRespaldo").ListColumns("Ajuste por %").DataBodyRange.PasteSpecial xlPasteValues, xlPasteSpecialOperationNone
        .ListColumns("Clave").DataBodyRange.ClearContents
        .ListColumns("Ajuste por %").DataBodyRange.ClearContents
    End With
    
    ' actualizamos los datos en el determinador
    i = 1
    With Datos.ListObjects("tDatos")
        For Each fila In .ListRows
            If .DataBodyRange(fila.Index, .ListColumns("Grupo").Index).Value = "51-FERRETERIA" Or .DataBodyRange(fila.Index, .ListColumns("Grupo").Index).Value = "61-PISOS-RECUB-ACC" Then
                If Not .DataBodyRange(fila.Index, .ListColumns("Tipo").Index).Value = "D" Then
                    If .DataBodyRange(fila.Index, .ListColumns("Costo").Index).Value > 0 Then
                        tDeterminador.DataBodyRange(i, tDeterminador.ListColumns("Clave").Index) = .DataBodyRange(fila.Index, .ListColumns("Clave").Index)
                        i = i + 1
                    End If
                End If
            End If
        Next fila
    End With

    ' obtenemos los datos del porcentaje de ajuste
    i = 1
    For Each fila In tDeterminador.ListRows
        If Datos.ListObjects("tRespaldo").DataBodyRange(i, 1) = tDeterminador.DataBodyRange(fila.Index, tDeterminador.ListColumns("Clave").Index) Then
            tDeterminador.DataBodyRange(fila.Index, tDeterminador.ListColumns("Ajuste por %").Index) = Datos.ListObjects("tRespaldo").DataBodyRange(i, 2)
            i = i + 1
        End If
    Next fila
    MsgBox "Proceso terminado con éxito", vbInformation
    Determinador.Protect AllowFiltering:=True
    Determinador.Select
    XlResetSettings
End Sub

' Macro usada para cargar los datos de los productos de ferretería en la tabla
' para lizar los tipos de productos.
' @author Antonio Alfredo Ramírez Ramírez
' @since 06/06/2020
' @version 1.0.0 06/06/2020
' @version 1.0.1 13/06/2020

Sub CargarTipo()
    FastWB True
    Dim fila As ListRow
    Dim cont As Integer
    Dim tipo As ListObject

    Set tipo = Grupo.ListObjects("tAnalisisTipo")
    Grupo.Unprotect
    Grupo.ListObjects("tAnalisisTipo").Sort.SortFields.Clear
    
    ' borramos los datos que tiene la tabla de Grupo
    With tipo
        .ListColumns("Clave").DataBodyRange.ClearContents
        .ListColumns("Descripción").DataBodyRange.ClearContents
        .ListColumns("UDM").DataBodyRange.ClearContents
        .ListColumns("Ventas").DataBodyRange.ClearContents
        .ListColumns("Venta acumulada").DataBodyRange.ClearContents
        .ListColumns("Utilidad").DataBodyRange.ClearContents
        .ListColumns("Utilidad acumulada").DataBodyRange.ClearContents
        .ListColumns("Tipo actual").DataBodyRange.ClearContents
        .ListColumns("Tipo nuevo").DataBodyRange.ClearContents
        .ListColumns("Tipo nuevo").DataBodyRange.Interior.ColorIndex = 0
        .ListColumns("% Acumulado").DataBodyRange.ClearContents
    End With
    MsgBox "Se cargarán los datos de productos para analizar su tipo", vbInformation

    ' cargamos los datos de los productos desde el determinador
    With Determinador.ListObjects("tDeterminador")
        .ListColumns("Clave").DataBodyRange.Copy
        tipo.ListColumns("Clave").DataBodyRange.PasteSpecial xlPasteValues, xlPasteSpecialOperationNone
        .ListColumns("Descripción").DataBodyRange.Copy
        tipo.ListColumns("Descripción").DataBodyRange.PasteSpecial xlPasteValues, xlPasteSpecialOperationNone
        .ListColumns("UDM").DataBodyRange.Copy
        tipo.ListColumns("UDM").DataBodyRange.PasteSpecial xlPasteValues, xlPasteSpecialOperationNone
        .ListColumns("Tipo").DataBodyRange.Copy
        tipo.ListColumns("Tipo actual").DataBodyRange.PasteSpecial xlPasteValues, xlPasteSpecialOperationNone
    End With
    
    ' cargamos las ventas y calculamos las ventas acumuladas asi como su utilidad
    cont = 1
    With Datos.ListObjects("tDatos")
        For Each fila In .ListRows
            If .DataBodyRange(fila.Index, .ListColumns("Clave").Index) = tipo.DataBodyRange(cont, tipo.ListColumns("Clave").Index) Then
                tipo.DataBodyRange(cont, tipo.ListColumns("Ventas").Index) = .DataBodyRange(fila.Index, .ListColumns("Importe").Index)
                tipo.DataBodyRange(cont, tipo.ListColumns("Utilidad").Index) = .DataBodyRange(fila.Index, .ListColumns("Utilidad").Index)
                cont = cont + 1
            End If
        Next fila
    End With
    XlResetSettings
    PrepararAnalisis
    MsgBox "Datos cargados correctamente" & vbNewLine & "Proceda a hacer el análisis", vbExclamation
    Grupo.Protect AllowFiltering:=True
End Sub

' Macro para cargar los datos de los precios actualizados en la tabla para
' realizar el análisis de los precios con los antiguos y los nuevos
' @author Antonio Alfredo Ramírez Ramírez
' @since 07/07/2020
' @version 1.0.0 07/07/2020

Sub CargarPrecios()
    FastWB True
    Dim cont As Integer
    Dim analisis As ListObject
    
    APrecios.Unprotect
    Set analisis = APrecios.ListObjects("tAnalisisPrecios")
    
    ' borramos los datos actuales de la tabla de analisis de precio
    With analisis
        .ListColumns("Clave").DataBodyRange.ClearContents
        .ListColumns("Descripción").DataBodyRange.ClearContents
        .ListColumns("Precio Antiguo Autoconstructor").DataBodyRange.ClearContents
        .ListColumns("Precio Nuevo Autoconstructor").DataBodyRange.ClearContents
        .ListColumns("Precio Antiguo Profesional").DataBodyRange.ClearContents
        .ListColumns("Precio Nuevo Profesional").DataBodyRange.ClearContents
        .ListColumns("Precio Antiguo Reventa").DataBodyRange.ClearContents
        .ListColumns("Precio Nuevo Reventa").DataBodyRange.ClearContents
        .ListColumns("Precio Antiguo Piso").DataBodyRange.ClearContents
        .ListColumns("Precio Nuevo Piso").DataBodyRange.ClearContents
        .ListColumns("Precio Antiguo Sucursal").DataBodyRange.ClearContents
        .ListColumns("Precio Nuevo Sucursal").DataBodyRange.ClearContents
    End With
    MsgBox "Se cargarán los precios actualizados para su análisis" & vbNewLine & "Por favor, espere", vbInformation
      
    ' cargamos los precios nuevos en la tabla de analisis de precios
    With Precios.ListObjects("tPrecios")
        .ListColumns("Clave").DataBodyRange.Copy
        analisis.ListColumns("Clave").DataBodyRange.PasteSpecial xlPasteValues, xlPasteSpecialOperationNone
        .ListColumns("Descripción").DataBodyRange.Copy
        analisis.ListColumns("Descripción").DataBodyRange.PasteSpecial xlPasteValues, xlPasteSpecialOperationNone
        .ListColumns("UDM").DataBodyRange.Copy
        analisis.ListColumns("UDM").DataBodyRange.PasteSpecial xlPasteValues, xlPasteSpecialOperationNone
        .ListColumns("Autoconstructor").DataBodyRange.Copy
        analisis.ListColumns("Precio Nuevo Autoconstructor").DataBodyRange.PasteSpecial xlPasteValues, xlPasteSpecialOperationNone
        .ListColumns("Profesional").DataBodyRange.Copy
        analisis.ListColumns("Precio Nuevo Profesional").DataBodyRange.PasteSpecial xlPasteValues, xlPasteSpecialOperationNone
        .ListColumns("Reventa").DataBodyRange.Copy
        analisis.ListColumns("Precio Nuevo Reventa").DataBodyRange.PasteSpecial xlPasteValues, xlPasteSpecialOperationNone
        .ListColumns("Piso").DataBodyRange.Copy
        analisis.ListColumns("Precio Nuevo Piso").DataBodyRange.PasteSpecial xlPasteValues, xlPasteSpecialOperationNone
        .ListColumns("Sucursal").DataBodyRange.Copy
        analisis.ListColumns("Precio Nuevo Sucursal").DataBodyRange.PasteSpecial xlPasteValues, xlPasteSpecialOperationNone
    End With
    
    ' cargamos los precios antiguos desde la tabla de datos
    cont = 1
    With Datos.ListObjects("tDatos")
        For Each fila In .ListRows
            If .DataBodyRange(fila.Index, .ListColumns("Clave").Index) = Precios.ListObjects("tPrecios").DataBodyRange(cont, 1) Then
                analisis.DataBodyRange(cont, analisis.ListColumns("Precio Antiguo Autoconstructor").Index) = .DataBodyRange(fila.Index, .ListColumns("Autoconstructor").Index)
                analisis.DataBodyRange(cont, analisis.ListColumns("Precio Antiguo Profesional").Index) = .DataBodyRange(fila.Index, .ListColumns("Profesional").Index)
                analisis.DataBodyRange(cont, analisis.ListColumns("Precio Antiguo Reventa").Index) = .DataBodyRange(fila.Index, .ListColumns("Reventa").Index)
                analisis.DataBodyRange(cont, analisis.ListColumns("Precio Antiguo Piso").Index) = .DataBodyRange(fila.Index, .ListColumns("Piso").Index)
                analisis.DataBodyRange(cont, analisis.ListColumns("Precio Antiguo Sucursal").Index) = .DataBodyRange(fila.Index, .ListColumns("Sucursal").Index)
                cont = cont + 1
            End If
        Next fila
    End With
    MsgBox "Proceso terminado con éxito.", vbInformation
    APrecios.Protect AllowFiltering:=True
    XlResetSettings
End Sub

' Macro para realizar el análisis de los datos de ventas
' y tener datos para cambiar o mantener el tipo de productos
' @author Antonio Alfredo Ramírez Ramírez
' @since 06/08/2020
' @version 1.0.2 06/08/2020

Sub PrepararAnalisis()
    FastWB True
    Dim tipo As ListObject
    Dim fila As Integer
    
    Set tipo = Grupo.ListObjects("tAnalisisTipo")
    Grupo.Unprotect
    ' ordenamos los datos por la descrición en orden alfabetico
    tipo.Sort.SortFields.Add Key:=Range("tAnalisisTipo[[#All],[Utilidad]]"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With tipo.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ' revisamos cada una de las cantidades de ventas para calcular el acumulado
    With tipo
        For fila = 7 To .ListRows.Count + 6
            If fila = 7 Then
                Grupo.Cells(fila, .ListColumns("Venta acumulada").Index) = Grupo.Cells(fila, .ListColumns("Ventas").Index)
                Grupo.Cells(fila, .ListColumns("Utilidad acumulada").Index) = Grupo.Cells(fila, .ListColumns("Utilidad").Index)
                Grupo.Cells(fila, .ListColumns("% Acumulado utilidad").Index) = Grupo.Cells(fila, .ListColumns("Ventas").Index) / Grupo.Range("VentasTotal")
            Else
                Grupo.Cells(fila, .ListColumns("Venta acumulada").Index) = Grupo.Cells(fila, .ListColumns("Ventas").Index) + Grupo.Cells(fila - 1, .ListColumns("Ventas").Index)
                Grupo.Cells(fila, .ListColumns("Utilidad acumulada").Index) = Grupo.Cells(fila, .ListColumns("Utilidad").Index) + Grupo.Cells(fila - 1, .ListColumns("Utilidad").Index)
                Grupo.Cells(fila, .ListColumns("% Acumulado utilidad").Index) = Grupo.Cells(fila, .ListColumns("Ventas").Index) / Grupo.Range("VentasTotal") + Grupo.Cells(fila - 1, .ListColumns("% Acumulado utilidad").Index)
            End If
        Next fila
    End With
    XlResetSettings
    Grupo.Protect AllowFiltering:=True
End Sub

Option Explicit

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
    
    With Determinador
        .Unprotect
        .cbGrupo.Clear
        .cbGrupo.AddItem "Ninguno"
        Determinador.cbGrupo.AddItem "51-FERRETERIA"
        Determinador.cbGrupo.AddItem "61-PISOS-RECUB-ACC"
        .ListObjects("tDeterminador").Range.AutoFilter
        .Protect AllowFiltering:=True
    End With
    Application.EnableEvents = True
End Sub