' Macro usada para agregar los elementos a un combo de lista para usarlo como filtro
' en la hoja Generador y poder revisar con mayor facilidad los productos
' @author Antonio Alfredo Ramírez Ramírez
' @since 20/05/2020
' @version 1.0.0 20/05/2020

Private Sub cbGrupo_Change()
    Dim lista() As String
    Dim aux As Integer, cont As Integer
    On Error GoTo handler
    Generador.Unprotect
    
    ' obtenemos la cantidad de elementos correspondientes al grupo
    aux = WorksheetFunction.CountIf(Datos.Range("tDatos[Grupo]"), cbGrupo.Value)
    
    ' cargamos los datos de los subgrupos correspondientes al grupo seleccionado
    If cbGrupo.Value = "Ninguno" Then
        cbSubgrupo.Clear
        txtBuscar.Text = ""
        cbSubgrupo.AddItem "Seleccione un grupo para filtrar"
        Generador.ListObjects("tGenerador").Range.AutoFilter
    Else
        Generador.cbSubgrupo.Clear
        With Separador.ListObjects("tSeparadores")
            For Each fila In .ListRows
                If .DataBodyRange(fila.Index, .ListColumns("Grupo").Index).Value = cbGrupo.Value Then
                    Generador.cbSubgrupo.AddItem .DataBodyRange(fila.Index, .ListColumns("SubGrupo").Index).Value
                End If
            Next fila
        End With
        ' filtramos los datos según el grupo seleccionado
        With Datos.ListObjects("tDatos")
            ReDim lista(0 To aux)
            cont = 1
            ' recorremos los datos de la hoja en Datos
            For Each fila In .ListRows
                If .DataBodyRange(fila.Index, .ListColumns("Grupo").Index) = cbGrupo.Value Then
                    lista(cont) = .DataBodyRange(fila.Index, .ListColumns("Clave").Index).Value
                    cont = cont + 1
                End If
            Next fila
            Generador.Unprotect
            Generador.ListObjects("tGenerador").Range.AutoFilter 1, lista, xlFilterValues
        End With
    End If
handler:
    Generador.Protect AllowFiltering:=True
End Sub

' Macro usada para agregar los elementos a un combo de lista para usarlo como filtro
' en la hoja Generador y poder revisar con mayor facilidad los productos
' @author Antonio Alfredo Ramírez Ramírez
' @since 20/05/2020
' @version 1.0.0 20/05/2020

Private Sub cbSubgrupo_Change()
    Dim inicio As Long, final As Long
    Dim lista() As String
    Dim fila As ListRow
    On Error GoTo handler
    Generador.Unprotect
    
    ' encontramos el rango inicial para el subgrupo seleccionado
    If cbGrupo.Value = "Ninguno" Then
        Generador.ListObjects("tGenerador").Range.AutoFilter
    Else
        With Separador.ListObjects("tSeparadores")
            For Each fila In .ListRows
                If .DataBodyRange(fila.Index, .ListColumns("Subgrupo").Index).Value = cbSubgrupo.Value Then
                    inicio = .DataBodyRange(fila.Index, .ListColumns("Clave").Index).Value
                    final = .DataBodyRange(fila.Index + 1, .ListColumns("Clave").Index).Value
                End If
            Next fila
        End With
        Generador.ListObjects("tGenerador").Range.AutoFilter Field:=1, Criteria1:=">=" & inicio, Operator:=xlAnd, Criteria2:="<=" & final
    End If
handler:
    Generador.Protect AllowFiltering:=True
End Sub

' Para buscar un producto al estar escribiendo la descripción del mismo'
' @author Antonio Alfredo Ramirez Ramirez'
' @since 14ene2020'
' @version 1.0.0 14ene2020
' @version 1.1.0 16may2020

Private Sub txtBuscar_Change()
    Dim criterio As String
    FastWB True
    Generador.Unprotect
    On Error GoTo manejo
    
    ' revisamos si el cuadro de texto esta vacia
    If txtBuscar.Text <> "" Then
        ' se coloca la cadena que se escribe como criterio del filtro'
        criterio = "*" & txtBuscar.Text & "*"
        Range("A8").AutoFilter Field:=2, Criteria1:=criterio
    Else
        ' se borra el contenido del campor de texto y se quita el filtro'
        txtBuscar.Text = ""
        Range("A8").CurrentRegion.AutoFilter
    End If
manejo:
    XlResetSettings
    Generador.Protect AllowFiltering:=True
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
        With Separador.ListObjects("tSeparadores")
            ' cargamos los datos de los grupos sin repeticiones
            For Each fila In .ListRows
                If fila.Index = 1 Then
                    Determinador.cbGrupo.AddItem .DataBodyRange(fila.Index, .ListColumns("Grupo").Index).Value
                ElseIf .DataBodyRange(fila.Index, .ListColumns("Grupo").Index) <> .DataBodyRange(fila.Index - 1, .ListColumns("Grupo").Index) Then
                    Determinador.cbGrupo.AddItem .DataBodyRange(fila.Index, .ListColumns("Grupo").Index).Value
                End If
            Next fila
        End With
        .ListObjects("tDeterminador").Range.AutoFilter
        .Protect AllowFiltering:=True
    End With
    Application.EnableEvents = True
End Sub
