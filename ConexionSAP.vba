' Función para realizar la conexión a SAP mediante el servidor SQL'
' Tomada de la versión para conectar a SAP creada por Ándres Carlos de AMASA'
' @author Antonio Alfredo Ramírez Ramírez'
' @since 19/05/2020'
' @version 1.0.0 19/05/2020'

' declaración de variables globales
Dim sap As SAPbobsCOM.Company
Dim producto As SAPbobsCOM.Items
Dim listaPrecio As SAPbobsCOM.Items_Prices
Dim error As Integer
Dim msgError As String

Sub ConectarSAP()
    FastWB True
    Set sap = New SAPbobsCOM.Company

    ' revisamos si existe la conexión'
    If sap.Connected = True Then
        sap.Disconnect
    End If
    
    With sap
        .Server = "SapSvr20" '"DamCa"
        .CompanyDB = "B1_Amasa_2019" '"X-PRUEBA"'
        .UserName = "manager"
        .Password = "SAPB1212"
        .LicenseServer = "SapSvr20" '"DamCa"
        .DbServerType = dst_MSSQL2012
        .Language = ln_Spanish_La
        error = .Connect()
        
        ' en caso de algún error se muestra el mensaje de error
        If error Then
            MsgBox "Error en la conexión:" & vbNewLine & .GetLastErrorDescription, vbCritical
        End If
    End With
    XlResetSettings
End Sub

' Procedimiento para actualizar los precios de acuerdo a lo generado en el determinador de precios
' para ferretería, adaptado de la función para actualziar precios creada por Ándres Carlos - AMASA
' @author Antonio Alfredo Ramírez Ramírez
' @since 23/05/2020
' @version 1.0.0 23/05/2020

Sub ActualizarPrecioSAP()
    Const COLS As Integer = 4
    Const MONEDA As String = "MXN"
    Dim fila As ListRow
    Dim columna As ListColumn
    Dim lista(4) As String
    lista(0) = "Autoconstructor"
    lista(1) = "Profesional"
    lista(2) = "Reventa"
    lista(3) = "Piso"
    lista(4) = "Sucursal"
    FastWB True
    
    ' obtenemos la conexion a SAP
    Call ConectarSAP
    Set producto = sap.GetBusinessObject(oItems)
    
    ' borramos el color interior de las celdas con precios
    With Precios.ListObjects("tPrecios")
        .ListColumns("Autoconstructor").DataBodyRange.Interior.ColorIndex = 0
        .ListColumns("Profesional").DataBodyRange.Interior.ColorIndex = 0
        .ListColumns("Reventa").DataBodyRange.Interior.ColorIndex = 0
        .ListColumns("Piso").DataBodyRange.Interior.ColorIndex = 0
        .ListColumns("Sucursal").DataBodyRange.Interior.ColorIndex = 0
    End With
    
    ' cargamos los nuevos precios en SAP
    For Each fila In Precios.ListObjects("tPrecios").ListRows
        With Precios.ListObjects("tPrecios")
            producto.GetByKey (CStr(.DataBodyRange(fila.Index, .ListColumns("Clave").Index).Value))
            Set listaPrecio = producto.PriceList
            For i = 0 To 4
                listaPrecio.SetCurrentLine (.ListColumns(lista(i)).Index - COLS)
                producto.PriceList.Currency = MONEDA
                producto.PriceList.Price = .DataBodyRange(fila.Index, .ListColumns(lista(i)).Index).Value
                error = producto.Update()
                If error <> 0 Then
                    .DataBodyRange(fila.Index, .ListColumns(lista(i)).Index).Interior.ColorIndex = 6
                Else
                    .DataBodyRange(fila.Index, .ListColumns(lista(i)).Index).Interior.ColorIndex = 4
                End If
            Next i
        End With
    Next fila
    
    ' nos desconectamos del servidor de la base de datos
    sap.Disconnect
    XlResetSettings
    MsgBox "Actualización completa con éxito", vbInformation
End Sub

' Procedimiento para actualizar el tipo de producto de ferretería, una vez que se han analizado
' los nuevos tipos que deben tener los productos de acuerdo con la demanda de dichos productos
' @author Antonio Alfredo Ramírez Ramírez
' @since 23/05/2020
' @version 1.0.0 23/05/2020
' @version 1.1.0 13/06/2020 <<se usa la tabla como objeto>>

Sub ActualizaTipoSAP()
    FastWB True
    Dim fila As ListRow
    Dim item As String
    
    ' obtenemos la conexión a SAP
    Call ConectarSAP
    Set producto = sap.GetBusinessObject(oItems)
    Analisis.ListObjects("tAnalisisTipo").ListColumns("Tipo nuevo").DataBodyRange.Interior.ColorIndex = 0
    
    ' cargamos los nuevos tipos desde Excel
    For Each fila In Analisis.ListObjects("tAnalisisTipo").ListRows
        With Analisis.ListObjects("tAnalisisTipo")
            producto.GetByKey (CStr(.DataBodyRange(fila.Index, .ListColumns("Clave").Index).Value))
            producto.UserFields.Fields.item("U_A_TIPO_PRODUCTO").Value = .DataBodyRange(fila.Index, .ListColumns("Tipo nuevo").Index).Value
            error = producto.Update()
            If error <> 0 Then
                .DataBodyRange(fila.Index, .ListColumns("Tipo nuevo").Index).Interior.ColorIndex = 6
            Else
                .DataBodyRange(fila.Index, .ListColumns("Tipo nuevo").Index).Interior.ColorIndex = 4
            End If
        End With
    Next fila
    XlResetSettings
End Sub
