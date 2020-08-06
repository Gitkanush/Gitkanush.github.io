//Consulta para obtener los datos resumidos desde SAP Business
//para los productos utilizados en AMASA, incluye las ventas, costos, 
//compras programadas y ventas programadas
//@author Antonio Alfredo Ramírez Ramírez
//@version 1.2.0 30/06/2020
//@version 1.2.1 01/07/2020
//@version 1.2.2 02/07/2020
//@version 1.2.3 03/07/2020

let
    //Parametros independientes y consultas SQL para obtener datos
    DireccionIP = "192.168.1.240",
    Base = "B1_Amasa_2019",
    Sucursal = "01",
    sqlFinal = "%'",
    sqlDatos = "select ItemCode as Clave, ItemName as Descripcion, InvntryUom as UDM, ItmsGrpNam as Grupo, U_A_TIPO_PRODUCTO as Tipo from OITM with (nolock) inner join OITB on OITM.ItmsGrpCod = OITB.ItmsGrpCod where (OITM.ItmsGrpCod <> 104 and OITM.ItmsGrpCod <> 107 and OITM.ItmsGrpCod <> 108 and OITM.ItmsGrpCod <> 109 and OITM.ItmsGrpCod <> 110 and OITM.ItmsGrpCod <> 120 and OITM.ItmsGrpCod <> 121) and ItemName not like '%----%' and U_A_TIPO_PRODUCTO is not null",
    sqlCompras = "select ItemCode as Clave, OpenQty * NumPerMsr as Compra from POR1 with (nolock) where DocDate < GETDATE() and LineStatus = 'O' and OpenQty > 0 and WhsCode like '" & Sucursal & sqlFinal,
    sqlCostos = "select OINM.ItemCode as Clave, OITW.AvgPrice as CostoPromedio from OINM with (nolock) inner join OITW on OINM.ItemCode = OITW.ItemCode and OINM.Warehouse = OITW.WhsCode inner join OITM on OINM.ItemCode = OITM.ItemCode inner join OWHS on OITW.WhsCode = OWHS.WhsCode where OITW.WhsCode like '" & Sucursal & sqlFinal,
    sqlPedidos = "select ItemCode as Clave, OpenQty * NumPerMsr as Pedido from RDR1 with (nolock) where DocDate < GETDATE() and LineStatus = 'O' and OpenQty > 0 and WhsCode like '" & Sucursal & sqlFinal,
    sqlInventario = "select ItemCode as Clave, InQty - OutQty as Existencias from OINM with (nolock) where CreateDate < GETDATE() and Warehouse like '" & Sucursal & sqlFinal,
    sqlNotas = "select RIN1.ItemCode as Clave, Quantity * NumPerMsr as Devolucion, Quantity * PriceAfVAT as Descuento from ORIN with (nolock) inner join RIN1 on ORIN.DocEntry = RIN1.DocEntry where ORIN.DocDate between DATEADD(day, - 120, GETDATE()) and GETDATE() and RIN1.ItemCode is not null and ORIN.CANCELED = 'N' and RIN1.BaseType in (13,-1) and BPLId = " & Sucursal,
    sqlVentas = "select INV1.ItemCode as Clave, Quantity * NumPerMsr as Salida, Quantity * PriceAfVAT as Venta from OINV with (nolock) inner join INV1 on OINV.DocEntry = INV1.DocEntry where OINV.DocDate between DATEADD(day, - 120, GETDATE()) and GETDATE() and INV1.ItemCode is not null and OINV.CANCELED = 'N' and BPLId = " & Sucursal,

    //Obtenemos los datos de las existencias de los productos
    dbInventario = Sql.Database(DireccionIP, Base, [Query=sqlInventario]),
    Inventario = Table.Group(dbInventario, {"Clave"},{{"Existencias", each List.Sum([Existencias]), Double.Type}}),

    //Obtenemos los datos de las ventas considerando las notas de crédito expedidas
    dbNotas = Sql.Database(DireccionIP, Base, [Query = sqlNotas]),
    totNotas = Table.Group(dbNotas, {"Clave"}, {{"Devolucion", each List.Sum([Devolucion]), type number}, {"Nota", each List.Sum([Descuento]), type number}}),
    dbVentas = Sql.Database(DireccionIP, Base, [Query=sqlVentas]),
    totVentas = Table.Group(dbVentas, {"Clave"}, {{"Salida", each List.Sum([Salida]), type number}, {"Venta", each List.Sum([Venta]), type number}}),
    cmbVentas = Table.NestedJoin(totVentas, {"Clave"}, totNotas, {"Clave"}, "Notas", JoinKind.LeftOuter),
    expComb = Table.ExpandTableColumn(cmbVentas, "Notas", {"Devolucion", "Nota"}),
    colDemanda = Table.AddColumn(expComb, "Demanda", each [Salida]-[Devolucion]),
    colImporte = Table.AddColumn(colDemanda, "Importe", each [Venta]-[Nota]),
    Ventas = Table.TransformColumnTypes(colImporte,{{"Demanda", Double.Type}, {"Importe", Double.Type}}),

    //Obtenemos los datos de los costos de los productos
    dbCostos = Sql.Database(DireccionIP, Base, [Query=sqlCostos]),
    Costos = Table.Group(dbCostos, {"Clave"}, {{"Costo", each List.Average([CostoPromedio]), Double.Type}}),

    //Obtenemos los datos de las compras programadas de productos
    dbCompras = Sql.Database(DireccionIP, Base, [Query=sqlCompras]),
    Compras = Table.Group(dbCompras, {"Clave"}, {{"Compras", each List.Sum([Compra]), Double.Type}}),

    //Obtenemos los datos de los pedidos de clientes programados
    dbPedidos = Sql.Database(DireccionIP, Base, [Query=sqlPedidos]),
    Pedidos = Table.Group(dbPedidos, {"Clave"}, {{"Pedidos", each List.Sum([Pedido]), Double.Type}}),

    //Obtenemos los datos de la venta atípica cargados en la tabla del generador de compras
    xlsAtipica = Excel.CurrentWorkbook(){[Name = "tGenerador"]}[Content],
    selAtipica = Table.SelectColumns(xlsAtipica, {"Clave", "Venta atípica"}),
    Atipica = Table.TransformColumnTypes(selAtipica, {{"Clave", Int64.Type}, {"Venta atípica", type number}}),

    //Obtenemos los parametros de los días de inventario mínimo y máximo
    xlsDias = Excel.CurrentWorkbook(){[Name = "tMaxMinDias"]}[Content],

    //Obtenemos los parametros extraordinarios para minimos y maximos
    xlsExtra = Excel.CurrentWorkbook(){[Name = "tExtraParam"]}[Content],

    //Obtenemos los parametros de capacidad de bodega para productos
    xlsCapacidad = Excel.CurrentWorkbook(){[Name = "tCapacidad"]}[Content],

    //Obtenemos los datos basicos de los productos
    dbDatos = Sql.Database(DireccionIP, Base, [Query=sqlDatos]),
    nulTipo = Table.ReplaceValue(dbDatos, null, "ST", Replacer.ReplaceValue, {"Tipo"}),

    //Agregamos los datos de las existencias a la tabla
    addInventario = Table.NestedJoin(nulTipo, {"Clave"}, Inventario, {"Clave"}, "Inventario", JoinKind.LeftOuter),
    expInventario = Table.ExpandTableColumn(addInventario, "Inventario", {"Existencias"}),

    //Agregamos los datos de los costos a la tabla
    addCostos = Table.NestedJoin(expInventario, {"Clave"}, Costos, {"Clave"}, "Costo", JoinKind.LeftOuter),
    expCostos = Table.ExpandTableColumn(addCostos, "Costo", {"Costo"}),

    //Agregamos los datos de las compras programadas a la tabla
    addCompras = Table.NestedJoin(expCostos, {"Clave"}, Compras, {"Clave"}, "Compras", JoinKind.LeftOuter),
    expCompras = Table.ExpandTableColumn(addCompras, "Compras", {"Compras"}),

    //Agregamos los datos de las ventas programadas a la tabla
    addPedidos = Table.NestedJoin(expCompras, {"Clave"}, Pedidos, {"Clave"}, "Pedidos", JoinKind.LeftOuter),
    expPedidos = Table.ExpandTableColumn(addPedidos, "Pedidos", {"Pedidos"}),

    //Agregamos los datos de ventas realizadas y la demanda a la tabla
    addVentas = Table.NestedJoin(expPedidos, {"Clave"}, Ventas, {"Clave"}, "Ventas", JoinKind.LeftOuter),
    expVentas = Table.ExpandTableColumn(addVentas, "Ventas", {"Demanda", "Importe"}),
    tpClave = Table.TransformColumnTypes(expVentas, {{"Clave", Int64.Type}}),

    //Agregamos los datos de las ventas atípicas para calcular la demanda diaria
    addAtipica = Table.NestedJoin(tpClave, {"Clave"}, Atipica, {"Clave"}, "Venta atípica", JoinKind.LeftOuter),
    expAtipica = Table.ExpandTableColumn(addAtipica, "Venta atípica", {"Venta atípica"}),
    nulAtipica = Table.ReplaceValue(expAtipica, null, 0, Replacer.ReplaceValue, {"Venta atípica"}),
    errParcial = Table.ReplaceErrorValues(nulAtipica, {{"Demanda", 0}, {"Importe", 0}}),

    //Obtenemos el valor de la demanda diaria considerando la venta atípica
    colDD = Table.AddColumn(errParcial, "Demanda diaria", each ([Demanda]-[Venta atípica])/120, Double.Type),
    nulDD = Table.ReplaceValue(colDD, null, 0, Replacer.ReplaceValue, {"Demanda diaria"}),
    invDias = Table.AddColumn(nulDD, "Inventario dias", each if [Existencias] = 0 then 0 else if [Demanda diaria] = 0 then 99999 else [Existencias]/[Demanda diaria], Double.Type),
    cmbDias = Table.NestedJoin(invDias, {"Grupo"}, xlsDias, {"Grupo"}, "Dias", JoinKind.LeftOuter),
    expDias = Table.ExpandTableColumn(cmbDias, "Dias", {"Minimo A", "Maximo A", "Minimo B", "Maximo B", "Minimo C", "Maximo C"}),
    minDias = Table.AddColumn(expDias, "Minimo dias", each if [Tipo] = "A" then [Minimo A] else if [Tipo] = "B" then [Minimo B] else [Minimo C], Double.Type),
    maxDias = Table.AddColumn(minDias, "Maximo dias", each if [Tipo] = "A" then [Maximo A] else if [Tipo] = "B" then [Maximo B] else [Maximo C], Double.Type),

    //Agregamos los datos extraordinarios sobre los minimos y maximos de productos seleccionados
    addExtra = Table.NestedJoin(maxDias, {"Clave"}, xlsExtra, {"Clave"}, "Extra", JoinKind.LeftOuter),
    expExtra = Table.ExpandTableColumn(addExtra, "Extra", {"Minimo", "Maximo"}),
    nulExtra = Table.ReplaceValue(expExtra, null, 0, Replacer.ReplaceValue, {"Minimo", "Maximo"}),
    tipMinimo = Table.TransformColumnTypes(nulExtra, {{"Minimo", Double.Type}}),
    tipMaximo = Table.TransformColumnTypes(tipMinimo, {{"Maximo", Double.Type}}),

    //Agregamos los datos sobre la capacidad de la bodega para cada producto
    addCapacidad = Table.NestedJoin(tipMaximo, {"Clave"}, xlsCapacidad, {"Clave"}, "Capacidad", JoinKind.LeftOuter),
    expCapacidad = Table.ExpandTableColumn(addCapacidad, "Capacidad", {"Capacidad"}),
    nulCapacidad = Table.ReplaceValue(expCapacidad, null, 0, Replacer.ReplaceValue, {"Capacidad"}),
    tipCapacidad = Table.TransformColumnTypes(nulCapacidad, {{"Capacidad", Double.Type}}),
    nulos = Table.ReplaceValue(tipCapacidad, null, 0, Replacer.ReplaceValue, {"Existencias", "Costo", "Compras", "Pedidos", "Demanda", "Importe", "Demanda diaria", "Inventario dias"}),
    minCant = Table.AddColumn(nulos, "Minimo cantidad", each if [Minimo] > 0 then [Minimo] else [Minimo dias]*[Demanda diaria], Double.Type),
    maxCant = Table.AddColumn(minCant, "Maximo cantidad", each if [Capacidad] > 0 then [Capacidad] else if [Maximo] > 0 then [Maximo] else [Maximo dias]*[Demanda diaria], Double.Type),
    acomodo = Table.ReorderColumns(maxCant, {"Clave", "Descripcion", "UDM", "Grupo", "Tipo", "Existencias", "Costo", "Compras", "Pedidos", "Demanda", "Importe", "Venta atípica", "Demanda diaria", "Inventario dias", "Minimo dias", "Maximo dias", "Minimo cantidad", "Maximo cantidad", "Minimo", "Maximo", "Capacidad", "Minimo A", "Maximo A", "Minimo B", "Maximo B", "Minimo C", "Maximo C"}),
    errores = Table.ReplaceErrorValues(acomodo, {{"Inventario dias", 0}, {"Minimo cantidad", 0}, {"Maximo cantidad", 0}}),
    Datos = Table.Sort(errores, {{"Clave", Order.Ascending}})
in
    Datos