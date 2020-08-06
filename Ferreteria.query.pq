//Consulta para obtener los datos resumidos desde SAP Business
//para los productos utilizados en AMASA, incluye las ventas, costos, 
//precios de productos de ferretería
//@author Antonio Alfredo Ramírez Ramírez
//@version 1.0.0 04/07/2020

let
    //Parametros independientes y consultas SQL para obtener datos
    DireccionIP = "192.168.1.240",
    Base = "B1_Amasa_2019",
    Sucursal = "04",
    sqlFinal = "%'",
    sqlDatos = "select ItemCode as Clave, ItemName as Descripcion, InvntryUom as UDM, ItmsGrpNam as Grupo, U_A_TIPO_PRODUCTO as Tipo from OITM with (nolock) inner join OITB on OITM.ItmsGrpCod = OITB.ItmsGrpCod where (OITM.ItmsGrpCod <> 104 and OITM.ItmsGrpCod <> 107 and OITM.ItmsGrpCod <> 108 and OITM.ItmsGrpCod <> 109 and OITM.ItmsGrpCod <> 110 and OITM.ItmsGrpCod <> 120 and OITM.ItmsGrpCod <> 121) and ItemName not like '%----%' and U_A_TIPO_PRODUCTO is not null",
    sqlCostos = "select OINM.ItemCode as Clave, OITW.AvgPrice as CostoPromedio from OINM with (nolock) inner join OITW on OINM.ItemCode = OITW.ItemCode and OINM.Warehouse = OITW.WhsCode inner join OITM on OINM.ItemCode = OITM.ItemCode inner join OWHS on OITW.WhsCode = OWHS.WhsCode where OITW.WhsCode like '" & Sucursal & sqlFinal,
    sqlInventario = "select ItemCode as Clave, InQty - OutQty as Existencias from OINM with (nolock) where CreateDate < GETDATE() and Warehouse like '" & Sucursal & sqlFinal,
    sqlNotas = "select RIN1.ItemCode as Clave, Quantity * NumPerMsr as Devolucion, Quantity * PriceAfVAT as Descuento from ORIN with (nolock) inner join RIN1 on ORIN.DocEntry = RIN1.DocEntry where ORIN.DocDate between DATEADD(day, - 120, GETDATE()) and GETDATE() and RIN1.ItemCode is not null and ORIN.CANCELED = 'N' and RIN1.BaseType in (13,-1) and BPLId = " & Sucursal,
    sqlVentas = "select INV1.ItemCode as Clave, Quantity * NumPerMsr as Salida, Quantity * PriceAfVAT as Venta from OINV with (nolock) inner join INV1 on OINV.DocEntry = INV1.DocEntry where OINV.DocDate between DATEADD(day, - 120, GETDATE()) and GETDATE() and INV1.ItemCode is not null and OINV.CANCELED = 'N' and BPLId = " & Sucursal,
    sqlAuto = "select ITM1.ItemCode as Clave, ITM1.Price as Precio from OITM with (nolock) inner join ITM1 on OITM.ItemCode = ITM1.ItemCode where ItemName not like '%-----%' and ITM1.PriceList ='01'",
    sqlProf = "select ITM1.ItemCode as Clave, ITM1.Price as Precio from OITM with (nolock) inner join ITM1 on OITM.ItemCode = ITM1.ItemCode where ItemName not like '%-----%' and ITM1.PriceList ='02'",
    sqlRvta = "select ITM1.ItemCode as Clave, ITM1.Price as Precio from OITM with (nolock) inner join ITM1 on OITM.ItemCode = ITM1.ItemCode where ItemName not like '%-----%' and ITM1.PriceList ='03'",
    sqlPiso = "select ITM1.ItemCode as Clave, ITM1.Price as Precio from OITM with (nolock) inner join ITM1 on OITM.ItemCode = ITM1.ItemCode where ItemName not like '%-----%' and ITM1.PriceList ='04'",
    sqlSucu = "select ITM1.ItemCode as Clave, ITM1.Price as Precio from OITM with (nolock) inner join ITM1 on OITM.ItemCode = ITM1.ItemCode where ItemName not like '%-----%' and ITM1.PriceList ='05'",

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

    //Obtenemos los datos de los precios de cada prodcuto
    PrecioAuto = Sql.Database(DireccionIP, Base, [Query=sqlAuto]),
    PrecioProf = Sql.Database(DireccionIP, Base, [Query=sqlProf]),
    PrecioRvta = Sql.Database(DireccionIP, Base, [Query=sqlRvta]),
    PrecioPiso = Sql.Database(DireccionIP, Base, [Query=sqlPiso]),
    PrecioSucu = Sql.Database(DireccionIP, Base, [Query=sqlSucu]),

    //Obtenemos los datos basicos de los productos
    dbDatos = Sql.Database(DireccionIP, Base, [Query=sqlDatos]),
    nulTipo = Table.ReplaceValue(dbDatos, null, "ST", Replacer.ReplaceValue, {"Tipo"}),

    //Agregamos los datos de las existencias a la tabla
    addInventario = Table.NestedJoin(nulTipo, {"Clave"}, Inventario, {"Clave"}, "Inventario", JoinKind.LeftOuter),
    expInventario = Table.ExpandTableColumn(addInventario, "Inventario", {"Existencias"}),

    //Agregamos los datos de los costos a la tabla
    addCostos = Table.NestedJoin(expInventario, {"Clave"}, Costos, {"Clave"}, "Costo", JoinKind.LeftOuter),
    expCostos = Table.ExpandTableColumn(addCostos, "Costo", {"Costo"}),

    //Agregamos los datos de ventas realizadas y la demanda a la tabla
    addVentas = Table.NestedJoin(expCostos, {"Clave"}, Ventas, {"Clave"}, "Ventas", JoinKind.LeftOuter),
    expVentas = Table.ExpandTableColumn(addVentas, "Ventas", {"Demanda", "Importe"}),
    Utilidad = Table.AddColumn(expVentas, "Utilidad", each [Importe] - ([Costo]*[Demanda])),

    //Agregamos los datos de los precios obtenidos por producto
    addAuto = Table.NestedJoin(Utilidad, {"Clave"}, PrecioAuto, {"Clave"}, "Auto", JoinKind.LeftOuter),
    expAuto = Table.ExpandTableColumn(addAuto, "Auto", {"Precio"}, {"Autoconstructor"}),
    addProf = Table.NestedJoin(expAuto, {"Clave"}, PrecioProf, {"Clave"}, "Prof", JoinKind.LeftOuter),
    expProf = Table.ExpandTableColumn(addProf, "Prof", {"Precio"}, {"Profesional"}),
    addRvta = Table.NestedJoin(expProf, {"Clave"}, PrecioRvta, {"Clave"}, "Rvta", JoinKind.LeftOuter),
    expRvta = Table.ExpandTableColumn(addRvta, "Rvta", {"Precio"}, {"Reventa"}),
    addPiso = Table.NestedJoin(expRvta, {"Clave"}, PrecioPiso, {"Clave"}, "Piso", JoinKind.LeftOuter),
    expPiso = Table.ExpandTableColumn(addPiso, "Piso", {"Precio"}, {"Piso"}),
    addSucu = Table.NestedJoin(expPiso, {"Clave"}, PrecioSucu, {"Clave"}, "Sucu", JoinKind.LeftOuter),
    expSucu = Table.ExpandTableColumn(addSucu, "Sucu", {"Precio"}, {"Sucursal"}),
    tpClave = Table.TransformColumnTypes(expSucu, {{"Clave", Int64.Type}}),

    //Ordenamos las columnas como se mostrarán en la tabla de Excel y quitamos los nulos
    nulos = Table.ReplaceValue(tpClave, null, 0, Replacer.ReplaceValue, {"Existencias", "Costo", "Demanda", "Importe", "Utilidad"}),
    errores = Table.ReplaceErrorValues(nulos, {{"Demanda", 0}, {"Importe", 0}, {"Utilidad", 0}}),
    Datos = Table.Sort(errores, {{"Clave", Order.Ascending}})
in
    Datos