//Consulta para obtener los datos de los separadores de subgrupos
//para los productos utilizados en AMASA 
//@author Antonio Alfredo Ramírez Ramírez
//@version 1.0.0 03/07/2020

let
    //Parametros para realizar la consulta SQL
    direccionIP = "192.168.1.240",
    base = "B1_Amasa_2019",
    sqlSeparador = "select ItemCode as Clave, ItmsGrpNam as Grupo, ItemName as Subgrupo from OITM with (nolock) inner join OITB on OITM.ItmsGrpCod = OITB.ItmsGrpCod where ItemName like '%----%' and (OITM.ItmsGrpCod <> 104 and OITM.ItmsGrpCod <> 107 and OITM.ItmsGrpCod <> 108 and OITM.ItmsGrpCod <> 109 and OITM.ItmsGrpCod <> 110 and OITM.ItmsGrpCod <> 120 and OITM.ItmsGrpCod <> 121) order by ItemCode asc",
    
    //Obtenemos los datos a través de una consulta SQL
    dbSeparador = Sql.Database(direccionIP, base, [Query=sqlSeparador]),
    inicio = Table.ReplaceValue(dbSeparador,"[-----", "", Replacer.ReplaceText, {"Subgrupo"}),
    final = Table.ReplaceValue(inicio, "-----]", "", Replacer.ReplaceText, {"Subgrupo"}),
    tipo = Table.TransformColumnTypes(final, {{"Clave", Int64.Type}})
in
    tipo