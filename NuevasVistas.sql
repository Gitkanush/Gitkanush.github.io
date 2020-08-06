-- Archivo con las consultas para datos desde SAP usadas para el archivo maestro de 
-- datos, mismo que obtiene datos desde una vista SQL creada en SAP.
-- @author Antonio Alfredo Ramírez Ramírez
-- @since 07dic2019.
-- @version 1.0.0 07/12/2019
-- @version 1.1.0 21/12/2019
-- @version 1.2.0 30/06/2020 

-- Consulta de datos principales desde SAP
-- @author Antonio Alfredo Ramirez Ramirez
-- @since 07dic2019

select
   ItemCode as Clave,
   ItemName as Descripcion,
   InvntryUom as UDM,
   ItmsGrpNam as Grupo,
   U_A_TIPO_PRODUCTO as Tipo 
from
   OITM 
   inner join
      OITB 
      on OITM.ItmsGrpCod = OITB.ItmsGrpCod 
where
   (
      OITM.ItmsGrpCod <> 104 
      and OITM.ItmsGrpCod <> 107 
      and OITM.ItmsGrpCod <> 108 
      and OITM.ItmsGrpCod <> 109 
      and OITM.ItmsGrpCod <> 110 
      and OITM.ItmsGrpCod <> 120 
      and OITM.ItmsGrpCod <> 121
   )
   and ItemName not like '%----%' 
   and U_A_TIPO_PRODUCTO is not null

-- Consulta para los separadores de productos SAP
-- @author Antonio Alfredo Ramirez Ramirez
-- @since 14dic2019

select
   ItemCode as Clave,
   ItmsGrpNam as Grupo,
   ItemName as Subgrupo 
from
   OITM 
   inner join
      OITB 
      on OITM.ItmsGrpCod = OITB.ItmsGrpCod 
where
   ItemName like '%----%' 
   and 
   (
      OITM.ItmsGrpCod <> 104 
      and OITM.ItmsGrpCod <> 107 
      and OITM.ItmsGrpCod <> 108 
      and OITM.ItmsGrpCod <> 109 
      and OITM.ItmsGrpCod <> 110 
      and OITM.ItmsGrpCod <> 120 
      and OITM.ItmsGrpCod <> 121
   )
order by
   ItemCode asc

-- Consulta para las existencias actualizadas desde SAP
-- @author Antonio Alfredo Ramirez Ramirez
-- @since 07dic2019

select
   ItemCode as Clave,
   InQty - OutQty as Existencias
from
   OINM 
where
   CreateDate < GETDATE() 
   and Warehouse like '01%'

-- Consulta para datos de ventas desde SAP
-- @author Antonio Alfredo Ramirez Ramirez
-- @since 14dic2019

select
   INV1.ItemCode as Clave,
   Quantity * NumPerMsr as Salida,
   Quantity * PriceAfVAT as Venta 
from
   OINV 
   inner join
      INV1 
      on OINV.DocEntry = INV1.DocEntry 
where
   OINV.DocDate between DATEADD(day, - 120, GETDATE()) and GETDATE() 
   and INV1.ItemCode is not null 
   and OINV.CANCELED = 'N' 
   and BPLId = 1

-- Consulta para las notas de crédito o descuentos sobre ventas
-- @author Antonio Alfredo Ramirez Ramirez
-- @since 14dic2019

select
   RIN1.ItemCode as Clave,
   Quantity * NumPerMsr as Devolucion,
   Quantity * PriceAfVAT as Descuento 
from
   ORIN 
   inner join
      RIN1 
      on ORIN.DocEntry = RIN1.DocEntry 
where
   ORIN.DocDate between DATEADD(day, - 120, GETDATE()) and GETDATE() 
   and RIN1.ItemCode is not null 
   and ORIN.CANCELED = 'N' 
   and RIN1.BaseType in 
   (13, - 1)
   and BPLId = 1

-- Consulta para obtener el costo promedio de los productos a 120 días
-- @author Antonio Alfredo Ramirez Ramirez
-- @since 14dic2019

select
   OINM.ItemCode as Clave,
   OITW.AvgPrice as CostoPromedio 
from
   OINM 
   inner join
      OITW 
      on OINM.ItemCode = OITW.ItemCode 
      and OINM.Warehouse = OITW.WhsCode 
   inner join
      OITM 
      on OINM.ItemCode = OITM.ItemCode 
   inner join
      OWHS 
      on OITW.WhsCode = OWHS.WhsCode 
where
   OITW.WhsCode like '01%'

-- Consulta para obtener las cantidades solicitadas de productos a proveedores
-- @author Antonio Alfredo Ramirez Ramirez
-- @since 21dic2019

select
   ItemCode as Clave,
   OpenQty * NumPerMsr as Compra 
from
   POR1 
where
   DocDate < GETDATE() 
   and LineStatus = 'O' 
   and OpenQty > 0 
   and WhsCode like '01%'

-- Consulta para obtener las cantidades solicitadas de productos por clientes
-- @author Antonio Alfredo Ramirez Ramirez
-- @since 21dic2019

select
   ItemCode as Clave,
   OpenQty * NumPerMsr as Pedido
from
   RDR1 
where
   DocDate < GETDATE() 
   and LineStatus = 'O' 
   and OpenQty > 0 
   and WhsCode like '01%'

-- Consulta para obtener los datos de los precios de productos cargados en SAP
-- La lista de precios puede variar según la lista que se desea analizar
-- 01-LP Zac-Gpe Autoconstructor.      02-LP Zac-Gpe Profesional.    03-LP Zac-Gpe Reventa.
-- 04-LP Zac-Gpe Piso.              05-LP Zac-Gpe Sucursal.       06-LP Zac-Gpe Santa Rita
-- 07-LP Zac-Gpe Municipio Guadalupe.  08-LP Zac-Gpe Casa Díaz.      09-LP Zac-Gpe Piso FORANEOS.
-- 11-LP Fllo Autoconstructor.         12-LP Fllo Profesional.       13-LP Fllo Reventa.
-- 14-LP Fllo Piso.                 16-LP Ags Autoconstructor.    17-LP Ags Profesional.
-- 18-LP Ags Reventa.               19-LP Ags Piso.               21-LP Ags Oliver Miranda.
-- @author Antonio Alfredo Ramirez Ramirez
-- @since 04/07/2020

select
   ITM1.ItemCode as Codigo,
   ITM1.Price as Precio
from
   OITM 
   inner join
      ITM1 
      on OITM.ItemCode = ITM1.ItemCode 
where
   ItemName not like '%-----%' 
   and ITM1.PriceList = '01'



