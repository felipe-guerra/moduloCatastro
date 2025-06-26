
--Remuneracion basica unificada
select top 1 Rbu_Año, Rbu_Valor from RBU order by Rbu_Valor desc

--Predios rurales,Tarifa municipal
select * from PARAMETROS_GENERALES 
where CoeDeE_Codigo = '0202' and CoeDeE_Estado = 'R'

--Predios rurales, por mil
select CoeDeE_Valor/1000 from PARAMETROS_GENERALES 
where CoeDeE_Codigo = '0202' and CoeDeE_Estado = 'R'

--Predios rurales, Tasa administrativa
select * from PARAMETROS_GENERALES where CoeDeE_Codigo = '0203' 

--Predios rurales, Calculos excenciones para predios rurales
select * from PARAMETROS_GENERALES where CoeDeE_Codigo = '0208'

--Predios rurales, Bombero, Si se tiene convenio para el cobro 
--del impuesto de bomberos
select * from PARAMETROS_GENERALES where CoeDeE_Codigo = '0210'
and CoeDeE_Estado = 'R'

--Predios rurales, Si se desea que las exenciones se apliquen 
--para el impuesto de bomberos
select * from PARAMETROS_GENERALES where CoeDeE_Codigo = '0211'
and CoeDeE_Estado = 'R'

--Nuevos valores a cobrar rural:
--VALOR 1
select * from PARAMETROS_GENERALES where CoeDeE_Codigo = '0215'
and CoeDeE_Estado = 'R'

--VALOR 2
select * from PARAMETROS_GENERALES where CoeDeE_Codigo = '0216'
and CoeDeE_Estado = 'R'

--Base porcentual de prestamos(RURAL / URBANO)
select * from PARAMETROS_GENERALES where CoeDeE_Codigo = '0207'

--Depreciacion de edificacion
--Text1 (2025 1/2)	1:Simulacion	2:Real
select * from PARAMETROS_GENERALES where CoeDeE_Codigo = '0902'

--Predios publicos/Exoneracion
--Valor1
select * from PARAMETROS_GENERALES where CoeDeE_Codigo = '1101'

--Valor2
select * from PARAMETROS_GENERALES where CoeDeE_Codigo = '1102'






