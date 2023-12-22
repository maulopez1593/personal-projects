SELECT 
	ser03.datum AS produced_date,
	ser03.uzeit AS produced_time,
	ser03.werk AS plant,
	objk.sernr AS serial_number,
	objk.matnr AS material,
	mvke.prodh AS prod_hierarchy,
	marc.dispo AS MRP
FROM sap_hana_be_ecc_b1_ecc_b1_a_usable.ser03 ser03
INNER JOIN sap_hana_be_ecc_b1_ecc_b1_a_usable.objk objk
	ON objk.obknr = ser03.obknr
	AND objk.obzae = '1'
	AND objk.taser = 'SER03'
LEFT JOIN sap_hana_be_ecc_b1_ecc_b1_a_usable.mvke mvke 
	ON mvke.matnr = objk.matnr 
	AND mvke.vkorg = 'US15'
	AND mvke.vtweg = 10
LEFT JOIN sap_hana_be_ecc_b1_ecc_b1_a_usable.marc marc 
	ON marc.matnr = objk.matnr
	AND marc.werks = ser03.werk 
WHERE 
	ser03.bwart = '101'
	AND ser03.werk IN ('1001','1002','1025')
	AND objk.matnr NOT LIKE 'HMH7%' 
	AND objk.matnr NOT LIKE 'MSH%' 
	AND objk.matnr NOT LIKE 'S1-%'
	AND from_unixtime(unix_timestamp(ser03.datum ,'yyyyMMdd'), 'yyyy-MM-dd') >= '2018-10-01' 
	AND ser03.datum < CONCAT(YEAR(CURRENT_DATE), LPAD(MONTH(CURRENT_DATE), 2, '0'))