SELECT
	REPLACE(LTRIM(REPLACE(likp.vbeln,'0',' ')),' ','0') AS delivery,
	lips.posnr AS delivery_line,
	lips.matnr AS material,
	lips.werks AS plant,
	marc.dispo AS mrpc,
	REPLACE(LTRIM(REPLACE(lips.vgbel,'0',' ')),' ','0') AS sales_order,
	lips.vgpos AS sales_item,
	REPLACE(LTRIM(REPLACE(likp.kunag,'0',' ')),' ','0') AS sold_to_party,
	kna1.name1 AS customer_name,
	--CONSIDER ADDING AN EXTRA COLUMN WITH CUSTOMER CONSOLIDATION LIKE JOHNSTONE WITH CASE STATEMENTS
	likp.wadat_ist AS ac_g_mvmt_date,
	likp.erdat AS created_date,
	vbap.zzopd AS promise_date,
	vbak.erdat AS sales_order_date,
	vbep.edatu AS delivery_schedule_date,
	marc.maabc AS abc_indicator,
	lips.lfimg AS delivery_quantity,
	lips.lfimg * vbap.netpr AS net_value,
	lips.prodh AS product_hierarchy,
	likp.lfdat AS delivery_date,
	vbak.zzjobsite AS jobsite,
	kna1.regio AS region,
	likp.bolnr AS bill_of_lading,
	likp.vsart AS shipping_type
FROM sap_hana_be_ecc_b1_ecc_b1_a_usable.likp likp
INNER JOIN sap_hana_be_ecc_b1_ecc_b1_a_usable.lips lips
	ON likp.vbeln = lips.vbeln
	AND lips.pstyv NOT IN ('ZXW2','ZXWM','ZXWN','ZRUN')
LEFT JOIN sap_hana_be_ecc_b1_ecc_b1_a_usable.marc marc
	ON lips.matnr = marc.matnr
	AND lips.werks = marc.werks
LEFT JOIN sap_hana_be_ecc_b1_ecc_b1_a_usable.kna1 kna1
	ON likp.kunag = kna1.kunnr
LEFT JOIN sap_hana_be_ecc_b1_ecc_b1_a_usable.vbak vbak
	ON lips.vgbel = vbak.vbeln
LEFT JOIN sap_hana_be_ecc_b1_ecc_b1_a_usable.vbap vbap 
	ON lips.vgbel = vbap.vbeln
	AND lips.vgpos = vbap.posnr
LEFT JOIN sap_hana_be_ecc_b1_ecc_b1_a_usable.vbep vbep
	ON lips.vgbel = vbep.vbeln
	AND lips.vgpos = vbep.posnr
	AND vbep.etenr = '0001'
LEFT JOIN sap_hana_be_ecc_b1_ecc_b1_a_usable.vttp vttp
	ON likp.vbeln = vttp.vbeln
LEFT JOIN sap_hana_be_ecc_b1_ecc_b1_a_usable.knvp knvp
	ON likp.kunag = knvp.kunnr
	AND knvp.vkorg = 'US15'
	AND knvp.vtweg = '10'
	AND knvp.parvw = 'SH'
	AND knvp.parza = ''
WHERE 
	lips.vkbur = 'US15'	
	AND lips.vtweg = '10'
	AND lips.werks IN ('1001','1002','1025')
	AND marc.dispo NOT IN ('800','240','538','539','542','100','700','270','500')
	AND from_unixtime(unix_timestamp(likp.wadat_ist ,'yyyyMMdd'), 'yyyy-MM-dd') >= '2022-10-01'
