SELECT
	afko.aufnr AS order_,
	afpo.matnr AS material_number,
	aufk.aenam AS changed_by,
	afpo.psmng AS target_qty,
	afpo.wemng AS delivered_qty,
	afko.igmng AS confirmed_qty,
	aufk.aezeit AS changed_at,
	aufk.aedat AS changed_date,
	makt.maktg AS material_description,
	afko.gstrp AS basic_start_date,
CASE
	WHEN afko.igmng >= afpo.psmng THEN 'TECO'
	WHEN afpo.wemng = afpo.psmng THEN 'DLV'
	ELSE ''
END AS system_status,		
	afpo.kdauf AS sales_order,
	afpo.kdpos AS sales_item,
	afko.gltrp AS basic_fin_date,
	afko.geuzi AS actual_fin_time,
	afko.getri AS actual_fin_date,
	afko.dispo AS mrp_controller,
	afpo.pwerk AS planning_plant,
	mvke.prodh AS product_hierarchy
FROM sap_hana_be_ecc_b1_ecc_b1_a_usable.afko afko
INNER JOIN sap_hana_be_ecc_b1_ecc_b1_a_usable.afpo afpo
	ON afpo.aufnr = afko.aufnr 
INNER JOIN sap_hana_be_ecc_b1_ecc_b1_a_usable.aufk aufk
	ON aufk.aufnr = afko.aufnr 
LEFT JOIN sap_hana_be_ecc_b1_ecc_b1_a_usable.makt makt
	ON makt.matnr = afpo.matnr
LEFT JOIN sap_hana_be_ecc_b1_ecc_b1_a_usable.mvke mvke
	ON mvke.matnr = afpo.matnr 
	AND mvke.vkorg = 'US15'
	AND mvke.vtweg = 10
WHERE afpo.pwerk IN ('1002','1025','1001')
	AND afpo.lgort = '0033'
	AND afko.dispo NOT IN ('004','500','531','538','539','542','800','TEC')
	AND afpo.wemng < afpo.psmng 	--TECO and DLV excluded
	AND afpo.dnrel = '' 			--not relevant for MRP filter
	AND from_unixtime(unix_timestamp(afko.gltrp ,'yyyyMMdd'), 'yyyy-MM-dd') >= '2022-10-01'
UNION 
SELECT
	plaf.plnum AS order_,
	plaf.matnr AS material_number,
	NULL AS changed_by,
	plaf.gsmng AS target_qty,
	NULL AS delivered_qty,
	NULL AS confirmed_qty,
	NULL AS changed_at,
	NULL AS changed_date,
	NULL AS material_description,
	plaf.psttr AS basic_start_date,
	NULL AS system_status,
	plaf.kdauf AS sales_order,
	plaf.kdpos AS sales_item,
	plaf.pedtr AS basic_fin_date,
	NULL AS actual_fin_time,
	NULL AS actual_fin_date,
	plaf.dispo AS mrp_controller,
	plaf.plwrk AS planning_plant,
	mvke.prodh AS product_hierarchy
FROM sap_hana_be_ecc_b1_ecc_b1_a_usable.plaf plaf
LEFT JOIN sap_hana_be_ecc_b1_ecc_b1_a_usable.mvke mvke
	ON mvke.matnr = plaf.matnr 
	AND mvke.vkorg = 'US15'
	AND mvke.vtweg = 10
WHERE plaf.plwrk IN ('1002','1025','1001')
	AND plaf.lgort = '0033'
	AND plaf.dispo NOT IN ('004','500','531','538','539','542','800','TEC')
	AND from_unixtime(unix_timestamp(plaf.pedtr ,'yyyyMMdd'), 'yyyy-MM-dd') >= '2022-10-01'
