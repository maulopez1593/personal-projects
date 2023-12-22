SELECT 
	mard.matnr AS material_number,
	makt.maktx AS material_description,
	mard.werks AS plant,
	mard.lgort AS storage_location,
	mvke.prodh AS product_hierarchy,
	mara.satnr AS cross_plant_mrpc,
	mard.labst AS unrestricted,
	mard.insme AS qc_hold,
	marc.trame AS in_transit,
	SUM(COALESCE(vbbe.omeng, 0)) AS deliveries
FROM sap_hana_be_ecc_b1_ecc_b1_a_usable.mard mard
LEFT JOIN sap_hana_be_ecc_b1_ecc_b1_a_usable.vbbe vbbe
	ON vbbe.matnr = mard.matnr
	AND vbbe.werks = mard.werks
	AND vbbe.lgort = mard.lgort
	AND vbbe.vbtyp = 'J'
LEFT JOIN sap_hana_be_ecc_b1_ecc_b1_a_usable.marc marc
	ON marc.matnr = mard.matnr
	AND marc.werks = mard.werks
LEFT JOIN sap_hana_be_ecc_b1_ecc_b1_a_usable.mvke mvke
	ON mvke.matnr = mard.matnr 
	AND mvke.vkorg = 'US15'
	AND mvke.vtweg = 10
LEFT JOIN sap_hana_be_ecc_b1_ecc_b1_a_usable.mara mara
	ON mara.matnr = mard.matnr
WHERE mard.werks IN ('1001','1002','1025')
	AND mard.lgort IN ('0030','0033')
	AND mard.labst + mard.insme + marc.trame + COALESCE(vbbe.omeng, 0) > 0
GROUP BY mard.matnr, mard.werks, mard.lgort, mvke.prodh, mard.labst, mard.insme, marc.trame