SELECT 
	ekpo.ebeln AS purchase_order,
	ekpo.ebelp AS item,
	ekko.aedat AS po_created_on,
	ekko.lifnr AS vendor,
	lfa1.name1 AS vendor_name,
	ekpo.matnr AS material_number,
	ekpo.werks AS plant,
	ekpo.lgort AS storage_location,
	ekpo.menge AS po_quantity,
	eket.wemng AS delivered_qty,
	ekpo.menge - eket.wemng AS po_open_qty,
	eket.eindt AS delivery_date,
	eket.etenr AS schedule_line,
	mvke.prodh AS product_hierarchy	
FROM sap_hana_be_ecc_b1_ecc_b1_a_usable.ekpo ekpo
LEFT JOIN sap_hana_be_ecc_b1_ecc_b1_a_usable.ekko ekko
	ON ekko.ebeln = ekpo.ebeln
LEFT JOIN sap_hana_be_ecc_b1_ecc_b1_a_usable.lfa1 lfa1
	ON lfa1.lifnr = ekko.lifnr
LEFT JOIN sap_hana_be_ecc_b1_ecc_b1_a_usable.eket eket
	ON eket.ebeln = ekpo.ebeln	
	AND eket.ebelp = ekpo.ebelp
	AND eket.etenr = '0001'
LEFT JOIN sap_hana_be_ecc_b1_ecc_b1_a_usable.mvke mvke 
	ON mvke.matnr = ekpo.matnr 
	AND mvke.vkorg = 'US15'
	AND mvke.vtweg = 10
WHERE ekpo.bukrs IN ('US13','US15')
	AND ekpo.werks IN ('1001','1002','1025')
	AND ekpo.lgort = '0033'
--	AND ekko.lifnr IN ('0005059227', '0005011756') -- filtered by purchased skus only (hisense and aspen coils)
	AND NOT ekpo.elikz = 'X'
	AND ekpo.loekz = ''
	AND ekko.lifnr NOT IN ('IC-MXC2','IC-1151')

