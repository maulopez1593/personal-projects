SELECT
	REPLACE(LTRIM(REPLACE(vbak.vbeln,'0',' ')),' ','0') AS sales_document,
	vbap.posnr AS sales_document_item,	
	vbap.erdat AS created_date,
	vbep.edatu AS requested_delivery_date, -- Line RDD
	--vbak.vdatu AS requested_delivery_date, -- Header RDD
	vbap.zzopd AS promise_date,
	vbuk.gbstk AS overall_status,
	vbup.gbsta AS line_status,
	vbuk.lfstk AS delivery_status,
	vbuk.abstk AS rejection_status,
	vbap.abgru AS rejection_reason,
	marc_subquery.maabc AS abc_indicator,
	REPLACE(LTRIM(REPLACE(vbak.kunnr,'0',' ')),' ','0') AS sold_to_party,
	vbap.mvgr2 as material_group_2,
	vbap.matnr AS material_number,
	vbap.arktx AS short_text_for_sales_order_item,
	vbap.prodh AS product_hierarchy,
	vbap.netwr AS net_order_value,
	vbap.netpr AS net_price,
	mbew.stprs AS item_standard_cost,
	mbew.stprv AS item_prev_standard_cost,
	vbap.kwmeng AS order_quantity,
	vbap.lsmeng AS req_delivery_quantity,
	COALESCE(vbbe.omeng, 0) AS open_quantity,
	lips.lfimg AS delivery_quantity,
	likp.wadat_ist AS delivery_date,
	vbep.etenr AS delivery_schedule_line_number,
	vbap.werks AS delivery_plant_nbr
FROM sap_hana_be_ecc_b1_ecc_b1_a_usable.vbak vbak
INNER JOIN sap_hana_be_ecc_b1_ecc_b1_a_usable.vbap vbap
	ON vbap.vbeln = vbak.vbeln
LEFT JOIN sap_hana_be_ecc_b1_ecc_b1_a_usable.vbuk vbuk
	ON vbuk.vbeln = vbap.vbeln
LEFT JOIN sap_hana_be_ecc_b1_ecc_b1_a_usable.vbup vbup
	ON vbup.vbeln = vbap.vbeln
	AND vbup.POSNR = vbap.posnr
INNER JOIN sap_hana_be_ecc_b1_ecc_b1_a_usable.vbep vbep
	ON vbap.vbeln = vbep.vbeln
	AND vbap.posnr = vbep.posnr
	AND vbep.etenr = '0001'
LEFT JOIN sap_hana_be_ecc_b1_ecc_b1_a_usable.mbew mbew
	ON mbew.matnr = vbap.matnr
	AND mbew.bwkey = vbap.werks
LEFT JOIN sap_hana_be_ecc_b1_ecc_b1_a_usable.vbbe vbbe
	ON vbap.vbeln = vbbe.vbeln 
	AND vbap.posnr = vbbe.posnr
	AND vbbe.etenr = '0001'
LEFT JOIN sap_hana_be_ecc_b1_ecc_b1_a_usable.lips lips
	ON vbap.vbeln = lips.vgbel
	AND vbap.posnr = lips.vgpos
LEFT JOIN sap_hana_be_ecc_b1_ecc_b1_a_usable.likp likp
	ON lips.vbeln = likp.vbeln
LEFT JOIN (
	SELECT marc.matnr, MAX(marc.maabc) AS maabc
	FROM sap_hana_be_ecc_b1_ecc_b1_a_usable.marc
	GROUP BY marc.matnr
) AS marc_subquery
	ON vbap.matnr = marc_subquery.matnr
WHERE
	vbak.vkorg = 'US15'
 	AND vbak.vtweg IN (10,40)
 	AND vbap.werks IN ('1001','1002','1025','1003')
 	AND vbap.abgru = ''
 	AND vbap.pstyv IN ('ZTAN','ZTAX','ZTUN','ZTUS')
 	--AND vbak.auart IN ('ZCR', 'ZCRC', 'ZCRU', 'ZDR', 'ZDRC', 'ZDRU', 'ZREU', 'ZST', 'ZSTU', 'ZWCU', 'ZWDU','')
	--AND vbap.pstyv NOT IN ('ZXW2','ZXWM','ZXWN','ZRUN')
	AND from_unixtime(unix_timestamp(vbep.edatu ,'yyyyMMdd'), 'yyyy-MM-dd') >= '2021-10-01' -- Line RDD
	--AND from_unixtime(unix_timestamp(vbak.vdatu ,'yyyyMMdd'), 'yyyy-MM-dd') >= '2021-10-01' -- Header RDD
ORDER BY vbap.erdat, vbap.posnr, vbak.vbeln
