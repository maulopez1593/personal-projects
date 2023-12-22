WITH LatestDeliveryDates AS (
    SELECT
        vbep.vbeln AS sales_document,
        vbep.posnr AS item,
        MAX(vbep.edatu) AS latest_delivery_date
    FROM
        sap_hana_be_ecc_b1_ecc_b1_a_usable.vbep vbep
    INNER JOIN sap_hana_be_ecc_b1_ecc_b1_a_usable.vbak vbak
        ON vbep.vbeln = vbak.vbeln 
    LEFT JOIN sap_hana_be_ecc_b1_ecc_b1_a_usable.vbap vbap 
        ON vbep.vbeln = vbap.vbeln 
        AND vbep.posnr = vbap.posnr 
        AND vbap.abgru = ''
        AND vbap.werks IN ('1001','1002','1025')
        AND vbap.pstyv IN ('ZTAN','ZTAX','ZTUN','ZTUS','ZTAC','ZCF1','ZXWN')
        AND from_unixtime(unix_timestamp(vbap.erdat,'yyyyMMdd'), 'yyyy-MM-dd') >= '2018-01-01'
    WHERE vbak.vkorg = 'US15'
        AND vbak.vtweg IN (10,40)
        AND NOT vbep.etenr  = ''
    GROUP BY vbep.vbeln, vbep.posnr
)
SELECT
    REPLACE(LTRIM(REPLACE(ld.sales_document,'0',' ')),' ','0') AS sales_document,
    MAX(ld.latest_delivery_date) AS max_dlv_date
FROM
    LatestDeliveryDates ld
GROUP BY ld.sales_document
