SELECT
	item_view.barcode AS "Item Barcode",
	bib_record_property.best_title AS "Title",
	bib_record_property.best_author AS "Author",
	TRIM(REGEXP_REPLACE(item_record_property.call_number,'\|.',' ','g')) AS "Call Number",
	item_view.location_code AS "Location Code",
	varfield_view.field_content AS "Variable Field Content",
    circ_trans.transaction_gmt AT TIME ZONE 'EST' AS "Date Counted"
FROM
	sierra_view.circ_trans
INNER JOIN
	sierra_view.item_view ON sierra_view.circ_trans.item_record_id = sierra_view.item_view.id
INNER JOIN
	sierra_view.bib_view ON sierra_view.circ_trans.bib_record_id = sierra_view.bib_view.id
INNER JOIN
	sierra_view.item_record_property ON sierra_view.circ_trans.item_record_id = sierra_view.item_record_property.item_record_id
INNER JOIN
	sierra_view.bib_record_property ON sierra_view.circ_trans.bib_record_id = sierra_view.bib_record_property.bib_record_id
LEFT OUTER JOIN
	sierra_view.varfield_view ON sierra_view.circ_trans.item_record_id = sierra_view.varfield_view.record_id AND sierra_view.varfield_view.varfield_type_code = 'v'
WHERE
	circ_trans.count_type_code_num = 1 AND
	(DATE(circ_trans.transaction_gmt) >= date_trunc('month', current_date - interval '1' month) AND DATE(circ_trans.transaction_gmt) < date_trunc('month', current_date)) AND
	item_view.location_code LIKE 'k____' AND
  	item_view.location_code != 'kelec'
ORDER BY
	"Title"