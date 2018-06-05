SELECT
	varfield_view.field_content AS "D Number",
	TO_CHAR(patron_view.expiration_date_gmt, 'YYYY-MM-DD') AS "Expiration Date",
	CONCAT(patron_record_fullname.last_name, ', ', patron_record_fullname.first_name, ' ', patron_record_fullname.middle_name) AS "Name",
	patron_record_address.addr1 AS "Address",
	fine.invoice_num AS "Invoice Number",
	fine.title AS "Description",
	(COALESCE(fine.item_charge_amt,0) + COALESCE(fine.processing_fee_amt,0) + COALESCE(fine.billing_fee_amt)) - COALESCE(fine.paid_amt) AS "Amount"
FROM
	sierra_view.fine
INNER JOIN
	sierra_view.varfield_view ON sierra_view.fine.patron_record_id = sierra_view.varfield_view.record_id
INNER JOIN
	sierra_view.patron_record_fullname ON sierra_view.fine.patron_record_id = sierra_view.patron_record_fullname.patron_record_id
INNER JOIN
	sierra_view.patron_record_address ON sierra_view.fine.patron_record_id = sierra_view.patron_record_address.patron_record_id
INNER JOIN
	sierra_view.patron_view ON sierra_view.fine.patron_record_id = sierra_view.patron_view.id
WHERE
	varfield_view.varfield_type_code = 'u' AND
	patron_record_address.patron_record_address_type_id = 1 AND
	patron_view.pcode1 = 'd' AND
	(patron_view.ptype_code = 0 OR patron_view.ptype_code = 6) AND
	DATE(patron_view.expiration_date_gmt) = '2018-05-24'
ORDER BY
	"Name"
LIMIT 5000;