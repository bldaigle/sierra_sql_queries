WITH
	charge_lookup (code, definition) AS (VALUES ('1', 'Manual Charge'), ('2', 'Overdue'), ('3', 'Replacement'), ('4', 'Adjustment'), ('5', 'Lost Book'), ('6', 'Overdue Renewed'), ('7', 'Rental'), ('8', 'Rental Adjustment'), ('9', 'Debit'), ('a', 'notice'), ('b', 'Credit Card'), ('p', 'Program'))
SELECT
	varfield_view.field_content AS "Student ID",
	concat(patron_record_fullname.last_name::text, ', ', patron_record_fullname.first_name::text, ' ', patron_record_fullname.middle_name::text) AS "Name",
	fine.invoice_num AS "Invoice Number",
	charge_lookup.definition AS "Charge Type",
	fine.title AS "Title",
	fine.checkout_gmt AT TIME ZONE 'EST' AS "Checkout Date",
	fine.due_gmt AT TIME ZONE 'EST' AS "Due Date",
	fine.returned_gmt AT TIME ZONE 'EST' AS "Returned Date",
	round((COALESCE(fine.item_charge_amt,0) + COALESCE(fine.processing_fee_amt,0) + COALESCE(fine.billing_fee_amt,0)) - COALESCE(fine.paid_amt,0),2) AS "Amount"
FROM
	sierra_view.fine
INNER JOIN
	charge_lookup ON sierra_view.fine.charge_code = charge_lookup.code
INNER JOIN
	sierra_view.varfield_view ON sierra_view.fine.patron_record_id = sierra_view.varfield_view.record_id
INNER JOIN
	sierra_view.patron_record_fullname ON sierra_view.fine.patron_record_id = sierra_view.patron_record_fullname.patron_record_id
INNER JOIN
	sierra_view.patron_view ON sierra_view.varfield_view.record_id = patron_view.id
WHERE
	varfield_view.varfield_type_code = 'u' AND
	patron_view.ptype_code IN ('15', '18', '19', '28', '29', '32', '34', '35') AND
	patron_view.owed_amt > 0
ORDER BY
	"Name" ASC
LIMIT 10000;