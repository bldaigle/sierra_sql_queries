SELECT
	varfield.field_content AS "Student ID",
    concat(patron_record_fullname.last_name::text, ', ', patron_record_fullname.first_name::text, ' ', patron_record_fullname.middle_name::text) AS "Name",
    round(patron_view.owed_amt,2) AS "Amount"
FROM
	sierra_view.varfield
INNER JOIN
	sierra_view.patron_record_fullname ON sierra_view.varfield.record_id = patron_record_fullname.patron_record_id
INNER JOIN
	sierra_view.patron_view ON sierra_view.varfield.record_id = patron_view.id
WHERE
	varfield.varfield_type_code = 'u' AND
    patron_view.ptype_code IN ('15', '18', '19', '28', '29', '32', '34', '35') AND
    patron_view.owed_amt > 0
ORDER BY
	"Name" ASC;