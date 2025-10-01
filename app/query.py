MERGED_DATA_QUERY = """
SELECT 
    po.user_id,
    concat(po.po_number, '-', po.po_line_no) AS po_id,
    acc.account_name,
    po.project_name,
    po.site_code,
    po.po_number AS po_no,
    po.po_line_no AS po_line,
    CASE
        WHEN po.item_description ILIKE '%Survey%' THEN 'Survey'
        WHEN po.item_description ILIKE '%Transportation%' THEN 'Transportation'
        WHEN po.item_description ILIKE '%Work Order%' THEN 'Site Engineer'
        ELSE 'Service'
    END AS category,
    po.item_description AS item_desc,
    CASE
        WHEN po.payment_terms::text LIKE '%COD%' THEN 'ACPAC 100%'
        WHEN po.payment_terms::text LIKE '%AC1%' AND po.payment_terms::text LIKE '%AC2%' THEN 'AC1 80 | PAC 20'
        WHEN po.payment_terms::text LIKE '%AC1%' THEN 'ACPAC 100%'
        ELSE ''
    END AS payment_terms,
    po.unit_price,
    po.requested_qty AS req_qty,
    po.line_amount,
    po.publish_date,
    ROUND(po.line_amount * 0.80, 2) AS ac_amount,
    a.ac_date,
    ROUND(po.line_amount * 0.20, 2) AS pac_amount,
    CASE
        WHEN (po.payment_terms::text LIKE '%COD%' OR (po.payment_terms::text LIKE '%AC1%' AND po.payment_terms::text NOT LIKE '%AC2%')) AND a.ac_date IS NOT NULL THEN a.ac_date
        ELSE a.pac_date
    END AS pac_date,
    CASE
        WHEN po.payment_terms::text LIKE '%COD%' OR (po.payment_terms::text LIKE '%AC1%' AND po.payment_terms::text NOT LIKE '%AC2%') THEN
        CASE
            WHEN po.requested_qty = 0 THEN 'CANCELLED'
            WHEN a.ac_date IS NOT NULL THEN 'CLOSED'
            WHEN a.ac_date IS NULL THEN 'Pending ACPAC'
            ELSE 'CLOSED'
        END
        WHEN po.payment_terms::text LIKE '%AC1%' AND po.payment_terms::text LIKE '%AC2%' THEN
        CASE
            WHEN po.po_status::text = 'CANCELLED' THEN 'CANCELLED'
            WHEN po.po_status::text = 'CLOSED' THEN 'CLOSED'
            WHEN a.ac_date IS NULL THEN 'Pending AC80%'
            WHEN a.pac_date IS NULL THEN 'Pending PAC20%'
            ELSE 'CLOSED'
        END
        ELSE 'Unknown'
    END AS status,
    CASE
        WHEN po.payment_terms::text LIKE '%COD%' OR (po.payment_terms::text LIKE '%AC1%' AND po.payment_terms::text NOT LIKE '%AC2%') THEN
        CASE
            WHEN po.requested_qty = 0 THEN 0
            WHEN a.ac_date IS NOT NULL THEN 0
            WHEN a.ac_date IS NULL THEN po.line_amount
            ELSE 0
        END
        WHEN po.payment_terms::text LIKE '%AC1%' AND po.payment_terms::text LIKE '%AC2%' THEN
        CASE
            WHEN po.po_status::text = 'CANCELLED' THEN 0
            WHEN po.po_status::text = 'CLOSED' THEN 0
            WHEN a.ac_date IS NULL THEN po.line_amount
            WHEN a.pac_date IS NULL THEN ROUND(po.line_amount * 0.20, 2)
            ELSE 0
        END
        ELSE 0
    END AS remaining
FROM purchase_orders po
LEFT JOIN (
    SELECT 
        acceptances.user_id,
        acceptances.po_number,
        acceptances.po_line_no,
        MIN(CASE WHEN acceptances.milestone_type::text = 'AC1' THEN acceptances.application_processed END) AS ac_date,
        MIN(CASE WHEN acceptances.milestone_type::text = 'AC2' THEN acceptances.application_processed END) AS pac_date
    FROM acceptances
    GROUP BY acceptances.user_id, acceptances.po_number, acceptances.po_line_no
) a ON po.user_id = a.user_id AND po.po_number::text = a.po_number::text AND po.po_line_no::text = a.po_line_no::text
LEFT JOIN accounts acc ON po.user_id = acc.user_id AND po.project_name::text = acc.project_name::text
WHERE {base_filter}
"""