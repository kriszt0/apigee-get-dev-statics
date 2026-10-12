COPY (
    SELECT
        client_received_start_timestamp::date AS datum,
        api_product,
        apiproxy,
        developer_app,
        is_error,
        SUM(COALESCE(message_count, 0)) AS message_count
    FROM analytics."ORG.ENV.fact"
    WHERE client_received_start_timestamp >= '2026-10-05'
      AND client_received_start_timestamp < '2026-10-12'
    GROUP BY
        client_received_start_timestamp::date,
        api_product,
        apiproxy,
        developer_app,
        is_error
    ORDER BY
        datum,
        api_product,
        apiproxy,
        developer_app,
        is_error
)
TO STDOUT
WITH (FORMAT CSV, HEADER TRUE, DELIMITER ';');