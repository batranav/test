# test
test cases
#for sql java query not execution
select * from(WITH all_events AS (
    SELECT 
        datalake_applicationid, 
        "datalake_date&timestamp" AS event_time,
        'Lead_creation' AS source
    FROM "clix_tezract_tables"."tezzract_leadcreation_transformed"

    UNION ALL

    SELECT 
        datalake_applicationid, 
        "datalake_date&timestamp" AS event_time,
        'Lead_creation' AS source
    FROM "clix_tezract_tables"."tezzract_onboarding_transformed"

    UNION ALL

    SELECT 
        datalake_applicationid, 
        "datalake_date&timestamp" AS event_time,
        'Credit_WIP' AS source
    FROM "clix_tezract_tables"."tezzract_creditreview_transformed"

    UNION ALL

    SELECT 
        datalake_applicationid, 
        "datalake_date&timestamp" AS event_time,
        'Ops' AS source
    FROM "clix_tezract_tables"."tezzract_opschecker_transformed"

    UNION ALL

    SELECT 
        datalake_applicationid, 
        "datalake_date&timestamp" AS event_time,
        'Onboarded' AS source
    FROM "clix_tezract_tables"."tezzract_lmsposting_transformed"
),

latest_events AS (
    SELECT *
    FROM (
        SELECT *,
               ROW_NUMBER() OVER (PARTITION BY datalake_applicationid ORDER BY "event_time" DESC) AS rn
        FROM all_events
    ) ranked_events
    WHERE rn = 1
)

SELECT
  json_extract_scalar(a.coapp1details, '$.applicantMobile') AS applicantMobile,
  json_extract_scalar(a.coapp1details, '$.applicantDOB') AS applicantDOB,
  json_extract_scalar(a.coapp1details, '$.applicantSalutation') AS applicantSalutation,
  a.entitynamelead AS entitynamelead,
  a.branchcode,
  a.application_id,
  b.entitycity,
  b.entitystate,
  a.leadprogramtype as Program_code,
  a.entityproductloanentry as Product_code,
  c.loan_finreferenceid AS loanApplicationId,
  json_extract_scalar(a.coapp1details, '$.panNumber') AS panNumber,
  json_extract_scalar(a.coapp1details, '$.customerId') AS customerId,
  json_extract_scalar(a.coapp1details, '$.applicantName') AS applicant_Name,
  le.event_time,
  le.source
FROM "clix_tezract_tables"."tezzract_leadcreation_transformed" a
JOIN "clix_tezract_tables"."tezzract_colender_transformed" b
  ON json_extract_scalar(a.coapp1details, '$.customerId') = json_extract_scalar(b.coapp1details, '$.customerId')
JOIN "clix_tezract_tables"."tezzract_lmsposting_transformed" c 
  ON json_extract_scalar(a.coapp1details, '$.customerId') = json_extract_scalar(c.coapp1details, '$.customerId')
JOIN latest_events le
  ON a.application_id = le.datalake_applicationid)
where panNumber IN ('AABTS7112E');
Query execution failed.
2025-07-23 18:25:08.163 DEBUG 15160 --- [nio-8080-exec-1] o.s.web.servlet.DispatcherServlet        : Failed to complete request: com.amazonaws.services.athena.model.InvalidRequestException: Query did not finish successfully. Final query state: FAILED (Service: AmazonAthena; Status Code: 400; Error Code: InvalidRequestException; Request ID: 5c5e6d7a-7f55-4579-bd15-c4f985640b1c)
2025-07-23 18:25:08.168 ERROR 15160 --- [nio-8080-exec-1] o.a.c.c.C.[.[.[.[dispatcherServlet]      : Servlet.service() for servlet [dispatcherServlet] in context with path [/mdmService] threw exception [Request processing failed; nested exception is com.amazonaws.services.athena.model.InvalidRequestException: Query did not finish successfully. Final query state: FAILED (Service: AmazonAthena; Status Code: 400; Error Code: InvalidRequestException; Request ID: 5c5e6d7a-7f55-4579-bd15-c4f985640b1c)] with root cause
