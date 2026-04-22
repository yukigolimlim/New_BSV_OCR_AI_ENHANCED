import sqlite3
import psycopg2

# Connect to SQLite
sqlite_conn = sqlite3.connect(
    r"C:\Users\Programmer 2\Desktop\NEW_BSV_OCR_ENHANCED\New_BSV_OCR_AI_ENHANCED\_NEWOCR_AI_ENHANCED\lookup_summary_results\applicants.db"
)
sqlite_cur = sqlite_conn.cursor()

# Connect to PostgreSQL
pg_conn = psycopg2.connect(
    host="localhost",
    port=5432,
    dbname="applicants",
    user="postgres",
    password="bsv_123"  # ← change this
)
pg_cur = pg_conn.cursor()

# Fetch all rows from SQLite
sqlite_cur.execute("SELECT * FROM applicants")
rows = sqlite_cur.fetchall()
print(f"Found {len(rows)} rows to migrate...")

# Insert into PostgreSQL
for row in rows:
    pg_cur.execute("""
        INSERT INTO applicants (
            id, session_id, processed_at, source_file, status,
            client_id, pn, applicant_name, residence_address, office_address,
            industry_name, income_items, income_total, business_items, business_total,
            household_items, household_total, net_income, petrol_risk, transport_risk,
            results_json, page_map, amort_current_total, loan_balance, amortized_cost,
            principal_loan, maturity, interest_rate, amort_history_total, branch,
            loan_class_name, product_name, industry_name_ploan, loan_date, term_unit,
            term, security, release_tag, loan_amount, loan_balance_ploan,
            amort_ploan, loan_status, ao_name
        ) VALUES (
            %s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,
            %s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s
        )
    """, row)

pg_conn.commit()
print(f"Done! {len(rows)} rows migrated successfully.")

sqlite_cur.close()
sqlite_conn.close()
pg_cur.close()
pg_conn.close()