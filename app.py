import streamlit as st
import pandas as pd
from datetime import date
import time
from report_builder import build_ar_aging_report
from logger import log_report_run

st.set_page_config(
    page_title="AR Aging Report",
    layout="centered",
)

st.title("AR Aging Report")
st.markdown("Upload an invoices export and select an **As of Date** to generate the AR Aging Summary.")

st.divider()

# Group selector
GROUP_OPTIONS = ["Y&S Group", "YS-SeatGeek", "The Ticket Guy", "Other"]
selected_group = st.selectbox("Group", GROUP_OPTIONS)

if selected_group == "Other":
    custom_group = st.text_input("Enter group name")
    group_name = custom_group.strip() if custom_group.strip() else "Other"
else:
    group_name = selected_group

as_of = st.date_input(
    "As of Date",
    value=date.today(),
    help="Aging buckets are calculated relative to this date.",
)

uploaded_file = st.file_uploader(
    "Upload Invoices Report (.xlsx)",
    type=["xlsx"],
    help="Export from your invoicing system. Must include columns: Paid, IsCancelled, Bal., Client, Company, Inv#, Ext Order #, Status, Created.",
)

if uploaded_file and as_of and group_name:
    if st.button("Generate AR Aging Report", type="primary", use_container_width=True):
        with st.spinner("Building report..."):
            try:
                start = time.time()

                excel_bytes, summary_df, grand_total, row_count = build_ar_aging_report(
                    uploaded_file, pd.Timestamp(as_of), group_name
                )

                elapsed = time.time() - start

                # Log to Supabase
                log_report_run(
                    as_of_date=as_of,
                    row_count=row_count,
                    grand_total=grand_total,
                    elapsed_seconds=elapsed,
                )

                st.success(f"Report generated — **{row_count:,}** unpaid invoices · Grand total: **${grand_total:,.2f}**")

                # Preview summary table
                st.subheader("AR Aging Summary Preview")
                st.dataframe(summary_df, use_container_width=True, hide_index=True)

                # Download button — filename includes group name and date
                file_name = f"AR Aging - {group_name} - {as_of.strftime('%m-%d-%Y')}.xlsx"
                st.download_button(
                    label="⬇️ Download AR Aging.xlsx",
                    data=excel_bytes,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

            except Exception as e:
                st.error(f"Error generating report: {e}")
                st.exception(e)
