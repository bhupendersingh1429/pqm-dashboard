# PQM Dashboard — Combined Report

Run: `pip install -r requirements.txt` then `streamlit run pq_dash.py`.

## Replace the existing dashboard

In the existing bhupendersingh1429/pqm-dashboard repository, replace these files:
- pq_dash.py
- requirements.txt
- Report.xlsx (the supplied Combined PQM Report, already renamed in this package)

Keep the Streamlit entry point as pq_dash.py. Reboot/redeploy the app after updating.
The older Report.xlsb is ignored. No changes to History.xlsx are needed.

## What is included

One summary dropdown: Zone, Region(AVP), DM, PQM. Original three-row grouped
headers, entity filters, text search, source report date, five headline metrics,
filtered CSV download and original Excel report download. No Staff KPI.
Values come directly from the four summary sheets and Excel's saved formula
results. Excel must calculate and save formulas before uploading the workbook.
Rates use their original numeric fractions, including values above 100%.
FTNR headline is calculated from summed FTNR counts / summed FTNR base, never
an average of row percentages. Excel totals are shown only in the full view.
CSV downloads contain the selected detail rows, without the Excel total row.

The supplied report has different DB amount totals across summaries:
Zone / Region 56.6685 Cr; DM 56.5744 Cr; PQM 56.4442 Cr. The dashboard preserves
each summary's results. It does not attempt to reconcile or alter the workbook.

## Shared uploads

Default: all users read Report.xlsx deployed with the app. Upload previews are
session-only and clearly labelled; they do not update other users.

For admin publishing, provide a persistent shared filesystem on the server and
set these environment variables or Streamlit secrets:

PQM_REPORT_DIR = "/your/persistent/shared/report-folder"
PQM_ADMIN_PASSWORD = "choose-a-strong-password"

Restart the app. Preview a workbook, enter the admin password, then click
Publish report for all users. Publication is atomic. Other users click Refresh
report or reload to see the new report. Never commit the password to GitHub.
Every replica must mount the same shared folder. Ephemeral hosting filesystem
storage, including a normal Streamlit Community Cloud app folder, is not a
durable upload store. On such hosting, update Report.xlsx in the repository,
or arrange persistent storage before enabling shared publication.

The existing hosting/access policy remains the operator's responsibility.
This package does not introduce login or publish the report to a new service.

## Verification

Tested all four summary switches, filter/search behavior, cached formula
availability, percent/count formatting, upload rejection of incompatible sheets,
and atomic report publication using a temporary test directory.
