# Employee Performance Forecast & Scenario Planner

Streamlit dashboard for employee-level Gross Sales / Net Sales scenario planning,
current run-rate forecasting, five-sheet workbook analysis, summary calculations,
and achievement-band reporting.

## Expected workbook sheets

- Summary
- Summary-Achievement
- RM Retail Sales
- RM DHNI
- VRM

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Community Cloud

1. Push this repository to GitHub.
2. Open Streamlit Community Cloud.
3. Create a new app and select this GitHub repository.
4. Set the main file path to `app.py`.
5. Deploy.

The Excel workbook is uploaded by the user at runtime, so no workbook needs to be committed.
