# Consolidated Resource Report Generator

A Streamlit web app that converts a weekly MIS Report into a formatted Consolidated Resource Report, matching the **Resource Allocation Master Sheet** design.

## Files in this repo

```
app.py              ← Streamlit app
Employee.xlsx       ← permanent employee reference (Role + Location)
Project.xlsx        ← permanent project-client reference
requirements.txt    ← Python dependencies
README.md
```

## How to use

1. Open the app URL
2. Upload your **MIS Report (.xlsx)**
3. Optionally upload updated Employee.xlsx or Project.xlsx to override the bundled versions
4. Click **Generate Report**
5. Download the consolidated output

## Local development

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Cloud

1. Push this repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Click **New app** → select your repo → set `app.py` as the main file
4. Click **Deploy** — done!

## Output design

Matches `Resource_Allocation_-_2026.xlsx → Master Sheet`:
- Columns: Department | Name | Role | Technology | Client | Project | Allocation | FTE | Billable (Y/N) | Location | Notes
- Light-blue header (Aptos Narrow 12pt bold), thin borders, autofilter
- Yellow flag on Role and FTE (data not in source — needs manual fill)
- Yellow flag on blank Location cells
