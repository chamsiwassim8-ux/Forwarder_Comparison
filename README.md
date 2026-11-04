# Freight comparison app

A Streamlit interface for comparing JAS and AIT freight offers in the supplied Excel workbook.

## Run locally

```bash
python -m pip install -r requirements.txt
streamlit run app.py
```

Keep `Forwarders comparison (1).xlsm` next to `app.py`. The app reads the `Freight Comparison` sheet directly, so updated workbook data appears on the next app rerun.