from pathlib import Path
import math
import re

import pandas as pd
import streamlit as st


WORKBOOK_PATH = Path(__file__).parent / "Forwarders comparison (1).xlsm"
SHEET_NAME = "Freight Comparison"

BASE_COLUMNS = [
    "Origin",
    "POL",
    "POD",
    "Destination",
    "20'DV JAS",
    "20'DV AIT",
    "20'RE JAS",
    "20'RE AIT",
    "40'DV/HC JAS",
    "40'DV/HC AIT",
    "40'HCRE JAS",
    "40'HCRE AIT",
    "Transshipment JAS",
    "Transshipment AIT",
    "Transit time JAS",
    "Transit time AIT",
    "Origin free time JAS",
    "Origin free time AIT",
    "Destination free time JAS",
    "Destination free time AIT",
    "Truck JAS",
    "Truck AIT",
]

CONTAINERS = {
    "20'DV": ("20'DV JAS", "20'DV AIT"),
    "20'RE": ("20'RE JAS", "20'RE AIT"),
    "40'DV/HC": ("40'DV/HC JAS", "40'DV/HC AIT"),
    "40'HCRE": ("40'HCRE JAS", "40'HCRE AIT"),
}


def load_freight_data() -> pd.DataFrame:
    data = pd.read_excel(
        WORKBOOK_PATH,
        sheet_name=SHEET_NAME,
        header=None,
        skiprows=3,
        usecols="A:V",
        engine="openpyxl",
    )
    data.columns = BASE_COLUMNS
    for column in BASE_COLUMNS:
        data[column] = data[column].map(clean_cell)
    return data[data["Origin"].ne("") & data["Destination"].ne("")].reset_index(
        drop=True
    )


def clean_cell(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).replace("\xa0", " ").strip()


def parse_number(value):
    if value is None or pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        number = float(value)
        return number if math.isfinite(number) else None

    text = re.sub(r"[^0-9,.-]", "", str(value).replace("\xa0", ""))
    if not text:
        return None
    if "," in text:
        text = text.replace(".", "").replace(",", ".")
    try:
        number = float(text)
        return number if math.isfinite(number) else None
    except ValueError:
        return None


def format_eur(value) -> str:
    number = parse_number(value)
    if number is None:
        return "—"
    formatted = f"{number:,.2f}".replace(",", " ").replace(".", ",")
    return f"€ {formatted}"


def format_days(value) -> str:
    number = parse_number(value)
    if number is None:
        return "—"
    shown = str(int(number)) if number.is_integer() else str(number).replace(".", ",")
    return f"{shown} days"


def style_best_offer(row):
    color = (
        "background-color: #d9eee3; color: #15563c; font-weight: 700"
        if row["Forwarder"] == "JAS"
        else "background-color: #ffeadb; color: #8a431e; font-weight: 700"
    )
    return [color] * len(row)


st.set_page_config(page_title="Freight comparison", layout="wide")
st.markdown(
    """
    <style>
    .stApp { background: #f5f7f4; color: #172a24; }
    .block-container { max-width: 1280px; padding-top: 2.2rem; }
    h1 { color: #173d30; font-size: 2.25rem !important; }
    [data-testid="stSelectbox"] label, [data-testid="stRadio"] label {
        color: #52665d;
    }
    div[data-testid="stDataFrame"] { border-top: 1px solid #d5ded8; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.caption("FREIGHT PROCUREMENT  /  JAS vs AIT")
st.title("Freight comparison")
st.write("Choose a destination, container, and objective to compare offers from every origin.")

if not WORKBOOK_PATH.exists():
    st.error(f"Workbook not found: {WORKBOOK_PATH.name}")
    st.stop()

try:
    freight = load_freight_data()
except Exception as error:
    st.error(f"Could not read the '{SHEET_NAME}' sheet: {error}")
    st.stop()

destination_col, container_col, objective_col = st.columns([1, 1, 1.4], gap="large")
destinations = sorted(freight["Destination"].unique(), key=str.casefold)
with destination_col:
    destination = st.selectbox("Destination", destinations)
with container_col:
    container = st.selectbox("Container type", list(CONTAINERS))
with objective_col:
    objective = st.radio("Minimize", ["Cost", "Transit time"], horizontal=True)
matching = freight[freight["Destination"].str.casefold() == destination.casefold()]

jas_rate_column, ait_rate_column = CONTAINERS[container]
offers = []
for _, route in matching.iterrows():
    for forwarder, rate_column, transit_column in (
        ("JAS", jas_rate_column, "Transit time JAS"),
        ("AIT", ait_rate_column, "Transit time AIT"),
    ):
        price = parse_number(route[rate_column])
        value_column = rate_column if objective == "Cost" else transit_column
        value = parse_number(route[value_column])
        if price is not None and value is not None:
            offers.append((value, route, forwarder, price))

best_value = min((offer[0] for offer in offers), default=None)
best_offers = [
    offer for offer in offers if math.isclose(offer[0], best_value)
] if best_value is not None else []
recommendations = [
    {
        "Price": format_eur(price),
        "Transit time": format_days(
            route["Transit time JAS" if forwarder == "JAS" else "Transit time AIT"]
        ),
        "Forwarder": forwarder,
        "Origin": route["Origin"],
        "POL": route["POL"],
        "POD": route["POD"],
    }
    for value, route, forwarder, price in best_offers
]

st.divider()
st.subheader(f"Best {container} option to {destination}")
st.caption(f"Minimizing {objective.lower()} · Prices in EUR")
if matching.empty:
    st.info("No matching freight rows were found.")
elif not recommendations:
    st.info("No priced offers are available for this destination, container, and objective.")
else:
    result = pd.DataFrame(recommendations)
    styled = result.style.apply(style_best_offer, axis=1)
    st.dataframe(styled, hide_index=True, width="stretch")
    st.caption(f"Compared {len(matching)} routes across {matching['Origin'].nunique()} origins")

st.caption(f"Source: {WORKBOOK_PATH.name} · {len(freight)} routes in Freight Comparison")