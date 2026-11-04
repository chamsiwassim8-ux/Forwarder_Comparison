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


def best_offers(routes: pd.DataFrame, container: str, objective: str) -> pd.DataFrame:
    jas_rate_column, ait_rate_column = CONTAINERS[container]
    offers = []
    for _, route in routes.iterrows():
        for forwarder, rate_column in (
            ("JAS", jas_rate_column),
            ("AIT", ait_rate_column),
        ):
            transit_column = f"Transit time {forwarder}"
            value_column = rate_column if objective == "Cost" else transit_column
            value = parse_number(route[value_column])
            if value is not None:
                offers.append((value, route, forwarder, rate_column, transit_column))

    best_value = min((offer[0] for offer in offers), default=None)
    if best_value is None:
        return pd.DataFrame()

    recommendations = []
    for value, route, forwarder, rate_column, transit_column in offers:
        if not math.isclose(value, best_value):
            continue
        recommendations.append(
            {
                "Forwarder": forwarder,
                "Price": format_eur(route[rate_column]),
                "Transit time": format_days(route[transit_column]),
                "Origin": route["Origin"],
                "Destination": route["Destination"],
                "POL": route["POL"],
                "POD": route["POD"],
                "Transshipment": route[f"Transshipment {forwarder}"] or "—",
                "Free time at origin": format_days(
                    route[f"Origin free time {forwarder}"]
                ),
                "Free time at destination": format_days(
                    route[f"Destination free time {forwarder}"]
                ),
            }
        )
    return pd.DataFrame(recommendations)


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
st.write(
    "Choose an origin, destination, and container to compare the best cost and "
    "transit-time offers."
)

if not WORKBOOK_PATH.exists():
    st.error(f"Workbook not found: {WORKBOOK_PATH.name}")
    st.stop()

try:
    freight = load_freight_data()
except Exception as error:
    st.error(f"Could not read the '{SHEET_NAME}' sheet: {error}")
    st.stop()

origin_col, destination_col, container_col = st.columns(3, gap="large")
origins = sorted(freight["Origin"].unique(), key=str.casefold)
with origin_col:
    origin = st.selectbox("Origin", origins)
origin_routes = freight[freight["Origin"].str.casefold() == origin.casefold()]
destinations = sorted(origin_routes["Destination"].unique(), key=str.casefold)
with destination_col:
    destination = st.selectbox("Destination", destinations)
with container_col:
    container = st.selectbox("Container type", list(CONTAINERS))
matching = origin_routes[
    origin_routes["Destination"].str.casefold() == destination.casefold()
]

st.divider()
for objective, title in (
    ("Cost", f"Best cost option · {container}"),
    ("Transit time", f"Best transit time option · {container}"),
):
    st.subheader(title)
    options = best_offers(matching, container, objective)
    if options.empty:
        st.info(
            f"No {objective.lower()} offers are available for this "
            "origin, destination, and container type."
        )
    else:
        st.dataframe(
            options.style.apply(style_best_offer, axis=1),
            hide_index=True,
            width="stretch",
        )

st.caption(f"Compared {len(matching)} matching route(s) · Prices in EUR")

st.caption(f"Source: {WORKBOOK_PATH.name} · {len(freight)} routes in Freight Comparison")