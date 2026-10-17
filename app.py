import streamlit as st
import pandas as pd
from pathlib import Path

st.set_page_config(
    page_title="Automotive Card Collector — Prototyp",
    page_icon="🏎️",
    layout="wide",
)

BASE_DIR = Path(__file__).resolve().parent
EXCEL_PATH = BASE_DIR / "Automotive_Card_Collector_Master_v0.5.xlsx"

SHEETS = [
    "Publishers", "Series", "Products", "Subsets", "Cards",
    "Variants", "Licences", "Price Rules", "Sources",
]


@st.cache_data
def load_data(excel_path: str, modified_time: float | None):
    """Load known workbook sheets. Return empty DataFrames if workbook is absent."""
    path = Path(excel_path)
    if not path.exists():
        return {sheet: pd.DataFrame() for sheet in SHEETS}

    try:
        xls = pd.ExcelFile(path)
        result = {}
        for sheet in SHEETS:
            if sheet in xls.sheet_names:
                result[sheet] = pd.read_excel(xls, sheet_name=sheet)
            else:
                result[sheet] = pd.DataFrame()
        return result
    except Exception as exc:
        st.error(f"Excel soubor se nepodařilo načíst: {exc}")
        return {sheet: pd.DataFrame() for sheet in SHEETS}


data = load_data(str(EXCEL_PATH), EXCEL_PATH.stat().st_mtime if EXCEL_PATH.exists() else None)

if "my_collection" not in st.session_state:
    st.session_state.my_collection = [
        {
            "Collection ID": 1,
            "Card ID": "TOPPS26-BASE-001",
            "Series": "2026 Topps Chrome Formula 1",
            "Card No": "001",
            "Subject": "Max Verstappen — Red Bull Racing",
            "Variant": "Gold Parallel (/10)",
            "Condition": "PSA 10",
            "Qty": 1,
            "Purchase Price": 350.0,
            "Est. Market Value": 450.0,
            "Currency": "USD",
        },
        {
            "Collection ID": 2,
            "Card ID": "TOPPS26-SW-01",
            "Series": "2026 Topps Chrome Formula 1",
            "Card No": "SW-01",
            "Subject": "Speed Wheels — Lewis Hamilton",
            "Variant": "Base Insert",
            "Condition": "Raw Mint",
            "Qty": 2,
            "Purchase Price": 15.0,
            "Est. Market Value": 25.0,
            "Currency": "USD",
        },
    ]

st.sidebar.title("🏎️ Auto Cards App")
st.sidebar.markdown("**Master v0.5 Prototyp**")
menu = st.sidebar.radio(
    "Navigace",
    [
        "📊 Dashboard & Portfolio",
        "🔍 Katalog karet (Browse)",
        "➕ Přidat kartu do sbírky",
        "📷 AI Scanner (Prototyp)",
        "📜 Pravidla oceňování (Price Rules)",
        "⚙️ Databáze (Raw Tables)",
    ],
)

# 1. Dashboard & portfolio
if menu == "📊 Dashboard & Portfolio":
    st.title("📊 Moje portfolio & statistiky")
    col_df = pd.DataFrame(st.session_state.my_collection)

    if col_df.empty:
        st.info("Sbírka je zatím prázdná.")
    else:
        total_items = int(col_df["Qty"].sum())
        total_invested = (col_df["Purchase Price"] * col_df["Qty"]).sum()
        total_value = (col_df["Est. Market Value"] * col_df["Qty"]).sum()
        profit = total_value - total_invested
        roi = profit / total_invested * 100 if total_invested > 0 else 0

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Celkem karet", f"{total_items} ks")
        c2.metric("Investováno", f"${total_invested:,.2f}")
        c3.metric("Odhadovaná hodnota", f"${total_value:,.2f}", f"{roi:+.1f}% ROI")
        c4.metric("Zisk / ztráta", f"${profit:+,.2f}")

        st.markdown("---")
        st.subheader("📋 Seznam vlastněných karet")
        st.dataframe(col_df, use_container_width=True, hide_index=True)

        csv = col_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "Exportovat sbírku do CSV",
            data=csv,
            file_name="moje_sbirka_karet.csv",
            mime="text/csv",
        )

# 2. Catalog
elif menu == "🔍 Katalog karet (Browse)":
    st.title("🔍 Katalog karet & variant")
    cards_df = data["Cards"]
    series_df = data["Series"]
    variants_df = data["Variants"]

    if cards_df.empty:
        st.warning(
            "List „Cards“ je prázdný nebo chybí Excel soubor. "
            "Vložte soubor Automotive_Card_Collector_Master_v0.5.xlsx "
            "do stejné složky jako app.py."
        )
    else:
        if "Series" in series_df.columns:
            series_list = ["Všechny série"] + sorted(
                series_df["Series"].dropna().astype(str).unique().tolist()
            )
        else:
            series_list = ["Všechny série"]

        selected_series = st.selectbox("Filtrovat podle série", series_list)
        filtered_cards = cards_df.copy()

        if selected_series != "Všechny série" and not series_df.empty:
            if {"Series", "Series ID"}.issubset(series_df.columns) and "Series ID" in cards_df.columns:
                matches = series_df.loc[series_df["Series"] == selected_series, "Series ID"]
                if not matches.empty:
                    filtered_cards = cards_df[cards_df["Series ID"].isin(matches)]

        st.subheader(f"Karty v katalogu ({len(filtered_cards)})")
        for _, card in filtered_cards.iterrows():
            card_no = card.get("Card No.", "—")
            card_name = card.get("Card Name / Subject", "Neznámá karta")
            discipline = card.get("Discipline", "—")
            with st.expander(f"🏎️ #{card_no} — {card_name} ({discipline})"):
                c1, c2 = st.columns([2, 1])
                with c1:
                    st.write(f"**Card ID:** `{card.get('Card ID', '—')}`")
                    st.write(f"**Series ID:** `{card.get('Series ID', '—')}`")
                    st.write(f"**Řidič / tým / vůz:** {card.get('Driver / Team / Car', '—')}")
                    st.write(f"**Třída rarity:** {card.get('Rarity Class', '—')}")
                    st.write(f"**Stav ověření:** `{card.get('Verification Status', '—')}`")
                with c2:
                    st.info("**Dostupné varianty**")
                    if variants_df.empty:
                        st.caption("V listu Variants nejsou data.")
                    else:
                        for _, variant_row in variants_df.head(4).iterrows():
                            st.write(
                                f"• **{variant_row.get('Variant', '—')}** "
                                f"({variant_row.get('Variant Type', '—')}) — "
                                f"`{variant_row.get('Serial / Print Run', '—')}`"
                            )

# 3. Add card
elif menu == "➕ Přidat kartu do sbírky":
    st.title("➕ Přidat kartu do sbírky")
    with st.form("add_card_form", clear_on_submit=True):
        card_id = st.text_input("Card ID", "TOPPS26-BASE-001")
        series_name = st.text_input("Série", "2026 Topps Chrome Formula 1")
        card_no = st.text_input("Číslo karty", "001")
        subject = st.text_input("Jméno / předmět", "Max Verstappen")
        variant = st.selectbox(
            "Varianta",
            ["Base", "Gold Parallel (/10)", "Black Parallel (/1)", "Autograph Gold (/10)"],
        )
        condition = st.selectbox(
            "Stav / grading", ["Raw Mint", "PSA 10", "PSA 9", "BGS 9.5", "Raw NM"]
        )
        qty = st.number_input("Počet kusů", min_value=1, value=1, step=1)
        price = st.number_input("Nákupní cena za kus ($)", min_value=0.0, value=50.0, step=1.0)
        est_val = st.number_input(
            "Odhadovaná tržní cena za kus ($)", min_value=0.0, value=75.0, step=1.0
        )

        submitted = st.form_submit_button("Vložit do mé sbírky")
        if submitted:
            st.session_state.my_collection.append(
                {
                    "Collection ID": max(
                        [item["Collection ID"] for item in st.session_state.my_collection],
                        default=0,
                    ) + 1,
                    "Card ID": card_id,
                    "Series": series_name,
                    "Card No": card_no,
                    "Subject": subject,
                    "Variant": variant,
                    "Condition": condition,
                    "Qty": int(qty),
                    "Purchase Price": float(price),
                    "Est. Market Value": float(est_val),
                    "Currency": "USD",
                }
            )
            st.success("Karta byla přidána do sbírky.")

# 4. AI scanner prototype
elif menu == "📷 AI Scanner (Prototyp)":
    st.title("📷 AI scanner karet (simulace)")
    st.caption("Toto je pouze demonstrační simulace; nejde o skutečné rozpoznávání pomocí AI.")
    uploaded_file = st.file_uploader("Nahrát fotografii karty", type=["jpg", "jpeg", "png"])
    if uploaded_file is not None:
        st.image(uploaded_file, caption="Nahraný snímek", width=300)
        st.info("Ukázkový výsledek simulace — údaje nejsou ověřené.")
        st.json(
            {
                "Publisher": "Topps",
                "Series": "2026 Topps Chrome Formula 1",
                "Card No": "001",
                "Subject": "Max Verstappen",
                "Detected Variant": "Gold Refractor Parallel (/10)",
            }
        )

# 5. Price rules
elif menu == "📜 Pravidla oceňování (Price Rules)":
    st.title("📜 Pravidla oceňování & trhu")
    if data["Price Rules"].empty:
        st.info("List „Price Rules“ je prázdný nebo není dostupný.")
    else:
        st.dataframe(data["Price Rules"], use_container_width=True, hide_index=True)

# 6. Raw tables
elif menu == "⚙️ Databáze (Raw Tables)":
    st.title("⚙️ Náhled databázových listů Excelu")
    selected_sheet = st.selectbox("Vyberte list", list(data.keys()))
    frame = data[selected_sheet]
    if frame.empty:
        st.info(f"List „{selected_sheet}“ je prázdný nebo není dostupný.")
    else:
        st.dataframe(frame, use_container_width=True, hide_index=True)
