import streamlit as st
import pandas as pd
import re
from io import BytesIO
from pathlib import Path

st.set_page_config(
    page_title="Teile-Preisvergleich | STAkis + TEXKAT",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------- Hilfsfunktionen ----------

def normalize_artikelnummer(value) -> str:
    """Artikelnummer normalisieren: Leerzeichen, Bindestriche, Punkte entfernen, Großbuchstaben."""
    if pd.isna(value):
        return ""
    s = str(value).strip().upper()
    s = re.sub(r"[\s\-\./_]", "", s)
    return s


def detect_columns(df: pd.DataFrame) -> dict:
    """Versucht automatisch Artikelnummer- und Preisspalten zu erkennen."""
    cols = {c: str(c).lower() for c in df.columns}
    
    artikel_candidates = []
    preis_candidates = []
    
    for orig, low in cols.items():
        # Artikelnummer
        if any(k in low for k in ["artikel", "artnr", "art-nr", "art_nr", "teilenummer", "teilnr", 
                                   "oe", "hersteller", "ref", "nummer", "sku", "item", "part"]):
            if "preis" not in low and "price" not in low and "euro" not in low:
                artikel_candidates.append(orig)
        
        # Preis
        if any(k in low for k in ["preis", "price", "einkauf", "vk", "netto", "brutto", 
                                   "euro", "€", "eur", "cost", "amount"]):
            preis_candidates.append(orig)
    
    # Fallback: erste Spalte als Artikel, zweite numerische als Preis
    if not artikel_candidates and len(df.columns) > 0:
        artikel_candidates = [df.columns[0]]
    
    if not preis_candidates:
        for col in df.columns:
            if pd.api.types.is_numeric_dtype(df[col]) or df[col].astype(str).str.replace(",", ".").str.match(r"^\d+[.,]?\d*$").any():
                preis_candidates.append(col)
                break
    
    return {
        "artikel": artikel_candidates[0] if artikel_candidates else None,
        "preis": preis_candidates[0] if preis_candidates else None,
        "all_artikel": artikel_candidates,
        "all_preis": preis_candidates,
    }


def load_file(uploaded_file) -> pd.DataFrame | None:
    """CSV oder Excel laden."""
    try:
        name = uploaded_file.name.lower()
        if name.endswith((".xlsx", ".xls")):
            df = pd.read_excel(uploaded_file, engine="openpyxl")
        else:
            # Verschiedene Trennzeichen und Encodings versuchen
            content = uploaded_file.read()
            for sep in [";", ",", "\t"]:
                for encoding in ["utf-8", "latin-1", "cp1252"]:
                    try:
                        df = pd.read_csv(BytesIO(content), sep=sep, encoding=encoding, decimal=",")
                        if len(df.columns) > 1:
                            return df
                    except Exception:
                        continue
            # Letzter Versuch
            df = pd.read_csv(BytesIO(content), sep=None, engine="python", encoding="utf-8", on_bad_lines="skip")
        return df
    except Exception as e:
        st.error(f"Fehler beim Laden von {uploaded_file.name}: {e}")
        return None


def clean_price(series: pd.Series) -> pd.Series:
    """Preise in float umwandeln (Komma → Punkt, €-Zeichen etc. entfernen)."""
    def to_float(x):
        if pd.isna(x):
            return None
        s = str(x).strip()
        s = re.sub(r"[€$£\s]", "", s)
        s = s.replace(".", "").replace(",", ".") if s.count(",") == 1 and s.count(".") > 1 else s.replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return None
    return series.apply(to_float)


# ---------- UI ----------

st.title("🔧 Teile-Preisvergleich")
st.caption("STAkis + TEXKAT Exporte vergleichen · Artikelnummern & Preise")

with st.sidebar:
    st.header("Einstellungen")
    st.markdown("""
    **So funktioniert's:**
    1. Exportiere in STAkis / TEXKAT die gewünschten Artikel (CSV oder Excel)
    2. Lade die Dateien hier hoch
    3. Ordne Artikelnummer- und Preisspalte zu
    4. Vergleiche die Preise
    """)
    st.divider()
    show_only_common = st.checkbox("Nur gemeinsame Artikel anzeigen", value=False)
    show_diff_only = st.checkbox("Nur Artikel mit Preisunterschied", value=False)
    min_diff = st.number_input("Mindest-Differenz (€)", min_value=0.0, value=0.0, step=0.5)
    st.divider()
    st.markdown("**Tipp:** In STAkis/TEXKAT meist über „Liste exportieren“ oder Warenkorb → Export.")

# Datei-Upload
st.subheader("1. Dateien hochladen")
col1, col2 = st.columns(2)

with col1:
    st.markdown("**Quelle A** (z. B. STAkis)")
    files_a = st.file_uploader("CSV / Excel", type=["csv", "xlsx", "xls"], key="files_a", accept_multiple_files=True)

with col2:
    st.markdown("**Quelle B** (z. B. TEXKAT)")
    files_b = st.file_uploader("CSV / Excel", type=["csv", "xlsx", "xls"], key="files_b", accept_multiple_files=True)

if not files_a and not files_b:
    st.info("👆 Lade mindestens eine Datei von STAkis und/oder TEXKAT hoch.")
    st.stop()

# Dateien laden und vorbereiten
sources = {}

def process_files(files, source_name):
    dfs = []
    for f in files:
        df = load_file(f)
        if df is not None and not df.empty:
            df["_quelle_datei"] = f.name
            dfs.append(df)
    if not dfs:
        return None
    combined = pd.concat(dfs, ignore_index=True)
    return combined

df_a = process_files(files_a, "A") if files_a else None
df_b = process_files(files_b, "B") if files_b else None

# Spaltenzuordnung
st.subheader("2. Spalten zuordnen")

mapping = {}

if df_a is not None:
    det_a = detect_columns(df_a)
    with st.expander(f"Quelle A – erkannte Spalten ({len(df_a)} Zeilen)", expanded=True):
        c1, c2 = st.columns(2)
        with c1:
            artikel_a = st.selectbox(
                "Artikelnummer-Spalte (A)",
                options=df_a.columns.tolist(),
                index=df_a.columns.tolist().index(det_a["artikel"]) if det_a["artikel"] in df_a.columns else 0,
                key="artikel_a"
            )
        with c2:
            preis_a = st.selectbox(
                "Preis-Spalte (A)",
                options=df_a.columns.tolist(),
                index=df_a.columns.tolist().index(det_a["preis"]) if det_a["preis"] in df_a.columns else min(1, len(df_a.columns)-1),
                key="preis_a"
            )
        mapping["A"] = {"artikel": artikel_a, "preis": preis_a, "df": df_a}
        st.dataframe(df_a.head(5), use_container_width=True)

if df_b is not None:
    det_b = detect_columns(df_b)
    with st.expander(f"Quelle B – erkannte Spalten ({len(df_b)} Zeilen)", expanded=True):
        c1, c2 = st.columns(2)
        with c1:
            artikel_b = st.selectbox(
                "Artikelnummer-Spalte (B)",
                options=df_b.columns.tolist(),
                index=df_b.columns.tolist().index(det_b["artikel"]) if det_b["artikel"] in df_b.columns else 0,
                key="artikel_b"
            )
        with c2:
            preis_b = st.selectbox(
                "Preis-Spalte (B)",
                options=df_b.columns.tolist(),
                index=df_b.columns.tolist().index(det_b["preis"]) if det_b["preis"] in df_b.columns else min(1, len(df_b.columns)-1),
                key="preis_b"
            )
        mapping["B"] = {"artikel": artikel_b, "preis": preis_b, "df": df_b}
        st.dataframe(df_b.head(5), use_container_width=True)

if not mapping:
    st.warning("Keine gültigen Dateien geladen.")
    st.stop()

# Daten vorbereiten
prepared = {}

for src, conf in mapping.items():
    df = conf["df"].copy()
    df["_artikel_norm"] = df[conf["artikel"]].apply(normalize_artikelnummer)
    df["_preis"] = clean_price(df[conf["preis"]])
    # Nur Zeilen mit gültiger Artikelnummer behalten
    df = df[df["_artikel_norm"] != ""]
    # Bei Duplikaten den niedrigsten Preis behalten
    df = df.sort_values("_preis").drop_duplicates(subset=["_artikel_norm"], keep="first")
    prepared[src] = df

# Vergleich bauen
st.subheader("3. Preisvergleich")

if len(prepared) == 1:
    # Nur eine Quelle → einfache Liste
    src = list(prepared.keys())[0]
    df = prepared[src][["_artikel_norm", "_preis"]].rename(columns={
        "_artikel_norm": "Artikelnummer",
        "_preis": f"Preis {src} (€)"
    })
    st.dataframe(df, use_container_width=True, height=500)
    
    # Download
    buffer = BytesIO()
    df.to_excel(buffer, index=False, engine="xlsxwriter")
    st.download_button("📥 Als Excel herunterladen", buffer.getvalue(), "teile_liste.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    # Zwei Quellen vergleichen
    a = prepared["A"][["_artikel_norm", "_preis"]].rename(columns={"_preis": "Preis_A"})
    b = prepared["B"][["_artikel_norm", "_preis"]].rename(columns={"_preis": "Preis_B"})
    
    merged = pd.merge(a, b, on="_artikel_norm", how="outer")
    
    merged["Differenz (€)"] = (merged["Preis_A"] - merged["Preis_B"]).round(2)
    merged["Differenz (%)"] = ((merged["Preis_A"] - merged["Preis_B"]) / merged["Preis_B"] * 100).round(1)
    
    def cheapest(row):
        if pd.isna(row["Preis_A"]) and pd.isna(row["Preis_B"]):
            return "–"
        if pd.isna(row["Preis_A"]):
            return "Nur B"
        if pd.isna(row["Preis_B"]):
            return "Nur A"
        if row["Preis_A"] < row["Preis_B"]:
            return "A günstiger"
        if row["Preis_B"] < row["Preis_A"]:
            return "B günstiger"
        return "Gleich"
    
    merged["Günstigste"] = merged.apply(cheapest, axis=1)
    
    # Filter anwenden
    result = merged.copy()
    if show_only_common:
        result = result.dropna(subset=["Preis_A", "Preis_B"])
    if show_diff_only:
        result = result[result["Differenz (€)"].abs() > 0]
    if min_diff > 0:
        result = result[result["Differenz (€)"].abs() >= min_diff]
    
    result = result.rename(columns={
        "_artikel_norm": "Artikelnummer",
        "Preis_A": "Preis A (€)",
        "Preis_B": "Preis B (€)",
    })
    
    # Statistik
    common = merged.dropna(subset=["Preis_A", "Preis_B"])
    only_a = merged[merged["Preis_B"].isna() & merged["Preis_A"].notna()]
    only_b = merged[merged["Preis_A"].isna() & merged["Preis_B"].notna()]
    
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Gemeinsame Artikel", len(common))
    m2.metric("Nur in A", len(only_a))
    m3.metric("Nur in B", len(only_b))
    if len(common) > 0:
        avg_diff = common["Differenz (€)"].mean()
        m4.metric("Ø Differenz (A−B)", f"{avg_diff:+.2f} €")
    
    # Tabelle
    st.dataframe(
        result.sort_values("Artikelnummer"),
        use_container_width=True,
        height=550,
        column_config={
            "Preis A (€)": st.column_config.NumberColumn(format="%.2f €"),
            "Preis B (€)": st.column_config.NumberColumn(format="%.2f €"),
            "Differenz (€)": st.column_config.NumberColumn(format="%+.2f €"),
            "Differenz (%)": st.column_config.NumberColumn(format="%+.1f %%"),
        }
    )
    
    # Download
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        result.to_excel(writer, sheet_name="Vergleich", index=False)
        if len(only_a) > 0:
            only_a[["_artikel_norm", "Preis_A"]].rename(columns={"_artikel_norm": "Artikelnummer", "Preis_A": "Preis A (€)"}).to_excel(writer, sheet_name="Nur in A", index=False)
        if len(only_b) > 0:
            only_b[["_artikel_norm", "Preis_B"]].rename(columns={"_artikel_norm": "Artikelnummer", "Preis_B": "Preis B (€)"}).to_excel(writer, sheet_name="Nur in B", index=False)
    
    st.download_button(
        "📥 Vergleich als Excel herunterladen",
        buffer.getvalue(),
        "preisvergleich_stakis_texkat.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.divider()
st.markdown("""
**Hinweise**
- Artikelnummern werden normalisiert (Leerzeichen, Bindestriche, Punkte entfernt).
- Bei doppelten Artikelnummern in einer Quelle wird der **niedrigste Preis** verwendet.
- Preise werden als Netto-Werte erwartet (je nach Export aus STAkis/TEXKAT prüfen).
- Die App speichert nichts – alles läuft lokal in deinem Browser/Session.
""")
