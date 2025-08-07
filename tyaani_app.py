
import streamlit as st
import pandas as pd

st.set_page_config(page_title="Tyaani Script Generator", layout="centered")
st.title("✨ Tyaani HeyGen Script Generator")
st.markdown("Upload your Tyaani Excel line sheet and get up to 50 ready-to-use marketing scripts.")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls", "csv"])

def spaced(code):
    return ' '.join(str(code))

def get_karat(row):
    weights = {
        '24Carit': row.get('GoldWeight24') or 0,
        '22Carit': row.get('GoldWeight22') or 0,
        '18Carit': row.get('GoldWeight18') or 0,
        '14Carit': row.get('GoldWeight14') or 0,
    }
    weights = {k: v for k, v in weights.items() if isinstance(v, (int, float)) and v > 0}
    return max(weights, key=weights.get) if weights else None

def studded_clause(row):
    parts = []
    if row.get('DiamondPc', 0) > 0:
        parts.append(f"studded with natural cut diamonds, {int(row['DiamondPc'])} pieces weighing {row['DiamondWt']} Carits")
    elif row.get('DiamondWt', 0) > 0:
        parts.append(f"studded with natural cut diamonds, weighing {row['DiamondWt']} Carits")

    if row.get('PolkiPc', 0) > 0:
        parts.append(f"and Natural Syndicate uncut diamonds also called polki, {int(row['PolkiPc'])} pieces weighing {row['PolkiWt']} Carits")
    elif row.get('PolkiWt', 0) > 0:
        parts.append(f"and Natural Syndicate uncut diamonds also called polki, weighing {row['PolkiWt']} Carits")
    return ', '.join(parts)

gem_map = {
    'ThaiRubyWt': 'Thhai rubies',
    'FreshWaterPearlWt': 'Fresshwater pearls',
    'SouthSeaPearlWt': 'South Sea pearls',
    'YellowSapphireWt': 'Yellow Supphires',
    'SapphireWt': 'Supphires',
    'CoralWt': 'Corals',
    'OnyxWt': 'Onyx',
    'MorganiteWt': 'Morganites',
    'IoliteWt': 'Iolites',
    'TanzanitesWt': 'Tanzanites',
    'NavratnaWt': 'Navratanas',
    'TurquoiseWt': 'Turquoises',
    'ZambianEmeraldWt': 'Emerelds',
    'RussianEmeraldWt': 'Emerelds',
}

def gemstone_clause(row):
    parts = []
    for col, label in gem_map.items():
        wt = row.get(col)
        if isinstance(wt, (int, float)) and wt > 0:
            parts.append(f"{label} weighing {wt} Carits")
    return ', '.join(parts)

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, sheet_name="Sheet1")
    except:
        st.error("Could not read 'Sheet1'. Please check your file.")
        st.stop()

    df = df.dropna(how='all')
    scripts = []
    missing = []

    for i, (_, row) in enumerate(df.iterrows()):
        if i >= 50:
            break
        jewelcode = str(row.get('JewelCode', '')).strip()
        netwt = row.get('TotNetwt')
        karat = get_karat(row)

        if not jewelcode or not netwt or netwt == 0 or not karat:
            missing.append(jewelcode)
            continue

        parts = [f"{len(scripts)+1}. This is {row['GrpGroupName']} {spaced(jewelcode)} handcrafted in {karat} hall marked gold"]
        stud = studded_clause(row)
        gems = gemstone_clause(row)
        if stud:
            parts.append(stud)
        if gems:
            parts.append(gems)
        parts.append(f"Net gold weight {netwt} grams. At Tyaani we believe in transparency, every detail is shared with utmost clarity.")
        scripts.append(' '.join(parts))

    txt_output = '\n\n'.join(scripts)
    txt_output += "\n\nMissing Data:\n" + '\n'.join(f"- {sku}" for sku in missing)

    st.download_button("📄 Download Scripts (.txt)", txt_output, file_name="tyaani_scripts.txt")
    st.text_area("Preview Output", txt_output, height=400)
else:
    st.info("Please upload an Excel file with a 'Sheet1' tab.")
