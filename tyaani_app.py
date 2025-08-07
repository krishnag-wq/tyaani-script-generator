
import streamlit as st
import pandas as pd
import io
import re

def generate_scripts(df, start_index):
    scripts = []
    missing_data = []

    gemstone_mapping = [
        ('ThaiRubyWt', 'Thhai rubies'),
        ('FreshWaterPearlWt', 'Fresshwater pearls'),
        ('SouthSeaPearlWt', 'South Sea pearls'),
        ('YellowSapphireWt', 'Yellow Supphires'),
        ('SapphireWt', 'Supphires'),
        ('CoralWt', 'Corals'),
        ('OnyxWt', 'Onyx'),
        ('MorganiteWt', 'Morganites'),
        ('IoliteWt', 'Iolites'),
        ('TanzanitesWt', 'Tanzanites'),
        ('NavratnaWt', 'Navratanas'),
        ('TurquoiseWt', 'Turquoises'),
        ('ZambianEmeraldWt', 'Emerelds'),
        ('RussianEmeraldWt', 'Emerelds')
    ]

    gold_columns = {
        'GoldWeight24': '24Carit',
        'GoldWeight22': '22Carit',
        'GoldWeight18': '18Carit',
        'GoldWeight14': '14Carit'
    }

    df = df.iloc[start_index:start_index+50].copy()

    for idx, row in df.iterrows():
        jewel_code = row.get('JewelCode', '')
        if not isinstance(jewel_code, str):
            jewel_code = str(jewel_code)
        spaced_jewel_code = ' '.join(jewel_code.upper())

        category = row.get('GrpGroupName', 'Jewelry')
        net_weight = row.get('TotNetwt')

        gold_weights = {col: row.get(col, 0) for col in gold_columns}
        valid_gold = {k: v for k, v in gold_weights.items() if pd.notna(v) and v > 0}
        gold_clause = ''

        if valid_gold:
            best_gold_col = max(valid_gold, key=lambda x: valid_gold[x])
            karat = gold_columns[best_gold_col]
            gold_clause = f'handcrafted in {karat} hall marked gold'
        else:
            missing_data.append(jewel_code)
            karat = ''

        diamond_clause = ''
        if row.get('DiamondPc', 0) > 0:
            diamond_clause = f"studded with natural cut diamonds, {int(row['DiamondPc'])} pieces weighing {row['DiamondWt']} Carits"
        elif row.get('DiamondWt', 0) > 0:
            diamond_clause = f"studded with natural cut diamonds, weighing {row['DiamondWt']} Carits"

        polki_clause = ''
        if row.get('PolkiPc', 0) > 0:
            polki_clause = f"and Natural Syndicate uncut diamonds also called polki, {int(row['PolkiPc'])} pieces weighing {row['PolkiWt']} Carits"
        elif row.get('PolkiWt', 0) > 0:
            polki_clause = f"and Natural Syndicate uncut diamonds also called polki, weighing {row['PolkiWt']} Carits"

        gemstone_clauses = []
        for col, label in gemstone_mapping:
            wt = row.get(col)
            if pd.notna(wt) and wt > 0:
                gemstone_clauses.append(f"{label} weighing {wt} Carits")

        all_gems = ', '.join(gemstone_clauses)

        if pd.isna(net_weight) or net_weight == 0:
            missing_data.append(jewel_code)

        parts = [
            f"This is {category} {spaced_jewel_code}"
        ]

        if gold_clause:
            parts.append(gold_clause)
        if diamond_clause:
            parts.append(diamond_clause)
        if polki_clause:
            parts.append(polki_clause)
        if all_gems:
            parts.append(all_gems)

        if pd.notna(net_weight) and net_weight != 0:
            parts.append(f"Net gold weight {net_weight} grams")

        parts.append("At Tyaani we believe in transparency, every detail is shared with utmost clarity.")

        final_script = '. '.join([p.strip().rstrip('.') for p in parts]) + '.'
        scripts.append(final_script)

    return scripts, missing_data

st.set_page_config(page_title="Tyaani Script Generator", layout="centered")
st.title("📜 Tyaani HeyGen Script Generator")

st.markdown("Upload your Excel file with a sheet named `Sheet1` and script columns.")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])
start_row = st.number_input("Start from SKU #", min_value=1, step=50, value=1)

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, sheet_name='Sheet1')
        scripts, missing = generate_scripts(df, start_row - 1)

        if scripts:
            result = "\n\n".join([f"{i+1}. {s}" for i, s in enumerate(scripts)])
            if missing:
                result += "\n\nMissing Data:\n" + "\n".join([f"- {sku}" for sku in missing])

            st.text_area("Generated Scripts", value=result, height=600)
            output_filename = f"Tyaani_Scripts_{start_row}_to_{start_row+49}.txt"
            st.download_button("📥 Download .txt", data=result, file_name=output_filename, mime="text/plain")

    except Exception as e:
        st.error(f"Error processing file: {e}")
