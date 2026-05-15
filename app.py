import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import hashlib
import json
import os
from datetime import datetime
import io

# הגדרות דף
st.set_page_config(page_title="Stack Monitoring Pro", layout="wide")

# --- CSS מעודכן ---
st.markdown("""
    <style>
    .stApp { background-color: #f8f9fa; }
    [data-testid="stFileUploaderFilesList"] { display: none; }
    
    .graph-card {
        background-color: white;
        padding: 10px 25px 25px 25px;
        border-radius: 15px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.05);
        border: 1px solid #eaeaea;
        margin-bottom: 25px;
    }
    
    .settings-header {
        color: #1e293b;
        border-bottom: 2px solid #3b82f6;
        padding-bottom: 10px;
        margin-bottom: 20px;
    }
    
    h1 { color: #1e293b; font-weight: 800 !important; }
    section[data-testid="stSidebar"] { background-color: #1e293b; color: white; }
    section[data-testid="stSidebar"] h1, section[data-testid="stSidebar"] h2 { color: white; }
    </style>
    """, unsafe_allow_html=True)

DB_FILE = "stack_db.json"

# --- פונקציות ליבה ---

def load_data():
    """
    Load the persisted database from disk.  The JSON structure holds the following keys:
      - samples: list of measurement dicts (pollutant, concentration, date, chimney_id, filename)
      - thresholds: nested dict of threshold values keyed by chimney_id and pollutant name
      - hashes: list of MD5 hashes of previously uploaded files (used to prevent duplicates)
      - chimney_names: mapping of chimney_id to a user friendly name
      - plants: mapping of chimney_id to the plant name it belongs to

    Returns a tuple (df, thresholds, file_hashes, chimney_names, plants).
    If the file does not exist or cannot be parsed, sensible defaults are returned.
    """
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                df = pd.DataFrame(data.get("samples", []))
                if not df.empty:
                    df['date'] = pd.to_datetime(df['date'], errors='coerce')
                    df = df.dropna(subset=['date'])
                    # Normalise pollutant names on load – collapse any consecutive whitespace so that
                    # "תחמוצות  גופרית" and "תחמוצות גופרית" are treated the same.
                    if 'pollutant' in df.columns:
                        df['pollutant'] = df['pollutant'].astype(str).apply(lambda x: " ".join(x.split()))
                thresholds = data.get("thresholds", {})
                file_hashes = set(data.get("hashes", []))
                chimney_names = data.get("chimney_names", {})
                plants = data.get("plants", {})
                return df, thresholds, file_hashes, chimney_names, plants
        except Exception:
            pass
    # Fallback defaults
    return pd.DataFrame(columns=["pollutant", "concentration", "date", "chimney_id", "filename"]), {}, set(), {}, {}

def save_data(df, thresholds, file_hashes, chimney_names, plants):
    """Persist the current in-memory state to disk."""
    df_to_save = df.copy()
    if not df_to_save.empty:
        # Store dates as strings to avoid JSON serialisation issues
        df_to_save['date'] = df_to_save['date'].dt.strftime('%Y-%m-%d %H:%M:%S')
    data = {
        "samples": df_to_save.to_dict(orient="records"),
        "thresholds": thresholds,
        "hashes": list(file_hashes),
        "chimney_names": chimney_names,
        "plants": plants,
    }
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

def to_excel_custom(df, chimney_name, pollutant_name):
    """מייצר קובץ אקסל עם כותרות מותאמות אישית"""
    output = io.BytesIO()
    
    # הכנת הנתונים למבנה המבוקש
    export_df = df.copy()
    export_df = export_df.rename(columns={
        'date': 'תאריך דיגום',
        'concentration': 'ריכוז (מ"ג/מק"ת)'
    })
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # כתיבת הנתונים החל משורה 3
        export_df.to_excel(writer, index=False, sheet_name='נתוני דיגום', startrow=2)
        
        # הוספת כותרת עליונה בשורה הראשונה
        workbook = writer.book
        worksheet = writer.sheets['נתוני דיגום']
        header_text = f"דוח נתוני דיגום: {chimney_name} | מזהם: {pollutant_name}"
        worksheet.merge_cells('A1:B1')
        worksheet['A1'] = header_text
        
    return output.getvalue()


# פונקציה ליצוא נתונים של מספר מזהמים ו/או ארובות לקובץ אקסל עם כותרת מותאמת
def to_excel_multi(df: pd.DataFrame, header_text: str):
    """
    Export a DataFrame containing multiple chimneys and pollutants to an Excel file with a custom header.

    The DataFrame is expected to have at least the columns: 'chimney_id', 'pollutant', 'date', 'concentration'.
    The resulting spreadsheet will include Hebrew column names and a merged header row.
    """
    output = io.BytesIO()
    if df.empty:
        # Create an empty Excel with header text only
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            workbook = writer.book
            sheet = workbook.create_sheet('נתוני דיגום')
            sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
            sheet['A1'] = header_text
        return output.getvalue()

    export_df = df.copy()
    # Rename columns to Hebrew for end users
    rename_map = {
        'chimney_id': 'ארובה',
        'pollutant': 'מזהם',
        'date': 'תאריך דיגום',
        'concentration': 'ריכוז (מ"ג/מק"ת)',
    }
    export_df = export_df.rename(columns={k: v for k, v in rename_map.items() if k in export_df.columns})
    # Format dates nicely
    if 'תאריך דיגום' in export_df.columns:
        export_df['תאריך דיגום'] = pd.to_datetime(export_df['תאריך דיגום'], errors='coerce').dt.strftime('%d/%m/%Y')
    # Write to Excel
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        export_df.to_excel(writer, index=False, sheet_name='נתוני דיגום', startrow=2)
        workbook = writer.book
        worksheet = writer.sheets['נתוני דיגום']
        # Merge cells across the number of columns to create a single header
        num_cols = export_df.shape[1]
        worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_cols)
        worksheet['A1'] = header_text
    return output.getvalue()

def process_excel(uploaded_file):
    """
    Parse a raw Excel file exported from the Ministry of Environment stack sampling form.

    This helper is deliberately resilient to minor variations in the underlying file:
    * The stack ID ("זיהוי ארובה:") is expected to appear on row 3, immediately to the right of the label.
    * The sampling date ("תאריך הדיגום:") is expected to appear on row 7, immediately to the right of the label.
    * Measurement results can appear in one of several columns (normalised, measured, LOQ or LOD).  The
      function will always try the normalised value first, then the measured concentration, then the
      quantification limit and finally the detection limit.  If the value is stored as a string beginning
      with "<" (e.g. "< LOQ" or "< LOD") the numeric part will be extracted and used.  Any row that
      cannot be converted into a float will be skipped rather than causing the entire import to fail.
    * Pollutant names are normalised by collapsing consecutive whitespace into a single space and
      trimming leading/trailing spaces.  This helps consolidate variants like "תחמוצות  גופרית" and
      "תחמוצות גופרית" under a single key.
    """
    try:
        # Read the file without assuming headers – the template uses fixed offsets.
        df_raw = pd.read_excel(uploaded_file, header=None)

        # Extract the chimney identifier and sampling date from their respective cells.
        chimney_id = str(df_raw.iloc[3, 7]).strip()
        sampling_date = pd.to_datetime(df_raw.iloc[7, 13], errors='coerce')
        if pd.isnull(sampling_date):
            return None, None

        data_rows = []
        # Iterate over potential result rows.  The results start at row 10 in the provided template.
        for i in range(10, len(df_raw)):
            raw_pollutant = df_raw.iloc[i, 2] if df_raw.shape[1] > 2 else None
            pollutant = str(raw_pollutant).strip() if raw_pollutant is not None else ""
            if not pollutant or pollutant.lower() == "nan":
                continue
            # Normalise pollutant name by squeezing extra spaces and stripping.
            pollutant = " ".join(pollutant.split())

            # Attempt to retrieve a numeric concentration from the preferred columns.
            concentration = None
            for col_idx in [16, 13, 15, 14]:
                # Guard against files with fewer columns
                if col_idx >= df_raw.shape[1]:
                    continue
                val = df_raw.iloc[i, col_idx]
                if pd.isnull(val) or str(val).strip() == "":
                    continue
                # Convert the value to a float.  Some values may be strings like "< LOQ" or "<0.5".
                val_str = str(val).strip()
                # If the value begins with '<' (below detection/quantification), strip the symbol and
                # attempt to interpret the remainder.  If conversion fails, skip this column.
                if val_str.startswith("<"):
                    val_str = val_str.lstrip("<").strip()
                try:
                    numeric_val = float(val_str)
                except Exception:
                    # If we can't coerce to a number then move on to the next candidate column.
                    continue
                concentration = round(numeric_val, 1)
                break
            # Only record rows where we were able to derive a concentration
            if concentration is not None:
                data_rows.append({
                    "pollutant": pollutant,
                    "concentration": concentration,
                    "date": sampling_date,
                    "chimney_id": chimney_id,
                    "filename": uploaded_file.name
                })
        # Build a DataFrame from all parsed rows.  If no rows were parsed return None
        if not data_rows:
            return None, chimney_id
        return pd.DataFrame(data_rows), chimney_id
    except Exception:
        # Suppress noisy traceback for end users – a silent failure will be reported upstream.
        return None, None

# --- Session State ---
if 'db' not in st.session_state:
    (
        st.session_state.db,
        st.session_state.thresholds,
        st.session_state.file_hashes,
        st.session_state.chimney_names,
        st.session_state.plants,
    ) = load_data()

# --- ממשק משתמש ---

st.title("📊 Stack Monitoring Analytics")

with st.sidebar:
    st.markdown("## 📥 העלאת נתונים")
    uploaded_files = st.file_uploader("בחר קבצי אקסל (XLSX)", type=["xlsx"], accept_multiple_files=True, label_visibility="collapsed")
    
    if uploaded_files:
        added = False
        for file in uploaded_files:
            f_hash = hashlib.md5(file.getvalue()).hexdigest()
            # Allow a file to be re‑uploaded if its name no longer exists in the DB.  This
            # accommodates deletions where the hash remains in the set but the rows have been purged.
            if f_hash not in st.session_state.file_hashes or file.name not in st.session_state.db.get('filename', pd.Series(dtype=str)).unique():
                new_data, _ = process_excel(file)
                if new_data is not None:
                    st.session_state.db = pd.concat([st.session_state.db, new_data], ignore_index=True)
                    st.session_state.file_hashes.add(f_hash)
                    added = True
        if added:
            st.session_state.db['date'] = pd.to_datetime(st.session_state.db['date'])
            save_data(
                st.session_state.db,
                st.session_state.thresholds,
                st.session_state.file_hashes,
                st.session_state.chimney_names,
                st.session_state.plants,
            )
            st.rerun()

tab1, tab2, tab3 = st.tabs(["📈 דשבורד", "⚙️ הגדרות ארובות", "💾 ניהול מסד נתונים"])

with tab1:
    if not st.session_state.db.empty:
        # Group chimneys by plant; default group for unassigned chimneys
        chimney_ids = sorted(st.session_state.db['chimney_id'].unique())
        # Determine plant name for each chimney; fallback to 'ללא מפעל'
        plants_mapping = st.session_state.plants if hasattr(st.session_state, 'plants') else {}
        plant_names = sorted({plants_mapping.get(cid, 'ללא מפעל') for cid in chimney_ids})
        # Select plant
        selected_plant = st.selectbox("בחר מפעל:", options=plant_names)
        # Filter chimneys belonging to selected plant
        chimneys_for_plant = [cid for cid in chimney_ids if plants_mapping.get(cid, 'ללא מפעל') == selected_plant]
        # Build label mapping for those chimneys
        id_to_label = {cid: st.session_state.chimney_names.get(cid, f"ארובה {cid}") for cid in chimneys_for_plant}
        if not id_to_label:
            st.warning("לא נמצאו ארובות במפעל הנבחר.")
        else:
            selected_label = st.selectbox("בחר ארובה לתצוגה:", options=list(id_to_label.values()))
            # Retrieve the corresponding chimney_id
            selected_id = [k for k, v in id_to_label.items() if v == selected_label][0]
            # Filter data for the selected chimney
            c_data = st.session_state.db[st.session_state.db['chimney_id'] == selected_id]
            # Iterate over pollutants and draw graphs
            for poll in sorted(c_data['pollutant'].unique()):
                poll_data = c_data[c_data['pollutant'] == poll].sort_values('date')
                threshold = int(st.session_state.thresholds.get(selected_id, {}).get(poll, 0))
                st.markdown(f'<div class="graph-card">', unsafe_allow_html=True)
                col_spacer, col_dl = st.columns([10, 1.5])
                with col_dl:
                    # Export only this pollutant and chimney
                    df_for_excel = poll_data[['date', 'concentration']].copy()
                    df_for_excel['date'] = df_for_excel['date'].dt.strftime('%d/%m/%Y')
                    excel_data = to_excel_custom(df_for_excel, selected_label, poll)
                    st.download_button("Excel 📥", data=excel_data, file_name=f"{selected_label}_{poll}.xlsx", key=f"dl_{selected_id}_{poll}")
                # Create figure
                fig = go.Figure()
                fig.add_trace(go.Scatter(
                    x=poll_data['date'].dt.strftime('%d/%m/%Y'),
                    y=poll_data['concentration'],
                    mode='lines+markers',
                    line=dict(color='#3b82f6', width=3),
                    marker=dict(size=8, color='#1e293b', line=dict(width=2, color='#3b82f6')),
                    fill='tozeroy',
                    fillcolor='rgba(59, 130, 246, 0.05)',
                    name=f"ריכוז (מ\"ג/מק\"ת)",
                    hovertemplate="תאריך: %{x}<br>ריכוז: %{y:.1f} מ\"ג/מק\"ת<extra></extra>"
                ))
                if threshold > 0:
                    fig.add_hline(y=threshold, line_dash="dash", line_color="#ef4444",
                                  annotation_text=f"סף תקן: {threshold} מ\"ג/מק\"ת", annotation_position="top left")
                fig.update_layout(
                    title={
                        'text': f"<b>{selected_label} - {poll}</b><br><span style='font-size:14px; color:gray;'>(מ\"ג/מק\"ת)</span>",
                        'y': 0.95, 'x': 0.5, 'xanchor': 'center', 'yanchor': 'top'
                    },
                    # Slightly smaller graph height to allow multiple plots to fit comfortably on screen
                    height=280,
                    margin=dict(l=20, r=20, t=70, b=20),
                    template="plotly_white",
                    xaxis=dict(type='category', gridcolor='#f0f2f6', title="תאריך דיגום"),
                    yaxis=dict(gridcolor='#f0f2f6', title="ריכוז (מ\"ג/מק\"ת)"),
                    hovermode="x unified",
                    dragmode="zoom"
                )
                st.plotly_chart(
                    fig,
                    use_container_width=True,
                    config={
                        'displaylogo': False,
                        'toImageButtonOptions': {
                            'format': 'png',
                            'filename': f'{selected_label}_{poll}',
                            'height': 800,
                            'width': 1200,
                            'scale': 2
                        }
                    },
                    key=f"chart_{selected_id}_{poll}"
                )
                st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.info("אנא העלה קבצים בתפריט הצד.")

# (שאר הקוד של לשונית הגדרות וניהול נתונים נשאר זהה...)
with tab2:
    st.markdown("<h2 class='settings-header'>🏷️ הגדרות ארובות וערכי סף</h2>", unsafe_allow_html=True)
    if not st.session_state.db.empty:
        ids = sorted(st.session_state.db['chimney_id'].unique())
        for cid in ids:
            with st.expander(f"⚙️ הגדרות ארובה: {st.session_state.chimney_names.get(cid, cid)}", expanded=True):
                # Two columns: edit chimney name and assign plant
                col1, col2 = st.columns([1, 1])
                with col1:
                    old_name = st.session_state.chimney_names.get(cid, "")
                    new_name = st.text_input(f"שם הארובה (ID: {cid})", value=old_name, key=f"edit_name_{cid}")
                    if new_name != old_name:
                        st.session_state.chimney_names[cid] = new_name
                        save_data(
                            st.session_state.db,
                            st.session_state.thresholds,
                            st.session_state.file_hashes,
                            st.session_state.chimney_names,
                            st.session_state.plants,
                        )
                        st.rerun()
                with col2:
                    old_plant = st.session_state.plants.get(cid, "") if hasattr(st.session_state, 'plants') else ""
                    new_plant = st.text_input(f"שם המפעל (ID: {cid})", value=old_plant, key=f"edit_plant_{cid}")
                    if new_plant != old_plant:
                        st.session_state.plants[cid] = new_plant
                        save_data(
                            st.session_state.db,
                            st.session_state.thresholds,
                            st.session_state.file_hashes,
                            st.session_state.chimney_names,
                            st.session_state.plants,
                        )
                        st.rerun()
                
                st.markdown("---")
                st.markdown("**הגדרת ערכי סף (מ\"ג/מק\"ת):**")
                chimney_pollutants = sorted(st.session_state.db[st.session_state.db['chimney_id'] == cid]['pollutant'].unique())
                cols = st.columns(3)
                for idx, p_name in enumerate(chimney_pollutants):
                    with cols[idx % 3]:
                        if cid not in st.session_state.thresholds:
                            st.session_state.thresholds[cid] = {}
                        curr_limit = int(st.session_state.thresholds[cid].get(p_name, 0))
                        new_limit = st.number_input(
                            f"סף {p_name}",
                            value=curr_limit,
                            key=f"limit_{cid}_{p_name}",
                            step=1
                        )
                        if new_limit != curr_limit:
                            st.session_state.thresholds[cid][p_name] = int(new_limit)
                            save_data(
                                st.session_state.db,
                                st.session_state.thresholds,
                                st.session_state.file_hashes,
                                st.session_state.chimney_names,
                                st.session_state.plants,
                            )
    else:
        st.info("העלה נתונים כדי להתחיל בהגדרות.")

with tab3:
    st.markdown("<h2 class='settings-header'>💾 ניהול נתונים</h2>", unsafe_allow_html=True)
    if not st.session_state.db.empty:
        # Build plant and chimney selection for file management
        chimney_ids_dm = sorted(st.session_state.db['chimney_id'].unique())
        plants_mapping_dm = st.session_state.plants if hasattr(st.session_state, 'plants') else {}
        plant_names_dm = sorted({plants_mapping_dm.get(cid, 'ללא מפעל') for cid in chimney_ids_dm})
        selected_plant_dm = st.selectbox("בחר מפעל לניהול קבצים:", options=plant_names_dm, key="dm_plant")
        chimneys_in_plant_dm = [cid for cid in chimney_ids_dm if plants_mapping_dm.get(cid, 'ללא מפעל') == selected_plant_dm]
        if not chimneys_in_plant_dm:
            st.info("אין ארובות במפעל זה.")
        else:
            id_to_label_dm = {cid: st.session_state.chimney_names.get(cid, f"ארובה {cid}") for cid in chimneys_in_plant_dm}
            selected_label_dm = st.selectbox("בחר ארובה לניהול קבצים:", options=list(id_to_label_dm.values()), key="dm_chimney")
            selected_chimney_dm = [k for k, v in id_to_label_dm.items() if v == selected_label_dm][0]
            # List files associated with this chimney and provide download/delete options
            sub_files = st.session_state.db[
                st.session_state.db['chimney_id'] == selected_chimney_dm
            ][['filename', 'date']].drop_duplicates().copy()
            if sub_files.empty:
                st.info("אין קבצי דיגום לארובה זו.")
            else:
                # Create a human‑readable label for each file including the date
                sub_files['display'] = sub_files.apply(
                    lambda x: f"{x['filename']} ({x['date'].strftime('%d/%m/%Y')})", axis=1
                )
                st.markdown("### 📁 הורדה או מחיקה של קובצי דיגום")
                # File selection for download
                selected_file_disp = st.selectbox(
                    "בחר קובץ להורדה:",
                    options=["--"] + sub_files['display'].tolist(),
                    key="dm_download_file"
                )
                if selected_file_disp and selected_file_disp != "--":
                    # Parse file name and date from display label
                    file_name_part = selected_file_disp.split(" (")[0]
                    date_part = selected_file_disp.split(" (")[1].rstrip(")")
                    # Filter database for the selected file (by chimney, file name and date)
                    export_file_df = st.session_state.db[
                        (st.session_state.db['chimney_id'] == selected_chimney_dm)
                        & (st.session_state.db['filename'] == file_name_part)
                        & (st.session_state.db['date'].dt.strftime('%d/%m/%Y') == date_part)
                    ]
                    # Build a descriptive header for this file export
                    header_file = f"דוח נתוני דיגום: {selected_plant_dm} | {selected_label_dm} | תאריך: {date_part}"
                    excel_file_dm = to_excel_multi(export_file_df, header_file)
                    st.download_button(
                        "הורד קובץ דיגום 📥",
                        data=excel_file_dm,
                        file_name=f"{selected_label_dm}_{date_part}.xlsx",
                        key=f"dl_file_{selected_file_disp}"
                    )
                # Multi‑select for deletion
                to_del_dm = st.multiselect(
                    "בחר קבצי דיגום להסרה מהמערכת:",
                    options=sub_files['display'].tolist(),
                    key="dm_files"
                )
                if st.button("מחק קבצים לצמיתות", type="primary"):
                    filenames_dm = [d.split(" (")[0] for d in to_del_dm]
                    dates_dm = [d.split(" (")[1].rstrip(")") for d in to_del_dm]
                    # Remove rows matching both file name and date for this chimney
                    mask = (
                        (st.session_state.db['chimney_id'] == selected_chimney_dm) &
                        (st.session_state.db['filename'].isin(filenames_dm)) &
                        (st.session_state.db['date'].dt.strftime('%d/%m/%Y').isin(dates_dm))
                    )
                    st.session_state.db = st.session_state.db[~mask]
                    save_data(
                        st.session_state.db,
                        st.session_state.thresholds,
                        st.session_state.file_hashes,
                        st.session_state.chimney_names,
                        st.session_state.plants,
                    )
                    st.rerun()
            # Export data section
            st.markdown("---")
            st.markdown("### 📤 יצוא נתונים לאקסל")
            # Available pollutants for export from the current plant selection
            pollutant_options_dm = sorted(
                st.session_state.db[
                    st.session_state.db['chimney_id'].isin(chimneys_in_plant_dm)
                ]['pollutant'].unique()
            )
            export_pollutants_dm = st.multiselect(
                "בחר מזהמים ליצוא:",
                options=pollutant_options_dm,
                default=pollutant_options_dm,
                key="export_pollutants_dm"
            )
            export_scope_dm = st.selectbox(
                "בחר טווח יצוא:",
                options=["ארובה ספציפית", "כל הארובות במפעל הנבחר", "כל הארובות"],
                key="export_scope_dm"
            )
            # Build DataFrame for export based on scope
            if export_scope_dm == "ארובה ספציפית":
                export_df_dm = st.session_state.db[
                    (st.session_state.db['chimney_id'] == selected_chimney_dm) &
                    (st.session_state.db['pollutant'].isin(export_pollutants_dm))
                ]
                header_dm = f"דוח נתוני דיגום: {selected_plant_dm} | {selected_label_dm} | מזהמים: {', '.join(export_pollutants_dm)}"
            elif export_scope_dm == "כל הארובות במפעל הנבחר":
                export_df_dm = st.session_state.db[
                    (st.session_state.db['chimney_id'].isin(chimneys_in_plant_dm)) &
                    (st.session_state.db['pollutant'].isin(export_pollutants_dm))
                ]
                header_dm = f"דוח נתוני דיגום: {selected_plant_dm} | כל הארובות | מזהמים: {', '.join(export_pollutants_dm)}"
            else:  # "כל הארובות"
                export_df_dm = st.session_state.db[
                    st.session_state.db['pollutant'].isin(export_pollutants_dm)
                ]
                header_dm = f"דוח נתוני דיגום: כל המפעלים | כל הארובות | מזהמים: {', '.join(export_pollutants_dm)}"
            if export_df_dm.empty:
                st.info("אין נתונים לייצוא לפי הסינון שנבחר.")
            else:
                excel_multi_dm = to_excel_multi(export_df_dm, header_dm)
                st.download_button(
                    "יצא לאקסל 📥", data=excel_multi_dm, file_name="exported_data.xlsx", key="export_multi_dm"
                )
