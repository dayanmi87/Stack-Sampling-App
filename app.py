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
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                df = pd.DataFrame(data.get("samples", []))
                if not df.empty:
                    df['date'] = pd.to_datetime(df['date'], errors='coerce')
                    df = df.dropna(subset=['date'])
                return df, data.get("thresholds", {}), set(data.get("hashes", [])), data.get("chimney_names", {})
        except: pass
    return pd.DataFrame(columns=["pollutant", "concentration", "date", "chimney_id", "filename"]), {}, set(), {}

def save_data(df, thresholds, file_hashes, chimney_names):
    df_to_save = df.copy()
    if not df_to_save.empty:
        df_to_save['date'] = df_to_save['date'].dt.strftime('%Y-%m-%d %H:%M:%S')
    data = {"samples": df_to_save.to_dict(orient="records"), "thresholds": thresholds, "hashes": list(file_hashes), "chimney_names": chimney_names}
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

def process_excel(uploaded_file):
    try:
        df_raw = pd.read_excel(uploaded_file, header=None)
        chimney_id = str(df_raw.iloc[3, 7]).strip()
        sampling_date = pd.to_datetime(df_raw.iloc[7, 13], errors='coerce')
        if pd.isnull(sampling_date): return None, None
        data_rows = []
        for i in range(10, len(df_raw)):
            pollutant = str(df_raw.iloc[i, 2]).strip()
            if not pollutant or pollutant.lower() == "nan": continue
            for col_idx in [16, 13, 15, 14]:
                val = df_raw.iloc[i, col_idx]
                if pd.notnull(val) and str(val).strip() != "":
                    data_rows.append({"pollutant": pollutant, "concentration": round(float(val), 1), 
                                     "date": sampling_date, "chimney_id": chimney_id, "filename": uploaded_file.name})
                    break
        return pd.DataFrame(data_rows), chimney_id
    except: return None, None

# --- Session State ---
if 'db' not in st.session_state:
    st.session_state.db, st.session_state.thresholds, st.session_state.file_hashes, st.session_state.chimney_names = load_data()

# --- ממשק משתמש ---

st.title("📊 Stack Monitoring Analytics")

with st.sidebar:
    st.markdown("## 📥 העלאת נתונים")
    uploaded_files = st.file_uploader("בחר קבצי אקסל (XLSX)", type=["xlsx"], accept_multiple_files=True, label_visibility="collapsed")
    
    if uploaded_files:
        added = False
        for file in uploaded_files:
            f_hash = hashlib.md5(file.getvalue()).hexdigest()
            if f_hash not in st.session_state.file_hashes:
                new_data, _ = process_excel(file)
                if new_data is not None:
                    st.session_state.db = pd.concat([st.session_state.db, new_data], ignore_index=True)
                    st.session_state.file_hashes.add(f_hash)
                    added = True
        if added:
            st.session_state.db['date'] = pd.to_datetime(st.session_state.db['date'])
            save_data(st.session_state.db, st.session_state.thresholds, st.session_state.file_hashes, st.session_state.chimney_names)
            st.rerun()

tab1, tab2, tab3 = st.tabs(["📈 דשבורד", "⚙️ הגדרות ארובות", "💾 ניהול מסד נתונים"])

with tab1:
    if not st.session_state.db.empty:
        raw_ids = sorted(st.session_state.db['chimney_id'].unique())
        id_to_label = {cid: st.session_state.chimney_names.get(cid, f"ארובה {cid}") for cid in raw_ids}
        
        selected_label = st.selectbox("בחר ארובה לתצוגה:", options=list(id_to_label.values()))
        selected_id = [k for k, v in id_to_label.items() if v == selected_label][0]
        
        c_data = st.session_state.db[st.session_state.db['chimney_id'] == selected_id]
        
        for poll in sorted(c_data['pollutant'].unique()):
            poll_data = c_data[c_data['pollutant'] == poll].sort_values('date')
            threshold = int(st.session_state.thresholds.get(selected_id, {}).get(poll, 0))
            
            st.markdown(f'<div class="graph-card">', unsafe_allow_html=True)
            
            col_spacer, col_dl = st.columns([10, 1.5])
            with col_dl:
                # שימוש בפונקציה החדשה לייצוא עם כותרות
                df_for_excel = poll_data[['date', 'concentration']].copy()
                df_for_excel['date'] = df_for_excel['date'].dt.strftime('%d/%m/%Y')
                excel_data = to_excel_custom(df_for_excel, selected_label, poll)
                st.download_button("Excel 📥", data=excel_data, file_name=f"{selected_label}_{poll}.xlsx", key=f"dl_{selected_id}_{poll}")

            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=poll_data['date'].dt.strftime('%d/%m/%Y'),
                y=poll_data['concentration'],
                mode='lines+markers',
                line=dict(color='#3b82f6', width=4),
                marker=dict(size=10, color='#1e293b', line=dict(width=2, color='#3b82f6')),
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
                height=450,
                margin=dict(l=20, r=20, t=80, b=20),
                template="plotly_white",
                xaxis=dict(type='category', gridcolor='#f0f2f6', title="תאריך דיגום"),
                yaxis=dict(gridcolor='#f0f2f6', title="ריכוז (מ\"ג/מק\"ת)"),
                hovermode="x unified",
                dragmode="zoom"
            )
            
            st.plotly_chart(fig, use_container_width=True, config={
                'displaylogo': False,
                'toImageButtonOptions': {
                    'format': 'png', 
                    'filename': f'{selected_label}_{poll}',
                    'height': 800,
                    'width': 1200,
                    'scale': 2
                }
            }, key=f"chart_{selected_id}_{poll}")
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
                col1, col2 = st.columns([1, 1])
                with col1:
                    old_name = st.session_state.chimney_names.get(cid, "")
                    new_name = st.text_input(f"שם הארובה (ID: {cid})", value=old_name, key=f"edit_name_{cid}")
                    if new_name != old_name:
                        st.session_state.chimney_names[cid] = new_name
                        save_data(st.session_state.db, st.session_state.thresholds, st.session_state.file_hashes, st.session_state.chimney_names)
                        st.rerun()
                
                st.markdown("---")
                st.markdown("**הגדרת ערכי סף (מ\"ג/מק\"ת):**")
                chimney_pollutants = sorted(st.session_state.db[st.session_state.db['chimney_id'] == cid]['pollutant'].unique())
                cols = st.columns(3)
                for idx, p_name in enumerate(chimney_pollutants):
                    with cols[idx % 3]:
                        if cid not in st.session_state.thresholds: st.session_state.thresholds[cid] = {}
                        curr_limit = int(st.session_state.thresholds[cid].get(p_name, 0))
                        new_limit = st.number_input(f"סף {p_name}", value=curr_limit, key=f"limit_{cid}_{p_name}", step=1)
                        if new_limit != curr_limit:
                            st.session_state.thresholds[cid][p_name] = int(new_limit)
                            save_data(st.session_state.db, st.session_state.thresholds, st.session_state.file_hashes, st.session_state.chimney_names)
    else:
        st.info("העלה נתונים כדי להתחיל בהגדרות.")

with tab3:
    st.markdown("<h2 class='settings-header'>💾 ניהול נתונים</h2>", unsafe_allow_html=True)
    if not st.session_state.db.empty:
        files = st.session_state.db[['filename', 'chimney_id', 'date']].drop_duplicates().copy()
        files['display'] = files.apply(lambda x: f"{x['filename']} ({st.session_state.chimney_names.get(x['chimney_id'], x['chimney_id'])} - {x['date'].strftime('%d/%m/%Y')})", axis=1)
        
        to_del = st.multiselect("בחר קבצים להסרה מהמערכת:", options=files['display'].tolist())
        if st.button("מחק לצמיתות", type="primary"):
            filenames = [d.split(" (")[0] for d in to_del]
            st.session_state.db = st.session_state.db[~st.session_state.db['filename'].isin(filenames)]
            save_data(st.session_state.db, st.session_state.thresholds, st.session_state.file_hashes, st.session_state.chimney_names)
            st.rerun()
