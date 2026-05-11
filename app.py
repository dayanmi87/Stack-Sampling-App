import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import hashlib
import json
import os
from datetime import datetime
import io

# הגדרות דף - מותאם גם לנייד
st.set_page_config(
    page_title="Stack Monitoring Pro", 
    layout="wide", 
    initial_sidebar_state="collapsed"
)

# --- CSS מעודכן למראה יוקרתי ומותאם לנייד ---
st.markdown("""
    <style>
    .stApp { background-color: #f8f9fa; }
    [data-testid="stFileUploaderFilesList"] { display: none; }
    
    .graph-card {
        background-color: white;
        padding: 15px 20px;
        border-radius: 15px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.05);
        border: 1px solid #eaeaea;
        margin-bottom: 20px;
    }
    
    .settings-header {
        color: #1e293b;
        border-bottom: 2px solid #3b82f6;
        padding-bottom: 10px;
        margin-bottom: 20px;
    }
    
    @media (max-width: 640px) {
        .graph-card { padding: 10px; }
        h1 { font-size: 1.5rem !important; }
    }
    
    h1 { color: #1e293b; font-weight: 800 !important; }
    section[data-testid="stSidebar"] { background-color: #1e293b; color: white; }
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
    output = io.BytesIO()
    export_df = df.copy().rename(columns={'date': 'תאריך דיגום', 'concentration': 'ריכוז (מ"ג/מק\"ת)'})
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        export_df.to_excel(writer, index=False, sheet_name='נתוני דיגום', startrow=2)
        worksheet = writer.sheets['נתוני דיגום']
        worksheet.merge_cells('A1:B1')
        worksheet['A1'] = f"דוח נתוני דיגום: {chimney_name} | מזהם: {pollutant_name}"
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

st.title("📊 Stack Monitoring Pro")

with st.sidebar:
    st.markdown("## 📥 העלאת נתונים")
    uploaded_files = st.file_uploader("בחר קבצי אקסל", type=["xlsx"], accept_multiple_files=True, label_visibility="collapsed")
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

tab1, tab2, tab3 = st.tabs(["📈 דשבורד", "⚙️ הגדרות", "💾 ניהול"])

with tab1:
    if not st.session_state.db.empty:
        raw_ids = sorted(st.session_state.db['chimney_id'].unique())
        id_to_label = {cid: st.session_state.chimney_names.get(cid, f"ארובה {cid}") for cid in raw_ids}
        selected_label = st.selectbox("בחר ארובה:", options=list(id_to_label.values()))
        selected_id = [k for k, v in id_to_label.items() if v == selected_label][0]
        c_data = st.session_state.db[st.session_state.db['chimney_id'] == selected_id]
        
        for poll in sorted(c_data['pollutant'].unique()):
            poll_data = c_data[c_data['pollutant'] == poll].sort_values('date')
            threshold = int(st.session_state.thresholds.get(selected_id, {}).get(poll, 0))
            st.markdown(f'<div class="graph-card">', unsafe_allow_html=True)
            
            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=poll_data['date'].dt.strftime('%d/%m/%Y'), y=poll_data['concentration'],
                mode='lines+markers', line=dict(color='#3b82f6', width=4),
                marker=dict(size=10, color='#1e293b', line=dict(width=2, color='#3b82f6')),
                fill='tozeroy', fillcolor='rgba(59, 130, 246, 0.05)',
                name="ריכוז"
            ))
            if threshold > 0:
                fig.add_hline(y=threshold, line_dash="dash", line_color="#ef4444", 
                              annotation_text=f"סף: {threshold}", annotation_position="top left")
            
            fig.update_layout(
                title={'text': f"<b>{selected_label} - {poll}</b>", 'y': 0.9, 'x': 0.5, 'xanchor': 'center'},
                height=400, margin=dict(l=10, r=10, t=50, b=10), template="plotly_white",
                yaxis=dict(title="מ\"ג/מק\"ת")
            )
            st.plotly_chart(fig, use_container_width=True, config={'displaylogo': False})
            st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.info("אנא העלה קבצים מהתפריט הצדדי.")

with tab2:
    st.markdown("### הגדרות ארובות וספים")
    if not st.session_state.db.empty:
        for cid in sorted(st.session_state.db['chimney_id'].unique()):
            with st.expander(f"הגדרות {st.session_state.chimney_names.get(cid, cid)}"):
                name = st.text_input("שם ארובה", value=st.session_state.chimney_names.get(cid, ""), key=f"n_{cid}")
                if name != st.session_state.chimney_names.get(cid, ""):
                    st.session_state.chimney_names[cid] = name
                    save_data(st.session_state.db, st.session_state.thresholds, st.session_state.file_hashes, st.session_state.chimney_names)
                    st.rerun()
                for p in sorted(st.session_state.db[st.session_state.db['chimney_id'] == cid]['pollutant'].unique()):
                    val = st.number_input(f"סף ל-{p}", value=int(st.session_state.thresholds.get(cid, {}).get(p, 0)), step=1, key=f"l_{cid}_{p}")
                    if val != int(st.session_state.thresholds.get(cid, {}).get(p, 0)):
                        if cid not in st.session_state.thresholds: st.session_state.thresholds[cid] = {}
                        st.session_state.thresholds[cid][p] = int(val)
                        save_data(st.session_state.db, st.session_state.thresholds, st.session_state.file_hashes, st.session_state.chimney_names)

with tab3:
    if st.button("מחיקת כל הנתונים"):
        if os.path.exists(DB_FILE): os.remove(DB_FILE)
        st.session_state.clear()
        st.rerun()
8
