import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import re

st.set_page_config(page_title="Stack Monitoring - Fixed Version", layout="wide")

# --- אתחול בסיס נתונים מקומי ---
if 'local_db' not in st.session_state:
    st.session_state.local_db = pd.DataFrame(columns=["pollutant", "concentration", "date", "chimney_id", "filename"])

# --- פונקציית ניקוי טקסט אגרסיבית ---
def clean_pollutant_name(name):
    if pd.isna(name):
        return ""
    # המרה לטקסט, הסרת רווחים מהקצוות
    name = str(name).strip()
    # החלפת כל סוגי הרווחים (כולל רווחים נסתרים של אקסל \xa0) ברווח רגיל
    name = name.replace('\xa0', ' ')
    # הפיכת רצף רווחים (כפול, משולש) לרווח יחיד
    name = " ".join(name.split())
    return name

# --- פונקציית עיבוד משופרת ---
def process_excel(uploaded_file):
    try:
        # קריאת כל הגיליונות
        all_sheets = pd.read_excel(uploaded_file, sheet_name=None, header=None)
        target_df = None
        
        # חיפוש הלשונית הנכונה לפי ביטוי ייחודי
        for name, df in all_sheets.items():
            if df.astype(str).apply(lambda x: x.str.contains('שעת תחילת הדיגום')).any().any():
                target_df = df
                break
        
        if target_df is None:
            st.error(f"לא נמצאה לשונית דיווח תקינה בקובץ {uploaded_file.name}")
            return None

        # חילוץ מזהה ארובה (H4) ותאריך (N8)
        chimney_id = str(target_df.iloc[3, 7]).strip()
        raw_date = target_df.iloc[7, 13]
        sampling_date = pd.to_datetime(raw_date, dayfirst=True, errors='coerce')
        
        if pd.isnull(sampling_date):
            st.warning(f"לא נמצא תאריך תקין בקובץ {uploaded_file.name}")
            return None

        data_rows = []
        # סריקת המזהמים מהשורה ה-11 והלאה
        for i in range(10, len(target_df)):
            raw_name = target_df.iloc[i, 2]
            pollutant = clean_pollutant_name(raw_name)
            
            # דילוג על שורות ריקות או כותרות
            if not pollutant or pollutant in ["מזהם", "המשך", "None", "nan"]:
                continue
            
            # חילוץ ריכוז מנורמל (עמודה Q - אינדקס 16)
            try:
                val = target_df.iloc[i, 16]
                if pd.notnull(val) and str(val).strip() != "":
                    data_rows.append({
                        "pollutant": pollutant, 
                        "concentration": round(float(val), 2), 
                        "date": sampling_date.strftime('%Y-%m-%d'), 
                        "chimney_id": chimney_id, 
                        "filename": uploaded_file.name
                    })
            except: continue
        
        return pd.DataFrame(data_rows)
    except Exception as e:
        st.error(f"שגיאה בעיבוד הקובץ {uploaded_file.name}: {e}")
        return None

# --- ממשק משתמש ---
st.title("📊 ניטור ארובות - גרסה מתוקנת")

with st.sidebar:
    st.header("📥 העלאת קבצים")
    files = st.file_uploader("העלה אקסלים לבדיקה", type=["xlsx"], accept_multiple_files=True)
    
    if files and st.button("עבד נתונים והצג גרפים"):
        new_data_list = []
        for f in files:
            res = process_excel(f)
            if res is not None and not res.empty:
                new_data_list.append(res)
        
        if new_data_list:
            combined_new = pd.concat(new_data_list, ignore_index=True)
            # איחוד והסרת כפילויות
            st.session_state.local_db = pd.concat([st.session_state.local_db, combined_new]).drop_duplicates(
                subset=['pollutant', 'date', 'chimney_id'], keep='last'
            )
            st.success(f"חולצו בהצלחה {len(combined_new)} שורות!")

    st.divider()
    if st.button("🗑️ נקה הכל (מומלץ אם יש גרפים כפולים)"):
        st.session_state.local_db = pd.DataFrame(columns=["pollutant", "concentration", "date", "chimney_id", "filename"])
        st.rerun()

# --- תצוגת נתונים וגרפים ---
db = st.session_state.local_db

if not db.empty:
    # בחירת ארובה
    chimney_list = sorted(db['chimney_id'].unique())
    selected_chimney = st.selectbox("בחר ארובה:", chimney_list)
    
    c_data = db[db['chimney_id'] == selected_chimney]
    
    # הצגת רשימת המזהמים שנמצאו (לביקורת)
    pollutants = sorted(c_data['pollutant'].unique())
    
    st.write(f"נמצאו {len(pollutants)} מזהמים ייחודיים לארובה {selected_chimney}")
    
    cols = st.columns(2)
    for i, poll in enumerate(pollutants):
        # סינון נתונים למזהם הספציפי ומיון לפי תאריך
        p_data = c_data[c_data['pollutant'] == poll].sort_values('date')
        
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=p_data['date'], 
            y=p_data['concentration'], 
            mode='lines+markers',
            name=poll,
            line=dict(width=3),
            marker=dict(size=8)
        ))
        
        fig.update_layout(
            title=f"מגמת {poll}",
            xaxis_title="תאריך",
            yaxis_title='מ"ג/מק"ת',
            height=400,
            template="plotly_white"
        )
        cols[i % 2].plotly_chart(fig, use_container_width=True)
    
    with st.expander("צפה בטבלה הגולמית"):
        st.dataframe(db)
else:
    st.info("העלה קבצי אקסל ולחץ על 'עבד נתונים' כדי להתחיל.")
