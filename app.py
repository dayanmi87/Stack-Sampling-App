import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from streamlit_gsheets import GSheetsConnection

# --- הגדרות דף ---
st.set_page_config(page_title="Stack Monitoring Pro", layout="wide")

# --- חיבור למסד נתונים (Google Sheets) ---
conn = st.connection("gsheets", type=GSheetsConnection)

def load_data():
    try:
        # קריאת נתונים - מתרענן מיד
        return conn.read(ttl=0)
    except Exception as e:
        st.error(f"שגיאה בקריאת הנתונים מהענן: {e}")
        return pd.DataFrame(columns=["pollutant", "concentration", "date", "chimney_id", "filename"])

# --- פונקציית עיבוד קבצי אקסל ---
def process_excel(uploaded_file):
    try:
        # טעינת כל הגיליונות מהקובץ
        all_sheets = pd.read_excel(uploaded_file, sheet_name=None, header=None)
        
        # ניסיון למצוא את הגיליון הנכון
        target_df = None
        sheet_name = "דיווח תוצאות דיגום ארובה"
        
        if sheet_name in all_sheets:
            target_df = all_sheets[sheet_name]
        else:
            # לקיחת הגיליון הראשון כברירת מחדל אם השם שונה
            target_df = list(all_sheets.values())[0]

        # חילוץ מזהה ארובה ותאריך
        chimney_id = str(target_df.iloc[3, 7]).strip()
        sampling_date = pd.to_datetime(target_df.iloc[7, 13], errors='coerce')
        
        if pd.isnull(sampling_date):
            return None

        data_rows = []
        for i in range(10, len(target_df)):
            raw_pollutant = str(target_df.iloc[i, 2]).strip()
            
            # התיקון שמאחד את המזהמים (מסיר רווחים כפולים)
            pollutant = " ".join(raw_pollutant.split()) 
            
            if not pollutant or pollutant.lower() in ["nan", "none", "", "מזהם"]:
                continue
            
            # חיפוש ריכוז בעמודות (Q, N, P, O)
            found_val = None
            for col_idx in [16, 13, 15, 14]:
                try:
                    val = target_df.iloc[i, col_idx]
                    if pd.notnull(val) and str(val).strip() != "":
                        found_val = round(float(val), 2)
                        break
                except: 
                    continue
            
            if found_val is not None:
                data_rows.append({
                    "pollutant": pollutant, 
                    "concentration": found_val, 
                    "date": sampling_date.strftime('%Y-%m-%d'), 
                    "chimney_id": chimney_id, 
                    "filename": uploaded_file.name
                })
        
        return pd.DataFrame(data_rows)
    except Exception:
        return None

# --- ממשק המשתמש (UI) ---
st.title("📊 מערכת ניטור ארובות")

# טעינת נתונים קיימים
df_db = load_data()

# תפריט צדדי להעלאת קבצים
with st.sidebar:
    st.header("📥 העלאת נתונים")
    files = st.file_uploader("בחר קבצי אקסל", type=["xlsx"], accept_multiple_files=True)
    
    if files and st.button("שמור ועדכן גרפים"):
        new_dfs = []
        for f in files:
            res = process_excel(f)
            if res is not None and not res.empty:
                new_dfs.append(res)
        
        if new_dfs:
            combined_new = pd.concat(new_dfs, ignore_index=True)
            # הוספת הנתונים החדשים למסד הקיים והסרת כפילויות
            final_df = pd.concat([df_db, combined_new]).drop_duplicates(
                subset=['pollutant', 'date', 'chimney_id'], keep='last'
            )
            # עדכון הגיליון בגוגל
            conn.update(data=final_df)
            st.success("הנתונים נשמרו בענן בהצלחה!")
            st.rerun()
        else:
            st.warning("לא נמצאו נתונים תקינים בקבצים שהועלו.")

# --- תצוגת הגרפים ---
if not df_db.empty:
    # ניקוי שורות ריקות למקרה שנשמרו בטעות
    df_db = df_db.dropna(subset=['pollutant', 'chimney_id'])
    
    # תיבת בחירה לארובה
    chimney_list = sorted(df_db['chimney_id'].unique())
    chimney = st.selectbox("בחר ארובה:", chimney_list)
    
    # סינון הנתונים לארובה שנבחרה
    c_data = df_db[df_db['chimney_id'] == chimney]
    pollutants = sorted(c_data['pollutant'].unique())
    
    # סידור הגרפים בשתי עמודות
    cols = st.columns(2)
    
    for i, poll in enumerate(pollutants):
        p_data = c_data[c_data['pollutant'] == poll].sort_values('date')
        
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=p_data['date'], 
            y=p_data['concentration'], 
            mode='lines+markers', 
            name=poll,
            line=dict(width=3),
            marker=dict(size=10)
        ))
        
        fig.update_layout(
            title=f"{poll} - ארובה {chimney}", 
            yaxis_title='מ"ג/מק"ת',
            hovermode="x unified",
            height=400
        )
        cols[i % 2].plotly_chart(fig, use_container_width=True)
else:
    st.info("אין נתונים להצגה. העלה קבצים בתפריט הצד.")
