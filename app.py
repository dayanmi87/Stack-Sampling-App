import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from streamlit_gsheets import GSheetsConnection

st.set_page_config(page_title="Stack Monitoring Pro", layout="wide")

# הגדרת עמודות ברירת מחדל למסד הנתונים
DEFAULT_COLS = ["pollutant", "concentration", "date", "chimney_id", "filename"]

# --- חיבור למסד נתונים (עם הגנה על גיליון ריק) ---
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
    try:
        df_db = conn.read(ttl=0)
        # אם הגיליון קיים אבל ריק מעמודות
        if df_db.empty and len(df_db.columns) < 2:
            df_db = pd.DataFrame(columns=DEFAULT_COLS)
    except:
        # אם הגיליון ריק לגמרי (שגיאת EmptyDataError)
        df_db = pd.DataFrame(columns=DEFAULT_COLS)
except Exception as e:
    st.error(f"שגיאה בחיבור לגוגל שיטס: {e}")
    df_db = pd.DataFrame(columns=DEFAULT_COLS)

# --- פונקציית עיבוד משופרת ---
def process_excel(uploaded_file):
    try:
        all_sheets = pd.read_excel(uploaded_file, sheet_name=None, header=None)
        target_df = None
        
        # חיפוש הלשונית לפי ביטוי ייחודי לדיווח: "שעת תחילת הדיגום"
        for name, df in all_sheets.items():
            if df.astype(str).apply(lambda x: x.str.contains('שעת תחילת הדיגום')).any().any():
                target_df = df
                break
        
        if target_df is None:
            st.error(f"לא נמצאה לשונית דיווח תקינה בקובץ {uploaded_file.name}")
            return None

        # חילוץ מזהה ארובה (H4) ותאריך (N8)
        try:
            chimney_id = str(target_df.iloc[3, 7]).strip()
            raw_date = target_df.iloc[7, 13]
            sampling_date = pd.to_datetime(raw_date, errors='coerce')
        except:
            st.error("לא הצלחתי לחלץ תאריך או מזהה ארובה מהמיקומים הסטנדרטיים.")
            return None
        
        data_rows = []
        # סריקה משורה 11 (אינדקס 10) והלאה
        for i in range(10, len(target_df)):
            pollutant = str(target_df.iloc[i, 2]).strip()
            # ניקוי רווחים כפולים (למשל עבור תחמוצות גופרית)
            pollutant = " ".join(pollutant.split())
            
            if not pollutant or pollutant.lower() in ["nan", "none", "", "מזהם", "המשך"]:
                continue
            
            # חיפוש הריכוז (עמודה Q - אינדקס 16)
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
        st.error(f"שגיאה בעיבוד הקובץ: {e}")
        return None

# --- ממשק משתמש ---
st.title("📊 מערכת ניטור ארובות")

with st.sidebar:
    st.header("📥 העלאת נתונים")
    files = st.file_uploader("העלה אקסלים של דיגום", type=["xlsx"], accept_multiple_files=True)
    
    if files and st.button("שמור ועדכן"):
        new_dfs = []
        for f in files:
            res = process_excel(f)
            if res is not None and not res.empty:
                new_dfs.append(res)
        
        if new_dfs:
            combined_new = pd.concat(new_dfs, ignore_index=True)
            # מיזוג ושמירה
            final_df = pd.concat([df_db, combined_new]).drop_duplicates(
                subset=['pollutant', 'date', 'chimney_id'], keep='last'
            )
            conn.update(data=final_df)
            st.success(f"נשמרו {len(combined_new)} שורות חדשות!")
            st.rerun()
        else:
            st.warning("לא נמצאו נתונים חדשים לחילוץ.")

# --- הצגת גרפים ---
if not df_db.empty and len(df_db) > 0:
    # ניקוי נתונים ריקים בבסיס הנתונים
    df_plot = df_db.dropna(subset=['pollutant', 'date'])
    
    chimney = st.selectbox("בחר ארובה:", sorted(df_plot['chimney_id'].unique()))
    c_data = df_plot[df_plot['chimney_id'] == chimney]
    
    pollutants = sorted(c_data['pollutant'].unique())
    cols = st.columns(2)
    
    for i, poll in enumerate(pollutants):
        p_data = c_data[c_data['pollutant'] == poll].sort_values('date')
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=p_data['date'], y=p_data['concentration'], mode='lines+markers', name=poll))
        fig.update_layout(title=f"ריכוז {poll}", yaxis_title='מ"ג/מק"ת', height=350)
        cols[i % 2].plotly_chart(fig, use_container_width=True)
else:
    st.info("בסיס הנתונים ריק. העלה קובץ כדי להתחיל.")
