import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from streamlit_gsheets import GSheetsConnection

# --- הגדרות דף ---
st.set_page_config(page_title="Stack Monitoring Pro", layout="wide")

# --- חיבור למסד נתונים (Google Sheets) ---
# החיבור מתבצע דרך ה-Secrets שהגדרת ב-Dashboard
conn = st.connection("gsheets", type=GSheetsConnection)

def load_data_from_gs():
    try:
        # קריאת נתונים - ttl=0 מבטיח רענון נתונים בכל טעינה
        df = conn.read(ttl=0)
        return df
    except:
        # מחזיר טבלה ריקה עם העמודות הנדרשות במקרה של שגיאה או גיליון ריק
        return pd.DataFrame(columns=["pollutant", "concentration", "date", "chimney_id", "filename"])

def save_data_to_gs(df):
    try:
        conn.update(data=df)
        return True
    except Exception as e:
        st.error(f"שגיאה בשמירת הנתונים: {e}")
        return False

# --- פונקציית עיבוד קבצי אקסל ---
def process_excel(uploaded_file):
    try:
        df_raw = pd.read_excel(uploaded_file, header=None)
        
        # חילוץ מזהה ארובה (שורה 4, עמודה H)
        chimney_id = str(df_raw.iloc[3, 7]).strip()
        
        # חילוץ תאריך (שורה 8, עמודה N)
        sampling_date = pd.to_datetime(df_raw.iloc[7, 13], errors='coerce')
        if pd.isnull(sampling_date):
            return None
            
        data_rows = []
        # רכיבת על שורות הנתונים החל משורה 11
        for i in range(10, len(df_raw)):
            raw_pollutant = str(df_raw.iloc[i, 2]).strip()
            
            # --- התיקון לבאג הרווחים ---
            # join(split()) מסיר את כל הרווחים הכפולים והמיותרים בתוך הטקסט
            pollutant = " ".join(raw_pollutant.split())
            
            if not pollutant or pollutant.lower() in ["nan", "none", ""]:
                continue
            
            # חיפוש ערך הריכוז בעמודות הנפוצות (Q, N, P, O)
            found_val = None
            for col_idx in [16, 13, 15, 14]:
                val = df_raw.iloc[i, col_idx]
                if pd.notnull(val) and str(val).strip() != "":
                    try:
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
    except Exception as e:
        st.error(f"שגיאה בעיבוד הקובץ {uploaded_file.name}: {e}")
        return None

# --- ממשק משתמש (UI) ---
st.title("📊 מערכת ניטור ארובות - גרסה 2.0")

# טעינת הנתונים הקיימים מהענן
df_db = load_data_from_gs()

# תפריט צדדי להעלאת קבצים
with st.sidebar:
    st.header("📥 הוספת נתונים")
    uploaded_files = st.file_uploader("בחר קבצי אקסל (XLSX)", type=["xlsx"], accept_multiple_files=True)
    
    if uploaded_files and st.button("שמור נתונים חדשים לענן"):
        all_new_data = []
        for file in uploaded_files:
            processed = process_excel(file)
            if processed is not None:
                all_new_data.append(processed)
        
        if all_new_data:
            new_df = pd.concat(all_new_data, ignore_index=True)
            
            # שילוב עם הנתונים הקיימים והסרת כפילויות (לפי מזהם, תאריך וארובה)
            final_df = pd.concat([df_db, new_df]).drop_duplicates(
                subset=['pollutant', 'date', 'chimney_id'], 
                keep='last'
            )
            
            if save_data_to_gs(final_df):
                st.success(f"נוספו {len(new_df)} שורות נתונים בהצלחה!")
                st.rerun()

# תצוגת נתונים וגרפים
if not df_db.empty:
    # בחירת ארובה
    chimney_list = sorted(df_db['chimney_id'].unique())
    selected_chimney = st.selectbox("בחר ארובה לצפייה:", chimney_list)
    
    # סינון נתונים לארובה הנבחרת
    c_data = df_db[df_db['chimney_id'] == selected_chimney]
    
    # יצירת גרף לכל מזהם
    st.subheader(f"מגמות פליטה - ארובה {selected_chimney}")
    
    pollutants = sorted(c_data['pollutant'].unique())
    
    # חלוקה לעמודות להצגה נוחה
    cols = st.columns(2)
    for idx, poll in enumerate(pollutants):
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
            title=f"ריכוז {poll}",
            xaxis_title="תאריך דיגום",
            yaxis_title="מ\"ג/מק\"ת",
            hovermode="x unified",
            height=400
        )
        
        cols[idx % 2].plotly_chart(fig, use_container_width=True)

else:
    st.info("מסד הנתונים ריק. אנא העלה קבצי אקסל דרך התפריט הצדדי.")
