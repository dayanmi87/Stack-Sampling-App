import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from streamlit_gsheets import GSheetsConnection

st.set_page_config(page_title="Stack Monitoring Pro - Debug", layout="wide")

st.title("🕵️‍♂️ מצב איתור תקלות")

# --- 1. בדיקת חיבור לגוגל ---
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
    df_db = conn.read(ttl=0)
    st.success("✅ שלב 1: החיבור לגוגל שיטס תקין!")
except Exception as e:
    st.error(f"❌ שלב 1 נכשל: בעיה בחיבור לגוגל. שגיאה: {e}")
    df_db = pd.DataFrame(columns=["pollutant", "concentration", "date", "chimney_id", "filename"])

# --- פונקציית עיבוד עם דיווח על המסך ---
def process_excel(uploaded_file):
    try:
        st.info(f"מתחיל לסרוק את: {uploaded_file.name}")
        all_sheets = pd.read_excel(uploaded_file, sheet_name=None, header=None)
        
        sheet_name = "דיווח תוצאות דיגום ארובה"
        if sheet_name in all_sheets:
            target_df = all_sheets[sheet_name]
            st.write("✅ מצאתי את הלשונית הנכונה באקסל.")
        else:
            target_df = list(all_sheets.values())[0]
            st.warning("⚠️ לא מצאתי את לשונית הדיווח הרשמית, מנסה לקרוא את הגיליון הראשון.")

        # חילוץ מזהה ותאריך
        chimney_id = str(target_df.iloc[3, 7]).strip()
        sampling_date = pd.to_datetime(target_df.iloc[7, 13], errors='coerce')
        st.write(f"🔍 זיהיתי ארובה: {chimney_id} | תאריך: {sampling_date}")
        
        data_rows = []
        for i in range(10, len(target_df)):
            raw_pollutant = str(target_df.iloc[i, 2]).strip()
            pollutant = " ".join(raw_pollutant.split()) 
            
            if not pollutant or pollutant.lower() in ["nan", "none", "", "מזהם"]:
                continue
            
            found_val = None
            for col_idx in [16, 13, 15, 14]:
                try:
                    val = target_df.iloc[i, col_idx]
                    if pd.notnull(val) and str(val).strip() != "":
                        found_val = round(float(val), 2)
                        break
                except: continue
            
            if found_val is not None:
                data_rows.append({
                    "pollutant": pollutant, 
                    "concentration": found_val, 
                    "date": sampling_date.strftime('%Y-%m-%d') if pd.notnull(sampling_date) else None,
                    "chimney_id": chimney_id, 
                    "filename": uploaded_file.name
                })
        
        st.success(f"✅ שלב 2: חילצתי {len(data_rows)} מזהמים מהקובץ!")
        return pd.DataFrame(data_rows)
    except Exception as e:
        st.error(f"❌ שלב 2 נכשל: התרסקה קריאת האקסל. שגיאה: {e}")
        return None

# --- ממשק ---
uploaded_files = st.file_uploader("העלה קובץ לבדיקה", type=["xlsx"], accept_multiple_files=True)

if uploaded_files and st.button("בדוק שמירה"):
    new_dfs = []
    for f in uploaded_files:
        res = process_excel(f)
        if res is not None and not res.empty:
            new_dfs.append(res)
            st.write("הנתונים שחולצו:")
            st.dataframe(res) # מראה לך את הטבלה בעיניים לפני השמירה
    
    if new_dfs:
        combined_new = pd.concat(new_dfs, ignore_index=True)
        try:
            final_df = pd.concat([df_db, combined_new]).drop_duplicates(
                subset=['pollutant', 'date', 'chimney_id'], keep='last'
            )
            conn.update(data=final_df)
            st.success("🎯 שלב 3: שמירת הנתונים לגוגל הצליחה! לך לבדוק בגיליון.")
        except Exception as e:
            st.error(f"❌ שלב 3 נכשל: לא הצלחתי לכתוב לגוגל שיטס. שגיאה: {e}")
