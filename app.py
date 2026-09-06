import streamlit as st
import pandas as pd
import os
from datetime import datetime

st.set_page_config(
    page_title="نظام الحضور والتقييم الذكي",
    page_icon="📚",
    layout="centered"
)

EXCEL_FILE = "سجل_الغياب_والحصص.xlsx"

def load_data():
    if os.path.exists(EXCEL_FILE):
        try:
            return pd.read_excel(EXCEL_FILE)
        except Exception:
            return pd.DataFrame(columns=["التاريخ", "الوقت", "اسم الطالب", "المرحلة الدراسية", "تقييم الحصة", "ملاحظات"])
    return pd.DataFrame(columns=["التاريخ", "الوقت", "اسم الطالب", "المرحلة الدراسية", "تقييم الحصة", "ملاحظات"])

def save_data(df):
    df.to_excel(EXCEL_FILE, index=False)

st.markdown(
    """
    <style>
    html, body, [class*="css"], .stMarkdown, p, span, label, div {
        direction: rtl;
        text-align: right;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }

    @media (prefers-color-scheme: light) {
        .stApp, p, span, label, h1, h2, h3, h4, h5, h6, .stMarkdown {
            color: #111111 !important;
            font-weight: 700 !important;
        }
        input, select, textarea {
            color: #000000 !important;
            background-color: #F8F9FA !important;
            font-weight: 600 !important;
        }
    }

    @media (prefers-color-scheme: dark) {
        .stApp, p, span, label, h1, h2, h3, h4, h5, h6, .stMarkdown {
            color: #FFFFFF !important;
            font-weight: 600 !important;
        }
        input, select, textarea {
            color: #FFFFFF !important;
            background-color: #262730 !important;
            font-weight: 500 !important;
        }
    }

    .stButton>button {
        width: 100%;
        border-radius: 8px;
        font-weight: 700;
        height: 3em;
    }
    </style>
    """,
    unsafe_allow_html=True
)

query_params = st.query_params
role_param = query_params.get("role", "teacher")

if role_param == "student":
    st.title("📝 استمارة تسجيل الحضور والتقييم")
    st.write("يرجى إدخال بياناتك وتقييم الحصة بعناية:")

    with st.form("student_form", clear_on_submit=True):
        student_name = st.text_input("اسم الطالب بالكامل:")
        grade_level = st.selectbox(
            "المرحلة الدراسية:",
            [
                "الصف الأول الإعدادي",
                "الصف الثاني الإعدادي",
                "الصف الثالث الإعدادي",
                "الصف الأول الثانوي",
                "الصف الثاني الثانوي",
                "الصف الثالث الثانوي"
            ]
        )
        rating = st.select_slider(
            "تقييمك لشرح واستفادة الحصة:",
            options=["⭐", "⭐⭐", "⭐⭐⭐", "⭐⭐⭐⭐", "⭐⭐⭐⭐⭐"],
            value="⭐⭐⭐⭐⭐"
        )
        notes = st.text_area("هل لديك استفسار أو اقتراح للحصة القادمة؟ (اختياري)")
        
        submitted = st.form_submit_button("إرسال وتأكيد الحضور")

        if submitted:
            if not student_name.strip():
                st.error("يرجى كتابة الاسم لتأكيد الحضور.")
            else:
                now = datetime.now()
                new_entry = {
                    "التاريخ": now.strftime("%Y-%m-%d"),
                    "الوقت": now.strftime("%H:%M:%S"),
                    "اسم الطالب": student_name.strip(),
                    "المرحلة الدراسية": grade_level,
                    "تقييم الحصة": rating,
                    "ملاحظات": notes.strip()
                }

                df = load_data()
                df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
                save_data(df)
                
                st.success(f"تم تسجيل حضورك بنجاح يا {student_name}! شكراً لتقييمك.")

else:
    st.title("📊 لوحة تحكم المعلم والإدارة")
    
    password = st.text_input("أدخل كلمة المرور لعرض البيانات:", type="password")
    
    if password == "1234":
        st.success("تم تأكيد الهوية بنجاح.")
        
        df = load_data()
        
        col1, col2 = st.columns(2)
        with col1:
            st.metric("إجمالي تسجيلات الحضور", len(df))
        with col2:
            if not df.empty and "تقييم الحصة" in df.columns:
                star_count = df["تقييم الحصة"].str.count("⭐").mean()
                st.metric("متوسط التقييم العام", f"{star_count:.1f} / 5")
        
        st.write("### سجل الحضور الشامل:")
        st.dataframe(df, use_container_width=True)
        
        if not df.empty:
            @st.cache_data
            def convert_df(df_to_export):
                return df_to_export.to_csv(index=False).encode('utf-8-sig')
            
            csv = convert_df(df)
            st.download_button(
                label="📥 تحميل السجل كملف Excel/CSV",
                data=csv,
                file_name=f"سجل_الطلاب_{datetime.now().strftime('%Y_%m_%d')}.csv",
                mime="text/csv"
            )
    elif password != "":
        st.error("كلمة المرور غير صحيحة.")
