import base64
import io
import os
from datetime import date, datetime
import pandas as pd
from PIL import Image
import streamlit as st

FILE_NAME = "سجل_الغياب_والحصص.xlsx"
IMG_NAME = "teacher.jpg"

st.set_page_config(
    page_title="البشمهندس X الرياضة | متابعة الطلاب",
    page_icon="📐",
    layout="wide",
)

CURRICULUM_DATA = {
    "المنهج المصري 🇪🇬": [
        "الصف الأول الإعدادي",
        "الصف الثاني الإعدادي",
        "الصف الثالث الإعدادي (الشهادة الإعدادية)",
        "الصف الأول الثانوي",
        "الصف الثاني الثانوي (علمي)",
        "الصف الثاني الثانوي (أدبي)",
        "الصف الثالث الثانوي (علمي رياضة)",
        "الصف الثالث الثانوي (علمي علوم)",
        "الصف الثالث الثانوي (أدبي)",
        "المرحلة الابتدائية",
    ],
    "المنهج القطري 🇶🇦": [
        "الصف السابع (إعدادي)",
        "الصف الثامن (إعدادي)",
        "الصف التاسع (إعدادي)",
        "الصف العاشر (المشترك)",
        "الصف الحادي عشر (المسار العلمي)",
        "الصف الحادي عشر (مسار الآداب والإنسانيات)",
        "الصف الحادي عشر (المسار التكنولوجي)",
        "الصف الثاني عشر (المسار العلمي - متقدم)",
        "الصف الثاني عشر (مسار الآداب - تأسيسي)",
        "الصف الثاني عشر (المسار التكنولوجي)",
        "المرحلة الابتدائية",
    ],
    "المنهج الإماراتي 🇦🇪": [
        "الحلقة الثانية (الصفوف 5 - 8)",
        "الصف التاسع (مسار عام)",
        "الصف التاسع (مسار متقدم)",
        "الصف العاشر (مسار عام)",
        "الصف العاشر (مسار متقدم)",
        "الصف العاشر (مسار النخبة)",
        "الصف الحادي عشر (مسار عام)",
        "الصف الحادي عشر (مسار متقدم)",
        "الصف الحادي عشر (مسار النخبة)",
        "الصف الثاني عشر (مسار عام)",
        "الصف الثاني عشر (مسار متقدم)",
        "الصف الثاني عشر (مسار النخبة)",
    ],
    "المنهج السعودي 🇸🇦": [
        "المرحلة المتوسطة (أول / ثاني / ثالث متوسط)",
        "السنة الأولى المشتركة (أول ثانوي)",
        "السنة الثانية (المسار العام)",
        "السنة الثانية (مسار علوم الحاسب والهندسة)",
        "السنة الثانية (مسار الصحة والحياة)",
        "السنة الثانية (مسار إدارة الأعمال / الشرعي)",
        "السنة الثالثة (المسار العام)",
        "السنة الثالثة (مسار علوم الحاسب والهندسة)",
        "السنة الثالثة (مسار الصحة والحياة)",
    ],
    "المنهج الكويتي 🇰🇼": [
        "المرحلة المتوسطة (الصفوف 6 - 9)",
        "الصف العاشر الثانوي (مشترك)",
        "الصف الحادي عشر (القسم العلمي)",
        "الصف الحادي عشر (القسم الأدبي)",
        "الصف الثاني عشر (القسم العلمي)",
        "الصف الثاني عشر (القسم الأدبي)",
    ],
    "المنهج السوداني 🇸🇩": [
        "المرحلة المتوسطة (أولى / ثانية / ثالثة متوسط)",
        "الصف الأول الثانوي",
        "الصف الثاني الثانوي (علمي)",
        "الصف الثاني الثانوي (أدبي)",
        "الصف الثالث الثانوي (علمي رياضيات - الشهادة السودانية)",
        "الصف الثالث الثانوي (علمي أحياء)",
        "الصف الثالث الثانوي (أدبي)",
    ],
}

possible_images = [
    "teacher.jpg",
    "teacher.png",
    "teacher.jpeg",
    "photo_2026-08-02_00-34-53.jpg",
]
found_img_path = None
for img_cand in possible_images:
    if os.path.exists(img_cand):
        found_img_path = img_cand
        break


def get_image_base64(path):
    if path and os.path.exists(path):
        try:
            with open(path, "rb") as img_file:
                return base64.b64encode(img_file.read()).decode()
        except Exception:
            return ""
    return ""


img_b64 = get_image_base64(found_img_path)

# ==================== التنسيق المتجاوب مع الوضع الفاتح والداكن ====================
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@600;700;800;900&display=swap');

    html, body, [class*="css"], p, span, label, div {
        font-family: 'Cairo', sans-serif !important;
        direction: rtl;
        text-align: right;
    }

    .brand-banner {
        background: linear-gradient(135deg, #0b1f3a, #1e3c72) !important;
        padding: 24px 30px;
        border-radius: 14px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 25px;
        box-shadow: 0 6px 18px rgba(11, 31, 58, 0.35);
    }
    
    .brand-title {
        font-size: 34px !important;
        font-weight: 900 !important;
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
        margin: 0 !important;
        text-shadow: 0 2px 6px rgba(0, 0, 0, 0.5);
    }

    .brand-subtitle {
        color: #e0edff !important;
        -webkit-text-fill-color: #e0edff !important;
        font-size: 16px !important;
        font-weight: 700 !important;
        margin-top: 6px !important;
    }

    .stButton>button {
        background: #0052cc !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 8px;
        font-weight: 900 !important;
        font-size: 16px !important;
        padding: 10px 24px;
        box-shadow: 0 4px 10px rgba(0, 82, 204, 0.25);
    }
    .stButton>button:hover {
        background: #003d99 !important;
    }

    /* صندوق التواصل وحفظ الحقوق */
    .social-footer-box {
        margin-top: 35px;
        padding: 20px 0;
        border-top: 1px solid rgba(150, 150, 150, 0.3);
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        width: 100%;
    }
    .social-icons-container {
        display: flex;
        gap: 15px;
        justify-content: center;
        align-items: center;
        margin-bottom: 12px;
    }
    .social-btn {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 44px;
        height: 44px;
        border-radius: 50%;
        text-decoration: none !important;
        box-shadow: 0 4px 10px rgba(0,0,0,0.2);
        transition: transform 0.2s ease;
    }
    .social-btn:hover {
        transform: scale(1.12);
    }
    .social-btn svg {
        width: 24px;
        height: 24px;
    }
    .facebook-bg { background-color: #1877F2; }
    .whatsapp-bg { background-color: #25D366; }
    .telegram-bg { background-color: #229ED9; }
    .tiktok-bg   { background-color: #000000; border: 1px solid #444; }
    .youtube-bg  { background-color: #FF0000; }

    .rights-text {
        font-size: 15px;
        font-weight: 900;
        margin-top: 10px;
        text-align: center !important;
        direction: rtl;
    }

    @media (prefers-color-scheme: light) {
        body, .stApp {
            background-color: #ffffff !important;
        }
        p, span, label, h1, h2, h3, h4, h5, h6, .stMarkdown {
            color: #111111 !important;
            font-weight: 800 !important;
        }
        div[data-testid="stMetric"] {
            background: #f8fafc !important;
            border: 2px solid #0052cc !important;
            border-radius: 12px;
            padding: 14px 18px;
        }
        div[data-testid="stMetric"] label {
            color: #1e293b !important;
            font-size: 15px !important;
            font-weight: 800 !important;
        }
        div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
            color: #0052cc !important;
            font-weight: 900 !important;
        }
        input, select, textarea {
            color: #000000 !important;
            background-color: #ffffff !important;
            border: 2px solid #94a3b8 !important;
            font-weight: 700 !important;
        }
        .rights-text {
            color: #1e293b !important;
        }
    }

    @media (prefers-color-scheme: dark) {
        body, .stApp {
            background-color: #0e1117 !important;
        }
        p, span, label, h1, h2, h3, h4, h5, h6, .stMarkdown {
            color: #ffffff !important;
            font-weight: 700 !important;
        }
        div[data-testid="stMetric"] {
            background: #1e232d !important;
            border: 2px solid #3b82f6 !important;
            border-radius: 12px;
            padding: 14px 18px;
        }
        div[data-testid="stMetric"] label {
            color: #e2e8f0 !important;
            font-size: 15px !important;
            font-weight: 800 !important;
        }
        div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
            color: #60a5fa !important;
            font-weight: 900 !important;
        }
        input, select, textarea {
            color: #ffffff !important;
            background-color: #262730 !important;
            border: 2px solid #475569 !important;
            font-weight: 700 !important;
        }
        .rights-text {
            color: #e2e8f0 !important;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

COLUMNS = [
    "التاريخ",
    "اسم الطالب",
    "المنهج/الدولة",
    "المجموعة/الصف",
    "الحالة",
    "سعر الحصة",
    "عدد الحصص الكلي",
    "نظام الدفع",
    "مستوى الطالب",
    "ملاحظات",
]


def load_data():
    if os.path.exists(FILE_NAME):
        try:
            df = pd.read_excel(FILE_NAME)
            for col in COLUMNS:
                if col not in df.columns:
                    df[col] = ""
            return df
        except Exception:
            return pd.DataFrame(columns=COLUMNS)
    return pd.DataFrame(columns=COLUMNS)


def save_data(df):
    df.to_excel(FILE_NAME, index=False)


if "data" not in st.session_state:
    st.session_state.data = load_data()

query_params = st.query_params
is_student_mode = query_params.get("role") == "student"

# ==============================================================================
# 1. واجهة الطالب
# ==============================================================================
if is_student_mode:
    st.markdown(
        """
        <style>
            [data-testid="stSidebar"] { display: none !important; }
            [data-testid="stSidebarCollapseButton"] { display: none !important; }
            [data-testid="collapsedControl"] { display: none !important; }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        f"""
    <div class="brand-banner" dir="rtl">
        <div>
            <h1 class="brand-title">البشمهندس X الرياضة 📐</h1>
            <p class="brand-subtitle">استمارة تسجيل حضور الطالب وتقييم الحصة</p>
        </div>
        {'<img src="data:image/jpeg;base64,' + img_b64 + '" style="width: 90px; height: 90px; border-radius: 50%; border: 3px solid #ffffff; object-fit: cover;">' if img_b64 else ''}
    </div>
    """,
        unsafe_allow_html=True,
    )

    st.subheader("سجل بيانات حضورك للحصة:")

    selected_curriculum = st.selectbox(
        "اختر المنهج الدراسي / الدولة:", list(CURRICULUM_DATA.keys())
    )
    available_grades = CURRICULUM_DATA[selected_curriculum]

    with st.form("student_self_registration", clear_on_submit=True):
        st_name = st.text_input(
            "اسم الطالب بالكامل (ثلاثي أو رباعي):", placeholder="مثال: أحمد محمد"
        )
        st_grade = st.selectbox("المرحلة / الصف الدراسي:", available_grades)
        st_date = st.date_input("تاريخ الحصة:", value=date.today())

        rating_options = [
            "⭐⭐⭐⭐⭐ (5/5) ممتاز جداً وفهم ممتاز",
            "⭐⭐⭐⭐ (4/5) جيد جداً",
            "⭐⭐⭐ (3/5) جيد ومفهوم",
            "⭐⭐ (2/5) متوسط ويحتاج توضيح",
            "⭐ (1/5) يحتاج إعادة شرح",
        ]
        selected_rating = st.selectbox(
            "⭐ قيّم الحصة مع البشمهندس (من 5 نجوم):",
            options=rating_options,
            index=0,
        )

        submit_btn = st.form_submit_button("✅ إرسال وتأكيد الحضور")

        if submit_btn:
            if not st_name.strip():
                st.error("يرجى كتابة اسمك بالكامل.")
            else:
                new_row = {
                    "التاريخ": str(st_date),
                    "اسم الطالب": st_name.strip(),
                    "المنهج/الدولة": selected_curriculum,
                    "المجموعة/الصف": st_grade,
                    "الحالة": "حاضر",
                    "سعر الحصة": 0.0,
                    "عدد الحصص الكلي": 1,
                    "نظام الدفع": "مؤجل",
                    "مستوى الطالب": "قيد التقييم",
                    "ملاحظات": f"تقييم الحصة: {selected_rating}",
                }
                st.session_state.data = pd.concat(
                    [st.session_state.data, pd.DataFrame([new_row])],
                    ignore_index=True,
                )
                save_data(st.session_state.data)
                st.success(
                    f"🎉 شكراً لك {st_name}! تم تسجيل حضورك وتقييمك بنجاح."
                )

    # شريط التواصل وحفظ الحقوق في واجهة الطالب
    st.markdown(
        """
        <div class="social-footer-box">
            <div class="social-icons-container">
                <!-- فيسبوك -->
                <a href="https://www.facebook.com/share/19fD41rV3H/" target="_blank" title="Facebook" class="social-btn facebook-bg">
                    <svg fill="#ffffff" viewBox="0 0 24 24"><path d="M24 12.073c0-6.627-5.373-12-12-12s-12 5.373-12 12c0 5.99 4.388 10.954 10.125 11.854v-8.385H7.078v-3.47h3.047V9.43c0-3.007 1.792-4.669 4.533-4.669 1.312 0 2.686.235 2.686.235v2.953H15.83c-1.491 0-1.956.925-1.956 1.874v2.25h3.328l-.532 3.47h-2.796v8.385C19.612 23.027 24 18.062 24 12.073z"/></svg>
                </a>
                <!-- واتساب -->
                <a href="https://wa.me/qr/YFB55GOW3WKZN1" target="_blank" title="WhatsApp" class="social-btn whatsapp-bg">
                    <svg fill="#ffffff" viewBox="0 0 24 24"><path d="M12.031 6.172c-3.181 0-5.767 2.586-5.768 5.766-.001 1.298.38 2.27 1.019 3.287l-.711 2.599 2.669-.699c.971.53 1.77.822 2.791.823h.002c3.18 0 5.767-2.586 5.768-5.766 0-3.18-2.587-5.766-5.77-5.766zm9.969 5.828c0 5.519-4.481 10-10 10-1.761 0-3.424-.46-4.881-1.267l-5.619 1.474 1.499-5.485c-.911-1.516-1.43-3.285-1.43-5.176 0-5.519 4.481-10 10-10 5.519 0 10 4.481 10 10z"/></svg>
                </a>
                <!-- تليجرام -->
                <a href="https://t.me/mrmaths22" target="_blank" title="Telegram" class="social-btn telegram-bg">
                    <svg fill="#ffffff" viewBox="0 0 24 24"><path d="M12 0c-6.627 0-12 5.373-12 12s5.373 12 12 12 12-5.373 12-12-5.373-12-12-12zm5.894 8.221l-1.97 9.28c-.145.658-.537.818-1.084.508l-3-2.21-1.446 1.394c-.16.16-.295.295-.605.295l.213-3.053 5.56-5.023c.242-.213-.054-.333-.373-.121l-6.871 4.326-2.962-.924c-.643-.204-.657-.643.136-.953l11.57-4.461c.537-.194 1.006.131.832.942z"/></svg>
                </a>
                <!-- تيك توك -->
                <a href="https://www.tiktok.com/@eng_mohamedghonaim?_r=1&_t=ZS-99VdklZPBUS" target="_blank" title="TikTok" class="social-btn tiktok-bg">
                    <svg fill="#ffffff" viewBox="0 0 24 24"><path d="M19.589 6.686a4.793 4.793 0 0 1-3.77-4.245V2h-3.445v13.672a2.896 2.896 0 0 1-5.201 1.743l-.068-.102a2.895 2.895 0 0 1 2.373-4.513c.277 0 .546.039.803.111V9.417a6.338 6.338 0 0 0-.803-.051C6.017 9.366 3.2 12.183 3.2 15.647 3.2 19.11 6.017 22 9.479 22c3.462 0 6.279-2.817 6.279-6.353V9.07c1.378.983 3.054 1.564 4.869 1.584V7.209a4.845 4.845 0 0 1-1.038-.523z"/></svg>
                </a>
                <!-- يوتيوب -->
                <a href="https://youtube.com/@engineermaths?si=8C6T808VuAU5OMOt" target="_blank" title="YouTube" class="social-btn youtube-bg">
                    <svg fill="#ffffff" viewBox="0 0 24 24"><path d="M23.498 6.186a3.016 3.016 0 0 0-2.122-2.136C19.505 3.545 12 3.545 12 3.545s-7.505 0-9.377.505A3.017 3.017 0 0 0 .502 6.186C0 8.07 0 12 0 12s0 3.93.502 5.814a3.016 3.016 0 0 0 2.122 2.136c1.871.505 9.376.505 9.376.505s7.505 0 9.377-.505a3.015 3.015 0 0 0 2.122-2.136C24 15.93 24 12 24 12s0-3.93-.502-5.814zM9.545 15.568V8.432L15.818 12l-6.273 3.568z"/></svg>
                </a>
            </div>
            <div class="rights-text">
                جميع الحقوق محفوظة لدي م / محمد غنيم 2026
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.stop()

# ==============================================================================
# 2. لوحة تحكم المعلم الرئيسية
# ==============================================================================
st.sidebar.markdown("### 📷 صورة الشعار والمعلم")
uploaded_photo = st.sidebar.file_uploader(
    "ارفع صورتك هنا إذا لم تظهر تلقائياً:", type=["jpg", "png", "jpeg"]
)
if uploaded_photo is not None:
    with open(IMG_NAME, "wb") as f:
        f.write(uploaded_photo.getbuffer())
    found_img_path = IMG_NAME
    st.sidebar.success("✓ تم حفظ صورتك بنجاح!")
    st.rerun()

if found_img_path and os.path.exists(found_img_path):
    st.sidebar.image(found_img_path, width=220)

st.sidebar.markdown(
    """
<div style="text-align: center; margin-top: 5px; margin-bottom: 20px;">
    <h2 style="margin: 0; color: #0052cc; font-weight: 900;">البشمهندس X الرياضة</h2>
    <p class="sidebar-desc" style="margin: 4px 0; font-weight: 800; font-size: 15px;">نظام إدارة ومتابعة الحصص</p>
</div>
""",
    unsafe_allow_html=True,
)

st.sidebar.header("🔗 رابط تسجيل الطلاب")
st.sidebar.markdown("شارك الرابط مع الطلاب ليسجلوا حضورهم بأنفسهم:")
student_url = "https://engmohamedghonaim.streamlit.app/?role=student"
st.sidebar.code(student_url, language="text")

st.markdown(
    f"""
<div class="brand-banner" dir="rtl">
    <div>
        <h1 class="brand-title">البشمهندس X الرياضة 📐</h1>
        <p class="brand-subtitle">إدارة الحصص • متابعة نسب الحضور والغياب • الكشوف المالية ومستوى الطلاب</p>
    </div>
    {'<img src="data:image/jpeg;base64,' + img_b64 + '" style="width: 90px; height: 90px; border-radius: 50%; border: 3px solid #ffffff; object-fit: cover;">' if img_b64 else ''}
</div>
""",
    unsafe_allow_html=True,
)

tab1, tab2, tab3, tab4 = st.tabs([
    "📝 رصد حصة جديدة",
    "✏️ تعديل ومراجعة السجلات",
    "📊 قاعدة البيانات الشاملة",
    "🖨️ كشف وطباعة بطاقة طالب",
])

# ----------------- تبويب 1: رصد الحصة -----------------
with tab1:
    st.subheader("إدخال بيانات الحصة")
    t_curriculum = st.selectbox(
        "اختر المنهج الدراسي:",
        list(CURRICULUM_DATA.keys()),
        key="teacher_curr_select",
    )
    t_grades = CURRICULUM_DATA[t_curriculum]

    with st.form("teacher_entry_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            session_date = st.date_input("تاريخ الحصة", value=date.today())
            student_name = st.text_input(
                "اسم الطالب", placeholder="مثال: أحمد محمد"
            )
            group_name = st.selectbox("المرحلة / الصف الدراسي:", t_grades)
            status = st.selectbox(
                "حالة الحضور", ["حاضر", "غائب", "متأخر", "بعذر"]
            )

        with col2:
            price = st.number_input(
                "سعر الحصة", min_value=0.0, step=10.0, value=100.0
            )
            total_sessions = st.number_input(
                "الحصص المنفذة حتى الآن", min_value=1, step=1, value=1
            )
            payment_type = st.selectbox(
                "نظام الدفع",
                ["اشتراك شهري", "مقدم", "مؤخر (بعد الحصة)", "مؤجل"],
            )
            student_level = st.selectbox(
                "المستوى الدراسي",
                [
                    "ممتاز ⭐⭐⭐",
                    "جيد جداً ⭐⭐",
                    "جيد ⭐",
                    "متوسط",
                    "يحتاج متابعة",
                ],
            )

        notes = st.text_area(
            "ملاحظات الواجب أو التقييم والدرجات",
            placeholder="أداء الحصة، حل الواجب...",
        )
        submitted = st.form_submit_button("💾 رصد وحفظ الحصة")

        if submitted:
            if not student_name.strip():
                st.error("يرجى كتابة اسم الطالب أولاً.")
            else:
                new_row = {
                    "التاريخ": str(session_date),
                    "اسم الطالب": student_name.strip(),
                    "المنهج/الدولة": t_curriculum,
                    "المجموعة/الصف": group_name,
                    "الحالة": status,
                    "سعر الحصة": price,
                    "عدد الحصص الكلي": total_sessions,
                    "نظام الدفع": payment_type,
                    "مستوى الطالب": student_level,
                    "ملاحظات": notes.strip(),
                }
                st.session_state.data = pd.concat(
                    [st.session_state.data, pd.DataFrame([new_row])],
                    ignore_index=True,
                )
                save_data(st.session_state.data)
                st.success(f"✓ تم حفظ سجل الطالب ({student_name}) بنجاح!")

# ----------------- تبويب 2: التعديل والمراجعة -----------------
with tab2:
    st.subheader("مراجعة وتعديل بيانات الطلاب والحصص")
    df = st.session_state.data

    if df.empty:
        st.info("لا توجد سجلات مسجلة بعد.")
    else:
        record_options = {
            idx: f"[{row['التاريخ']}] - {row['اسم الطالب']} ({row.get('المنهج/الدولة', '')} | {row['المجموعة/الصف']}) - {row['الحالة']}"
            for idx, row in df.iterrows()
        }
        selected_idx = st.selectbox(
            "اختر السجل المراد تعديل بياناته:",
            options=list(record_options.keys()),
            format_func=lambda x: record_options[x],
        )

        selected_row = df.loc[selected_idx]
        try:
            curr_date = datetime.strptime(
                str(selected_row["التاريخ"]), "%Y-%m-%d"
            ).date()
        except Exception:
            curr_date = date.today()

        with st.form("edit_record_form"):
            c1, c2 = st.columns(2)
            with c1:
                edit_name = st.text_input(
                    "اسم الطالب:", value=str(selected_row["اسم الطالب"])
                )

                saved_curr = selected_row.get("المنهج/الدولة", "")
                curr_list = list(CURRICULUM_DATA.keys())
                curr_idx = (
                    curr_list.index(saved_curr) if saved_curr in curr_list else 0
                )
                edit_curr = st.selectbox(
                    "المنهج / الدولة:", curr_list, index=curr_idx
                )

                edit_group = st.text_input(
                    "المرحلة / الصف الدراسي:",
                    value=str(selected_row["المجموعة/الصف"]),
                )
                edit_date = st.date_input("التاريخ:", value=curr_date)
                s_opts = ["حاضر", "غائب", "متأخر", "بعذر"]
                edit_status = st.selectbox(
                    "الحالة:",
                    s_opts,
                    index=(
                        s_opts.index(selected_row["الحالة"])
                        if selected_row["الحالة"] in s_opts
                        else 0
                    ),
                )

            with c2:
                try:
                    p_val = float(selected_row["سعر الحصة"])
                except Exception:
                    p_val = 0.0
                edit_price = st.number_input(
                    "سعر الحصة:", min_value=0.0, step=10.0, value=p_val
                )

                try:
                    s_val = int(selected_row["عدد الحصص الكلي"])
                except Exception:
                    s_val = 1
                edit_sessions = st.number_input(
                    "عدد الحصص الكلي:", min_value=1, step=1, value=s_val
                )

                p_opts = ["اشتراك شهري", "مقدم", "مؤخر (بعد الحصة)", "مؤجل"]
                edit_pay = st.selectbox(
                    "نظام الدفع:",
                    p_opts,
                    index=(
                        p_opts.index(selected_row["نظام الدفع"])
                        if selected_row["نظام الدفع"] in p_opts
                        else 0
                    ),
                )

                l_opts = [
                    "ممتاز ⭐⭐⭐",
                    "جيد جداً ⭐⭐",
                    "جيد ⭐",
                    "متوسط",
                    "يحتاج متابعة",
                    "قيد التقييم",
                ]
                edit_level = st.selectbox(
                    "المستوى:",
                    l_opts,
                    index=(
                        l_opts.index(selected_row["مستوى الطالب"])
                        if selected_row["مستوى الطالب"] in l_opts
                        else 0
                    ),
                )

            edit_notes = st.text_area(
                "الملاحظات والتقييم:", value=str(selected_row["ملاحظات"])
            )

            b1, b2 = st.columns([1, 4])
            with b1:
                update_btn = st.form_submit_button("💾 تحديث البيانات")
            with b2:
                delete_btn = st.form_submit_button("🗑️ حذف السجل")

            if update_btn:
                df.at[selected_idx, "التاريخ"] = str(edit_date)
                df.at[selected_idx, "اسم الطالب"] = edit_name.strip()
                df.at[selected_idx, "المنهج/الدولة"] = edit_curr
                df.at[selected_idx, "المجموعة/الصف"] = edit_group.strip()
                df.at[selected_idx, "الحالة"] = edit_status
                df.at[selected_idx, "سعر الحصة"] = edit_price
                df.at[selected_idx, "عدد الحصص الكلي"] = edit_sessions
                df.at[selected_idx, "نظام الدفع"] = edit_pay
                df.at[selected_idx, "مستوى الطالب"] = edit_level
                df.at[selected_idx, "ملاحظات"] = edit_notes.strip()

                save_data(df)
                st.session_state.data = df
                st.success("✓ تم تحديث بيانات الحصة بنجاح!")
                st.rerun()

            if delete_btn:
                df = df.drop(selected_idx).reset_index(drop=True)
                save_data(df)
                st.session_state.data = df
                st.warning("⚠️ تم حذف السجل.")
                st.rerun()

# ----------------- تبويب 3: السجلات العامة -----------------
with tab3:
    st.subheader("نظرة شاملة على السجلات والإحصائيات")
    current_df = st.session_state.data

    if current_df.empty:
        st.info("لا توجد بيانات متاحة.")
    else:
        c_f1, c_f2, c_f3 = st.columns(3)
        with c_f1:
            all_currs = ["الكل"] + [
                c
                for c in current_df["المنهج/الدولة"].dropna().unique()
                if str(c).strip()
            ]
            filter_curr = st.selectbox("تصفية حسب المنهج / الدولة:", all_currs)
        with c_f2:
            all_groups = ["الكل"] + [
                g
                for g in current_df["المجموعة/الصف"].dropna().unique()
                if str(g).strip()
            ]
            filter_group = st.selectbox("تصفية حسب المرحلة / الصف:", all_groups)
        with c_f3:
            search_name = st.text_input("بحث باسم الطالب:")

        filtered = current_df.copy()
        if filter_curr != "الكل":
            filtered = filtered[filtered["المنهج/الدولة"] == filter_curr]
        if filter_group != "الكل":
            filtered = filtered[filtered["المجموعة/الصف"] == filter_group]
        if search_name.strip():
            filtered = filtered[
                filtered["اسم الطالب"].str.contains(
                    search_name.strip(), na=False
                )
            ]

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("إجمالي الحصص", len(filtered))
        m2.metric(
            "عدد مرات الحضور", len(filtered[filtered["الحالة"] == "حاضر"])
        )
        m3.metric("عدد مرات الغياب", len(filtered[filtered["الحالة"] == "غائب"]))
        total_cash = (
            filtered["سعر الحصة"]
            .astype(float, errors="ignore")
            .sum(numeric_only=True)
        )
        m4.metric("إجمالي المبالغ المستحقة", f"{total_cash:,.1f}")

        st.dataframe(filtered, width="stretch")

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            current_df.to_excel(writer, index=False)
        st.download_button(
            label="📥 تصدير السجل بالكامل إلى Excel",
            data=buf.getvalue(),
            file_name=FILE_NAME,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ----------------- تبويب 4: الكشف والطباعة -----------------
with tab4:
    st.subheader("كشف متابعة وحساب الطالب الجاهز للطباعة")
    current_df = st.session_state.data

    if current_df.empty:
        st.info("السجل فارغ حالياً.")
    else:
        st_names = sorted(
            list(
                set([
                    s
                    for s in current_df["اسم الطالب"].dropna().unique()
                    if str(s).strip()
                ])
            )
        )
        selected_student = st.selectbox("اختر الطالب لإصدار الكشف:", st_names)

        if selected_student:
            st_records = (
                current_df[current_df["اسم الطالب"] == selected_student]
                .copy()
                .sort_values(by="التاريخ")
            )
            latest = st_records.iloc[-1]

            level_val = latest.get("مستوى الطالب", "غير محدد")
            pay_val = latest.get("نظام الدفع", "غير محدد")
            curr_val = latest.get("المنهج/الدولة", "-")
            group_val = latest.get("المجموعة/الصف", "-")
            price_val = latest.get("سعر الحصة", 0)

            att_cnt = len(st_records[st_records["الحالة"] == "حاضر"])
            abs_cnt = len(st_records[st_records["الحالة"] == "غائب"])
            total_cnt = len(st_records)
            total_due = (
                st_records["سعر الحصة"]
                .astype(float, errors="ignore")
                .sum(numeric_only=True)
            )

            col_k1, col_k2, col_k3, col_k4 = st.columns(4)
            col_k1.metric("المنهج والمرحلة", f"{curr_val} - {group_val}")
            col_k2.metric("المستوى", str(level_val))
            col_k3.metric("نظام الدفع", str(pay_val))
            col_k4.metric("سعر الحصة", f"{price_val}")

            col_k5, col_k6, col_k7, col_k8 = st.columns(4)
            col_k5.metric("الحصص المنفذة", total_cnt)
            col_k6.metric("مرات الحضور", att_cnt)
            col_k7.metric("مرات الغياب", abs_cnt)
            col_k8.metric("إجمالي الحساب", f"{total_due:,.1f}")

            rows_html = ""
            for _, r in st_records.iterrows():
                st_color = (
                    "#0f766e"
                    if r["الحالة"] == "حاضر"
                    else ("#b91c1c" if r["الحالة"] == "غائب" else "#b45309")
                )
                rows_html += f"""
                <tr>
                    <td style="padding: 10px; border: 2px solid #000; font-weight: 800;">{r['التاريخ']}</td>
                    <td style="padding: 10px; border: 2px solid #000; color:{st_color}; font-weight: 900;">{r['الحالة']}</td>
                    <td style="padding: 10px; border: 2px solid #000; font-weight: 800;">{r['سعر الحصة']}</td>
                    <td style="padding: 10px; border: 2px solid #000; font-weight: 800;">{r['نظام الدفع']}</td>
                    <td style="padding: 10px; border: 2px solid #000; font-weight: 900; color: #0052cc;">{r['مستوى الطالب']}</td>
                    <td style="padding: 10px; border: 2px solid #000; font-weight: 800;">{r['ملاحظات']}</td>
                </tr>
                """

            teacher_img_tag = (
                f'<img src="data:image/jpeg;base64,{img_b64}" style="width: 95px; height: 95px; border-radius: 50%; border: 3px solid #0052cc; object-fit: cover;">'
                if img_b64
                else ""
            )

            printable_html = f"""<!DOCTYPE html>
            <html dir="rtl" lang="ar">
            <head>
                <meta charset="utf-8">
                <title>كشف متابعة - {selected_student}</title>
                <style>
                    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@700;800;900&display=swap');
                    body {{ 
                        font-family: 'Cairo', Tahoma, Arial, sans-serif; 
                        padding: 35px; 
                        color: #000000 !important; 
                        background-color: #ffffff !important;
                        font-weight: 800;
                    }}
                    .header-box {{
                        border-bottom: 3px solid #0052cc;
                        padding-bottom: 18px;
                        margin-bottom: 25px;
                        display: flex;
                        justify-content: space-between;
                        align-items: center;
                    }}
                    .brand-name {{ color: #0052cc; margin: 0; font-size: 32px; font-weight: 900; }}
                    .brand-sub {{ color: #000000; margin: 5px 0 0 0; font-size: 16px; font-weight: 800; }}
                    .stats-box {{ width: 100%; border-collapse: collapse; margin-bottom: 25px; }}
                    .stats-box th, .stats-box td {{ border: 2px solid #000000; padding: 12px; text-align: center; font-size: 16px; font-weight: 800; }}
                    .stats-box th {{ background: #ffffff; color: #000000; font-weight: 900; }}
                    .details-table {{ width: 100%; border-collapse: collapse; text-align: center; margin-top: 15px; }}
                    .details-table th {{ background-color: #ffffff; color: #000000; font-weight: 900; font-size: 16px; padding: 12px; border: 2px solid #000000; }}
                    .footer-note {{ margin-top: 35px; text-align: center; color: #000000; font-size: 15px; font-weight: 800; border-top: 2px solid #000000; padding-top: 15px; }}
                </style>
            </head>
            <body onload="window.print()">
                <div class="header-box">
                    <div style="display: flex; align-items: center; gap: 20px;">
                        {teacher_img_tag}
                        <div>
                            <h2 class="brand-name">البشمهندس X الرياضة 📐</h2>
                            <p class="brand-sub">كشف التقييم الأكاديمي والحساب المالي الدوري</p>
                            <p style="margin: 4px 0; color: #000000; font-size: 16px; font-weight: 800;"><b>المنهج الدراسي:</b> {curr_val} — {group_val}</p>
                        </div>
                    </div>
                    <div style="text-align: left;">
                        <h2 style="color: #0052cc; margin: 0; font-size: 26px; font-weight: 900;">{selected_student}</h2>
                        <p style="margin: 5px 0 0 0; color: #000000; font-size: 15px; font-weight: 800;">تاريخ الكشف: {date.today()}</p>
                    </div>
                </div>

                <table class="stats-box">
                    <tr>
                        <th>المستوى الدراسي</th>
                        <th>نظام الدفع</th>
                        <th>سعر الحصة</th>
                        <th>إجمالي الحصص</th>
                        <th>مرات الحضور</th>
                        <th>مرات الغياب</th>
                        <th>المبلغ المستحق</th>
                    </tr>
                    <tr>
                        <td style="color: #0052cc; font-weight: 900;">{level_val}</td>
                        <td>{pay_val}</td>
                        <td>{price_val}</td>
                        <td style="font-weight: 900;">{total_cnt}</td>
                        <td style="color: #0f766e; font-weight: 900;">{att_cnt}</td>
                        <td style="color: #b91c1c; font-weight: 900;">{abs_cnt}</td>
                        <td style="font-weight: 900; font-size: 18px; color: #000000;">{total_due:,.1f}</td>
                    </tr>
                </table>

                <h3 style="margin-top: 30px; margin-bottom: 12px; color: #0052cc; font-size: 20px; font-weight: 900;">سجل تفاصيل الحصص والواجبات:</h3>
                <table class="details-table">
                    <tr>
                        <th>التاريخ</th>
                        <th>الحالة</th>
                        <th>سعر الحصة</th>
                        <th>نظام الدفع</th>
                        <th>المستوى</th>
                        <th>ملاحظات المعلم وتقييم الطالب</th>
                    </tr>
                    {rows_html}
                </table>

                <div class="footer-note">
                    مع تحيات: <b>البشمهندس X الرياضة</b> — متابعة مستمرة نحو التفوق والدرجة النهائية 🌟
                </div>
            </body>
            </html>"""

            st.download_button(
                label=f"🖨️ تحميل وطباعة كشف ({selected_student}) كـ PDF",
                data=printable_html.encode("utf-8"),
                file_name=f"كشف_{selected_student}.html",
                mime="text/html",
            )

            st.write("---")
            st.dataframe(
                st_records[[
                    "التاريخ",
                    "الحالة",
                    "سعر الحصة",
                    "نظام الدفع",
                    "مستوى الطالب",
                    "ملاحظات",
                ]],
                width="stretch",
            )
