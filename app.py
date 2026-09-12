
import os as _os
import io
import json
import base64
import urllib.request
import urllib.error
from datetime import date, datetime, time
import pandas as pd
from PIL import Image
import streamlit as st
try:
    from weasyprint import HTML as WeasyHTML
except Exception:
    WeasyHTML = None

st.set_page_config(
    page_title="م/ محمد غنيم | منصة شرح الرياضيات والإحصاء",
    page_icon="📐",
    layout="wide",
)

try:
    from streamlit_paste_button import paste_image_button
except ImportError:
    paste_image_button = None

try:
    from streamlit_cropper import st_cropper
except ImportError:
    st_cropper = None

FILE_NAME = "سجل_الغياب_والحصص.xlsx"
IMG_NAME = "teacher.jpg"
TEACHER_PHONE = "01016361440"

# تخزين دائم اختياري على Supabase لمنع ضياع بيانات المنصة عند إعادة تشغيل Streamlit Cloud.
# إذا لم يتم ضبط الأسرار، يستمر التطبيق في العمل بالطريقة المحلية القديمة.
SUPABASE_URL = ""
SUPABASE_KEY = ""
try:
    SUPABASE_URL = str(st.secrets.get("SUPABASE_URL", "")).strip()
    SUPABASE_KEY = str(st.secrets.get("SUPABASE_KEY", "")).strip()
except Exception:
    pass
SUPABASE_TABLE = "platform_storage"
SUPABASE_INTERFACE_TABLE = "student_interface_storage"
SUPABASE_RECORD_ID = "main"

def _cloud_storage_enabled():
    return bool(SUPABASE_URL and SUPABASE_KEY)

def _supabase_headers():
    return {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/json",
        "Prefer": "return=representation",
    }

def _cloud_load_excel_bytes():
    """قراءة نسخة Excel الكاملة من التخزين الدائم، إن كانت موجودة."""
    if not _cloud_storage_enabled():
        return None
    try:
        url = f"{SUPABASE_URL.rstrip('/')}/rest/v1/{SUPABASE_TABLE}?id=eq.{SUPABASE_RECORD_ID}&select=payload"
        req = urllib.request.Request(url, headers=_supabase_headers(), method="GET")
        with urllib.request.urlopen(req, timeout=20) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        if data and data[0].get("payload"):
            return base64.b64decode(data[0]["payload"])
    except Exception:
        pass
    return None

def _cloud_save_excel_bytes(excel_bytes):
    """حفظ نسخة Excel الكاملة في سجل واحد دائم على Supabase."""
    if not _cloud_storage_enabled():
        return False
    try:
        payload = base64.b64encode(excel_bytes).decode("ascii")
        body = json.dumps({"id": SUPABASE_RECORD_ID, "payload": payload}).encode("utf-8")
        url = f"{SUPABASE_URL.rstrip('/')}/rest/v1/{SUPABASE_TABLE}"
        headers = _supabase_headers()
        headers["Prefer"] = "resolution=merge-duplicates,return=minimal"
        req = urllib.request.Request(url, data=body, headers=headers, method="POST")
        with urllib.request.urlopen(req, timeout=30) as resp:
            resp.read()
        return True
    except Exception as exc:
        # لا نوقف المنصة بالكامل إذا حدث عطل مؤقت في التخزين السحابي.
        try:
            st.session_state["cloud_storage_last_error"] = str(exc)
        except Exception:
            pass
        return False


def _cloud_load_student_interface():
    """قراءة إعدادات واجهة الطالب وحدها من التخزين الدائم، بدون الاعتماد على ملف Excel الكبير."""
    if not _cloud_storage_enabled():
        return None
    try:
        url = f"{SUPABASE_URL.rstrip('/')}/rest/v1/{SUPABASE_INTERFACE_TABLE}?id=eq.{SUPABASE_RECORD_ID}&select=payload"
        req = urllib.request.Request(url, headers=_supabase_headers(), method="GET")
        with urllib.request.urlopen(req, timeout=20) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        if data and data[0].get("payload"):
            return json.loads(data[0]["payload"])
    except Exception as exc:
        try:
            st.session_state["cloud_interface_load_error"] = str(exc)
        except Exception:
            pass
    return None

def _cloud_save_student_interface(interface_df):
    """حفظ واجهة الطالب (النصوص والصور) في سجل مستقل صغير."""
    if not _cloud_storage_enabled():
        return False
    try:
        row = interface_df.iloc[0].to_dict() if interface_df is not None and not interface_df.empty else {}
        clean = {}
        for key, value in row.items():
            if pd.isna(value):
                value = ""
            clean[str(key)] = str(value)
        payload = json.dumps(clean, ensure_ascii=False)
        body = json.dumps({"id": SUPABASE_RECORD_ID, "payload": payload}, ensure_ascii=False).encode("utf-8")
        url = f"{SUPABASE_URL.rstrip('/')}/rest/v1/{SUPABASE_INTERFACE_TABLE}"
        headers = _supabase_headers()
        headers["Prefer"] = "resolution=merge-duplicates,return=minimal"
        req = urllib.request.Request(url, data=body, headers=headers, method="POST")
        with urllib.request.urlopen(req, timeout=30) as resp:
            resp.read()
        st.session_state["cloud_interface_last_saved"] = True
        return True
    except Exception as exc:
        try:
            st.session_state["cloud_interface_last_error"] = str(exc)
            st.session_state["cloud_interface_last_saved"] = False
        except Exception:
            pass
        return False

def _get_excel_source():
    """يفضل التخزين الدائم، ثم يرجع للملف المحلي القديم كخطة احتياطية."""
    cloud_bytes = _cloud_load_excel_bytes()
    if cloud_bytes:
        return io.BytesIO(cloud_bytes)
    if _os.path.exists(FILE_NAME):
        return FILE_NAME
    return None

CURRICULUM_DATA = {
    "المنهج المصري 🇪🇬": [
        "الصف الأول الإعدادي", "الصف الثاني الإعدادي", "الصف الثالث الإعدادي (الشهادة الإعدادية)",
        "الصف الأول الثانوي", "الصف الثاني الثانوي (علمي)", "الصف الثاني الثانوي (أدبي)",
        "الصف الثالث الثانوي (علمي رياضة)", "الصف الثالث الثانوي (علمي علوم)", "الصف الثالث الثانوي (أدبي)",
        "المرحلة الابتدائية"
    ],
    "المنهج القطري 🇶🇦": [
        "الصف السابع (إعدادي)", "الصف الثامن (إعدادي)", "الصف التاسع (إعدادي)",
        "الصف العاشر (المشترك)", "الصف الحادي عشر (المسار العلمي)", "الصف الحادي عشر (مسار الآداب والإنسانيات)",
        "الصف الحادي عشر (المسار التكنولوجي)", "الصف الثاني عشر (المسار العلمي - متقدم)",
        "الصف الثاني عشر (مسار الآداب - تأسيسي)", "الصف الثاني عشر (المسار التكنولوجي)",
        "المرحلة الابتدائية"
    ],
    "المنهج الإماراتي 🇦🇪": [
        "الحلقة الثانية (الصفوف 5 - 8)", "الصف التاسع (مسار عام)", "الصف التاسع (مسار متقدم)",
        "الصف العاشر (مسار عام)", "الصف العاشر (مسار متقدم)", "الصف العاشر (مسار النخبة)",
        "الصف الحادي عشر (مسار عام)", "الصف الحادي عشر (مسار متقدم)", "الصف الحادي عشر (مسار النخبة)",
        "الصف الثاني عشر (مسار عام)", "الصف الثاني عشر (مسار متقدم)", "الصف الثاني عشر (مسار النخبة)"
    ],
    "المنهج السعودي 🇸🇦": [
        "المرحلة المتوسطة (أول / ثاني / ثالث متوسط)", "السنة الأولى المشتركة (أول ثانوي)",
        "السنة الثانية (المسار العام)", "السنة الثانية (مسار علوم الحاسب والهندسة)",
        "السنة الثانية (مسار الصحة والحياة)", "السنة الثانية (مسار إدارة الأعمال / الشرعي)",
        "السنة الثالثة (المسار العام)", "السنة الثالثة (مسار علوم الحاسب والهندسة)",
        "السنة الثالثة (مسار الصحة والحياة)"
    ],
    "المنهج الكويتي 🇰🇼": [
        "المرحلة المتوسطة (الصفوف 6 - 9)", "الصف العاشر الثانوي (مشترك)",
        "الصف الحادي عشر (القسم العلمي)", "الصف الحادي عشر (القسم الأدبي)",
        "الصف الثاني عشر (القسم العلمي)", "الصف الثاني عشر (القسم الأدبي)"
    ],
    "المنهج السوداني 🇸🇩": [
        "المرحلة المتوسطة (أولى / ثانية / ثالثة متوسط)", "الصف الأول الثانوي",
        "الصف الثاني الثانوي (علمي)", "الصف الثاني الثانوي (أدبي)",
        "الصف الثالث الثانوي (علمي رياضيات - الشهادة السودانية)", "الصف الثالث الثانوي (علمي أحياء)",
        "الصف الثالث الثانوي (أدبي)"
    ]
}

possible_images = ["teacher.jpg", "teacher.png", "teacher.jpeg", "photo_2026-08-02_00-34-53.jpg"]
found_img_path = None
for img_cand in possible_images:
    if __import__("os").path.exists(img_cand):
        found_img_path = img_cand
        break

def get_image_base64(path):
    if path and _os.path.exists(path):
        try:
            with open(path, "rb") as img_file:
                return base64.b64encode(img_file.read()).decode()
        except Exception:
            return ""
    return ""

def pil_to_base64(pil_img):
    if pil_img is None:
        return ""
    buf = io.BytesIO()
    pil_img.save(buf, format="PNG")
    return base64.b64encode(buf.getvalue()).decode()

def base64_to_pil(b64_str):
    if not b64_str:
        return None
    try:
        img_bytes = base64.b64decode(b64_str)
        return Image.open(io.BytesIO(img_bytes))
    except Exception:
        return None

img_b64 = get_image_base64(found_img_path)

# صورة ثابتة لواجهة الطالب: مدمجة داخل التطبيق نفسه حتى تظهر دائمًا على المنصة.
STUDENT_FIXED_IMAGE_B64 = "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TRwCQyAwygyQxTBRighhAyjBjjhjgQhCyiQjBgABDSAAADSgQzAhgjDAAQxBCTxjwxRSCQgAACAAhQASCABiAAgATAgQDCBAAAAAAAAAAgBCgABBQBAggABQACgSDQDihxyigRzxxTwTwzQwyzhQTTRATBSQxzDjBBzgQSDDCCASgBAjgDDzBiwwRCDDAAjjyCwBDBQAghQQBARCTTRAAgAAAACgQAAABAAAAAAQwAQwAAAQBQhAgQAAAAgADQRxRBwxDjxCzRxADSjCQgAhhzhxRxzgzQxzjDgBRxzygAAwxAhAgjQgigjgAASzCjyzgjQACSAigABDSRBBAgQAQBQAgAACAAAAAAABAAjAAAAQQABwBByzwwAgACgBQjjzCQMwFysO4TiyBgSz5+xzTizQzCAzhwihjgBjAxCCxgQgBARByzzThRzAQBDRBARSCyACAgAQiSRABgggigAACAAAAAAAgQAAgCQwAACgABAQAAAAAAgAAAAAAAACyCyRhSTMZecZs7X8vyMQ4zA+yhzxCTQRxThiAxCBBBQyBQQACgAAAxTxxgRggADACTzRzTgCBAgRwwADRATBCAgQRABhAAADAiAAABAAAAgARwSwAAAghADAhAhAAAAgRSc+J0B9zs96P55YBhRGiQDSfhTzzjiSwAzDCQDQBRyRgiACQjihhxySxjTwCghhhCzDAQxTBQASASShCSBhADAABAAAAAAAAAAQBRDAAAxgBAAQAAAQASQQQAAgQACCUtkufzPDxge/DCKFbpZhABhDJijBzhADQzSAwjASARCQgQwjwBwBhBRRhTCBQABAjiigwBATQiBgAAQAhDAzBQBAAAAAAgBAgAAAQABQhgAACAAABgDggARgABSyCNS91bQ6mOC1BwiBt7RQUXMnqkPvuOwwCTwTRyxghzgyAQASgwBDgwgijzTySgChCCBRzyDBTiAAAggAACQRwCATgCCCAAAAAAgQAAAAAQgRBBBAAhxAACABAQAAgCXkqgFe5kw2lvaqRCTUyy9OOYMQDOxfUATPdxBjzyTjwhDAwhSBABhzxDDwSCzjSCAQABDCQhRBxhQQhCADwBCABAAARABAAAAABAAAAAABBwgASAAAAAAQAAABAQAADkqQuBBhCBo6NcACiyiCSkIR7RACMxhyM9njlyRxwTzDjCAxCACAAjBSRBBz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CAENGKIGCCEAAAIAADGKEEGEIICFGHOCDBAIAKKEOGJNOMEIDIAAAAABACCIBHHHOAKDAAAICAAICCGAMLAADtmlvhb0gEhIOFGGfuER7gomxxMmFQHiYopma92ejGqwELCDPICBNAIAAKJDICBKBDLBJIDNCHNPKOCLIIAALAPMABIAAIBDAEAAACIBHGDBLPACOOKFIAAFICACABAGIsLMUHD0sSSIBKBM3UrgMjOtV0x1m/wgimeBon0SLKouPCEIKONOJEKCNCHFIEACHIEAJBMHNCIODBFPIDNENEKIAAAAABGFAIEEAAHCBCCAAAMFOIAABAEGAIAIADBHBOCU2aW0D3ZNNFPIATg3qwLELNIPGtTpsNQK2Upbh2yAIAAEECBBJKEEBICEBADMACAAAIJCGEAMGPFHMGIEAMCHMBJAIEOEJFAAIEDAAABABAEBMKBEAAAEBBIACGADGDXOCiWkptSPLDLN4YobHdD/AIoSFxWD2iDigUqXHzzkQQAAABQDiTAADQggAAAAgDQBCCgBAAwBSiQAABSABRQABADAABAQBSxAAADQAAAAAAACSCRAiBBgBAAAAAAACABJAMspNfd3jRyxHHgzEzh3GHpW7TaSRCXRjShz4xKaRAAAAAAAACABAAAAggQAABQwQQABiAyAhRwSiAAAADAgAAgACAAAACDAAAjTyAAhRDiBiSQAABAADSgAQCASgADLgemIhS/3vBhUB2h7Yk02WHuhwZ3QiHBj5TMfITqCAAAAgRBAAAAAAACBABggSRRQCAgSAQDyxggABBCAAggAASAQAwAABSAAAAhSAiCAAAgDxgShgjgDDAgAQwQQwQDR3OzSDl1JKFeUesYUzejUVL5iai1V3BATClPcg9mAACAAABAABQAAgAABBSCBCBSCASBAAAATgRCQACBAASwAAQgAAAQABBAgRhAQiRgAzSATyiwwSyQBCBDDABQxDARCV4cC8r1ol0HThRwF3zqTkW53K2nQACjDqAegNcCAAACiQCQBAiAAAAgAACAgBCDgACACCACChTwAgAQAQBAiQAgAQAACAgDyjgAATRgASDCwCCyBRiAhAAAQACDhiigQYs/zccB9+pK5QCqh32ADw9xLOLDQCAT7GE82d0wAAAAgRAAAAAAAAAAAgAgABgQACQAAADgDSQAQAhQhAgAQBAxw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"


def teacher_image_data_uri(b64_str):
    """تحويل صورة المعلم المحفوظة Base64 إلى Data URI صحيح مع قص المساحات البيضاء الزائدة."""
    if not b64_str or str(b64_str).strip().lower() == "nan":
        return ""
    try:
        raw = base64.b64decode(str(b64_str))
        im = Image.open(io.BytesIO(raw)).convert("RGBA")

        # قص الحواف البيضاء/الشفافة الزائدة حتى لا يظهر الوجه في نصف الدائرة فقط.
        bbox = None
        pix = im.load()
        xs, ys = [], []
        for y in range(im.height):
            for x in range(im.width):
                r, g, b, a = pix[x, y]
                if a > 18 and not (r > 245 and g > 245 and b > 245):
                    xs.append(x)
                    ys.append(y)
        if xs and ys:
            left, top, right, bottom = min(xs), min(ys), max(xs) + 1, max(ys) + 1
            pad = max(8, int(max(right-left, bottom-top) * 0.10))
            left = max(0, left-pad); top = max(0, top-pad)
            right = min(im.width, right+pad); bottom = min(im.height, bottom+pad)
            bbox = (left, top, right, bottom)
        if bbox:
            im = im.crop(bbox)

        # نجعل الصورة مربعة مع الحفاظ على الوجه في المنتصف.
        side = max(im.width, im.height)
        canvas = Image.new("RGBA", (side, side), (255, 255, 255, 0))
        canvas.paste(im, ((side-im.width)//2, (side-im.height)//2), im)
        buf = io.BytesIO()
        canvas.save(buf, format="PNG")
        return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode("utf-8")
    except Exception:
        # fallback: تحديد نوع الصورة من التوقيع بدل افتراض JPEG
        try:
            raw = base64.b64decode(str(b64_str))
            if raw.startswith(b"\x89PNG"):
                mime = "image/png"
            elif raw.startswith(b"\xff\xd8\xff"):
                mime = "image/jpeg"
            elif raw.startswith(b"RIFF") and raw[8:12] == b"WEBP":
                mime = "image/webp"
            else:
                mime = "image/png"
            return f"data:{mime};base64,{b64_str}"
        except Exception:
            return ""

# بطاقات الاشتراكات الموحدة للصفحة الرئيسية وصفحة الطالب
DARSSLY_COURSES = [
    {"icon": "📘", "title": "رياضيات أول إعدادي", "price": 200, "link": "https://darssly.com/courses/mathematics-for-preparatory-stage-mr-mohamed-ghonaim-3/plans", "tag": "أولى إعدادي"},
    {"icon": "📗", "title": "رياضيات ثاني إعدادي", "price": 200, "link": "https://darssly.com/courses/mathematics-for-preparatory-stage-mr-mohamed-ghonaim/plans", "tag": "ثانية إعدادي"},
    {"icon": "📐", "title": "رياضيات ثالث إعدادي", "price": 200, "link": "https://darssly.com/courses/mathematics-for-preparatory-stage-mr-mohamed-ghonaim-2/plans", "tag": "ثالثة إعدادي"},
    {"icon": "📊", "title": "إحصاء ثالثة ثانوي", "price": 250, "link": "https://darssly.com/courses/mohamed-ghoneim-statistics/plans", "tag": "ثالثة ثانوي – إحصاء"},
]

def render_darssly_cards(section_key="home"):
    interface_df = st.session_state.get("student_interface_df", pd.DataFrame())
    row = interface_df.iloc[0] if interface_df is not None and not interface_df.empty else {}
    sub_title = str(row.get("عنوان_الاشتراكات", "📢 اشتراكات درسلي"))
    sub_desc = str(row.get("وصف_الاشتراكات", "اختر المرحلة وشاهد نظام الشرح والمتابعة والسعر الشهري"))
    sub_img = STUDENT_FIXED_IMAGE_B64
    sub_uri = teacher_image_data_uri(sub_img) if sub_img else ""
    st.markdown("<div class='darssly-box'>", unsafe_allow_html=True)
    st.markdown(f"<h3 style='color:#fff;text-align:center;margin:0 0 5px;font-size:23px;font-weight:900;'>{sub_title}</h3>", unsafe_allow_html=True)
    st.markdown(f"<p style='color:#fff;text-align:center;margin:0 0 20px;font-size:15px;font-weight:800;'>{sub_desc}</p>", unsafe_allow_html=True)
    cols = st.columns(2)
    for idx, course in enumerate(DARSSLY_COURSES):
        with cols[idx % 2]:
            photo = (f"<img class='subscription-photo' src='{sub_uri}'>" if sub_uri else f"<div class='subscription-icon'>{course['icon']}</div>")
            st.markdown(f"""
                <div class='subscription-card' dir='rtl'>
                    <div class='subscription-badge'>باقـة {course['tag']}</div>
                    {photo}
                    <div class='subscription-title'>{course['title']}</div>
                    <div class='subscription-price'>{course['price']} جنيه / شهر</div>
                    <div class='subscription-features'>✓ فيديوهات شرح مسجلة<br>✓ حصص Zoom مباشرة<br>✓ متابعة مستمرة<br>✓ حل وتدريب على الأسئلة</div>
                </div>
            """, unsafe_allow_html=True)
            st.link_button(f"🔴 معرفة تفاصيل الاشتراك — {course['title']}", course['link'], use_container_width=True)
            st.write("")
    st.markdown('</div>', unsafe_allow_html=True)


COL_SESSIONS = ["التاريخ", "اسم الطالب", "المنهج/الدولة", "المجموعة/الصف", "الحالة", "سعر الحصة", "عدد الحصص الكلي", "نظام الدفع", "مستوى الطالب", "ملاحظات"]
COL_USERS = ["اسم الطالب", "رقم الهاتف", "كلمة المرور", "المنهج/الدولة", "المجموعة/الصف", "اسم ولي الأمر", "رقم ولي الأمر", "تاريخ التسجيل", "الحالة_حظر", "حالة_الاشتراك_البنك"]
COL_ASSESSMENTS = ["التاريخ", "اسم الطالب", "النوع", "عنوان التكليف", "الدرجة المحصلة", "الدرجة العظمى", "حالة التسليم", "ملاحظات وتوجيهات"]
COL_MESSAGES = ["التاريخ_والوقت", "اسم الطالب", "المرسل", "نص الرسالة", "الصورة_base64"]
COL_EXAMS = ["معرف_الامتحان", "عنوان الامتحان", "وصف الامتحان", "كلمة المرور", "المنهج/الدولة", "المجموعة/الصف", "المادة", "الفصل الدراسي", "مدة الامتحان بالدقائق", "الأسئلة_JSON", "تاريخ الإنشاء"]
COL_ESSAYS = ["معرف_الحل", "معرف_الامتحان", "عنوان الامتحان", "اسم الطالب", "رقم السؤال", "نص السؤال", "إجابة الطالب النصية", "صورة الحل_base64", "درجة السؤال", "الدرجة المرصودة", "حالة التصحيح", "ملاحظات المعلم", "تاريخ الحل"]
COL_BOOKINGS = ["تاريخ_الحجز", "اسم الطالب", "المنهج_الدولة", "المرحلة_الصف", "رقم_الهاتف", "رقم_ولي_الأمر", "الحالة"]
COL_BANK_REQUESTS = ["تاريخ_الطلب", "اسم الطالب", "رقم_الهاتف", "كود_OTP", "حالة_الدفع", "إيصال_الدفع_base64"]
COL_QUESTION_BANK = ["معرف_السؤال", "المنهج/الدولة", "المجموعة/الصف", "المادة", "نوع_السؤال", "بيانات_السؤال_JSON"]
COL_VIDEOS = ["معرف_الفيديو", "عنوان_الفيديو", "المنهج/الدولة", "المجموعة/الصف", "رابط_الفيديو", "فيديو_base64", "تاريخ_الرفع"]
COL_VIDEO_COMMENTS = ["التاريخ_والوقت", "عنوان_الفيديو", "اسم الطالب", "نص_التعليق"]
COL_ABQARY = ["معرف_عبقري", "عنوان_الإمتحان", "المنهج/الدولة", "المجموعة/الصف", "رابط_الإمتحان", "كود_HTML", "رابط_النتيجة", "الرقم_السري_للنتيجة", "تاريخ_النشر"]
COL_ONLINE_SCHEDULE = ["اسم الطالب", "اسم الأكاديمية", "المنهج/الدولة", "المجموعة/الصف", "رقم الطالب", "رقم مشرف الأكاديمية", "سعر الحصة", "تاريخ الحصة", "ساعة الحصة", "رابط زوم", "حالة فتح الحصة"]
COL_WEEKLY_SCHEDULE = ["اسم الطالب", "اسم الأكاديمية", "المنهج/الدولة", "المجموعة/الصف", "رقم الطالب", "رقم مشرف الأكاديمية", "سعر الحصة", "اليوم", "الموعد", "اللون", "حالة الموعد"]
COL_TEACHER_PROFILE = ["اسم المعلم", "الصورة_base64"]
COL_STUDENT_INTERFACE = ["عنوان_الواجهة", "الشارة", "الوصف", "صورة_الواجهة_base64", "عنوان_الاشتراكات", "وصف_الاشتراكات", "عنوان_الحجز", "نص_الحجز", "نص_الفوتر", "صورة_الاشتراكات_base64", "صورة_البانر_base64"]
COL_PAYMENT_RECORDS = ["التاريخ", "الشهر", "اسم الطالب", "المبلغ", "طريقة الدفع", "حالة الدفع", "ملاحظات"]

def load_teacher_profile():
    profile = pd.DataFrame(columns=COL_TEACHER_PROFILE)
    excel_source = _get_excel_source()
    if excel_source is not None:
        try:
            with pd.ExcelFile(excel_source) as xls:
                if "TeacherProfile" in xls.sheet_names:
                    profile = pd.read_excel(xls, "TeacherProfile")
        except Exception:
            pass
    for c in COL_TEACHER_PROFILE:
        if c not in profile.columns:
            profile[c] = ""
    if profile.empty:
        profile = pd.DataFrame([{"اسم المعلم":"م/ محمد غنيم","الصورة_base64":img_b64}])
    elif not str(profile.iloc[0].get("الصورة_base64", "")).strip() and img_b64:
        profile.at[0,"الصورة_base64"] = img_b64
    return profile

def load_student_interface():
    defaults = {
        "عنوان_الواجهة": "أهلاً بيكم منورين المنصة! 🚀",
        "الشارة": "منصة شرح الرياضيات والإحصاء",
        "الوصف": "مع م / محمد غنيم. خبرة متميزة في تدريس الرياضيات والإحصاء للثانوية العامة والمرحلة الإعدادية. آلاف الطلاب حققوا التفوق والدرجات النهائية.",
        "صورة_الواجهة_base64": STUDENT_FIXED_IMAGE_B64, "عنوان_الاشتراكات": "📢 اشتراكات درسلي",
        "وصف_الاشتراكات": "اختر المرحلة وشاهد نظام الشرح والمتابعة والسعر الشهري",
        "عنوان_الحجز": "📅 حجز دروس أونلاين مباشرة مع م / محمد غنيم",
        "نص_الحجز": "احجز درس أونلاين مباشر وحدد المنهج والمرحلة ورقم الهاتف للتواصل معك.",
        "نص_الفوتر": "جميع الحقوق محفوظة لدي م / محمد غنيم 2026",
        "صورة_الاشتراكات_base64": "", "صورة_البانر_base64": ""
    }
    df = pd.DataFrame([defaults])
    # واجهة الطالب لها تخزين سحابي مستقل حتى لو كان ملف Excel كبيراً أو لم يتم تحميله.
    cloud_interface = _cloud_load_student_interface()
    if isinstance(cloud_interface, dict) and cloud_interface:
        for key, default in defaults.items():
            value = cloud_interface.get(key, default)
            if value is None or str(value).lower() == "nan":
                value = default
            df.at[0, key] = value
        if not str(df.at[0, "صورة_الواجهة_base64"]).strip() and STUDENT_FIXED_IMAGE_B64:
            df.at[0, "صورة_الواجهة_base64"] = STUDENT_FIXED_IMAGE_B64
        if not str(df.at[0, "صورة_الاشتراكات_base64"]).strip() and STUDENT_FIXED_IMAGE_B64:
            df.at[0, "صورة_الاشتراكات_base64"] = STUDENT_FIXED_IMAGE_B64
        return df[COL_STUDENT_INTERFACE]
    source = _get_excel_source()
    if source is not None:
        try:
            with pd.ExcelFile(source) as xls:
                if "StudentInterface" in xls.sheet_names:
                    loaded = pd.read_excel(xls, "StudentInterface")
                    if not loaded.empty:
                        row = loaded.iloc[0].to_dict()
                        for key, default in defaults.items():
                            value = row.get(key, default)
                            if pd.isna(value): value = default
                            df.at[0, key] = value
        except Exception:
            pass
    if not str(df.at[0, "صورة_الواجهة_base64"]).strip() and img_b64:
        df.at[0, "صورة_الواجهة_base64"] = img_b64
    if not str(df.at[0, "صورة_الاشتراكات_base64"]).strip() and img_b64:
        df.at[0, "صورة_الاشتراكات_base64"] = img_b64
    return df[COL_STUDENT_INTERFACE]


def _get_print_profile():
    """بيانات المعلم المستخدمة في رأس ملفات الطباعة."""
    name = "م/ محمد غنيم"
    photo_b64 = img_b64
    try:
        profile_df = st.session_state.get("teacher_profile_df", pd.DataFrame())
        if profile_df is not None and not profile_df.empty:
            saved_name = str(profile_df.iloc[0].get("اسم المعلم", "")).strip()
            saved_photo = str(profile_df.iloc[0].get("الصورة_base64", "")).strip()
            if saved_name and saved_name.lower() != "nan":
                name = saved_name
            if saved_photo and saved_photo.lower() != "nan":
                photo_b64 = saved_photo
    except Exception:
        pass
    return name, photo_b64, TEACHER_PHONE

def make_print_html(title, rows_html, headers_html, subtitle=""):
    teacher_name, teacher_photo_b64, teacher_phone = _get_print_profile()
    photo_html = ""
    if teacher_photo_b64:
        photo_html = f"<img class='teacher-photo' src='data:image/png;base64,{teacher_photo_b64}' alt='صورة المعلم'>"
    return f"""<!DOCTYPE html>
<html dir='rtl' lang='ar'>
<head>
<meta charset='utf-8'>
<title>{title}</title>
<style>
@page{{size:A4 landscape;margin:10mm}}
*{{box-sizing:border-box}}
body{{font-family:'Noto Kufi Arabic','Noto Sans Arabic','Amiri',Tahoma,Arial,sans-serif;font-weight:700;color:#10233f;padding:8px;background:#fff}}
.header{{border:2px solid #0f766e;border-radius:18px;padding:14px 18px;margin-bottom:14px;background:linear-gradient(135deg,#effcf8,#eef7ff);display:flex;align-items:center;gap:16px;direction:rtl}}
.teacher-photo{{width:82px;height:82px;border-radius:50%;object-fit:cover;border:4px solid #0f766e;background:#fff;display:block}}
.brand{{flex:1;text-align:right}}
.brand h2{{margin:0 0 4px;font-size:22px;color:#075985;font-weight:900}}
.brand .phone{{font-size:13px;color:#334155;margin-top:4px}}
.doc-title{{text-align:center;margin:10px 0 4px;font-size:21px;color:#0f172a;font-weight:900}}
.subtitle{{text-align:center;margin:0 0 8px;color:#475569;font-size:12px}}
.badge{{display:inline-block;background:#0f766e;color:white;border-radius:999px;padding:4px 12px;font-size:11px;margin-top:4px}}
table{{width:100%;border-collapse:separate;border-spacing:0;margin-top:12px;overflow:hidden;border:1.5px solid #94a3b8;border-radius:10px}}
th,td{{border-left:1px solid #cbd5e1;border-bottom:1px solid #cbd5e1;padding:7px 6px;text-align:center;font-size:10.5px;vertical-align:middle}}
th{{background:#0f766e;color:#fff;font-size:11px;font-weight:900}}
tr:last-child td{{border-bottom:0}}
th:last-child,td:last-child{{border-left:0}}
.student{{color:#fff;border-radius:9px;padding:6px 5px;margin:2px 0;line-height:1.45;box-shadow:0 2px 5px rgba(15,23,42,.12)}}
.student b{{display:block;font-size:11px}}
.student span,.student small{{display:block;font-size:8.5px}}
.time{{font-weight:900;background:#f8fafc;font-size:11px;white-space:nowrap}}
.footer{{margin-top:14px;padding-top:8px;border-top:1px solid #cbd5e1;text-align:center;font-size:9.5px;color:#475569}}
@media print{{body{{padding:0}}}}
</style>
</head>
<body>
<div class='header'>{photo_html}<div class='brand'><h2>{teacher_name}</h2><div>منصة شرح الرياضيات والإحصاء</div><div class='phone'>📞 {teacher_phone}</div></div></div>
<div class='doc-title'>{title}</div>
<p class='subtitle'>{subtitle}</p>
<table><thead><tr>{headers_html}</tr></thead><tbody>{rows_html}</tbody></table>
<div class='footer'>إعداد ومتابعة: {teacher_name} &nbsp; | &nbsp; 📞 {teacher_phone} &nbsp; | &nbsp; جميع البيانات خاصة بالمنصة التعليمية</div>
</body>
</html>"""

def html_to_pdf_bytes(html_text):
    if WeasyHTML is None:
        return None
    try:
        return WeasyHTML(string=html_text, base_url=_os.getcwd()).write_pdf()
    except Exception:
        return None

def build_student_roster_html(names, title="كشف الطلاب المسجلين"):
    rows=[]
    for nm in names:
        ur=st.session_state.users_df[st.session_state.users_df["اسم الطالب"].astype(str).str.strip()==nm]
        wr=st.session_state.weekly_schedule_df[st.session_state.weekly_schedule_df["اسم الطالب"].astype(str).str.strip()==nm]
        src=ur.iloc[0].to_dict() if not ur.empty else (wr.iloc[0].to_dict() if not wr.empty else {})
        rows.append(f"<tr><td>{nm}</td><td>{src.get('رقم الهاتف',src.get('رقم الطالب',''))}</td><td>{src.get('اسم ولي الأمر','')}</td><td>{src.get('رقم ولي الأمر','')}</td><td>{src.get('المنهج/الدولة','')}</td><td>{src.get('المجموعة/الصف','')}</td></tr>")
    return make_print_html(title,''.join(rows) or "<tr><td colspan='6'>لا توجد بيانات</td></tr>","<th>اسم الطالب</th><th>رقم الطالب</th><th>اسم ولي الأمر</th><th>رقم ولي الأمر</th><th>المنهج</th><th>المرحلة</th>",f"إجمالي الطلاب: {len(names)}")


def format_schedule_time_ampm(value):
    """عرض وقت الحصة بصيغة 12 ساعة مع AM/PM بدون تغيير القيمة المخزنة."""
    try:
        text = str(value).strip()
        if not text or text.lower() == "nan":
            return text
        dt = pd.to_datetime(text, format="%H:%M", errors="coerce")
        if pd.isna(dt):
            dt = pd.to_datetime(text, errors="coerce")
        if pd.isna(dt):
            return text
        return dt.strftime("%I:%M %p").lstrip("0")
    except Exception:
        return str(value)


def build_weekly_schedule_print_html(df, title="الجدول الأسبوعي لمواعيد الطلاب"):
    """إنشاء نسخة طباعة/PDF مطابقة للجدول الأسبوعي بالأيام أعمدة، مع لون مستقل لكل طالب."""
    days = ["السبت", "الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة"]
    if df is None or df.empty:
        return make_print_html(title, "<tr><td colspan='8'>لا توجد مواعيد</td></tr>", "<th>الساعة</th>" + "".join(f"<th>{d}</th>" for d in days))
    work = df.copy()
    work["_time"] = work["الموعد"].astype(str).str[:5]
    work["_day"] = work["اليوم"].astype(str)
    times = sorted([t for t in work["_time"].dropna().unique() if t and t != "nan"])
    rows=[]
    for tm in times:
        cells=[f"<td class='time'>{tm}</td>"]
        for d in days:
            matches=work[(work["_time"]==tm) & (work["_day"]==d) & (work["حالة الموعد"].astype(str)!="متوقف")]
            parts=[]
            for _,r in matches.iterrows():
                color=str(r.get("اللون","#2563eb"))
                if not color.startswith("#"): color="#2563eb"
                student=str(r.get("اسم الطالب",""))
                grade=str(r.get("المجموعة/الصف",""))
                academy=str(r.get("اسم الأكاديمية",""))
                parts.append(f"<div class='student' style='background:{color};'><b>👤 {student}</b><span>{grade}</span><small>{academy}</small></div>")
            cells.append("<td>"+("".join(parts) if parts else "—")+"</td>")
        rows.append("<tr>"+"".join(cells)+"</tr>")
    headers="<th>الساعة</th>"+"".join(f"<th>{d}</th>" for d in days)
    css="""<style>body{font-family:Tahoma,Arial,sans-serif;font-weight:700;color:#111;direction:rtl}h1{text-align:center;color:#075985}.subtitle{text-align:center;margin-bottom:14px}table{width:100%;border-collapse:collapse;table-layout:fixed}th,td{border:1.5px solid #111;padding:6px;text-align:center;vertical-align:top;font-size:10px;min-height:55px}th{background:#e2e8f0;font-size:12px}.time{width:55px;font-size:12px;vertical-align:middle}.student{color:#fff;border-radius:8px;padding:7px 4px;margin:2px 0;line-height:1.25}.student b,.student span,.student small{display:block}.student span{font-size:9px;margin-top:3px}.student small{font-size:8px;margin-top:2px}</style>"""
    html=make_print_html(title, ''.join(rows), headers, f"إجمالي المواعيد: {len(work)}")
    return html.replace("</head>", css+"</head>")

def load_all_data():
    users_df = pd.DataFrame(columns=COL_USERS)
    sessions_df = pd.DataFrame(columns=COL_SESSIONS)
    assessments_df = pd.DataFrame(columns=COL_ASSESSMENTS)
    messages_df = pd.DataFrame(columns=COL_MESSAGES)
    exams_df = pd.DataFrame(columns=COL_EXAMS)
    essays_df = pd.DataFrame(columns=COL_ESSAYS)
    bookings_df = pd.DataFrame(columns=COL_BOOKINGS)
    bank_requests_df = pd.DataFrame(columns=COL_BANK_REQUESTS)
    question_bank_df = pd.DataFrame(columns=COL_QUESTION_BANK)
    videos_df = pd.DataFrame(columns=COL_VIDEOS)
    video_comments_df = pd.DataFrame(columns=COL_VIDEO_COMMENTS)
    abqary_df = pd.DataFrame(columns=COL_ABQARY)
    online_schedule_df = pd.DataFrame(columns=COL_ONLINE_SCHEDULE)
    weekly_schedule_df = pd.DataFrame(columns=COL_WEEKLY_SCHEDULE)
    payment_records_df = pd.DataFrame(columns=COL_PAYMENT_RECORDS)

    excel_source = _get_excel_source()
    if excel_source is not None:
        try:
            with pd.ExcelFile(excel_source) as xls:
                if "Users" in xls.sheet_names: users_df = pd.read_excel(xls, "Users")
                if "Sessions" in xls.sheet_names: sessions_df = pd.read_excel(xls, "Sessions")
                elif "Sheet1" in xls.sheet_names: sessions_df = pd.read_excel(xls, "Sheet1")
                if "Assessments" in xls.sheet_names: assessments_df = pd.read_excel(xls, "Assessments")
                if "Messages" in xls.sheet_names: messages_df = pd.read_excel(xls, "Messages")
                if "Exams" in xls.sheet_names: exams_df = pd.read_excel(xls, "Exams")
                if "Essays" in xls.sheet_names: essays_df = pd.read_excel(xls, "Essays")
                if "Bookings" in xls.sheet_names: bookings_df = pd.read_excel(xls, "Bookings")
                if "BankRequests" in xls.sheet_names: bank_requests_df = pd.read_excel(xls, "BankRequests")
                if "QuestionBank" in xls.sheet_names: question_bank_df = pd.read_excel(xls, "QuestionBank")
                if "Videos" in xls.sheet_names: videos_df = pd.read_excel(xls, "Videos")
                if "VideoComments" in xls.sheet_names: video_comments_df = pd.read_excel(xls, "VideoComments")
                if "AbqaryExams" in xls.sheet_names:
                    abqary_df = pd.read_excel(xls, "AbqaryExams")
                    # دعم البيانات القديمة التي لا تحتوي على كود HTML
                    if "كود_HTML" not in abqary_df.columns:
                        abqary_df["كود_HTML"] = ""
                    for _c in COL_ABQARY:
                        if _c not in abqary_df.columns:
                            abqary_df[_c] = ""
                    abqary_df = abqary_df[COL_ABQARY]
                if "OnlineSchedule" in xls.sheet_names: online_schedule_df = pd.read_excel(xls, "OnlineSchedule")
                if "WeeklySchedule" in xls.sheet_names: weekly_schedule_df = pd.read_excel(xls, "WeeklySchedule")
                if "PaymentRecords" in xls.sheet_names: payment_records_df = pd.read_excel(xls, "PaymentRecords")
        except Exception:
            pass

    for col in COL_USERS:
        if col not in users_df.columns:
            if col == "رقم الهاتف": users_df[col] = ""
            elif col in ["اسم ولي الأمر", "رقم ولي الأمر"]: users_df[col] = ""
            elif col == "الحالة_حظر": users_df[col] = "نشط"
            elif col == "حالة_الاشتراك_البنك": users_df[col] = "غير مشترك"
            else: users_df[col] = ""
    for col in COL_WEEKLY_SCHEDULE:
        if col not in weekly_schedule_df.columns:
            if col == "اللون": weekly_schedule_df[col] = "#2563eb"
            elif col == "حالة الموعد": weekly_schedule_df[col] = "نشط"
            else: weekly_schedule_df[col] = ""

    for col in COL_PAYMENT_RECORDS:
        if col not in payment_records_df.columns:
            if col == "حالة الدفع": payment_records_df[col] = "مؤكد"
            elif col == "المبلغ": payment_records_df[col] = 0.0
            else: payment_records_df[col] = ""

    for col in COL_ONLINE_SCHEDULE:
        if col not in online_schedule_df.columns:
            if col == "رابط زوم": online_schedule_df[col] = "https://us05web.zoom.us/j/83526892910?pwd=2jWRgATgBRPbXttdnm0QpLwBApsZL4.1"
            elif col == "حالة فتح الحصة": online_schedule_df[col] = "مغلقة"
            elif col == "اسم الأكاديمية": online_schedule_df[col] = "أكاديمية البشمهندس"
            else: online_schedule_df[col] = ""

    return users_df, sessions_df, assessments_df, messages_df, exams_df, essays_df, bookings_df, bank_requests_df, question_bank_df, videos_df, video_comments_df, abqary_df, online_schedule_df, weekly_schedule_df, payment_records_df

def save_all_data(users_df, sessions_df, assessments_df, messages_df, exams_df, essays_df, bookings_df, bank_requests_df, question_bank_df, videos_df, video_comments_df, abqary_df, online_schedule_df, weekly_schedule_df=None, payment_records_df=None):
    if weekly_schedule_df is None:
        weekly_schedule_df = st.session_state.get("weekly_schedule_df", pd.DataFrame(columns=COL_WEEKLY_SCHEDULE))
    if payment_records_df is None:
        payment_records_df = st.session_state.get("payment_records_df", pd.DataFrame(columns=COL_PAYMENT_RECORDS))
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        users_df.to_excel(writer, sheet_name="Users", index=False)
        sessions_df.to_excel(writer, sheet_name="Sessions", index=False)
        assessments_df.to_excel(writer, sheet_name="Assessments", index=False)
        messages_df.to_excel(writer, sheet_name="Messages", index=False)
        exams_df.to_excel(writer, sheet_name="Exams", index=False)
        essays_df.to_excel(writer, sheet_name="Essays", index=False)
        bookings_df.to_excel(writer, sheet_name="Bookings", index=False)
        bank_requests_df.to_excel(writer, sheet_name="BankRequests", index=False)
        question_bank_df.to_excel(writer, sheet_name="QuestionBank", index=False)
        videos_df.to_excel(writer, sheet_name="Videos", index=False)
        video_comments_df.to_excel(writer, sheet_name="VideoComments", index=False)
        abqary_df.to_excel(writer, sheet_name="AbqaryExams", index=False)
        online_schedule_df.to_excel(writer, sheet_name="OnlineSchedule", index=False)
        weekly_schedule_df.to_excel(writer, sheet_name="WeeklySchedule", index=False)
        payment_records_df.to_excel(writer, sheet_name="PaymentRecords", index=False)
        st.session_state.get("student_interface_df", load_student_interface()).to_excel(writer, sheet_name="StudentInterface", index=False)
        st.session_state.get("teacher_profile_df", pd.DataFrame([{"اسم المعلم":"م/ محمد غنيم","الصورة_base64":img_b64}])).to_excel(writer, sheet_name="TeacherProfile", index=False)
    excel_bytes = excel_buffer.getvalue()
    # احفظ محلياً أيضاً عندما يكون ذلك ممكناً، ثم ارفع نفس الملف للتخزين الدائم.
    try:
        with open(FILE_NAME, "wb") as _local_file:
            _local_file.write(excel_bytes)
    except Exception:
        pass
    _cloud_save_excel_bytes(excel_bytes)

if "users_df" not in st.session_state:
    u_df, s_df, a_df, m_df, e_df, es_df, b_df, br_df, qb_df, v_df, vc_df, ab_df, os_df, ws_df, pr_df = load_all_data()
    st.session_state.users_df = u_df
    st.session_state.sessions_df = s_df
    st.session_state.assessments_df = a_df
    st.session_state.messages_df = m_df
    st.session_state.exams_df = e_df
    st.session_state.essays_df = es_df
    st.session_state.bookings_df = b_df
    st.session_state.bank_requests_df = br_df
    st.session_state.question_bank_df = qb_df
    st.session_state.videos_df = v_df
    st.session_state.video_comments_df = vc_df
    st.session_state.abqary_df = ab_df
    st.session_state.online_schedule_df = os_df
    st.session_state.weekly_schedule_df = ws_df
    st.session_state.payment_records_df = pr_df
    st.session_state.teacher_profile_df = load_teacher_profile()
    # استخدم صورة المعلم المحفوظة داخل TeacherProfile/التخزين السحابي في كل صفحات الطالب
    # بدلاً من الاعتماد على ملف teacher.jpg الموجود محلياً فقط.
    try:
        _saved_teacher_photo = str(st.session_state.teacher_profile_df.iloc[0].get("الصورة_base64", "")).strip() if not st.session_state.teacher_profile_df.empty else ""
        if _saved_teacher_photo and _saved_teacher_photo.lower() != "nan":
            img_b64 = _saved_teacher_photo
    except Exception:
        pass

# تأكد من وجود جدول المواعيد حتى لو كانت جلسة Streamlit قديمة قبل إضافة الميزة
if "weekly_schedule_df" not in st.session_state:
    try:
        _, _, _, _, _, _, _, _, _, _, _, _, _, ws_df = load_all_data()
        st.session_state.weekly_schedule_df = ws_df
    except Exception:
        st.session_state.weekly_schedule_df = pd.DataFrame(columns=COL_WEEKLY_SCHEDULE)
if "payment_records_df" not in st.session_state:
    st.session_state.payment_records_df = pd.DataFrame(columns=COL_PAYMENT_RECORDS)
    if _os.path.exists(FILE_NAME):
        try:
            with pd.ExcelFile(FILE_NAME) as _xls_pay:
                if "PaymentRecords" in _xls_pay.sheet_names:
                    st.session_state.payment_records_df = pd.read_excel(_xls_pay, "PaymentRecords")
        except Exception:
            pass
    for _pc in COL_PAYMENT_RECORDS:
        if _pc not in st.session_state.payment_records_df.columns:
            st.session_state.payment_records_df[_pc] = 0.0 if _pc == "المبلغ" else ("مؤكد" if _pc == "حالة الدفع" else "")

if "teacher_profile_df" not in st.session_state:
    st.session_state.teacher_profile_df = load_teacher_profile()

# مزامنة صورة المعلم من الملف السحابي مع المتغير المستخدم في الصفحة الرئيسية وبطاقات درسلي
try:
    _saved_teacher_photo = str(st.session_state.teacher_profile_df.iloc[0].get("الصورة_base64", "")).strip() if not st.session_state.teacher_profile_df.empty else ""
    if _saved_teacher_photo and _saved_teacher_photo.lower() != "nan":
        img_b64 = _saved_teacher_photo
except Exception:
    pass

if "student_interface_df" not in st.session_state:
    st.session_state.student_interface_df = load_student_interface()

if "page_view" not in st.session_state:
    st.session_state.page_view = "home"

if "student_sub_page" not in st.session_state:
    st.session_state.student_sub_page = "dashboard"

if "teacher_page" not in st.session_state:
    st.session_state.teacher_page = "dashboard"

if "dark_mode" not in st.session_state:
    st.session_state.dark_mode = False

query_params = st.query_params
is_student_mode = query_params.get("role") == "student"

if "logged_student" not in st.session_state:
    st.session_state.logged_student = None

saved_student_phone = query_params.get("st_phone")
if not st.session_state.logged_student and saved_student_phone:
    matched_st = st.session_state.users_df[st.session_state.users_df["رقم الهاتف"].astype(str).str.strip() == str(saved_student_phone).strip()]
    if not matched_st.empty:
        st.session_state.logged_student = matched_st.iloc[0].to_dict()

def delete_student_completely(student_name_to_del):
    target = student_name_to_del.strip()
    st.session_state.users_df = st.session_state.users_df[st.session_state.users_df["اسم الطالب"].astype(str).str.strip() != target].reset_index(drop=True)
    st.session_state.sessions_df = st.session_state.sessions_df[st.session_state.sessions_df["اسم الطالب"].astype(str).str.strip() != target].reset_index(drop=True)
    st.session_state.assessments_df = st.session_state.assessments_df[st.session_state.assessments_df["اسم الطالب"].astype(str).str.strip() != target].reset_index(drop=True)
    st.session_state.messages_df = st.session_state.messages_df[st.session_state.messages_df["اسم الطالب"].astype(str).str.strip() != target].reset_index(drop=True)
    st.session_state.essays_df = st.session_state.essays_df[st.session_state.essays_df["اسم الطالب"].astype(str).str.strip() != target].reset_index(drop=True)
    st.session_state.bookings_df = st.session_state.bookings_df[st.session_state.bookings_df["اسم الطالب"].astype(str).str.strip() != target].reset_index(drop=True)
    st.session_state.bank_requests_df = st.session_state.bank_requests_df[st.session_state.bank_requests_df["اسم الطالب"].astype(str).str.strip() != target].reset_index(drop=True)
    st.session_state.video_comments_df = st.session_state.video_comments_df[st.session_state.video_comments_df["اسم الطالب"].astype(str).str.strip() != target].reset_index(drop=True)
    st.session_state.online_schedule_df = st.session_state.online_schedule_df[st.session_state.online_schedule_df["اسم الطالب"].astype(str).str.strip() != target].reset_index(drop=True)
    st.session_state.weekly_schedule_df = st.session_state.weekly_schedule_df[st.session_state.weekly_schedule_df["اسم الطالب"].astype(str).str.strip() != target].reset_index(drop=True)
    st.session_state.payment_records_df = st.session_state.payment_records_df[st.session_state.payment_records_df["اسم الطالب"].astype(str).str.strip() != target].reset_index(drop=True)
    save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)

def _money_sum(series):
    try:
        return float(pd.to_numeric(series, errors="coerce").fillna(0).sum())
    except Exception:
        return 0.0


def get_student_financials(student_name):
    """حساب إجمالي الحصص والمدفوعات والرصيد المتبقي للطالب."""
    target = str(student_name).strip()
    s_df = st.session_state.get("sessions_df", pd.DataFrame())
    p_df = st.session_state.get("payment_records_df", pd.DataFrame())
    if s_df.empty or "اسم الطالب" not in s_df.columns:
        due = 0.0
    else:
        rows = s_df[s_df["اسم الطالب"].astype(str).str.strip() == target]
        due = _money_sum(rows["سعر الحصة"]) if "سعر الحصة" in rows.columns else 0.0
    if p_df.empty or "اسم الطالب" not in p_df.columns:
        paid = 0.0
    else:
        if "حالة الدفع" in p_df.columns:
            rows = p_df[(p_df["اسم الطالب"].astype(str).str.strip() == target) & (p_df["حالة الدفع"].astype(str).str.strip().isin(["مؤكد", "مدفوع"]))]
        else:
            rows = p_df[p_df["اسم الطالب"].astype(str).str.strip() == target]
        paid = _money_sum(rows["المبلغ"]) if "المبلغ" in rows.columns else 0.0
    return due, paid, due - paid


def get_all_financial_totals():
    p_df = st.session_state.get("payment_records_df", pd.DataFrame())
    s_df = st.session_state.get("sessions_df", pd.DataFrame())
    total_due = _money_sum(s_df["سعر الحصة"]) if not s_df.empty and "سعر الحصة" in s_df.columns else 0.0
    if not p_df.empty and "المبلغ" in p_df.columns:
        if "حالة الدفع" in p_df.columns:
            paid_rows = p_df[p_df["حالة الدفع"].astype(str).str.strip().isin(["مؤكد", "مدفوع"]) ]
        else:
            paid_rows = p_df
        total_paid = _money_sum(paid_rows["المبلغ"])
    else:
        total_paid = 0.0
    return total_due, total_paid, total_due - total_paid


def _month_from_value(value):
    try:
        return pd.to_datetime(value).strftime("%Y-%m")
    except Exception:
        return ""

def get_student_monthly_financials(student_name, month_key):
    target = str(student_name).strip()
    s_df = st.session_state.get("sessions_df", pd.DataFrame()).copy()
    p_df = st.session_state.get("payment_records_df", pd.DataFrame()).copy()
    if not s_df.empty and "اسم الطالب" in s_df.columns:
        rows = s_df[s_df["اسم الطالب"].astype(str).str.strip() == target].copy()
        if "التاريخ" in rows.columns:
            rows = rows[rows["التاريخ"].apply(_month_from_value) == month_key]
        due = _money_sum(rows["سعر الحصة"]) if "سعر الحصة" in rows.columns else 0.0
    else:
        due = 0.0
    if not p_df.empty and "اسم الطالب" in p_df.columns:
        rows = p_df[p_df["اسم الطالب"].astype(str).str.strip() == target].copy()
        if "الشهر" in rows.columns:
            rows = rows[rows["الشهر"].astype(str).str.strip() == month_key]
        elif "التاريخ" in rows.columns:
            rows = rows[rows["التاريخ"].apply(_month_from_value) == month_key]
        if "حالة الدفع" in rows.columns:
            rows = rows[rows["حالة الدفع"].astype(str).str.strip().isin(["مؤكد", "مدفوع"])]
        paid = _money_sum(rows["المبلغ"]) if "المبلغ" in rows.columns else 0.0
    else:
        paid = 0.0
    return due, paid, due - paid

def get_monthly_financial_totals(month_key):
    s_df = st.session_state.get("sessions_df", pd.DataFrame()).copy()
    p_df = st.session_state.get("payment_records_df", pd.DataFrame()).copy()
    sm = s_df[s_df["التاريخ"].apply(_month_from_value) == month_key] if not s_df.empty and "التاريخ" in s_df.columns else pd.DataFrame()
    due = _money_sum(sm["سعر الحصة"]) if not sm.empty and "سعر الحصة" in sm.columns else 0.0
    if not p_df.empty:
        if "الشهر" in p_df.columns:
            pm = p_df[p_df["الشهر"].astype(str).str.strip() == month_key]
        else:
            pm = p_df[p_df["التاريخ"].apply(_month_from_value) == month_key] if "التاريخ" in p_df.columns else pd.DataFrame()
        if "حالة الدفع" in pm.columns:
            pm = pm[pm["حالة الدفع"].astype(str).str.strip().isin(["مؤكد", "مدفوع"])]
        paid = _money_sum(pm["المبلغ"]) if "المبلغ" in pm.columns else 0.0
    else:
        paid = 0.0
    return due, paid, due - paid


if st.session_state.dark_mode:
    bg_color = "#0f172a"
    text_color = "#f8fafc"
    card_bg = "#1e293b"
    card_border = "#334155"
else:
    bg_color = "#ffffff"
    text_color = "#0f172a"
    card_bg = "#f8fafc"
    card_border = "#cbd5e1"

st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@700;800;900&display=swap');
    
    html, body, [class*="css"], p, span, label, div, button, h1, h2, h3, h4, h5, h6 {{
        font-family: 'Cairo', sans-serif !important;
        font-weight: 900 !important;
        letter-spacing: 0.3px !important;
        color: {text_color} !important;
    }}

    .stApp {{
        background-color: {bg_color} !important;
    }}

    input, textarea, select, div[data-baseweb="select"] > div {{
        font-family: 'Cairo', sans-serif !important;
        font-weight: 800 !important;
        color: {text_color} !important;
        background-color: {card_bg} !important;
    }}

    div[data-baseweb="popover"], div[role="dialog"], div[aria-label*="Choose a date"], div[aria-label*="Calendar"], div[data-baseweb="calendar"] {{
        background-color: {card_bg} !important;
        color: {text_color} !important;
        border-radius: 12px !important;
        box-shadow: 0 10px 25px rgba(0,0,0,0.15) !important;
    }}
    div[data-baseweb="calendar"] *, div[data-baseweb="popover"] * {{
        color: {text_color} !important;
        background-color: transparent !important;
    }}
    div[data-baseweb="calendar"] button {{
        color: {text_color} !important;
        background-color: {card_bg} !important;
        border-radius: 6px !important;
        font-weight: 900 !important;
    }}
    div[data-baseweb="calendar"] button:hover {{
        background-color: #059669 !important;
        color: #ffffff !important;
    }}

    div[data-baseweb="popover"], div[role="listbox"] div {{
        background-color: {card_bg} !important;
        color: {text_color} !important;
    }}
    div[role="option"] span, div[role="option"] div {{
        color: {text_color} !important;
    }}

    body, h1, h2, h3, h4, h5, h6, .stMarkdown, .stSelectbox, .stTextInput, .stTextArea {{
        direction: rtl;
        text-align: right;
    }}
    
    img {{
        max-width: 100% !important;
        height: auto !important;
        object-fit: contain !important;
        border-radius: 8px;
    }}

    .darssly-navbar {{
        background: {card_bg};
        padding: 12px 25px;
        border-radius: 14px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 25px;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.08);
        border: 1px solid {card_border};
    }}
    .navbar-brand {{ display: flex; align-items: center; gap: 15px; }}
    .navbar-title-group h3 {{ margin: 0 !important; color: #059669 !important; font-size: 18px !important; }}
    .navbar-title-group p {{ margin: 2px 0 0 0 !important; color: {text_color} !important; font-size: 12px !important; opacity: 0.8; }}

    .exam-builder-header {{
        background-color: #f59e0b; color: #ffffff !important; padding: 16px 22px; border-radius: 10px 10px 0 0; font-size: 22px; font-weight: 900; margin-bottom: 15px; display: flex; align-items: center; gap: 12px;
    }}
    .vertical-section-header {{
        background: linear-gradient(135deg, #0284c7, #0369a1); color: #ffffff !important; padding: 14px 20px; border-radius: 10px; margin-top: 25px; margin-bottom: 15px; font-size: 20px; font-weight: 900; display: flex; align-items: center; gap: 10px;
    }}
    
    .stButton>button {{
        background: #10b981 !important; color: #ffffff !important; -webkit-text-fill-color: #ffffff !important; border: none !important; border-radius: 10px; font-weight: 900 !important; font-size: 16px !important; padding: 12px 22px; box-shadow: 0 4px 10px rgba(16, 185, 129, 0.3); width: 100% !important;
    }}
    .stButton>button:hover {{ background: #059669 !important; }}

    div.stFormSubmitButton > button, div.row-widget.stButton > button:nth-of-type(1) {{
        background-color: #dc2626 !important;
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
    }}
    div.stFormSubmitButton > button:hover {{
        background-color: #b91c1c !important;
    }}

    .stLinkButton>a {{
        background-color: #dc2626 !important;
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
        font-family: 'Cairo', sans-serif !important;
        font-weight: 900 !important;
        font-size: 16px !important;
        border-radius: 10px !important;
        text-align: center !important;
        display: block !important;
        padding: 10px 15px !important;
        box-shadow: 0 4px 12px rgba(220, 38, 38, 0.3);
        border: none !important;
        text-decoration: none !important;
    }}
    .stLinkButton>a:hover {{
        background-color: #b91c1c !important;
        color: #ffffff !important;
    }}

    .darssly-box {{
        background: linear-gradient(135deg, #059669, #10b981); border-radius: 16px; padding: 24px; margin-bottom: 25px; box-shadow: 0 8px 25px rgba(16, 185, 129, 0.3); border: 2px solid rgba(255, 255, 255, 0.25); color: #ffffff !important;
    }}
    .darssly-box * {{ color: #ffffff !important; }}

    .course-card {{
        background: {card_bg}; border: 2px solid {card_border}; border-radius: 16px; padding: 20px; text-align: center; box-shadow: 0 6px 16px rgba(0,0,0,0.08); margin-bottom: 20px; display: flex; flex-direction: column; justify-content: space-between; height: 100%;
    }}
    .course-icon-box {{
        font-size: 45px; background: #065f46; width: 80px; height: 80px; display: flex; align-items: center; justify-content: center; border-radius: 50%; margin: 0 auto 15px auto; border: 2px solid #10b981;
    }}
    .course-title {{ color: #059669 !important; font-size: 20px !important; font-weight: 900 !important; margin-bottom: 8px !important; }}
    .course-desc {{ color: {text_color} !important; font-size: 14px !important; font-weight: 800 !important; opacity: 0.9; margin-bottom: 15px !important; }}

    .social-top-container {{
        display: flex; gap: 12px; align-items: center; margin-top: 10px; justify-content: center;
    }}
    .social-btn-top {{
        display: inline-flex; align-items: center; justify-content: center; width: 40px; height: 40px; border-radius: 50%; text-decoration: none !important; box-shadow: 0 3px 8px rgba(0,0,0,0.25); transition: transform 0.2s ease;
    }}
    .social-btn-top:hover {{ transform: scale(1.12); }}
    .social-btn-top svg {{ width: 20px; height: 20px; fill: #ffffff; }}

    .facebook-bg {{ background-color: #1877F2; }}
    .whatsapp-bg {{ background-color: #25D366; }}
    .telegram-bg {{ background-color: #229ED9; }}
    .tiktok-bg   {{ background-color: #000000; border: 1px solid #444; }}
    .youtube-bg  {{ background-color: #FF0000; }}

    .subscription-card {{
        background: {card_bg}; border: 2px solid #dbe4ef; border-radius: 22px; padding: 18px 18px 12px; text-align: center; box-shadow: 0 10px 28px rgba(15,23,42,.10); margin-bottom: 18px; overflow: hidden; position: relative;
    }}
    .subscription-card:hover {{ border-color: #10b981; box-shadow: 0 14px 34px rgba(5,150,105,.16); }}
    .subscription-photo {{ width: 118px; height: 118px; border-radius: 50%; object-fit: cover; border: 5px solid #10b981; box-shadow: 0 7px 18px rgba(5,150,105,.22); margin: 2px auto 10px; display:block; }}
    .subscription-icon {{ width: 118px; height: 118px; border-radius: 50%; background: #ecfdf5; display:flex; align-items:center; justify-content:center; font-size:50px; margin: 2px auto 10px; border:5px solid #10b981; }}
    .subscription-title {{ color:#0f172a; font-size:20px; font-weight:900; margin:7px 0 4px; }}
    .subscription-price {{ color:#ea580c; font-size:25px; font-weight:1000; margin:4px 0 10px; }}
    .subscription-features {{ color:{text_color}; font-size:13px; font-weight:800; line-height:1.9; margin-bottom:8px; }}
    .subscription-badge {{ display:inline-block; background:#ecfdf5; color:#047857; border:1px solid #a7f3d0; border-radius:999px; padding:4px 12px; font-size:12px; font-weight:900; margin-bottom:5px; }}

    .social-top-container {{
        display: flex; gap: 12px; align-items: center; margin-top: 10px; justify-content: center;
    }}
    .social-btn-top {{
        display: inline-flex; align-items: center; justify-content: center; width: 40px; height: 40px; border-radius: 50%; text-decoration: none !important; box-shadow: 0 3px 8px rgba(0,0,0,0.25); transition: transform 0.2s ease;
    }}
    .social-btn-top:hover {{ transform: scale(1.12); }}
    .social-btn-top svg {{ width: 20px; height: 20px; fill: #ffffff; }}

    .facebook-bg {{ background-color: #1877F2; }}
    .whatsapp-bg {{ background-color: #25D366; }}
    .telegram-bg {{ background-color: #229ED9; }}
    .tiktok-bg   {{ background-color: #000000; border: 1px solid #444; }}
    .youtube-bg  {{ background-color: #FF0000; }}


    .call-btn-container {{ display: flex; justify-content: center; margin-top: 25px; margin-bottom: 15px; width: 100%; }}
    .call-btn {{
        display: inline-flex; align-items: center; justify-content: center; gap: 12px; background: linear-gradient(135deg, #059669, #10b981); color: #ffffff !important; padding: 14px 28px; border-radius: 50px; font-size: 18px; font-weight: 900; text-decoration: none !important; border: 2px solid #ffffff;
    }}
    .social-footer-box {{
        margin-top: 25px; padding: 20px 0; border-top: 1px solid rgba(150, 150, 150, 0.3); display: flex; flex-direction: column; align-items: center; justify-content: center; width: 100%;
    }}
    .rights-text {{ font-size: 16px; font-weight: 900; margin-top: 12px; text-align: center; color: {text_color}; }}
    </style>
""", unsafe_allow_html=True)

# ==============================================================================
# 1. واجهة الطالب الشاملة
# ==============================================================================
if is_student_mode:
    st.markdown("""
        <style>
            [data-testid="stSidebar"], [data-testid="stSidebarCollapseButton"], [data-testid="collapsedControl"] { display: none !important; }
        </style>
    """, unsafe_allow_html=True)

    nav_avatar_tag = f'<img src="{teacher_image_data_uri(img_b64)}" style="width: 45px; height: 45px; border-radius: 50%; border: 2px solid #10b981; object-fit: cover;">' if img_b64 else '<div style="width:45px;height:45px;border-radius:50%;background:#d1fae5;display:flex;align-items:center;justify-content:center;">👨‍🏫</div>'
    
    col_nav1, col_nav2 = st.columns([2, 1.5])
    with col_nav1:
        st.markdown(f"""
        <div class="darssly-navbar" dir="rtl" style="margin-bottom: 0px;">
            <div class="navbar-brand">
                {nav_avatar_tag}
                <div class="navbar-title-group">
                    <h3>م/ محمد غنيم</h3>
                    <p>منصة شرح الرياضيات والإحصاء</p>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
    with col_nav2:
        st.write("")
        c_btn0, c_btn1, c_btn2, c_btn3, c_btn_guest = st.columns(5)
        with c_btn0:
            if st.button("🏠 الرئيسية"):
                st.session_state.page_view = "home"
                st.session_state.logged_student = None
                st.query_params.clear()
                st.query_params["role"] = "student"
                st.rerun()
        with c_btn1:
            if st.button("👤 دخول"):
                st.session_state.page_view = "login"
                st.rerun()
        with c_btn2:
            if st.button("✨ حساب"):
                st.session_state.page_view = "register"
                st.rerun()
        with c_btn_guest:
            if st.button("👥 ضيف"):
                st.session_state.page_view = "guest_reg"
                st.rerun()
        with c_btn3:
            mode_label = "☀️ فاتح" if st.session_state.dark_mode else "🌙 داكن"
            if st.button(mode_label):
                st.session_state.dark_mode = not st.session_state.dark_mode
                st.rerun()

    # --- أيقونات التواصل أعلى صفحة الطالب (إضافة جديدة بدون حذف الفوتر القديم) ---
    st.markdown(f"""
        <div class="social-top-container" style="justify-content:flex-start; margin-top:0; margin-bottom:12px; padding:8px 12px; background:{card_bg}; border:1px solid {card_border}; border-radius:14px;">
            <a href="https://www.facebook.com/share/19fD41rV3H/" target="_blank" title="Facebook" class="social-btn-top facebook-bg"><svg viewBox="0 0 24 24"><path d="M24 12.073c0-6.627-5.373-12-12-12s-12 5.373-12 12c0 5.99 4.388 10.954 10.125 11.854v-8.385H7.078v-3.47h3.047V9.43c0-3.007 1.792-4.669 4.533-4.669 1.312 0 2.686.235 2.686.235v2.953H15.83c-1.491 0-1.956.925-1.956 1.874v2.25h3.328l-.532 3.47h-2.796v8.385C19.612 23.027 24 18.062 24 12.073z"/></svg></a>
            <a href="https://wa.me/201016361440" target="_blank" title="WhatsApp" class="social-btn-top whatsapp-bg"><svg viewBox="0 0 24 24"><path d="M12.031 6.172c-3.181 0-5.767 2.586-5.768 5.766-.001 1.298.38 2.27 1.019 3.287l-.711 2.599 2.669-.699c.971.53 1.77.822 2.791.823h.002c3.18 0 5.767-2.586 5.768-5.766 0-3.18-2.587-5.766-5.77-5.766zm9.969 5.828c0 5.519-4.481 10-10 10-1.761 0-3.424-.46-4.881-1.267l-5.619 1.474 1.499-5.485c-.911-1.516-1.43-3.285-1.43-5.176 0-5.519 4.481-10 10-10 5.519 0 10 4.481 10 10z"/></svg></a>
            <a href="https://t.me/mrmaths22" target="_blank" title="Telegram" class="social-btn-top telegram-bg"><svg viewBox="0 0 24 24"><path d="M12 0c-6.627 0-12 5.373-12 12s5.373 12 12 12 12-5.373 12-12S18.627 0 12 0zm5.894 8.221l-1.97 9.28c-.145.658-.537.818-1.084.508l-3-2.21-1.446 1.394c-.16.16-.295.295-.605.295l.213-3.053 5.56-5.023c.242-.213-.054-.333-.373-.121l-6.871 4.326-2.962-.924c-.643-.204-.657-.643.136-.953l11.57-4.461c.537-.194 1.006.131.832.942z"/></svg></a>
            <a href="https://www.tiktok.com/@eng_mohamedghonaim?_r=1&_t=ZS-99VdklZPBUS" target="_blank" title="TikTok" class="social-btn-top tiktok-bg"><svg viewBox="0 0 24 24"><path d="M19.589 6.686a4.793 4.793 0 0 1-3.77-4.245V2h-3.445v13.672a2.896 2.896 0 0 1-5.201 1.743l-.068-.102a2.895 2.895 0 0 1 2.373-4.513c.277 0 .546.039.803.111V9.417a6.338 6.338 0 0 0-.803-.051C6.017 9.366 3.2 12.183 3.2 15.647 3.2 19.11 6.017 22 9.479 22c3.462 0 6.279-2.817 6.279-6.353V9.07c1.378.983 3.054 1.564 4.869 1.584V7.209a4.845 4.845 0 0 1-1.038-.523z"/></svg></a>
            <a href="https://youtube.com/@engineermaths?si=8C6T808VuAU5OMOt" target="_blank" title="YouTube" class="social-btn-top youtube-bg"><svg viewBox="0 0 24 24"><path d="M23.498 6.186a3.016 3.016 0 0 0-2.122-2.136C19.505 3.545 12 3.545 12 3.545s-7.505 0-9.377.505A3.017 3.017 0 0 0 .502 6.186C0 8.07 0 12 0 12s0 3.93.502 5.814a3.016 3.016 0 0 0 2.122 2.136c1.871.505 9.376.505 9.376.505s7.505 0 9.377-.505a3.016 3.016 0 0 0 2.122-2.136C24 15.93 24 12 24 12s0-3.93-.502-5.814zM9.545 15.568V8.432L15.818 12l-6.273 3.568z"/></svg></a>
            <span style="font-weight:900; margin-right:8px; color:{text_color};">تابعنا وتواصل معنا</span>
        </div>
    """, unsafe_allow_html=True)

    st.write("---")

    if st.session_state.page_view == "guest_reg":
        st.markdown(f"<h3 style='color: {text_color}; font-size: 24px;'>👥 تسجيل بياناتك كـ (ضيف جديد):</h3>", unsafe_allow_html=True)
        st.markdown(f"<p style='color: {text_color}; font-size: 16px;'>أدخل اسمك، رقم هاتفك، ومنهجك لتظهر في سجلات المعلم فوراً:</p>", unsafe_allow_html=True)
        
        with st.form("guest_registration_form"):
            g_name = st.text_input("اسم الطالب بالكامل:*", placeholder="مثال: كريم أحمد محمود")
            g_phone = st.text_input("رقم الهاتف المحمول:*", placeholder="010XXXXXXXX")
            g_curr = st.selectbox("المنهج الدراسي / الدولة:*", list(CURRICULUM_DATA.keys()))
            g_grade = st.selectbox("المرحلة / الصف الدراسي:*", CURRICULUM_DATA[g_curr])
            g_pass = st.text_input("اختر رقماً سرياً خاصاً بك:*", type="password", value="1234")
            
            c_g1, c_g2 = st.columns(2)
            with c_g1:
                submit_guest = st.form_submit_button("🚀 حفظ والدخول للمنصة")
            with c_g2:
                if st.form_submit_button("العودة للرئيسية"):
                    st.session_state.page_view = "home"
                    st.rerun()

            if submit_guest:
                if not g_name.strip() or not g_phone.strip():
                    st.error("يرجى كتابة الاسم ورقم الهاتف على الأقل.")
                else:
                    new_guest_user = {
                        "اسم الطالب": g_name.strip(),
                        "رقم الهاتف": g_phone.strip(),
                        "كلمة المرور": g_pass.strip(),
                        "المنهج/الدولة": g_curr,
                        "المجموعة/الصف": g_grade,
                        "تاريخ التسجيل": str(date.today()),
                        "الحالة_حظر": "نشط",
                        "حالة_الاشتراك_البنك": "غير مشترك"
                    }
                    existing = st.session_state.users_df[st.session_state.users_df["رقم الهاتف"].astype(str).str.strip() == g_phone.strip()]
                    if existing.empty:
                        st.session_state.users_df = pd.concat([st.session_state.users_df, pd.DataFrame([new_guest_user])], ignore_index=True)
                    else:
                        st.session_state.users_df.loc[st.session_state.users_df["رقم الهاتف"].astype(str).str.strip() == g_phone.strip(), ["اسم الطالب", "المنهج/الدولة", "المجموعة/الصف"]] = [g_name.strip(), g_curr, g_grade]
                    
                    save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                    
                    st.session_state.logged_student = new_guest_user
                    st.query_params["role"] = "student"
                    st.query_params["st_phone"] = g_phone.strip()
                    st.success(f"أهلاً بك يا {g_name.strip()}! تم تسجيل بياناتك بنجاح في سجلات المعلم.")
                    st.rerun()

    elif not st.session_state.logged_student:
        if st.session_state.page_view == "home":
            _ui_df = st.session_state.get("student_interface_df", pd.DataFrame())
            si = _ui_df.iloc[0].to_dict() if not _ui_df.empty else {}
            ui_title = str(si.get("عنوان_الواجهة", "أهلاً بيكم منورين المنصة! 🚀"))
            ui_badge = str(si.get("الشارة", "منصة شرح الرياضيات والإحصاء"))
            ui_desc = str(si.get("الوصف", ""))
            ui_main_b64 = STUDENT_FIXED_IMAGE_B64
            ui_main_uri = teacher_image_data_uri(ui_main_b64) if ui_main_b64 else ""
            ui_booking_title = str(si.get("عنوان_الحجز", "📅 حجز دروس أونلاين مباشرة مع م / محمد غنيم"))
            ui_booking_text = str(si.get("نص_الحجز", ""))
            col_hero_txt, col_hero_img = st.columns([1.3, 1])
            with col_hero_txt:
                st.markdown(f"<h1 style='color: #059669; font-size: 38px; font-weight: 900; margin-bottom: 10px;'>{ui_title}</h1>", unsafe_allow_html=True)
                st.markdown(f"<div style='background: #059669; color: #ffffff; padding: 6px 16px; border-radius: 20px; display: inline-block; font-size: 15px; font-weight: 900; margin-bottom: 15px;'>{ui_badge}</div>", unsafe_allow_html=True)
                st.markdown(f"<p style='font-size: 17px; font-weight: 800; line-height: 1.8; color: {text_color};'>{ui_desc}</p>", unsafe_allow_html=True)
                
                c_home_b1, c_home_b2 = st.columns(2)
                with c_home_b1:
                    if st.button("🔐 تسجيل الدخول الآن"):
                        st.session_state.page_view = "login"
                        st.rerun()
                with c_home_b2:
                    if st.button("✨ إنشاء حساب جديد"):
                        st.session_state.page_view = "register"
                        st.rerun()

            with col_hero_img:
                if ui_main_uri:
                    st.markdown(f"""
                        <div style="display: flex; justify-content: center; align-items: center; position: relative; margin-top: 10px;">
                            <img src="{ui_main_uri}" style="width: 230px; height: 230px; border-radius: 50%; border: 5px solid #059669; object-fit: cover; box-shadow: 0 10px 25px rgba(0,0,0,0.15);">
                        </div>
                    """, unsafe_allow_html=True)

            st.write("---")
            # --- اشتراكات درسلي الجديدة (إضافة فقط، مع الإبقاء على القسم القديم) ---
            render_darssly_cards("home")

            # --- حجز الدروس ---
            st.markdown(f"<div class='vertical-section-header'>{ui_booking_title}</div>", unsafe_allow_html=True)
            if ui_booking_text.strip():
                st.markdown(f"<p style='text-align:center;font-weight:800;line-height:1.8;'>{ui_booking_text}</p>", unsafe_allow_html=True)
            with st.form("online_booking_form", clear_on_submit=True):
                book_name = st.text_input("اسم الطالب بالكامل:")
                book_curr = st.selectbox("اختر المنهج الدراسي / الدولة:", list(CURRICULUM_DATA.keys()), key="book_c")
                book_grade = st.selectbox("المرحلة / الصف الدراسي:", CURRICULUM_DATA[book_curr], key="book_g")
                book_phone = st.text_input("رقم هاتف الطالب:")
                book_parent_phone = st.text_input("رقم تليفون ولي الأمر:")
                
                if st.form_submit_button("🚀 إرسال طلب حجز الدرس أونلاين"):
                    if not book_name.strip() or not book_phone.strip():
                        st.error("يرجى كتابة اسم الطالب ورقم الهاتف على الأقل.")
                    else:
                        new_booking = {
                            "تاريخ_الحجز": str(date.today()),
                            "اسم الطالب": book_name.strip(),
                            "المنهج_الدولة": book_curr,
                            "المرحلة_الصف": book_grade,
                            "رقم_الهاتف": book_phone.strip(),
                            "رقم_ولي_الأمر": book_parent_phone.strip(),
                            "الحالة": "قيد المتابعة"
                        }
                        st.session_state.bookings_df = pd.concat([st.session_state.bookings_df, pd.DataFrame([new_booking])], ignore_index=True)
                        save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                        st.success("✓ تم إرسال طلب الحجز بنجاح!")

        elif st.session_state.page_view == "login":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>🔐 تسجيل دخول الطالب (برقم الهاتف):</h3>", unsafe_allow_html=True)
            with st.form("student_login_form"):
                login_phone = st.text_input("رقم الهاتف المحمول المسجل:")
                login_pass = st.text_input("الرقم السري الخاص بك:", type="password")
                c_l1, c_l2 = st.columns(2)
                with c_l1:
                    submit_login = st.form_submit_button("دخول إلى حسابي")
                with c_l2:
                    if st.form_submit_button("العودة للرئيسية"):
                        st.session_state.page_view = "home"
                        st.rerun()

                if submit_login:
                    users_match = st.session_state.users_df[
                        (st.session_state.users_df["رقم الهاتف"].astype(str).str.strip() == login_phone.strip())
                        & (st.session_state.users_df["كلمة المرور"].astype(str).str.strip() == login_pass.strip())
                    ]
                    if not users_match.empty:
                        user_info = users_match.iloc[0].to_dict()
                        if user_info.get("الحالة_حظر") == "محظور":
                            st.error("🚫 تم حظر هذا الحساب من قبل المعلم.")
                        else:
                            st.session_state.logged_student = user_info
                            st.query_params["role"] = "student"
                            st.query_params["st_phone"] = user_info["رقم الهاتف"]
                            st.success(f"مرحباً بك مجدداً يا {user_info['اسم الطالب']}!")
                            st.rerun()
                    else:
                        st.error("رقم الهاتف أو الرقم السري غير صحيح.")

        elif st.session_state.page_view == "register":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>✨ إنشاء حساب طالب جديد:</h3>", unsafe_allow_html=True)
            with st.form("student_register_form"):
                reg_name = st.text_input("اسمك بالكامل:")
                reg_phone = st.text_input("رقم الهاتف المحمول (لتسجيل الدخول به لاحقاً):*")
                reg_curr = st.selectbox("المنهج الدراسي / الدولة:", list(CURRICULUM_DATA.keys()))
                reg_grade = st.selectbox("المرحلة / الصف الدراسي:", CURRICULUM_DATA[reg_curr])
                reg_pass = st.text_input("اختر رقماً سرياً خاصاً بك:", type="password")
                c_r1, c_r2 = st.columns(2)
                with c_r1:
                    submit_reg = st.form_submit_button("تأكيد إنشاء الحساب")
                with c_r2:
                    if st.form_submit_button("العودة للرئيسية"):
                        st.session_state.page_view = "home"
                        st.rerun()

                if submit_reg:
                    if not reg_name.strip() or not reg_phone.strip() or not reg_pass.strip():
                        st.error("يرجى ملء جميع الحقول المطلوبة.")
                    else:
                        existing = st.session_state.users_df[st.session_state.users_df["رقم الهاتف"].astype(str).str.strip() == reg_phone.strip()]
                        if not existing.empty:
                            st.warning("رقم الهاتف هذا مسجل بالفعل! يرجى تسجيل الدخول مباشرة برقم هاتفك.")
                        else:
                            new_user = {
                                "اسم الطالب": reg_name.strip(),
                                "رقم الهاتف": reg_phone.strip(),
                                "كلمة المرور": reg_pass.strip(),
                                "المنهج/الدولة": reg_curr,
                                "المجموعة/الصف": reg_grade,
                                "تاريخ التسجيل": str(date.today()),
                                "الحالة_حظر": "نشط",
                                "حالة_الاشتراك_البنك": "غير مشترك"
                            }
                            st.session_state.users_df = pd.concat([st.session_state.users_df, pd.DataFrame([new_user])], ignore_index=True)
                            save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                            st.session_state.logged_student = new_user
                            st.query_params["role"] = "student"
                            st.query_params["st_phone"] = new_user["رقم الهاتف"]
                            st.success(f"تم إنشاء حسابك بنجاح يا {reg_name}!")
                            st.rerun()

        st.markdown("<div class='vertical-section-header'>🎥 حصص سريعة وملخصات هامة في أقل من دقيقة</div>", unsafe_allow_html=True)
        col_vid1, col_vid2, col_vid3 = st.columns([1, 2, 1])
        with col_vid2:
            st.video("https://www.youtube.com/watch?v=6PleAxZCNZM")

    else:
        st_user = st.session_state.logged_student
        
        # فلترة ذكية ومرنة لمرحلة الطالب بصرف النظر عن أي فروق بسيطة في الأحرف أو الأقواس
        user_grade_raw = str(st_user.get("المجموعة/الصف", "")).strip()
        user_grade_clean = user_grade_raw.split("(")[0].strip().lower()

        fresh_user = st.session_state.users_df[st.session_state.users_df["رقم الهاتف"].astype(str).str.strip() == str(st_user.get("رقم الهاتف", "")).strip()]
        if not fresh_user.empty and fresh_user.iloc[0].get("الحالة_حظر") == "محظور":
            st.error("🚫 عذراً، تم حظر حسابك.")
            st.session_state.logged_student = None
            st.query_params.clear()
            st.query_params["role"] = "student"
            st.rerun()

        col_u1, col_u2 = st.columns([4, 1])
        with col_u1:
            st.markdown(f"""
                <div style="background: {card_bg}; padding: 14px 20px; border-radius: 12px; border-right: 5px solid #10b981; margin-bottom: 20px; border: 1px solid {card_border};">
                    <h3 style="margin: 0; color: #059669; font-size: 20px;">أهلاً بك: {st_user['اسم الطالب']} 🌟</h3>
                    <p style="margin: 4px 0 10px 0; font-weight: 900; color: {text_color}; font-size: 16px;">{st_user.get('المنهج/الدولة', '')} | {st_user.get('المجموعة/الصف', '')} | الهاتف: {st_user.get('رقم الهاتف', '')}</p>
                </div>
            """, unsafe_allow_html=True)
        with col_u2:
            st.write("")
            if st.button("🚪 خروج"):
                st.session_state.logged_student = None
                st.session_state.page_view = "home"
                st.query_params.clear()
                st.query_params["role"] = "student"
                st.rerun()

        # --- جدول الأونلاين وزوم مع التايمر ---
        st.markdown("<div class='vertical-section-header'>💻 حصص الأونلاين وجدول زوم الخاص بي</div>", unsafe_allow_html=True)
        student_name_str = str(st_user.get("اسم الطالب", "")).strip()
        os_df = st.session_state.online_schedule_df
        st_sched = os_df[os_df["اسم الطالب"].astype(str).str.strip().str.lower() == student_name_str.lower()]

        if st_sched.empty:
            st.info("لا توجد حصص أونلاين مسجلة في الجدول المخصص لك حالياً.")
        else:
            for _, s_row in st_sched.iterrows():
                academy_name = s_row.get("اسم الأكاديمية", "أكاديمية البشمهندس")
                sched_grade = s_row["المجموعة/الصف"]
                sched_sup_phone = s_row["رقم مشرف الأكاديمية"]
                sched_price = s_row["سعر الحصة"]
                sched_date = s_row.get("تاريخ الحصة", str(date.today()))
                sched_time = s_row.get("ساعة الحصة", "18:00")
                zoom_link = s_row["رابط زوم"]
                zoom_status = s_row["حالة فتح الحصة"]

                try:
                    target_dt = datetime.strptime(f"{sched_date} {sched_time}", "%Y-%m-%d %H:%M")
                    now_dt = datetime.now()
                    diff_seconds = int((target_dt - now_dt).total_seconds())
                except Exception:
                    diff_seconds = -1

                timer_text = ""
                if diff_seconds > 0:
                    d_days = diff_seconds // 86400
                    d_hours = (diff_seconds % 86400) // 3600
                    d_mins = (diff_seconds % 3600) // 60
                    d_secs = diff_seconds % 60
                    if d_days > 0:
                        timer_text = f"⏳ باقي على الحصة: {d_days} يوم و {d_hours} ساعة و {d_mins} دقيقة"
                    else:
                        timer_text = f"⏳ باقي على الحصة: {d_hours:02d}:{d_mins:02d}:{d_secs:02d}"
                elif diff_seconds == 0:
                    timer_text = "🟢 وقت الحصة الآن!"
                else:
                    timer_text = "📌 موعد الحصة قد حان أو انتهى."

                st.markdown(f"""
                    <div style="background:{card_bg}; border:2px solid #0284c7; border-radius:14px; padding:20px; margin-bottom:15px;">
                        <h4 style="color:#0284c7; margin-top:0;">🏛️ الأكاديمية: {academy_name} | المرحلة: {sched_grade}</h4>
                        <p style="font-size:16px; margin:4px 0;"><b>📅 موعد الحصة:</b> {sched_date} في تمام الساعة {sched_time}</p>
                        <p style="font-size:16px; margin:4px 0; color:#d97706;"><b>{timer_text}</b></p>
                        <p style="font-size:16px; margin:4px 0;"><b>💰 سعر الحصة:</b> {sched_price} جنيه</p>
                        <p style="font-size:16px; margin:4px 0;"><b>📞 رقم مشرف الأكاديمية:</b> {sched_sup_phone}</p>
                        <p style="font-size:16px; margin:4px 0;"><b>حالة الحصة الحالية:</b> <span style="color: {'#16a34a' if zoom_status == 'مفتوحة' else '#dc2626'};"><b>{zoom_status}</b></span></p>
                    </div>
                """, unsafe_allow_html=True)

                if zoom_status == "مفتوحة":
                    if zoom_link and zoom_link != "nan":
                        st.link_button("🚀 انضم الآن إلى حصة زوم (الحصة مفتوحة) 🟢", zoom_link, use_container_width=True)
                    else:
                        st.link_button("🚀 انضم الآن إلى حصة زوم 🟢", "https://us05web.zoom.us/j/83526892910?pwd=2jWRgATgBRPbXttdnm0QpLwBApsZL4.1", use_container_width=True)
                else:
                    st.warning("⏳ الحصة مغلقة حالياً. سيتم فتحها من قبل المعلم في موعدها المحدد.")

        # --- موعد الطالب الأسبوعي المرتبط بملف الطالب ---
        student_weekly = st.session_state.weekly_schedule_df[st.session_state.weekly_schedule_df["اسم الطالب"].astype(str).str.strip().str.lower() == student_name_str.lower()]
        if not student_weekly.empty:
            st.markdown("<div class='vertical-section-header'>🗓️ مواعيدي الأسبوعية</div>", unsafe_allow_html=True)
            for _, wr in student_weekly.iterrows():
                w_color = str(wr.get("اللون", "#2563eb"))
                st.markdown(f"""<div style='border-right:6px solid {w_color};background:{card_bg};border:1px solid {card_border};border-radius:12px;padding:14px;margin-bottom:10px;'>
                <b>📅 {wr.get('اليوم','')} — ⏰ {wr.get('الموعد','')}</b><br>
                🏛️ الأكاديمية: {wr.get('اسم الأكاديمية','')} | 📚 {wr.get('المنهج/الدولة','')} — {wr.get('المجموعة/الصف','')}<br>
                💰 سعر الحصة: {wr.get('سعر الحصة',0)} جنيه | 📞 مشرف الأكاديمية: {wr.get('رقم مشرف الأكاديمية','')}
                </div>""", unsafe_allow_html=True)

        # --- اشتراكات درسلي داخل منصة الطالب ---
        # تصميم بطاقات مختصر وواضح، مع الأسعار المحددة لكل مرحلة وروابط الاشتراك الحالية.
        darssly_subscriptions = [
            {
                "badge": "باقة شهرية",
                "title": "باقة أولى إعدادي",
                "grade": "الصف الأول الإعدادي",
                "price": "200",
                "icon": "📘",
                "accent": "#f59e0b",
                "link": "https://darssly.com/courses/mathematics-for-preparatory-stage-mr-mohamed-ghonaim-3/plans",
            },
            {
                "badge": "باقة شهرية",
                "title": "باقة ثانية إعدادي",
                "grade": "الصف الثاني الإعدادي",
                "price": "200",
                "icon": "📗",
                "accent": "#10b981",
                "link": "https://darssly.com/courses/mathematics-for-preparatory-stage-mr-mohamed-ghonaim/plans",
            },
            {
                "badge": "باقة شهرية",
                "title": "باقة ثالثة إعدادي",
                "grade": "الصف الثالث الإعدادي",
                "price": "200",
                "icon": "📕",
                "accent": "#8b5cf6",
                "link": "https://darssly.com/courses/mathematics-for-preparatory-stage-mr-mohamed-ghonaim-2/plans",
            },
            {
                "badge": "باقة شهرية",
                "title": "إحصاء ثالثة ثانوي",
                "grade": "الصف الثالث الثانوي — إحصاء",
                "price": "250",
                "icon": "📊",
                "accent": "#ef4444",
                "link": "https://darssly.com/courses/mohamed-ghoneim-statistics/plans",
            },
        ]

        # إظهار الباقات فقط إذا كانت مرحلة الطالب من الباقات المحددة، مع إظهار الكل عند عدم وجود تطابق.
        grade_for_package = str(user_grade_raw or user_grade_clean or "").strip().lower()
        grade_aliases = {
            "باقة أولى إعدادي": ["الأول الإعدادي", "اول اعدادي", "أولى إعدادي", "اولى اعدادي", "1 اعدادي", "الأول اعدادي"],
            "باقة ثانية إعدادي": ["الثاني الإعدادي", "ثاني اعدادي", "ثانية إعدادي", "ثانيه اعدادي", "2 اعدادي", "الثاني اعدادي"],
            "باقة ثالثة إعدادي": ["الثالث الإعدادي", "ثالث اعدادي", "ثالثة إعدادي", "ثالثه اعدادي", "3 اعدادي", "الثالث اعدادي"],
            "إحصاء ثالثة ثانوي": ["ثالثة ثانوي", "ثالث ثانوي", "الثالث الثانوي", "إحصاء", "احصاء"],
        }
        matched_packages = []
        for pkg in darssly_subscriptions:
            aliases = [str(x).lower() for x in grade_aliases.get(pkg["title"], [])]
            if any(a in grade_for_package for a in aliases) or any(grade_for_package in a for a in aliases if grade_for_package):
                matched_packages.append(pkg)
        if not matched_packages:
            matched_packages = darssly_subscriptions

        st.markdown("<div class='vertical-section-header'>💳 اشتراكات درسلي</div>", unsafe_allow_html=True)
        st.markdown(f"""
        <div style="text-align:center; margin: -4px 0 18px; color:{text_color}; font-weight:800; font-size:16px;">
            اشترك في باقتك الشهرية واستمتع بنظام شرح ومتابعة متكامل
        </div>
        """, unsafe_allow_html=True)

        package_cols = st.columns(len(matched_packages))
        for p_idx, pkg in enumerate(matched_packages):
            with package_cols[p_idx]:
                st.markdown(f"""
                <div style="background:{card_bg}; border:1.5px solid {pkg['accent']}; border-radius:20px; padding:18px 16px 14px; min-height:365px; box-shadow:0 8px 24px rgba(15,23,42,.08); direction:rtl; text-align:right; margin-bottom:10px;">
                    <div style="display:inline-block; background:{pkg['accent']}18; color:{pkg['accent']}; border:1px solid {pkg['accent']}55; border-radius:20px; padding:5px 12px; font-size:13px; font-weight:900;">{pkg['badge']}</div>
                    {f"<img src='{teacher_image_data_uri(STUDENT_FIXED_IMAGE_B64)}' style='width:108px;height:108px;border-radius:50%;object-fit:cover;display:block;margin:14px auto 10px;border:5px solid {pkg['accent']};box-shadow:0 8px 20px rgba(15,23,42,.18);'>" if img_b64 else f"<div style='font-size:46px; text-align:center; margin:14px 0 8px;'>{pkg['icon']}</div>"}
                    <h3 style="color:{text_color}; text-align:center; font-size:20px; margin:4px 0 6px;">{pkg['title']}</h3>
                    <p style="color:{text_color}; opacity:.82; text-align:center; font-weight:800; font-size:14px; margin-bottom:16px;">{pkg['grade']}</p>
                    <div style="font-size:30px; font-weight:950; color:{pkg['accent']}; text-align:center; margin-bottom:12px;">{pkg['price']} جنيه <span style="font-size:13px; color:{text_color};">/ شهر</span></div>
                    <div style="background:{pkg['accent']}0d; border-radius:14px; padding:10px 12px; color:{text_color}; font-size:13px; line-height:1.9; font-weight:700;">
                        ✓ فيديوهات شرح مسجلة<br>
                        ✓ حصص Zoom مباشرة<br>
                        ✓ متابعة مستمرة<br>
                        ✓ حل وتدريب على الأسئلة
                    </div>
                </div>
                """, unsafe_allow_html=True)
                st.link_button("🔴 معرفة تفاصيل الباقة", pkg["link"], use_container_width=True)

        st.write("")

        st.markdown("<div class='vertical-section-header'>🗂️ لوحة خدمات الطالب التفاعلية</div>", unsafe_allow_html=True)
        
        if st.button("🎥 الفيديوهات والشروحات التعليمية", use_container_width=True):
            st.session_state.student_sub_page = "videos"
            st.rerun()
        if st.button("📚 بنك الأسئلة الشامل (الاشتراك والدفع)", use_container_width=True):
            st.session_state.student_sub_page = "bank"
            st.rerun()
        if st.button("🧠 اختبارات ونتائج موقع عبقري 💡", use_container_width=True):
            st.session_state.student_sub_page = "abqary"
            st.rerun()
        if st.button("✍️ الاختبارات الإلكترونية التفاعلية", use_container_width=True):
            st.session_state.student_sub_page = "exams"
            st.rerun()
        if st.button("📝 تسجيل حضور حصة اليوم", use_container_width=True):
            st.session_state.student_sub_page = "attendance"
            st.rerun()
        if st.button("📊 متابعة درجات الواجبات", use_container_width=True):
            st.session_state.student_sub_page = "hw_grades"
            st.rerun()
        if st.button("📈 متابعة درجات الاختبارات", use_container_width=True):
            st.session_state.student_sub_page = "exam_grades"
            st.rerun()
        if st.button("💬 مركز الدردشة والدعم المباشر", use_container_width=True):
            st.session_state.student_sub_page = "chat"
            st.rerun()

        sub_page = st.session_state.student_sub_page
        if sub_page != "dashboard":
            if st.button("⬅️ رجوع إلى خدمات الطالب", key="student_subpage_back", use_container_width=True):
                st.session_state.student_sub_page = "dashboard"
                st.rerun()
        st.write("---")

        # --- قسم اختبارات عبقري مع فلترة مرنة ومطابقة شاملة ---
        if sub_page == "abqary":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>🧠 اختبارات ونتائج موقع عبقري</h3>", unsafe_allow_html=True)
            abq_df = st.session_state.abqary_df
            
            # فلترة مرنة جداً تتأكد من ظهور الامتحان بغض النظر عن الأقواس أو التشكيل
            st_abq = abq_df[abq_df["المجموعة/الصف"].astype(str).str.strip().str.lower().apply(
                lambda x: user_grade_clean in x or user_grade_raw.lower() in x or x in user_grade_clean
            )]
            if st_abq.empty:
                st_abq = abq_df.copy() # لو مفيش تطابق حرفي تام، اعرض الكل عشان الطالب ميتعطلش

            if st_abq.empty:
                st.info("لا توجد اختبارات منشورة عبر موقع عبقري لمرحلتك حالياً.")
            else:
                for ab_i, ab_row in st_abq.iterrows():
                    ab_title = ab_row["عنوان_الإمتحان"]
                    ab_link = ab_row.get("رابط_الإمتحان", "")
                    ab_html = str(ab_row.get("كود_HTML", "") or "").strip()
                    res_link = ab_row.get("رابط_النتيجة", "")
                    secret_code = str(ab_row.get("الرقم_السري_للنتيجة", "")).strip()

                    st.markdown(f"""
                        <div style="background:{card_bg}; border:2px solid #059669; border-radius:14px; padding:20px; margin-bottom:15px;">
                            <h4 style="color:#059669; margin-top:0;">💡 امتحان عبقري: {ab_title}</h4>
                            <p style="font-size:15px;">إليك امتحان موقع عبقري معروض بالكامل داخل المنصة:</p>
                        </div>
                    """, unsafe_allow_html=True)

                    # إذا كان المعلم قد أضاف كود HTML، يظهر الاختبار تفاعلياً داخل المنصة.
                    if ab_html and ab_html.lower() != "nan":
                        st.markdown("### 📝 الاختبار الإلكتروني")
                        st.components.v1.html(ab_html, height=800, scrolling=True)
                        st.write("")

                    # الرابط القديم يظل موجوداً كما هو كخيار بديل.
                    if ab_link and str(ab_link).lower() != "nan" and str(ab_link).strip() != "":
                        st.components.v1.iframe(str(ab_link).strip(), height=650, scrolling=True)
                        st.write("")
                        st.link_button(f"🔗 فتح امتحان عبقري في نافذة جديدة 🚀", str(ab_link).strip(), use_container_width=True)

                    st.write("")
                    st.markdown("##### 🔍 استعلام عن نتيجة هذا الاختبار برقم سري:")
                    with st.form(f"abqary_result_form_{ab_i}"):
                        entered_pass = st.text_input("أدخل الرقم السري المخصص لإظهار النتيجة:", type="password", key=f"pass_input_{ab_i}")
                        if st.form_submit_button("🔓 إظهار النتيجة"):
                            if secret_code and entered_pass.strip() == secret_code:
                                st.success("✓ الرقم السري صحيح!")
                                if res_link and res_link != "nan":
                                    st.link_button("📊 اضغط هنا لعرض نتيجة امتحان عبقري 🏆", res_link, use_container_width=True)
                                else:
                                    st.info("النتيجة متاحة وسيتم إضافتها قريباً.")
                            else:
                                st.error("❌ الرقم السري غير صحيح.")
                    st.write("---")

        # --- الفيديوهات بتصميم درسلي (Sidebar يمين، والفيديو شمال) ---
        elif sub_page == "videos":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>🎥 محتوى الشروحات والفيديوهات التعليمية</h3>", unsafe_allow_html=True)
            v_df = st.session_state.videos_df
            
            # فلترة مرنة للفيديوهات حسب المرحلة
            st_videos = v_df[v_df["المجموعة/الصف"].astype(str).str.strip().str.lower().apply(
                lambda x: user_grade_clean in x or user_grade_raw.lower() in x or x in user_grade_clean
            )]
            if st_videos.empty:
                st_videos = v_df.copy()

            if st_videos.empty:
                st.info("لا توجد فيديوهات مرفوعة لمرحلتك الدراسية حالياً.")
            else:
                if "selected_video_idx" not in st.session_state:
                    st.session_state.selected_video_idx = 0

                col_main_v, col_side_v = st.columns([2.5, 1])

                with col_side_v:
                    st.markdown(f"""
                        <div style="background:{card_bg}; border:1px solid {card_border}; border-radius:12px; padding:15px; margin-bottom:15px;">
                            <h4 style="margin:0 0 10px 0; color:#059669; font-size:18px;">📚 محتوى الدرس</h4>
                        </div>
                    """, unsafe_allow_html=True)

                    search_vid_query = st.text_input("🔍 ابحث في الحصص...", placeholder="اكتب اسم الدرس...")
                    filtered_videos = st_videos
                    if search_vid_query.strip():
                        filtered_videos = st_videos[st_videos["عنوان_الفيديو"].astype(str).str.contains(search_vid_query, case=False, na=False)]

                    if filtered_videos.empty:
                        st.info("لا توجد نتائج مطابقة للبحث.")
                    else:
                        for v_i, v_row in filtered_videos.reset_index(drop=True).iterrows():
                            v_title_btn = f"🟢 {v_row['عنوان_الفيديو']}"
                            if st.button(v_title_btn, key=f"darssly_sidebar_v_{v_i}", use_container_width=True):
                                st.session_state.selected_video_idx = v_i
                                st.rerun()

                with col_main_v:
                    selected_row = filtered_videos.iloc[0] if filtered_videos.empty else (filtered_videos.iloc[st.session_state.selected_video_idx] if st.session_state.selected_video_idx < len(filtered_videos) else filtered_videos.iloc[0])

                    st.markdown(f"""
                        <div style="background:linear-gradient(135deg, #059669, #10b981); color:#ffffff; padding:15px 20px; border-radius:10px; margin-bottom:15px;">
                            <h3 style="margin:0; color:#ffffff; font-size:20px;">📺 {selected_row['عنوان_الفيديو']}</h3>
                        </div>
                    """, unsafe_allow_html=True)

                    v_link = str(selected_row.get("رابط_الفيديو", "")).strip()
                    v_bytes = selected_row.get("فيديو_base64", "")

                    if v_link and v_link != "nan" and v_link != "":
                        st.video(v_link)
                    elif pd.notnull(v_bytes) and str(v_bytes).strip() and str(v_bytes) != "nan":
                        try:
                            vid_bytes_dec = base64.b64decode(v_bytes)
                            st.video(vid_bytes_dec)
                        except Exception:
                            st.error("⚠️ يتعذر تشغيل ملف الفيديو.")
                    else:
                        st.info("لا يوجد فيديو متاح لهذا الدرس.")

                    st.write("---")
                    st.markdown(f"#### 📌 {selected_row['عنوان_الفيديو']}")
                    st.caption(f"المرحلة الدراسية: {selected_row.get('المجموعة/الصف', '')} | أكملت هذه الحصة")

                    st.write("---")
                    current_vid_title = selected_row['عنوان_الفيديو']
                    vc_df = st.session_state.video_comments_df
                    vid_comments = vc_df[vc_df["عنوان_الفيديو"].astype(str).str.strip() == current_vid_title.strip()]

                    st.markdown(f"<h4 style='font-size:18px;'>الأسئلة والتعليقات ({len(vid_comments)})</h4>", unsafe_allow_html=True)
                    if not vid_comments.empty:
                        for _, c_row in vid_comments.iterrows():
                            st.markdown(f"""
                                <div style="background:{card_bg}; border:1px solid {card_border}; border-radius:10px; padding:12px 15px; margin-bottom:10px;">
                                    <p style="margin:0; font-size:14px; color:#0284c7;"><b>{c_row['اسم الطالب']}</b> — <span style="font-size:12px; opacity:0.7;">{c_row['التاريخ_والوقت']}</span></p>
                                    <p style="margin:5px 0 0 0; font-size:16px;">{c_row['نص_التعليق']}</p>
                                </div>
                            """, unsafe_allow_html=True)

                    with st.form(f"comment_form_{selected_row['معرف_الفيديو']}"):
                        user_comment_text = st.text_area("أكتب سؤالك أو تعليقك هنا:", placeholder="اطرح سؤالك على المعلم...")
                        if st.form_submit_button("إرسال التعليق"):
                            if user_comment_text.strip():
                                new_comment = {
                                    "التاريخ_والوقت": datetime.now().strftime("%Y-%m-%d %H:%M"),
                                    "عنوان_الفيديو": current_vid_title.strip(),
                                    "اسم الطالب": st_user["اسم الطالب"],
                                    "نص_التعليق": user_comment_text.strip()
                                }
                                st.session_state.video_comments_df = pd.concat([st.session_state.video_comments_df, pd.DataFrame([new_comment])], ignore_index=True)
                                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                                st.success("✓ تم إرسال تعليقك بنجاح!")
                                st.rerun()

        # --- بنك الأسئلة والاشتراك والدفع عبر انستا باي ---
        elif sub_page == "bank":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>📚 بنك الأسئلة الشامل (مرحلتك الدراسية)</h3>", unsafe_allow_html=True)
            
            curr_user_row = st.session_state.users_df[st.session_state.users_df["رقم الهاتف"].astype(str).str.strip() == str(st_user.get("رقم الهاتف", "")).strip()]
            sub_status = curr_user_row.iloc[0].get("حالة_الاشتراك_البنك", "غير مشترك") if not curr_user_row.empty else "غير مشترك"

            if sub_status != "مشترك":
                st.warning("🔒 عذراً، بنك الأسئلة مغلق ويحتاج إلى اشتراك خاص بمرحلتك الدراسية بقيمة **100 جنيه** فقط.")
                st.markdown(f"""
                    <div style="background:{card_bg}; border:2px solid #dc2626; border-radius:15px; padding:25px; margin-bottom:20px;">
                        <h3 style="color:#dc2626; margin-top:0;">💳 تعليمات الاشتراك في بنك الأسئلة:</h3>
                        <p style="font-size:18px;">1. قم بالدفع بقيمة <b>100 جنيه</b> عبر تطبيق InstaPay باستخدام زر الدفع السريع بالأسفل.</p>
                        <p style="font-size:18px;">2. اكتب رقم هاتفك المحول منه، وكود التحقق (OTP)، وارفق صورة اسكرين (إيصال) الدفع لتأكيد طلبك.</p>
                    </div>
                """, unsafe_allow_html=True)

                with st.form("bank_subscription_form"):
                    pay_phone = st.text_input("رقم الهاتف المحول منه:", placeholder="010XXXXXXXX")
                    pay_otp = st.text_input("كود التحقق الخاص بـ InstaPay (OTP):", placeholder="مثال: 4589")
                    pay_receipt = st.file_uploader("📷 رفع اسكرين (إيصال) الدفع:", type=["jpg", "png", "jpeg"])
                    
                    if st.form_submit_button("📤 إرسال طلب الاشتراك لتأكيد المعلم"):
                        if not pay_phone.strip() or not pay_otp.strip() or pay_receipt is None:
                            st.error("يرجى إدخال رقم الهاتف، وكود الـ OTP، وإرفاق صورة إيصال الدفع.")
                        else:
                            rcpt_str = base64.b64encode(pay_receipt.read()).decode()
                            new_req = {
                                "تاريخ_الطلب": str(date.today()),
                                "اسم الطالب": st_user["اسم الطالب"],
                                "رقم_الهاتف": pay_phone.strip(),
                                "كود_OTP": pay_otp.strip(),
                                "حالة_الدفع": "قيد المراجعة",
                                "إيصال_الدفع_base64": rcpt_str
                            }
                            st.session_state.bank_requests_df = pd.concat([st.session_state.bank_requests_df, pd.DataFrame([new_req])], ignore_index=True)
                            save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                            st.success("✓ تم إرسال طلب اشتراكك بنجاح! سيقوم المعلم بمراجعة الإيصال وتفعيل حسابك خلال دقائق.")
                
                st.link_button("📲 اضغط هنا للدفع من خلال انستا باي", "https://ipn.eg/S/moghonem2002/instapay/6EyvZs")
            else:
                st.success("🎉 أهلاً بك! حسابك مفعل ومسجل في بنك الأسئلة الخاص بمرحلتك الدراسية.")
                qb_df = st.session_state.question_bank_df
                st_qb = qb_df[qb_df["المجموعة/الصف"].astype(str).str.strip().str.lower().apply(
                    lambda x: user_grade_clean in x or user_grade_raw.lower() in x or x in user_grade_clean
                )]
                if st_qb.empty: st_qb = qb_df.copy()

                if st_qb.empty:
                    st.info("لا توجد أسئلة مضافة في بنك الأسئلة لمرحلتك حالياً.")
                else:
                    for qb_i, qb_r in st_qb.iterrows():
                        raw_json_data = qb_r.get("بيانات_السؤال_JSON", "{}")
                        try:
                            q_data = json.loads(raw_json_data) if pd.notnull(raw_json_data) and str(raw_json_data).strip() else {}
                        except Exception:
                            q_data = {}

                        if not q_data: continue

                        st.markdown(f"""
                            <div style="background:{card_bg}; border:1px solid {card_border}; border-radius:12px; padding:20px; margin-bottom:15px;">
                                <p style="font-size:18px; color:{text_color};"><b>سؤال ({qb_i+1}) — الدرجة: {q_data.get('points', 1.0)}</b></p>
                                <p style="font-size:17px; color:{text_color};">{q_data.get('text', '')}</p>
                            </div>
                        """, unsafe_allow_html=True)
                        
                        if q_data.get("q_img"):
                            st.image(f"data:image/jpeg;base64,{q_data['q_img']}", use_container_width=True)

                        q_type_val = q_data.get("type", "اختيار من متعدد")
                        if "اختيار" in q_type_val or q_type_val == "موضوعي":
                            opts = ["أ", "ب", "ج", "د"]
                            ans_choice = st.radio(f"اختر الإجابة الصحيحة للسؤال ({qb_i+1}):", opts, key=f"qb_radio_{qb_i}")
                            
                            opt1 = q_data.get('opt1', '')
                            opt2 = q_data.get('opt2', '')
                            opt3 = q_data.get('opt3', '')
                            opt4 = q_data.get('opt4', '')
                            if opt1 or opt2 or opt3 or opt4:
                                st.markdown(f"""
                                    <div style="padding: 10px; background: #f1f5f9; border-radius: 8px; margin-bottom: 10px;">
                                        أ) {opt1} &nbsp;&nbsp;|&nbsp;&nbsp; ب) {opt2} &nbsp;&nbsp;|&nbsp;&nbsp; ج) {opt3} &nbsp;&nbsp;|&nbsp;&nbsp; د) {opt4}
                                    </div>
                                """, unsafe_allow_html=True)

                            if st.button(f"تحقق من إجابة السؤال ({qb_i+1})", key=f"check_qb_{qb_i}"):
                                correct_idx = int(q_data.get("correct", 1))
                                correct_letter = opts[correct_idx - 1] if 0 <= correct_idx - 1 < 4 else 'أ'
                                if ans_choice == correct_letter:
                                    st.success("إجابة صحيحة تماماً! أحسنت ✅")
                                else:
                                    st.error(f"إجابة خاطئة ❌. الإجابة الصحيحة هي: الخيار ({correct_letter})")

        elif sub_page == "exams":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>✍️ الاختبارات الإلكترونية التفاعلية</h3>", unsafe_allow_html=True)
            st.info("لا توجد اختبارات تفاعلية نشطة حالياً.")

        elif sub_page == "attendance":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>📝 تسجيل حضور حصة اليوم</h3>", unsafe_allow_html=True)
            with st.form("att_form"):
                st.date_input("تاريخ الحصة:")
                if st.form_submit_button("تأكيد الحضور"):
                    st.success("تم تسجيل الحضور بنجاح!")

        elif sub_page == "hw_grades":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>📊 متابعة درجات الواجبات المنزلية</h3>", unsafe_allow_html=True)
            my_ass = st.session_state.assessments_df[st.session_state.assessments_df["اسم الطالب"].astype(str).str.strip() == st_user["اسم الطالب"].strip()]
            if my_ass.empty: st.info("لا توجد درجات مرصودة.")
            else: st.dataframe(my_ass, use_container_width=True)

        elif sub_page == "exam_grades":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>📈 متابعة درجات الاختبارات والكويزات</h3>", unsafe_allow_html=True)
            st.info("لا توجد اختبارات سابقة.")

        elif sub_page == "chat":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>💬 مركز الدردشة والدعم المباشر</h3>", unsafe_allow_html=True)
            st.info("تواصل مع البشمهندس مباشرة عبر مركز الدردشة.")

    st.markdown("""
        <div class="call-btn-container">
            <a href="tel:01016361440" class="call-btn">
                <svg viewBox="0 0 24 24" style="width:24px;height:24px;fill:#ffffff;"><path d="M6.62 10.79c1.44 2.83 3.76 5.14 6.59 6.59l2.2-2.2c.27-.27.67-.36 1.02-.24 1.12.37 2.33.57 3.57.57.55 0 1 .45 1 1V20c0 .55-.45 1-1 1-9.39 0-17-7.61-17-17 0-.55.45-1 1-1h3.5c.55 0 1 .45 1 1 0 1.25.2 2.45.57 3.57.11.35.03.74-.25 1.02l-2.2 2.2z"/></svg>
                <span>للتواصل مع م / محمد غنيم: 01016361440</span>
            </a>
        </div>
        <div class="social-footer-box">
            <div class="social-footer-container">
                <a href="https://www.facebook.com/share/19fD41rV3H/" target="_blank" title="Facebook" class="social-btn-top facebook-bg"><svg viewBox="0 0 24 24"><path d="M24 12.073c0-6.627-5.373-12-12-12s-12 5.373-12 12c0 5.99 4.388 10.954 10.125 11.854v-8.385H7.078v-3.47h3.047V9.43c0-3.007 1.792-4.669 4.533-4.669 1.312 0 2.686.235 2.686.235v2.953H15.83c-1.491 0-1.956.925-1.956 1.874v2.25h3.328l-.532 3.47h-2.796v8.385C19.612 23.027 24 18.062 24 12.073z"/></svg></a>
                <a href="https://wa.me/201016361440" target="_blank" title="WhatsApp" class="social-btn-top whatsapp-bg"><svg viewBox="0 0 24 24"><path d="M12.031 6.172c-3.181 0-5.767 2.586-5.768 5.766-.001 1.298.38 2.27 1.019 3.287l-.711 2.599 2.669-.699c.971.53 1.77.822 2.791.823h.002c3.18 0 5.767-2.586 5.768-5.766 0-3.18-2.587-5.766-5.77-5.766zm9.969 5.828c0 5.519-4.481 10-10 10-1.761 0-3.424-.46-4.881-1.267l-5.619 1.474 1.499-5.485c-.911-1.516-1.43-3.285-1.43-5.176 0-5.519 4.481-10 10-10 5.519 0 10 4.481 10 10z"/></svg></a>
                <a href="https://t.me/mrmaths22" target="_blank" title="Telegram" class="social-btn-top telegram-bg"><svg viewBox="0 0 24 24"><path d="M12 0c-6.627 0-12 5.373-12 12s5.373 12 12 12 12-5.373 12-12-5.373-12-12-12zm5.894 8.221l-1.97 9.28c-.145.658-.537.818-1.084.508l-3-2.21-1.446 1.394c-.16.16-.295.295-.605.295l.213-3.053 5.56-5.023c.242-.213-.054-.333-.373-.121l-6.871 4.326-2.962-.924c-.643-.204-.657-.643.136-.953l11.57-4.461c.537-.194 1.006.131.832.942z"/></svg></a>
                <a href="https://www.tiktok.com/@eng_mohamedghonaim?_r=1&_t=ZS-99VdklZPBUS" target="_blank" title="TikTok" class="social-btn-top tiktok-bg"><svg viewBox="0 0 24 24"><path d="M19.589 6.686a4.793 4.793 0 0 1-3.77-4.245V2h-3.445v13.672a2.896 2.896 0 0 1-5.201 1.743l-.068-.102a2.895 2.895 0 0 1 2.373-4.513c.277 0 .546.039.803.111V9.417a6.338 6.338 0 0 0-.803-.051C6.017 9.366 3.2 12.183 3.2 15.647 3.2 19.11 6.017 22 9.479 22c3.462 0 6.279-2.817 6.279-6.353V9.07c1.378.983 3.054 1.564 4.869 1.584V7.209a4.845 4.845 0 0 1-1.038-.523z"/></svg></a>
                <a href="https://youtube.com/@engineermaths?si=8C6T808VuAU5OMOt" target="_blank" title="YouTube" class="social-btn-top youtube-bg"><svg viewBox="0 0 24 24"><path d="M23.498 6.186a3.016 3.016 0 0 0-2.122-2.136C19.505 3.545 12 3.545 12 3.545s-7.505 0-9.377.505A3.017 3.017 0 0 0 .502 6.186C0 8.07 0 12 0 12s0 3.93.502 5.814a3.016 3.016 0 0 0 2.122 2.136c1.871.505 9.376.505 9.376.505s7.505 0 9.377-.505a3.015 3.015 0 0 0 2.122-2.136C24 15.93 24 12 24 12s0-3.93-.502-5.814zM9.545 15.568V8.432L15.818 12l-6.273 3.568z"/></svg></a>
            </div>
            <div class="rights-text">جميع الحقوق محفوظة لدي م / محمد غنيم 2026</div>
         </div>
    """, unsafe_allow_html=True)
    st.stop()


# ==============================================================================
# 2. لوحة تحكم المعلم (الشاملة بجميع الأقسام وزوم الأونلاين والتقارير المالية)
# ==============================================================================
total_exams_count = len(st.session_state.exams_df)
total_students_count = len(st.session_state.users_df)

st.sidebar.markdown(f"""
    <div style="text-align: center; margin-top: 5px; margin-bottom: 20px;">
        <h2 style="margin: 0; color: #059669; font-weight: 900;">م/ محمد غنيم</h2>
        <p style="margin: 4px 0; font-weight: 900; font-size: 14px;">لوحة تحكم المعلم الاحترافية</p>
    </div>
""", unsafe_allow_html=True)

if st.sidebar.button("📊 نظرة عامة (الرئيسية)", use_container_width=True):
    st.session_state.teacher_page = "dashboard"
    st.rerun()
if st.sidebar.button("💻 جدول حصص الأونلاين (Zoom)", use_container_width=True):
    st.session_state.teacher_page = "online_schedule"
    st.rerun()
if st.sidebar.button("🗓️ لوحة مواعيد الطلاب", use_container_width=True):
    st.session_state.teacher_page = "weekly_schedule"
    st.rerun()
if st.sidebar.button("⚙️ صانع الامتحانات", use_container_width=True):
    st.session_state.teacher_page = "exam_maker"
    st.rerun()
if st.sidebar.button("📚 بنك الأسئلة", use_container_width=True):
    st.session_state.teacher_page = "question_bank"
    st.rerun()
if st.sidebar.button("🎥 إدارة الفيديوهات", use_container_width=True):
    st.session_state.teacher_page = "videos"
    st.rerun()
if st.sidebar.button("💡 امتحانات عبقري", use_container_width=True):
    st.session_state.teacher_page = "abqary"
    st.rerun()
if st.sidebar.button("📈 درجات الاختبارات", use_container_width=True):
    st.session_state.teacher_page = "grades"
    st.rerun()
if st.sidebar.button("📝 تصحيح المقالي", use_container_width=True):
    st.session_state.teacher_page = "essays"
    st.rerun()
if st.sidebar.button("💬 الرسائل والدردشة", use_container_width=True):
    st.session_state.teacher_page = "chat"
    st.rerun()
if st.sidebar.button("👥 بطاقات الطلاب", use_container_width=True):
    st.session_state.teacher_page = "students"
    st.rerun()
if st.sidebar.button("📝 رصد حصة جديدة", use_container_width=True):
    st.session_state.teacher_page = "add_session"
    st.rerun()
if st.sidebar.button("📚 رصد واجب يدوي", use_container_width=True):
    st.session_state.teacher_page = "add_hw"
    st.rerun()
if st.sidebar.button("✏️ تعديل السجلات", use_container_width=True):
    st.session_state.teacher_page = "edit_records"
    st.rerun()
if st.sidebar.button("📊 السجلات الشاملة", use_container_width=True):
    st.session_state.teacher_page = "all_records"
    st.rerun()
if st.sidebar.button("🖨️ تقرير ولي الأمر", use_container_width=True):
    st.session_state.teacher_page = "parent_report"
    st.rerun()
if st.sidebar.button("💰 حسابات ومدفوعات الطلاب", use_container_width=True):
    st.session_state.teacher_page = "payments"
    st.rerun()
if st.sidebar.button("🎨 واجهة الطالب", use_container_width=True):
    st.session_state.teacher_page = "student_interface"
    st.rerun()

st.sidebar.write("---")
st.sidebar.code("https://engmohamedghonaim.streamlit.app/?role=student", language="text")

t_page = st.session_state.teacher_page

# زر رجوع موحد يظهر في أعلى أي قائمة/صفحة داخل لوحة المعلم
if t_page != "dashboard":
    _back_col, _ = st.columns([1, 5])
    with _back_col:
        if st.button("⬅️ رجوع", key="global_teacher_back", use_container_width=True):
            st.session_state.teacher_page = "dashboard"
            st.rerun()

if t_page == "student_interface":
    st.subheader("🎨 تصميم واجهة الطالب")
    st.caption("قسم مستقل لتعديل الصور والنصوص التي يراها الطالب في الصفحة الرئيسية، بدون المساس ببيانات الطلاب.")
    sidf = st.session_state.student_interface_df
    si = sidf.iloc[0].to_dict() if not sidf.empty else {}
    with st.container(border=True):
        st.markdown("### 📝 نصوص الواجهة")
        si_title = st.text_input("عنوان الواجهة:", value=str(si.get("عنوان_الواجهة", "أهلاً بيكم منورين المنصة! 🚀")), key="si_title")
        si_badge = st.text_input("الشارة تحت العنوان:", value=str(si.get("الشارة", "منصة شرح الرياضيات والإحصاء")), key="si_badge")
        si_desc = st.text_area("وصف الواجهة:", value=str(si.get("الوصف", "")), height=110, key="si_desc")
        a, b = st.columns(2)
        with a:
            si_sub_title = st.text_input("عنوان قسم الاشتراكات:", value=str(si.get("عنوان_الاشتراكات", "📢 اشتراكات درسلي")), key="si_sub_title")
            si_sub_desc = st.text_input("وصف قسم الاشتراكات:", value=str(si.get("وصف_الاشتراكات", "")), key="si_sub_desc")
        with b:
            si_booking_title = st.text_input("عنوان قسم الحجز:", value=str(si.get("عنوان_الحجز", "📅 حجز دروس أونلاين مباشرة مع م / محمد غنيم")), key="si_booking_title")
            si_booking_text = st.text_input("نص قسم الحجز:", value=str(si.get("نص_الحجز", "")), key="si_booking_text")
        si_footer = st.text_input("نص أسفل الواجهة:", value=str(si.get("نص_الفوتر", "")), key="si_footer")
    with st.container(border=True):
        st.markdown("### 🖼️ صور واجهة الطالب")
        c1, c2 = st.columns(2)
        current_main = str(si.get("صورة_الواجهة_base64", "") or "").strip()
        current_sub = str(si.get("صورة_الاشتراكات_base64", "") or "").strip()
        with c1:
            st.markdown("**الصورة الرئيسية أعلى الصفحة**")
            if current_main and current_main.lower() != "nan": st.image(base64_to_pil(current_main), width=220)
            upload_main = st.file_uploader("رفع صورة الواجهة", type=["png","jpg","jpeg","webp"], key="si_upload_main")
        with c2:
            st.markdown("**صورة اشتراكات درسلي**")
            if current_sub and current_sub.lower() != "nan": st.image(base64_to_pil(current_sub), width=180)
            upload_sub = st.file_uploader("رفع صورة الاشتراكات", type=["png","jpg","jpeg","webp"], key="si_upload_sub")
        main_b64 = base64.b64encode(upload_main.getvalue()).decode("utf-8") if upload_main is not None else (current_main or STUDENT_FIXED_IMAGE_B64)
        sub_b64 = base64.b64encode(upload_sub.getvalue()).decode("utf-8") if upload_sub is not None else (current_sub or STUDENT_FIXED_IMAGE_B64)
        st.markdown("### 👀 معاينة")
        preview_uri = teacher_image_data_uri(main_b64) if main_b64 else ""
        if preview_uri: st.markdown(f"<div style='text-align:center;'><img src='{preview_uri}' style='width:180px;height:180px;border-radius:50%;object-fit:cover;border:5px solid #059669;'></div>", unsafe_allow_html=True)
        st.markdown(f"<h2 style='text-align:center;color:#059669'>{si_title}</h2>", unsafe_allow_html=True)
        st.markdown(f"<div style='text-align:center'><span style='background:#059669;color:white;padding:6px 14px;border-radius:20px;font-weight:900'>{si_badge}</span></div>", unsafe_allow_html=True)
        st.markdown(f"<p style='text-align:center;font-weight:800'>{si_desc}</p>", unsafe_allow_html=True)
    if st.button("💾 حفظ واجهة الطالب", key="save_student_interface", use_container_width=True):
        st.session_state.student_interface_df = pd.DataFrame([{
            "عنوان_الواجهة": si_title.strip(), "الشارة": si_badge.strip(), "الوصف": si_desc.strip(), "صورة_الواجهة_base64": main_b64,
            "عنوان_الاشتراكات": si_sub_title.strip(), "وصف_الاشتراكات": si_sub_desc.strip(), "عنوان_الحجز": si_booking_title.strip(),
            "نص_الحجز": si_booking_text.strip(), "نص_الفوتر": si_footer.strip(), "صورة_الاشتراكات_base64": sub_b64, "صورة_البانر_base64": str(si.get("صورة_البانر_base64", "") or "")
        }], columns=COL_STUDENT_INTERFACE)
        save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
        interface_cloud_ok = _cloud_save_student_interface(st.session_state.student_interface_df)
        if _cloud_storage_enabled() and not interface_cloud_ok:
            st.error("⚠️ تم حفظ الواجهة محلياً، لكن لم يتم حفظها في التخزين الدائم. راجع إعدادات Supabase وجدول student_interface_storage.")
        elif not _cloud_storage_enabled():
            st.warning("⚠️ التخزين الدائم غير مفعّل حالياً؛ الصورة ستظل محفوظة في هذه النسخة فقط حتى يتم إعداد Supabase.")
        else:
            st.success("✓ تم حفظ واجهة الطالب والصور في التخزين الدائم بنجاح")
        st.rerun()

elif t_page == "dashboard":
    dashboard_students = sorted(list(set([str(x).strip() for x in st.session_state.users_df["اسم الطالب"].dropna().unique() if str(x).strip()] + [str(x).strip() for x in st.session_state.weekly_schedule_df["اسم الطالب"].dropna().unique() if str(x).strip()] + [str(x).strip() for x in st.session_state.online_schedule_df["اسم الطالب"].dropna().unique() if str(x).strip()])))
    profile_b64 = str(st.session_state.teacher_profile_df.iloc[0].get("الصورة_base64", "")) if not st.session_state.teacher_profile_df.empty else img_b64
    col_main_top, col_stat_sidebar = st.columns([3, 1])
    
    with col_main_top:
        profile_tag = f'<img src="data:image/png;base64,{profile_b64}" style="width:90px;height:90px;border-radius:50%;object-fit:cover;border:4px solid #10b981;">' if profile_b64 and profile_b64 != "nan" else '<div style="width:90px;height:90px;border-radius:50%;background:#e2e8f0;display:flex;align-items:center;justify-content:center;font-size:40px">👨‍🏫</div>'
        st.markdown(f"""<div style="background:linear-gradient(135deg,#0f766e,#0284c7);padding:28px;border-radius:22px;margin-bottom:20px;display:flex;align-items:center;gap:22px;direction:rtl;box-shadow:0 12px 30px rgba(2,132,199,.18)">{profile_tag}<div><h2 style="color:white!important;margin:0 0 8px">أهلاً، م/ محمد غنيم 👨‍🏫</h2><p style="color:white!important;margin:0">لوحة التحكم الاحترافية لإدارة الطلاب والمواعيد والحصص والواجبات والاختبارات والتقارير.</p></div></div>""", unsafe_allow_html=True)
        # قسم صورة المعلم بشكل منفصل ونظيف؛ لا يظهر مربع رفع الصورة إلا عند طلب تعديله
        with st.container(border=True):
            st.markdown("<div style='text-align:right;font-weight:800;font-size:18px;margin-bottom:8px'>📷 صورة المعلم</div>", unsafe_allow_html=True)
            if profile_b64 and profile_b64 != "nan":
                st.markdown("<div style='text-align:right;color:#64748b;font-size:13px;margin-bottom:8px'>صورتك الحالية محفوظة وتظهر في لوحة التحكم والتقارير.</div>", unsafe_allow_html=True)
            change_photo = st.checkbox("✏️ أريد تغيير الصورة", key="change_teacher_photo")
            if change_photo:
                teacher_photo = st.file_uploader("اختر صورة جديدة", type=["png","jpg","jpeg"], key="teacher_profile_upload", label_visibility="visible")
                if teacher_photo is not None and st.button("💾 حفظ صورتي الجديدة", key="save_teacher_photo", use_container_width=True):
                    b64=base64.b64encode(teacher_photo.getvalue()).decode("utf-8")
                    st.session_state.teacher_profile_df=pd.DataFrame([{"اسم المعلم":"م/ محمد غنيم","الصورة_base64":b64}])
                    save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                    st.success("✓ تم حفظ الصورة الجديدة")
                    st.rerun()

    with col_stat_sidebar:
        st.markdown(f"""
            <div style="background: {card_bg}; padding: 15px; border-radius: 14px; border: 1px solid {card_border}; text-align: center;">
                <p style="margin:0; font-size:14px;">عدد الامتحانات</p>
                <h3 style="color: #059669; margin: 5px 0;">{total_exams_count}</h3>
                <hr style="margin: 8px 0;">
                <p style="margin:0; font-size:14px;">المتقدمين (إجمالي)</p>
                <h3 style="color: #0284c7; margin: 5px 0;">{len(dashboard_students)}</h3>
            </div>
        """, unsafe_allow_html=True)

    total_due_all, total_paid_all, total_balance_all = get_all_financial_totals()
    f1, f2, f3 = st.columns(3)
    f1.metric("💳 إجمالي الحساب المستحق", f"{total_due_all:,.0f} جنيه")
    f2.metric("✅ إجمالي المدفوع", f"{total_paid_all:,.0f} جنيه")
    f3.metric("💰 الرصيد المتبقي", f"{max(total_balance_all, 0):,.0f} جنيه")
    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    d1,d2,d3,d4=st.columns(4)
    d1.metric("👥 الطلاب",len(dashboard_students))
    d2.metric("🗓️ المواعيد الأسبوعية",len(st.session_state.weekly_schedule_df))
    d3.metric("💻 مواعيد Zoom",len(st.session_state.online_schedule_df))
    d4.metric("📝 الحصص المرصودة",len(st.session_state.sessions_df))
    st.markdown("### أدواتي السريعة")
    
    bc1, bc2, bc3 = st.columns(3)
    with bc1:
        st.markdown(f"""
            <div style="background: {card_bg}; border: 1px solid {card_border}; border-radius: 14px; padding: 20px; margin-bottom: 15px; height: 160px; display: flex; flex-direction: column; justify-content: space-between;">
                <div>
                    <h4 style="margin: 0; color: #059669;">جدول حصص الأونلاين</h4>
                    <p style="margin: 5px 0 0 0; font-size: 13px; opacity: 0.8;">إدارة جداول زوم والطلاب</p>
                </div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("ادارة زوم", key="card_btn_zoom"):
            st.session_state.teacher_page = "online_schedule"
            st.rerun()

    with bc2:
        st.markdown(f"""
            <div style="background: {card_bg}; border: 1px solid {card_border}; border-radius: 14px; padding: 20px; margin-bottom: 15px; height: 160px; display: flex; flex-direction: column; justify-content: space-between;">
                <div>
                    <h4 style="margin: 0; color: #059669;">امتحاناتي</h4>
                    <p style="margin: 5px 0 0 0; font-size: 13px; opacity: 0.8;">إنشاء وإدارة الامتحانات</p>
                </div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("ادارة الامتحانات", key="card_btn_ex"):
            st.session_state.teacher_page = "exam_maker"
            st.rerun()

    with bc3:
        st.markdown(f"""
            <div style="background: {card_bg}; border: 1px solid {card_border}; border-radius: 14px; padding: 20px; margin-bottom: 15px; height: 160px; display: flex; flex-direction: column; justify-content: space-between;">
                <div>
                    <h4 style="margin: 0; color: #059669;">امتحانات عبقري</h4>
                    <p style="margin: 5px 0 0 0; font-size: 13px; opacity: 0.8;">ربط اختبارات موقع عبقري</p>
                </div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("ادارة عبقري", key="card_btn_abq"):
            st.session_state.teacher_page = "abqary"
            st.rerun()

    bc4, bc5, bc6 = st.columns(3)
    with bc4:
        st.markdown(f"""
            <div style="background: {card_bg}; border: 1px solid {card_border}; border-radius: 14px; padding: 20px; margin-bottom: 15px; height: 160px; display: flex; flex-direction: column; justify-content: space-between;">
                <div>
                    <h4 style="margin: 0; color: #059669;">سجل الحضور والغياب</h4>
                    <p style="margin: 5px 0 0 0; font-size: 13px; opacity: 0.8;">رصد ومتابعة الحصص</p>
                </div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("رصد الحصص", key="card_btn_ses"):
            st.session_state.teacher_page = "add_session"
            st.rerun()

    with bc5:
        st.markdown(f"""
            <div style="background: {card_bg}; border: 1px solid {card_border}; border-radius: 14px; padding: 20px; margin-bottom: 15px; height: 160px; display: flex; flex-direction: column; justify-content: space-between;">
                <div>
                    <h4 style="margin: 0; color: #059669;">الرسائل والدعم</h4>
                    <p style="margin: 5px 0 0 0; font-size: 13px; opacity: 0.8;">صندوق محادثات الطلاب</p>
                </div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("عرض الرسائل", key="card_btn_chat"):
            st.session_state.teacher_page = "chat"
            st.rerun()

    with bc6:
        st.markdown(f"""
            <div style="background: {card_bg}; border: 1px solid {card_border}; border-radius: 14px; padding: 20px; margin-bottom: 15px; height: 160px; display: flex; flex-direction: column; justify-content: space-between;">
                <div>
                    <h4 style="margin: 0; color: #059669;">تقارير ولي الأمر</h4>
                    <p style="margin: 5px 0 0 0; font-size: 13px; opacity: 0.8;">طباعة التقارير بصيغة PDF</p>
                </div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("عرض التقارير", key="card_btn_rep"):
            st.session_state.teacher_page = "parent_report"
            st.rerun()


    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    with st.container(border=True):
        st.markdown(f"<div style='display:flex;align-items:center;justify-content:space-between;gap:15px;direction:rtl'><div><h3 style='margin:0;color:#059669'>💰 حسابات ومدفوعات الطلاب</h3><p style='margin:5px 0 0 0'>متابعة المستحق والمدفوع والرصيد المتبقي وتسجيل طريقة الدفع وتاريخها.</p></div><div style='font-size:22px;font-weight:900;color:#dc2626'>الرصيد المتبقي: {max(total_balance_all,0):,.0f} جنيه</div></div>", unsafe_allow_html=True)
        if st.button("فتح الحسابات والمدفوعات", key="card_btn_payments"):
            st.session_state.teacher_page = "payments"
            st.rerun()
elif t_page == "payments":
    st.markdown("<div class='vertical-section-header'>💰 حسابات ومدفوعات الطلاب</div>", unsafe_allow_html=True)
    st.caption("سجل المبالغ المستحقة من الحصص، وأكد المدفوعات عند استلامها. كل دفعة تحفظ بتاريخها وطريقة الدفع في سجل مستقل.")

    total_due_all, total_paid_all, total_balance_all = get_all_financial_totals()
    pc1, pc2, pc3 = st.columns(3)
    pc1.metric("💳 إجمالي المستحق", f"{total_due_all:,.0f} جنيه")
    pc2.metric("✅ إجمالي المدفوع", f"{total_paid_all:,.0f} جنيه")
    pc3.metric("💰 إجمالي الرصيد المتبقي", f"{max(total_balance_all, 0):,.0f} جنيه")

    st.markdown("### 📆 ملخص الحساب حسب الشهر")
    _months = set()
    if not st.session_state.sessions_df.empty and "التاريخ" in st.session_state.sessions_df.columns:
        _months.update([m for m in st.session_state.sessions_df["التاريخ"].apply(_month_from_value) if m])
    if not st.session_state.payment_records_df.empty:
        if "الشهر" in st.session_state.payment_records_df.columns:
            _months.update([str(m).strip() for m in st.session_state.payment_records_df["الشهر"].dropna() if str(m).strip()])
        elif "التاريخ" in st.session_state.payment_records_df.columns:
            _months.update([m for m in st.session_state.payment_records_df["التاريخ"].apply(_month_from_value) if m])
    _month_options = sorted(_months, reverse=True) or [date.today().strftime("%Y-%m")]
    selected_fin_month = st.selectbox("اختر الشهر:", _month_options, key="financial_month_filter")
    mdue, mpaid, mbal = get_monthly_financial_totals(selected_fin_month)
    mc1, mc2, mc3 = st.columns(3)
    mc1.metric("مستحق الشهر", f"{mdue:,.0f} جنيه")
    mc2.metric("مدفوع الشهر", f"{mpaid:,.0f} جنيه")
    mc3.metric("متبقي الشهر", f"{max(mbal,0):,.0f} جنيه")

    payment_students = sorted(list(set(
        [str(x).strip() for x in st.session_state.users_df["اسم الطالب"].dropna().unique() if str(x).strip()] +
        [str(x).strip() for x in st.session_state.sessions_df["اسم الطالب"].dropna().unique() if str(x).strip()] +
        [str(x).strip() for x in st.session_state.weekly_schedule_df["اسم الطالب"].dropna().unique() if str(x).strip()]
    )))

    if not payment_students:
        st.info("لا يوجد طلاب مسجلون أو حصص مرصودة حتى الآن.")
    else:
        st.markdown("### 👤 تسجيل دفعة جديدة")
        pay_c1, pay_c2 = st.columns(2)
        with pay_c1:
            selected_pay_student = st.selectbox("اختر الطالب:", payment_students, key="payment_student_select")
            due_now, paid_now, balance_now = get_student_financials(selected_pay_student)
            st.markdown(f"<div style='background:{card_bg};border:2px solid #10b981;border-radius:14px;padding:14px;text-align:center'><b>المستحق: {due_now:,.0f} جنيه</b> &nbsp; | &nbsp; <b style='color:#059669'>المدفوع: {paid_now:,.0f} جنيه</b> &nbsp; | &nbsp; <b style='color:#dc2626'>المتبقي: {max(balance_now,0):,.0f} جنيه</b></div>", unsafe_allow_html=True)
        with pay_c2:
            pay_date = st.date_input("تاريخ الدفع:", value=date.today(), key="payment_date")
            payment_month = pay_date.strftime("%Y-%m")
            st.caption(f"📆 شهر الدفعة: {payment_month}")
            pay_method = st.selectbox("طريقة الدفع:", ["محفظة كاش", "InstaPay"], key="payment_method")

        max_pay = max(float(balance_now), 0.0)
        pay_amount = st.number_input("المبلغ المدفوع (جنيه):", min_value=0.0, max_value=max_pay if max_pay > 0 else 1000000.0, value=0.0, step=50.0, key="payment_amount")
        pay_note = st.text_input("ملاحظات الدفع (اختياري):", key="payment_note")

        if st.button("✅ تأكيد استلام الدفع", key="confirm_payment", use_container_width=True):
            if pay_amount <= 0:
                st.error("اكتب قيمة الدفع أولاً.")
            elif balance_now <= 0:
                st.warning("الطالب لا يوجد عليه رصيد مستحق حالياً.")
            else:
                new_payment = {
                    "التاريخ": str(pay_date),
                    "الشهر": pay_date.strftime("%Y-%m"),
                    "اسم الطالب": selected_pay_student,
                    "المبلغ": float(pay_amount),
                    "طريقة الدفع": pay_method,
                    "حالة الدفع": "مؤكد",
                    "ملاحظات": pay_note.strip(),
                }
                st.session_state.payment_records_df = pd.concat([st.session_state.payment_records_df, pd.DataFrame([new_payment])], ignore_index=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                st.success(f"✓ تم تأكيد دفع {pay_amount:,.0f} جنيه للطالب {selected_pay_student} عن طريق {pay_method}.")
                st.rerun()

    st.write("---")
    st.markdown("### 📊 رصيد كل طالب")
    financial_rows = []
    for nm in payment_students:
        due, paid, balance = get_student_financials(nm)
        urow = st.session_state.users_df[st.session_state.users_df["اسم الطالب"].astype(str).str.strip() == nm]
        grade = str(urow.iloc[0].get("المجموعة/الصف", "")) if not urow.empty else ""
        sess_count = int(len(st.session_state.sessions_df[st.session_state.sessions_df["اسم الطالب"].astype(str).str.strip() == nm]))
        financial_rows.append({"اسم الطالب": nm, "المرحلة": grade, "إجمالي الحصص": sess_count, "إجمالي المستحق": round(due,2), "إجمالي المدفوع": round(paid,2), "الرصيد المتبقي": round(max(balance,0),2), "الحالة": "مدفوع بالكامل" if balance <= 0 else "عليه رصيد"})
    financial_df = pd.DataFrame(financial_rows)
    if not financial_df.empty:
        st.dataframe(financial_df, use_container_width=True)

    st.markdown("### 📜 سجل المدفوعات")
    pay_log = st.session_state.payment_records_df.copy()
    if pay_log.empty:
        st.info("لا توجد دفعات مؤكدة مسجلة حتى الآن.")
    else:
        pay_log = pay_log.sort_values(by="التاريخ", ascending=False, kind="stable")
        if "الشهر" in pay_log.columns:
            pay_log = pay_log[pay_log["الشهر"].astype(str).str.strip() == selected_fin_month]
        st.dataframe(pay_log, use_container_width=True)
        st.caption("حذف سجل الدفع يعكس العملية من الرصيد، ولا يحذف الطالب أو الحصص.")
        # نحتفظ بالفهرس الأصلي منفصلاً عن النص المعروض، حتى لا يعتمد التراجع
        # على تحويل النص إلى رقم إذا تغير شكل الفهرس أو تنسيق السجل.
        payment_log_indices = list(pay_log.index)
        # نستخدم رقم الصف الحقيقي كقيمة للـ selectbox، ونستخدم format_func للعرض.
        # بهذه الطريقة لا نعتمد على النص المعروض ولا يحدث خطأ إذا تغيّر ترتيب/شهر السجل.
        payment_log_indices = list(pay_log.index)
        def _payment_log_label(idx):
            r = pay_log.loc[idx]
            return f"{idx} — {r.get('اسم الطالب','')} — {float(r.get('المبلغ',0) or 0):,.0f} جنيه — {r.get('التاريخ','')} — {r.get('طريقة الدفع','')}"
        chosen_idx = st.selectbox(
            "اختر عملية لتراجعها:",
            payment_log_indices,
            format_func=_payment_log_label,
            key="payment_delete_select_idx"
        )
        if st.button("↩️ تراجع عن عملية الدفع المحددة", key="reverse_payment"):
            original_idx = chosen_idx
            st.session_state.payment_records_df = st.session_state.payment_records_df.drop(index=original_idx).reset_index(drop=True)
            save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
            st.success("✓ تم التراجع عن عملية الدفع وعاد المبلغ إلى الرصيد المستحق.")
            st.rerun()

    st.markdown("### 📥 تصدير الحسابات")
    pay_export = io.BytesIO()
    with pd.ExcelWriter(pay_export, engine="openpyxl") as writer:
        financial_df.to_excel(writer, sheet_name="StudentBalances", index=False)
        st.session_state.payment_records_df.to_excel(writer, sheet_name="PaymentRecords", index=False)
    st.download_button("📥 تصدير كشف الحسابات والمدفوعات Excel", data=pay_export.getvalue(), file_name="حسابات_ومدفوعات_الطلاب.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="payment_export_excel")


elif t_page == "weekly_schedule":
    st.markdown("<div class='vertical-section-header'>🗓️ لوحة مواعيد الطلاب الأسبوعية</div>", unsafe_allow_html=True)
    st.caption("لوحة مستقلة لإضافة الطلاب ومواعيدهم. كل طالب يظهر بلون مختلف، ويمكن إضافة أكثر من موعد للطالب نفسه.")

    ws_df = st.session_state.weekly_schedule_df
    known_students = sorted(list(set(
        [str(x).strip() for x in st.session_state.users_df["اسم الطالب"].dropna().unique() if str(x).strip()] +
        [str(x).strip() for x in ws_df["اسم الطالب"].dropna().unique() if str(x).strip()]
    )))

    # ألوان جاهزة ومتعددة للطلاب
    palette = ["#2563eb", "#16a34a", "#dc2626", "#9333ea", "#ea580c", "#0891b2", "#db2777", "#65a30d", "#7c3aed", "#0f766e"]
    # لون ثابت لكل طالب، ومختلف تلقائياً عن لون أي طالب آخر.
    # إذا كانت البيانات القديمة أعطت أكثر من طالب نفس اللون، نعيد توزيع الألوان مرة واحدة.
    color_map = {}
    used_colors = set()
    if not ws_df.empty:
        for i, name in enumerate(known_students):
            old_colors = [str(x).strip() for x in ws_df.loc[ws_df["اسم الطالب"].astype(str).str.strip() == name, "اللون"].dropna().tolist()]
            old_color = next((c for c in old_colors if c.startswith("#")), "")
            if old_color and old_color not in used_colors:
                chosen = old_color
            else:
                chosen = palette[i % len(palette)]
                # ضمان عدم تكرار اللون حتى لو عدد الطلاب أكبر من الألوان الجاهزة
                if chosen in used_colors:
                    chosen = palette[next(j for j in range(len(palette)) if palette[j] not in used_colors)] if len(used_colors) < len(palette) else f"hsl({(i*47)%360}, 75%, 48%)"
            color_map[name] = chosen
            used_colors.add(chosen)
    else:
        color_map = {name: palette[i % len(palette)] for i, name in enumerate(known_students)}

    # تحديث ألوان الصفوف الموجودة لضمان أن كل اسم طالب له لون مختلف في الجدول والحفظ.
    if not ws_df.empty and color_map:
        ws_df = ws_df.copy()
        ws_df["اللون"] = ws_df["اسم الطالب"].astype(str).str.strip().map(color_map).fillna(ws_df["اللون"])
        st.session_state.weekly_schedule_df = ws_df

    st.markdown("### 👥 إدارة طلاب لوحة المواعيد")
    wm1,wm2,wm3=st.columns(3)
    with wm1:
        with st.expander("➕ إضافة طالب جديد", expanded=False):
            # اختيار المنهج والمرحلة خارج الفورم حتى تتغير قائمة المراحل فوراً عند تغيير المنهج
            ms_curr=st.selectbox("المنهج",list(CURRICULUM_DATA.keys()),key="wm_curr")
            ms_grade=st.selectbox("المرحلة",CURRICULUM_DATA[ms_curr],key="wm_grade")
            with st.form("weekly_manual_student_form", clear_on_submit=True):
                ms_name=st.text_input("اسم الطالب")
                ms_phone=st.text_input("رقم الطالب")
                ms_parent=st.text_input("اسم ولي الأمر")
                ms_parent_phone=st.text_input("رقم ولي الأمر")
                ms_pass=st.text_input("كلمة المرور",value="123456")
                if st.form_submit_button("💾 إضافة الطالب"):
                    target=ms_name.strip()
                    if not target: st.error("اكتب اسم الطالب.")
                    elif target in known_students: st.warning("الطالب موجود بالفعل.")
                    else:
                        row={"اسم الطالب":target,"رقم الهاتف":ms_phone.strip(),"كلمة المرور":ms_pass.strip() or "123456","المنهج/الدولة":ms_curr,"المجموعة/الصف":ms_grade,"اسم ولي الأمر":ms_parent.strip(),"رقم ولي الأمر":ms_parent_phone.strip(),"تاريخ التسجيل":str(date.today()),"الحالة_حظر":"نشط","حالة_الاشتراك_البنك":"غير مشترك"}
                        st.session_state.users_df=pd.concat([st.session_state.users_df,pd.DataFrame([row])],ignore_index=True)
                        save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                        st.success("✓ تمت إضافة الطالب لكشف المسجلين وبطاقات الطلاب والسجلات.")
                        st.rerun()
    with wm2:
        if known_students:
            wd=st.selectbox("اختر طالباً",known_students,key="weekly_delete_student")
            if st.button("🗑️ حذف الطالب نهائياً",key="weekly_delete_student_btn"):
                delete_student_completely(wd); st.success(f"تم حذف {wd} من جميع السجلات."); st.rerun()
    with wm3:
        wh=build_student_roster_html(known_students,"كشف طلاب لوحة المواعيد")
        wp=html_to_pdf_bytes(wh)
        if wp: st.download_button("📄 طباعة الطلاب PDF",wp,file_name="كشف_طلاب_لوحة_المواعيد.pdf",mime="application/pdf",key="weekly_students_pdf")
        else: st.download_button("🖨️ طباعة الطلاب",wh.encode("utf-8"),file_name="كشف_طلاب_لوحة_المواعيد.html",mime="text/html",key="weekly_students_html")
    st.write("---")
    # بيانات الموعد السابق: عند اختيار "إضافة موعد آخر" لنفس الطالب، يتم الاحتفاظ بكل بياناته
    # ونحتاج فقط لتغيير اليوم والساعة (ويمكن تعديل أي بيان قبل الحفظ).
    prefill_student = st.session_state.pop("schedule_prefill_student", "")
    prefill_record = st.session_state.pop("schedule_prefill_record", None)
    prefill_student = str(prefill_student).strip()

    with st.expander("➕ إضافة موعد طالب جديد / إضافة موعد آخر", expanded=True):
        # اختيار الطالب خارج الفورم حتى تتحدث بياناته المرتبطة فوراً عند تغييره
        if known_students:
            default_student_idx = known_students.index(prefill_student) if prefill_student in known_students else 0
            ws_student = st.selectbox("اسم الطالب:", known_students, index=default_student_idx, key="weekly_student")
        else:
            ws_student = st.text_input("اسم الطالب:", value=prefill_student, key="weekly_student_text")

        # بيانات الطالب الأساسية تؤخذ تلقائياً من سجل الطالب الذي أضافه المعلم سابقاً
        student_clean_for_lookup = str(ws_student).strip()
        registered_student = st.session_state.users_df[
            st.session_state.users_df["اسم الطالب"].astype(str).str.strip() == student_clean_for_lookup
        ] if student_clean_for_lookup else pd.DataFrame()

        # نستخدم سجل الطالب الرئيسي أولاً، ثم آخر موعد له كاحتياط للبيانات القديمة
        student_existing = pd.DataFrame()
        if student_clean_for_lookup and not ws_df.empty:
            student_existing = ws_df[ws_df["اسم الطالب"].astype(str).str.strip() == student_clean_for_lookup]
        base_record = prefill_record if isinstance(prefill_record, dict) else (student_existing.iloc[-1].to_dict() if not student_existing.empty else {})

        if not registered_student.empty:
            user_record = registered_student.iloc[0]
            ws_curr = str(user_record.get("المنهج/الدولة", "")).strip()
            ws_grade = str(user_record.get("المجموعة/الصف", "")).strip()
            if not ws_curr or ws_curr.lower() == "nan":
                ws_curr = str(base_record.get("المنهج/الدولة", list(CURRICULUM_DATA.keys())[0]))
            if not ws_grade or ws_grade.lower() == "nan":
                ws_grade = str(base_record.get("المجموعة/الصف", "")).strip()
            registered_phone = str(user_record.get("رقم الهاتف", "")).strip()
        else:
            ws_curr = str(base_record.get("المنهج/الدولة", list(CURRICULUM_DATA.keys())[0]))
            ws_grade = str(base_record.get("المجموعة/الصف", "")).strip()
            registered_phone = ""

        if ws_curr.lower() == "nan":
            ws_curr = list(CURRICULUM_DATA.keys())[0]
        if ws_grade.lower() == "nan":
            ws_grade = ""

        with st.form("weekly_schedule_add_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                ws_academy = st.text_input("اسم الأكاديمية:", value=str(base_record.get("اسم الأكاديمية", "أكاديمية البشمهندس")))
                st.text_input("المنهج الدراسي / الدولة (تلقائي):", value=ws_curr, disabled=True)
                st.text_input("المرحلة / الصف (تلقائي):", value=ws_grade, disabled=True)
                ws_phone = st.text_input("رقم الطالب:", value=registered_phone or str(base_record.get("رقم الطالب", "")))
            with c2:
                ws_sup = st.text_input("رقم مشرف الأكاديمية:", value=str(base_record.get("رقم مشرف الأكاديمية", "")))
                try:
                    base_price = float(base_record.get("سعر الحصة", 100.0))
                except Exception:
                    base_price = 100.0
                ws_price = st.number_input("سعر الحصة (جنيه):", min_value=0.0, step=10.0, value=base_price)
                days_options = ["السبت", "الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة"]
                base_day = str(base_record.get("اليوم", "السبت"))
                day_index = days_options.index(base_day) if base_day in days_options else 0
                ws_day = st.selectbox("يوم الحصة:", days_options, index=day_index)
                # اختيار موعد الحصة بنظام 12 ساعة مع AM / PM بشكل واضح
                try:
                    base_time = datetime.strptime(str(base_record.get("الموعد", "18:00"))[:5], "%H:%M").time()
                except Exception:
                    base_time = time(18, 0)
                base_hour_12 = base_time.hour % 12 or 12
                base_ampm = "AM" if base_time.hour < 12 else "PM"
                base_minute = (base_time.minute // 15) * 15
                if base_minute not in [0, 15, 30, 45]:
                    base_minute = 0
                th1, th2, th3 = st.columns(3)
                with th1:
                    ws_hour = st.selectbox("ساعة الحصة:", list(range(1, 13)), index=list(range(1, 13)).index(base_hour_12))
                with th2:
                    ws_minute = st.selectbox("الدقائق:", [0, 15, 30, 45], index=[0, 15, 30, 45].index(base_minute), format_func=lambda x: f"{x:02d}")
                with th3:
                    ws_ampm = st.selectbox("الفترة:", ["AM", "PM"], index=0 if base_ampm == "AM" else 1)
                hour24 = ws_hour % 12 + (12 if ws_ampm == "PM" else 0)
                ws_time = time(hour24, ws_minute)
                st.caption(f"🕐 الموعد المختار: **{ws_hour}:{ws_minute:02d} {ws_ampm}**")

            st.caption("🔗 المنهج والمرحلة مرتبطان تلقائياً ببيانات الطالب المسجلة في حسابه، ولا يحتاجان لإعادة الاختيار هنا.")
            if prefill_student:
                st.info(f"📌 يتم الآن إضافة موعد آخر للطالب: **{prefill_student}** — غيّر اليوم والساعة ثم اضغط حفظ.")

            if st.form_submit_button("💾 إضافة الموعد إلى الجدول"):
                if not str(ws_student).strip():
                    st.error("يرجى اختيار أو كتابة اسم الطالب.")
                else:
                    student_clean = str(ws_student).strip()
                    new_ws = {
                        "اسم الطالب": student_clean,
                        "اسم الأكاديمية": ws_academy.strip(),
                        "المنهج/الدولة": ws_curr,
                        "المجموعة/الصف": ws_grade,
                        "رقم الطالب": ws_phone.strip(),
                        "رقم مشرف الأكاديمية": ws_sup.strip(),
                        "سعر الحصة": ws_price,
                        "اليوم": ws_day,
                        "الموعد": ws_time.strftime("%H:%M"),
                        "اللون": color_map.get(student_clean, palette[len(color_map) % len(palette)]),
                        "حالة الموعد": "نشط"
                    }
                    # ربط الموعد بسجل الطالب الرئيسي حتى يظهر تلقائياً في كشف المسجلين والبطاقات والسجلات والتقارير والواجبات
                    existing_user = st.session_state.users_df[st.session_state.users_df["اسم الطالب"].astype(str).str.strip() == student_clean]
                    if existing_user.empty:
                        auto_user = {"اسم الطالب":student_clean,"رقم الهاتف":ws_phone.strip(),"كلمة المرور":"123456","المنهج/الدولة":ws_curr,"المجموعة/الصف":ws_grade,"اسم ولي الأمر":"","رقم ولي الأمر":"","تاريخ التسجيل":str(date.today()),"الحالة_حظر":"نشط","حالة_الاشتراك_البنك":"غير مشترك"}
                        st.session_state.users_df = pd.concat([st.session_state.users_df, pd.DataFrame([auto_user])], ignore_index=True)
                    else:
                        ui = existing_user.index[0]
                        if not str(st.session_state.users_df.at[ui,"رقم الهاتف"]).strip() and ws_phone.strip(): st.session_state.users_df.at[ui,"رقم الهاتف"] = ws_phone.strip()
                        st.session_state.users_df.at[ui,"المنهج/الدولة"] = ws_curr
                        st.session_state.users_df.at[ui,"المجموعة/الصف"] = ws_grade
                    st.session_state.weekly_schedule_df = pd.concat([st.session_state.weekly_schedule_df, pd.DataFrame([new_ws])], ignore_index=True)
                    save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                    st.success(f"✓ تم إضافة موعد {student_clean} يوم {ws_day} الساعة {ws_time.strftime('%I:%M %p').lstrip('0')}.")
                    st.rerun()

    # جدول أسبوعي حقيقي: الأيام أعمدة والطلاب داخل الخلايا حسب الموعد
    st.markdown("### 📅 الجدول الأسبوعي")
    days = ["السبت", "الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة"]
    if ws_df.empty:
        st.info("لا توجد مواعيد مضافة حتى الآن.")
    else:
        rows = []
        times = sorted([str(x) for x in ws_df["الموعد"].dropna().unique()])
        for tm in times:
            row = {"الساعة": format_schedule_time_ampm(tm)}
            for d in days:
                matches = ws_df[(ws_df["الموعد"].astype(str) == tm) & (ws_df["اليوم"].astype(str) == d) & (ws_df["حالة الموعد"].astype(str) != "متوقف")]
                parts = []
                for _, r in matches.iterrows():
                    col = str(r.get("اللون", "#2563eb"))
                    parts.append(f"<div style='background:{col};color:#fff;padding:7px;border-radius:8px;margin:2px 0;font-weight:900;'>👤 {r['اسم الطالب']}<br><small>{r['المجموعة/الصف']} | {r['اسم الأكاديمية']}</small></div>")
                row[d] = "".join(parts) if parts else "—"
            rows.append(row)
        schedule_html = "<table style='width:100%;border-collapse:collapse;text-align:center;direction:rtl'><tr style='background:#f1f5f9'><th style='padding:10px;border:1px solid #cbd5e1'>الساعة</th>" + "".join([f"<th style='padding:10px;border:1px solid #cbd5e1'>{d}</th>" for d in days]) + "</tr>"
        for row in rows:
            schedule_html += f"<tr><td style='padding:10px;border:1px solid #cbd5e1;font-weight:900'>{row['الساعة']}</td>" + "".join([f"<td style='padding:5px;border:1px solid #cbd5e1;vertical-align:top'>{row[d]}</td>" for d in days]) + "</tr>"
        schedule_html += "</table>"
        st.markdown(schedule_html, unsafe_allow_html=True)

        st.write("---")
        st.markdown("### 👥 تفاصيل المواعيد وإدارتها")
        for ws_idx, r in ws_df.iterrows():
            student_nm = str(r.get("اسم الطالب", ""))
            with st.expander(f"{student_nm} — {r.get('اليوم','')} {r.get('الموعد','')} | {r.get('اسم الأكاديمية','')}"):
                c1, c2, c3 = st.columns(3)
                c1.write(f"**المنهج:** {r.get('المنهج/الدولة','')}\n\n**المرحلة:** {r.get('المجموعة/الصف','')}")
                c2.write(f"**سعر الحصة:** {r.get('سعر الحصة',0)} جنيه\n\n**مشرف الأكاديمية:** {r.get('رقم مشرف الأكاديمية','')}")
                c3.write(f"**رقم الطالب:** {r.get('رقم الطالب','')}\n\n**الحالة:** {r.get('حالة الموعد','نشط')}")
                b1, b2, b3, b4 = st.columns(4)
                with b1:
                    if st.button("➕ إضافة موعد آخر", key=f"ws_add_another_{ws_idx}"):
                        st.session_state.schedule_prefill_student = student_nm
                        st.session_state.schedule_prefill_record = r.to_dict()
                        st.rerun()
                with b2:
                    if st.button("📝 رصد حصة الطالب", key=f"ws_session_{ws_idx}"):
                        st.session_state.prefill_student = student_nm
                        st.session_state.prefill_schedule_idx = ws_idx
                        st.session_state.teacher_page = "add_session"
                        st.rerun()
                with b3:
                    if st.button("📚 رصد واجب الطالب", key=f"ws_hw_{ws_idx}"):
                        st.session_state.prefill_student = student_nm
                        st.session_state.teacher_page = "add_hw"
                        st.rerun()
                with b4:
                    if st.button("🗑️ حذف الموعد", key=f"ws_del_{ws_idx}"):
                        st.session_state.weekly_schedule_df = ws_df.drop(ws_idx).reset_index(drop=True)
                        save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                        st.success("تم حذف الموعد فقط، ولن يتم حذف الطالب أو سجلاته.")
                        st.rerun()

    st.write("---")
    st.markdown("### 📋 كشف المواعيد القابل للطباعة والتصدير")
    st.dataframe(st.session_state.weekly_schedule_df, use_container_width=True)
    # نسخة الطباعة تكون بنفس شكل الجدول الأسبوعي: الأيام أعمدة والطلاب داخل الخلايا بألوانهم.
    ws_html=build_weekly_schedule_print_html(st.session_state.weekly_schedule_df, "الجدول الأسبوعي لمواعيد الطلاب")
    ws_pdf=html_to_pdf_bytes(ws_html)
    e1,e2,e3=st.columns(3)
    with e1:
        if ws_pdf: st.download_button("📄 طباعة الجدول الأسبوعي PDF",ws_pdf,file_name="الجدول_الأسبوعي_للطلاب.pdf",mime="application/pdf",key="weekly_schedule_pdf")
        else: st.download_button("🖨️ طباعة الجدول الأسبوعي",ws_html.encode("utf-8"),file_name="الجدول_الأسبوعي_للطلاب.html",mime="text/html",key="weekly_schedule_html")
    with e2:
        # كشف تفصيلي منفصل لمن يريد طباعته كقائمة
        detail_rows="".join(f"<tr><td>{r.get('اسم الطالب','')}</td><td>{r.get('اليوم','')}</td><td>{r.get('الموعد','')}</td><td>{r.get('اسم الأكاديمية','')}</td><td>{r.get('المنهج/الدولة','')}</td><td>{r.get('المجموعة/الصف','')}</td><td>{r.get('سعر الحصة','')}</td></tr>" for _,r in st.session_state.weekly_schedule_df.iterrows())
        detail_html=make_print_html("كشف مواعيد الطلاب",detail_rows,"<th>الطالب</th><th>اليوم</th><th>الموعد</th><th>الأكاديمية</th><th>المنهج</th><th>المرحلة</th><th>السعر</th>")
        detail_pdf=html_to_pdf_bytes(detail_html)
        if detail_pdf: st.download_button("📋 طباعة كشف المواعيد PDF",detail_pdf,file_name="كشف_مواعيد_الطلاب.pdf",mime="application/pdf",key="weekly_schedule_detail_pdf")
        else: st.download_button("🖨️ طباعة كشف المواعيد",detail_html.encode("utf-8"),file_name="كشف_مواعيد_الطلاب.html",mime="text/html",key="weekly_schedule_detail_html")
    with e3:
        buf_ws=io.BytesIO()
        with pd.ExcelWriter(buf_ws,engine="openpyxl") as writer: st.session_state.weekly_schedule_df.to_excel(writer,sheet_name="WeeklySchedule",index=False)
        st.download_button("📥 تصدير جدول المواعيد Excel",data=buf_ws.getvalue(),file_name="جدول_مواعيد_الطلاب.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif t_page == "online_schedule":
    st.subheader("💻 إدارة جدول حصص الأونلاين (Zoom) وتحديد مواعيد الطلاب:")

    all_registered_names = sorted(list(set(
        [str(s).strip() for s in st.session_state.users_df["اسم الطالب"].dropna().unique() if str(s).strip()]
        + [str(s).strip() for s in st.session_state.sessions_df["اسم الطالب"].dropna().unique() if str(s).strip()]
        + [str(s).strip() for s in st.session_state.weekly_schedule_df["اسم الطالب"].dropna().unique() if str(s).strip()]
        + [str(s).strip() for s in st.session_state.online_schedule_df["اسم الطالب"].dropna().unique() if str(s).strip()]
    )))

    st.markdown("### 👥 إدارة طلاب Zoom")
    zm1,zm2,zm3=st.columns(3)
    with zm1:
        with st.expander("➕ إضافة طالب",expanded=False):
            # اختيار المنهج والمرحلة خارج الفورم حتى تتغير قائمة المراحل فوراً عند تغيير المنهج
            zc=st.selectbox("المنهج",list(CURRICULUM_DATA.keys()),key="zm_curr")
            zg=st.selectbox("المرحلة",CURRICULUM_DATA[zc],key="zm_grade")
            with st.form("zoom_manual_student_form",clear_on_submit=True):
                zn=st.text_input("اسم الطالب"); zp=st.text_input("رقم الطالب"); zpn=st.text_input("اسم ولي الأمر"); zpp=st.text_input("رقم ولي الأمر")
                zpass=st.text_input("كلمة المرور",value="123456")
                if st.form_submit_button("💾 إضافة الطالب"):
                    target=zn.strip()
                    if not target: st.error("اكتب اسم الطالب.")
                    elif target in all_registered_names: st.warning("الطالب موجود بالفعل.")
                    else:
                        row={"اسم الطالب":target,"رقم الهاتف":zp.strip(),"كلمة المرور":zpass.strip() or "123456","المنهج/الدولة":zc,"المجموعة/الصف":zg,"اسم ولي الأمر":zpn.strip(),"رقم ولي الأمر":zpp.strip(),"تاريخ التسجيل":str(date.today()),"الحالة_حظر":"نشط","حالة_الاشتراك_البنك":"غير مشترك"}
                        st.session_state.users_df=pd.concat([st.session_state.users_df,pd.DataFrame([row])],ignore_index=True)
                        save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                        st.success("✓ تمت إضافة الطالب لكشف المسجلين وبطاقات الطلاب."); st.rerun()
    with zm2:
        if all_registered_names:
            zd=st.selectbox("اختر طالباً",all_registered_names,key="zoom_delete_student")
            if st.button("🗑️ حذف الطالب نهائياً",key="zoom_delete_student_btn"):
                delete_student_completely(zd); st.success(f"تم حذف {zd} من جميع السجلات."); st.rerun()
    with zm3:
        zh=build_student_roster_html(all_registered_names,"كشف طلاب Zoom"); zpdata=html_to_pdf_bytes(zh)
        if zpdata: st.download_button("📄 طباعة الطلاب PDF",zpdata,file_name="كشف_طلاب_Zoom.pdf",mime="application/pdf",key="zoom_roster_pdf")
        else: st.download_button("🖨️ طباعة الطلاب",zh.encode("utf-8"),file_name="كشف_طلاب_Zoom.html",mime="text/html",key="zoom_roster_html")

    # المنهج والمرحلة خارج الفورم لأن Streamlit لا يعيد تشغيل widgets داخل form عند تغييرها
    os_curr = st.selectbox("المنهج الدراسي / الدولة:", list(CURRICULUM_DATA.keys()), key="zoom_schedule_curr")
    os_grade = st.selectbox("المرحلة / الصف الدراسي:", CURRICULUM_DATA[os_curr], key="zoom_schedule_grade")
    with st.form("add_online_sched_form", clear_on_submit=True):
        col_os1, col_os2 = st.columns(2)
        with col_os1:
            os_student = st.selectbox("اختر الطالب أو اكتبه:", all_registered_names) if all_registered_names else st.text_input("اسم الطالب:")
            os_academy = st.text_input("اسم الأكاديمية:", value="أكاديمية البشمهندس")
            os_phone = st.text_input("رقم الهاتف المحمول للطالب:")
        with col_os2:
            os_sup_phone = st.text_input("رقم مشرف الأكاديمية:")
            os_price = st.number_input("سعر الحصة (جنيه):", min_value=0.0, step=10.0, value=100.0)
            os_date = st.date_input("تقويم موعد الحصة:", value=date.today())
            os_time = st.time_input("ساعة الحصة:", value=time(18, 0))
            os_zoom = st.text_input("رابط زوم المخصص:", value="https://us05web.zoom.us/j/83526892910?pwd=2jWRgATgBRPbXttdnm0QpLwBApsZL4.1")

        if st.form_submit_button("💾 حفظ وإدراج الطالب في جدول الأونلاين"):
            if not str(os_student).strip() or not os_phone.strip():
                st.error("يرجى كتابة اسم الطالب ورقم الهاتف.")
            else:
                new_sched_row = {
                    "اسم الطالب": str(os_student).strip(),
                    "اسم الأكاديمية": os_academy.strip(),
                    "المنهج/الدولة": os_curr,
                    "المجموعة/الصف": os_grade,
                    "رقم الطالب": os_phone.strip(),
                    "رقم مشرف الأكاديمية": os_sup_phone.strip(),
                    "سعر الحصة": os_price,
                    "تاريخ الحصة": str(os_date),
                    "ساعة الحصة": os_time.strftime("%H:%M"),
                    "رابط زوم": os_zoom.strip(),
                    "حالة فتح الحصة": "مغلقة"
                }
                online_student_clean = str(os_student).strip()
                existing_user = st.session_state.users_df[st.session_state.users_df["اسم الطالب"].astype(str).str.strip() == online_student_clean]
                if existing_user.empty:
                    auto_user = {"اسم الطالب":online_student_clean,"رقم الهاتف":os_phone.strip(),"كلمة المرور":"123456","المنهج/الدولة":os_curr,"المجموعة/الصف":os_grade,"اسم ولي الأمر":"","رقم ولي الأمر":"","تاريخ التسجيل":str(date.today()),"الحالة_حظر":"نشط","حالة_الاشتراك_البنك":"غير مشترك"}
                    st.session_state.users_df = pd.concat([st.session_state.users_df, pd.DataFrame([auto_user])], ignore_index=True)
                st.session_state.online_schedule_df = pd.concat([st.session_state.online_schedule_df, pd.DataFrame([new_sched_row])], ignore_index=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                st.success(f"✓ تم إضافة الطالب ({os_student}) إلى جدول الأونلاين بنجاح!")

    st.write("---")
    st.markdown("### 📋 جدول الطلاب المسجلين أونلاين والتحكم بحالة الحصة (فتح/إغلاق):")
    os_state = st.session_state.online_schedule_df
    if os_state.empty:
        st.info("لا توجد حصص أونلاين مسجلة حتى الآن.")
    else:
        for os_idx, os_row in os_state.iterrows():
            st_n = os_row["اسم الطالب"]
            ac_n = os_row.get("اسم الأكاديمية", "أكاديمية")
            dt_n = os_row.get("تاريخ الحصة", "")
            tm_n = os_row.get("ساعة الحصة", "")
            cur_status = os_row["حالة فتح الحصة"]

            with st.expander(f"طالب: {st_n} — أكاديمية: ({ac_n}) | الموعد: {dt_n} {tm_n} | الحالة: [{cur_status}]"):
                with st.form(f"update_zoom_status_{os_idx}"):
                    new_status = st.selectbox("تغيير حالة الحصة (ليراها الطالب زر فتح):", ["مغلقة", "مفتوحة"], index=0 if cur_status == "مغلقة" else 1)
                    edit_zoom_link = st.text_input("تحديث رابط زوم:", value=str(os_row["رابط زوم"]))
                    
                    c_os1, c_os2 = st.columns(2)
                    with c_os1:
                        update_sub = st.form_submit_button("🔄 تحديث حالة الحصة ولينك زوم")
                    with c_os2:
                        del_sub = st.form_submit_button("🗑️ حذف من جدول الأونلاين")

                    if update_sub:
                        os_state.at[os_idx, "حالة فتح الحصة"] = new_status
                        os_state.at[os_idx, "رابط زوم"] = edit_zoom_link.strip()
                        st.session_state.online_schedule_df = os_state
                        save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                        st.success("✓ تم تحديث حالة الحصة بنجاح وسيظهر للطالب فوراً!")
                        st.rerun()

                    if del_sub:
                        st.session_state.online_schedule_df = os_state.drop(os_idx).reset_index(drop=True)
                        save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                        st.warning("تم حذف السجل.")
                        st.rerun()
                oa,ob,oc=st.columns(3)
                with oa:
                    if st.button("📝 رصد حصة الطالب",key=f"zoom_session_{os_idx}"):
                        st.session_state.prefill_student=str(st_n); st.session_state.teacher_page="add_session"; st.rerun()
                with ob:
                    if st.button("📚 رصد واجب الطالب",key=f"zoom_hw_{os_idx}"):
                        st.session_state.prefill_student=str(st_n); st.session_state.teacher_page="add_hw"; st.rerun()
                with oc:
                    if st.button("➕ إضافة موعد آخر",key=f"zoom_add_other_{os_idx}"):
                        st.session_state.schedule_prefill_student=str(st_n); st.session_state.schedule_prefill_record=os_row.to_dict(); st.session_state.teacher_page="weekly_schedule"; st.rerun()

elif t_page == "exam_maker":
    st.markdown("<div class='exam-builder-header'>➕ إضافة امتحان جديد / محرر وقص الصور وميزة الطباعة PDF</div>", unsafe_allow_html=True)

    if "temp_questions" not in st.session_state: st.session_state.temp_questions = []
    if "q_img_ver" not in st.session_state: st.session_state.q_img_ver = 0
    if "opt1_ver" not in st.session_state: st.session_state.opt1_ver = 0
    if "opt2_ver" not in st.session_state: st.session_state.opt2_ver = 0
    if "opt3_ver" not in st.session_state: st.session_state.opt3_ver = 0
    if "opt4_ver" not in st.session_state: st.session_state.opt4_ver = 0

    if "pasted_q_img" not in st.session_state: st.session_state.pasted_q_img = ""
    if "pasted_opt1_img" not in st.session_state: st.session_state.pasted_opt1_img = ""
    if "pasted_opt2_img" not in st.session_state: st.session_state.pasted_opt2_img = ""
    if "pasted_opt3_img" not in st.session_state: st.session_state.pasted_opt3_img = ""
    if "pasted_opt4_img" not in st.session_state: st.session_state.pasted_opt4_img = ""

    with st.expander("⚙️ 1. الإعدادات الرئيسية للامتحان", expanded=True):
        c_ex1, c_ex2 = st.columns(2)
        with c_ex1:
            ex_title_input = st.text_input("عنوان الامتحان:*", placeholder="مثال: اختبار الجبر المصور - شهر أكتوبر")
            ex_desc_input = st.text_area("وصف الامتحان:*", placeholder="تعليمات هامة للطلاب قبل البدء...")
            ex_pass_input = st.text_input("كلمة المرور (اختياري):", placeholder="اتركه فارغاً إن لم ترغب بكلمة سر")
        with c_ex2:
            ex_curr_input = st.selectbox("اختر الدولة / المنهج:*", list(CURRICULUM_DATA.keys()), key="mk_c")
            ex_grade_input = st.selectbox("اختر الصف الدراسي:*", CURRICULUM_DATA[ex_curr_input], key="mk_g")
            ex_subject_input = st.selectbox("اختر مادة الامتحان:*", ["الرياضيات (عام)", "الجبر والإحصاء", "الهندسة وحساب المثلثات", "التفاضل والتكامل", "الاستاتيكا والديناميكا", "الرياضيات التطبيقية"])
            ex_term_input = st.selectbox("اختر الفصل الدراسي:*", ["الفصل الدراسي الأول", "الفصل الدراسي الثاني", "مراجعة نهائية"])
            ex_time_input = st.number_input("مدة الامتحان بالدقائق:", min_value=5, max_value=180, value=45, step=5)

    with st.expander("📐 2. إدخال الأسئلة وتحرير وقص الصور", expanded=True):
        q_count = len(st.session_state.temp_questions) + 1
        st.markdown(f"#### إضافة سؤال رقم {q_count}")

        q_type_choice = st.selectbox("نوع السؤال:", ["سؤال موضوعي (اختيار من متعدد)", "سؤال مقالي (خطوات حل ورفع صورة)"])
        is_mcq = "موضوعي" in q_type_choice

        st.markdown("##### 📌 أ) صورة السؤال:")
        col_qp1, col_qp2 = st.columns([1, 1])
        with col_qp1:
            if paste_image_button:
                paste_res = paste_image_button("📋 لصق صورة السؤال من الحافظة", key=f"paste_q_{q_count}_{st.session_state.q_img_ver}")
                if paste_res.image_data is not None:
                    st.session_state.pasted_q_img = pil_to_base64(paste_res.image_data)
        with col_qp2:
            q_file_up = st.file_uploader("أو اسحب صورة السؤال هنا:", type=["jpg", "png", "jpeg"], key=f"upload_q_{q_count}_{st.session_state.q_img_ver}")
            if q_file_up is not None:
                st.session_state.pasted_q_img = base64.b64encode(q_file_up.read()).decode()

        if st.session_state.pasted_q_img:
            st.write("---")
            col_preview, col_del = st.columns([3, 1])
            with col_preview:
                st.markdown("**المعاينة الحالية لصورة السؤال:**")
                st.image(f"data:image/jpeg;base64,{st.session_state.pasted_q_img}", width=400)
            with col_del:
                st.write("")
                st.write("")
                if st.button("🗑️ مسح الصورة بالكامل", key=f"del_q_main_btn_{q_count}_{st.session_state.q_img_ver}"):
                    st.session_state.pasted_q_img = ""
                    st.session_state.q_img_ver += 1
                    st.rerun()

            st.markdown("##### ✂️ أداة قص الصورة بالماوس مباشرة:")
            pil_q = base64_to_pil(st.session_state.pasted_q_img)
            if pil_q:
                st.markdown('<div class="crop-container">', unsafe_allow_html=True)
                st.info("👇 حرّك المربع بالماوس لتحديد الجزء المطلوب قصه ثم اضغط زر 'اعتماد القص':")
                if st_cropper:
                    cropped_img = st_cropper(pil_q, realtime_update=True, box_color='#0052cc', aspect_ratio=None, key=f"cropper_q_{q_count}_{st.session_state.q_img_ver}")
                    if st.button("✂️ اعتماد وحفظ الجزء المقصوص", key=f"btn_confirm_crop_{q_count}"):
                        st.session_state.pasted_q_img = pil_to_base64(cropped_img)
                        st.session_state.q_img_ver += 1
                        st.success("✓ تم قص وتحديث صورة السؤال!")
                        st.rerun()
                else:
                    st.warning("يرجى التأكد من إضافة `streamlit-cropper` في requirements.txt.")
                st.markdown('</div>', unsafe_allow_html=True)

        q_text_input = st.text_area("نص السؤال المكتوب (اختياري):", placeholder="مثال: أوجد قيمة س الموضحة بالرسم:")

        opt1_txt, opt2_txt, opt3_txt, opt4_txt = "", "", "", ""
        correct_num = 1

        if is_mcq:
            st.write("---")
            st.markdown("##### 🎯 ب) خيارات الإجابة الأربعة ومسح صورها:")

            st.markdown("**الخيار الأول (أ):**")
            co1_a, co1_b, co1_c = st.columns([3, 2, 1])
            with co1_a: opt1_txt = st.text_input("نص الخيار (أ):", key=f"t_opt1_{q_count}")
            with co1_b:
                if paste_image_button:
                    p1 = paste_image_button("📋 لصق صورة (أ)", key=f"p_opt1_{q_count}_{st.session_state.opt1_ver}")
                    if p1.image_data is not None: st.session_state.pasted_opt1_img = pil_to_base64(p1.image_data)
                f1 = st.file_uploader("ارفع صورة (أ):", type=["jpg", "png", "jpeg"], key=f"u_opt1_{q_count}_{st.session_state.opt1_ver}")
                if f1 is not None: st.session_state.pasted_opt1_img = base64.b64encode(f1.read()).decode()
            with co1_c:
                st.write("")
                if st.session_state.pasted_opt1_img and st.button("🗑️ مسح", key=f"del_o1_{q_count}_{st.session_state.opt1_ver}"):
                    st.session_state.pasted_opt1_img = ""
                    st.session_state.opt1_ver += 1
                    st.rerun()
            if st.session_state.pasted_opt1_img: st.image(f"data:image/jpeg;base64,{st.session_state.pasted_opt1_img}", width=180)

            st.markdown("**الخيار الثاني (ب):**")
            co2_a, co2_b, co2_c = st.columns([3, 2, 1])
            with co2_a: opt2_txt = st.text_input("نص الخيار (ب):", key=f"t_opt2_{q_count}")
            with co2_b:
                if paste_image_button:
                    p2 = paste_image_button("📋 لصق صورة (ب)", key=f"p_opt2_{q_count}_{st.session_state.opt2_ver}")
                    if p2.image_data is not None: st.session_state.pasted_opt2_img = pil_to_base64(p2.image_data)
                f2 = st.file_uploader("ارفع صورة (ب):", type=["jpg", "png", "jpeg"], key=f"u_opt2_{q_count}_{st.session_state.opt2_ver}")
                if f2 is not None: st.session_state.pasted_opt2_img = base64.b64encode(f2.read()).decode()
            with co2_c:
                st.write("")
                if st.session_state.pasted_opt2_img and st.button("🗑️ مسح", key=f"del_o2_{q_count}_{st.session_state.opt2_ver}"):
                    st.session_state.pasted_opt2_img = ""
                    st.session_state.opt2_ver += 1
                    st.rerun()
            if st.session_state.pasted_opt2_img: st.image(f"data:image/jpeg;base64,{st.session_state.pasted_opt2_img}", width=180)

            st.markdown("**الخيار الثالث (ج):**")
            co3_a, co3_b, co3_c = st.columns([3, 2, 1])
            with co3_a: opt3_txt = st.text_input("نص الخيار (ج):", key=f"t_opt3_{q_count}")
            with co3_b:
                if paste_image_button:
                    p3 = paste_image_button("📋 لصق صورة (ج)", key=f"p_opt3_{q_count}_{st.session_state.opt3_ver}")
                    if p3.image_data is not None: st.session_state.pasted_opt3_img = pil_to_base64(p3.image_data)
                f3 = st.file_uploader("ارفع صورة (ج):", type=["jpg", "png", "jpeg"], key=f"u_opt3_{q_count}_{st.session_state.opt3_ver}")
                if f3 is not None: st.session_state.pasted_opt3_img = base64.b64encode(f3.read()).decode()
            with co3_c:
                st.write("")
                if st.session_state.pasted_opt3_img and st.button("🗑️ مسح", key=f"del_o3_{q_count}_{st.session_state.opt3_ver}"):
                    st.session_state.pasted_opt3_img = ""
                    st.session_state.opt3_ver += 1
                    st.rerun()
            if st.session_state.pasted_opt3_img: st.image(f"data:image/jpeg;base64,{st.session_state.pasted_opt3_img}", width=180)

            st.markdown("**الخيار الرابع (د):**")
            co4_a, co4_b, co4_c = st.columns([3, 2, 1])
            with co4_a: opt4_txt = st.text_input("نص الخيار (د):", key=f"t_opt4_{q_count}")
            with co4_b:
                if paste_image_button:
                    p4 = paste_image_button("📋 لصق صورة (د)", key=f"p_opt4_{q_count}_{st.session_state.opt4_ver}")
                    if p4.image_data is not None: st.session_state.pasted_opt4_img = pil_to_base64(p4.image_data)
                f4 = st.file_uploader("ارفع صورة (د):", type=["jpg", "png", "jpeg"], key=f"u_opt4_{q_count}_{st.session_state.opt4_ver}")
                if f4 is not None: st.session_state.pasted_opt4_img = base64.b64encode(f4.read()).decode()
            with co4_c:
                st.write("")
                if st.session_state.pasted_opt4_img and st.button("🗑️ مسح", key=f"del_o4_{q_count}_{st.session_state.opt4_ver}"):
                    st.session_state.pasted_opt4_img = ""
                    st.session_state.opt4_ver += 1
                    st.rerun()
            if st.session_state.pasted_opt4_img: st.image(f"data:image/jpeg;base64,{st.session_state.pasted_opt4_img}", width=180)

            st.write("---")
            correct_num = st.selectbox("الإجابة الصحيحة هي:*", [1, 2, 3, 4], format_func=lambda x: f"الخيار ({['أ', 'ب', 'ج', 'د'][x-1]})", key=f"cor_{q_count}")

        st.write("---")
        q_pts_val = st.selectbox("علامة هذا السؤال (الدرجة):*", [0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 4.0, 5.0, 10.0], index=1, key=f"pts_{q_count}")

        if st.button(f"➕ حفظ وإدراج السؤال رقم {q_count} في الامتحان"):
            if not q_text_input.strip() and not st.session_state.pasted_q_img:
                st.error("يرجى كتابة نص السؤال أو لصق صورة للسؤال أولاً.")
            else:
                st.session_state.temp_questions.append({
                    "type": "موضوعي" if is_mcq else "مقالي",
                    "text": q_text_input.strip(),
                    "q_img": st.session_state.pasted_q_img,
                    "opt1": opt1_txt.strip(), "opt1_img": st.session_state.pasted_opt1_img,
                    "opt2": opt2_txt.strip(), "opt2_img": st.session_state.pasted_opt2_img,
                    "opt3": opt3_txt.strip(), "opt3_img": st.session_state.pasted_opt3_img,
                    "opt4": opt4_txt.strip(), "opt4_img": st.session_state.pasted_opt4_img,
                    "correct": correct_num,
                    "points": q_pts_val,
                })
                st.session_state.pasted_q_img = ""
                st.session_state.pasted_opt1_img = ""
                st.session_state.pasted_opt2_img = ""
                st.session_state.pasted_opt3_img = ""
                st.session_state.pasted_opt4_img = ""
                st.session_state.q_img_ver += 1
                st.session_state.opt1_ver += 1
                st.session_state.opt2_ver += 1
                st.session_state.opt3_ver += 1
                st.session_state.opt4_ver += 1
                st.success(f"✓ تم حفظ السؤال رقم {q_count} بنجاح!")
                st.rerun()

        if st.session_state.temp_questions:
            st.write(f"**الأسئلة الجاهزة في هذا الاختبار: ({len(st.session_state.temp_questions)}) سؤال**")
            for idx_q, q_item in enumerate(st.session_state.temp_questions):
                st.markdown(f"- **س {idx_q+1} ({q_item['type']}):** {q_item['text'] if q_item['text'] else '[سؤال مصور]'} — *(الدرجة: {q_item['points']})*")
            if st.button("🗑️ مسح كل الأسئلة والبدء من جديد"):
                st.session_state.temp_questions = []
                st.rerun()

    st.write("---")
    if st.button("🚀 حفظ ونشر الامتحان للطلاب"):
        if not ex_title_input.strip():
            st.error("يرجى كتابة عنوان الامتحان أولاً.")
        elif not st.session_state.temp_questions:
            st.error("يرجى إضافة سؤال واحد على الأقل قبل النشر.")
        else:
            new_ex = {
                "معرف_الامتحان": f"EX_{datetime.now().strftime('%Y%m%d%H%M%S')}",
                "عنوان الامتحان": ex_title_input.strip(),
                "وصف الامتحان": ex_desc_input.strip(),
                "كلمة المرور": str(ex_pass_input).strip(),
                "المنهج/الدولة": ex_curr_input,
                "المجموعة/الصف": ex_grade_input,
                "المادة": ex_subject_input,
                "الفصل الدراسي": ex_term_input,
                "مدة الامتحان بالدقائق": int(ex_time_input),
                "الأسئلة_JSON": json.dumps(st.session_state.temp_questions, ensure_ascii=False),
                "تاريخ الإنشاء": str(date.today()),
            }
            st.session_state.exams_df = pd.concat([st.session_state.exams_df, pd.DataFrame([new_ex])], ignore_index=True)
            save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
            st.session_state.temp_questions = []
            st.success(f"✓ تم نشر امتحان ({ex_title_input}) بنجاح لصف ({ex_grade_input})!")
            st.rerun()

elif t_page == "question_bank":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("📚 بنك الأسئلة الشامل (إضافة أسئلة اختر ومقالي بصور ومواصفات كاملة):")

    if "qb_q_img" not in st.session_state: st.session_state.qb_q_img = ""
    if "qb_ver" not in st.session_state: st.session_state.qb_ver = 0

    qb_curr = st.selectbox("المنهج الدراسي / الدولة:", list(CURRICULUM_DATA.keys()), key="qbc_c")
    qb_grade = st.selectbox("المرحلة / الصف الدراسي المستهدف:", CURRICULUM_DATA[qb_curr], key="qbc_g")
    qb_subject = st.selectbox("المادة:", ["الرياضيات (عام)", "الجبر والإحصاء", "الهندسة وحساب المثلثات", "التفاضل والتكامل", "الاستاتيكا والديناميكا"], key="qbc_s")
    
    qb_type = st.selectbox("نوع السؤال:", ["اختيار من متعدد", "مقالي"], key="qbc_t", on_change=lambda: st.rerun())
    is_qbank_mcq = (qb_type == "اختيار من متعدد")

    with st.form("add_qbank_form"):
        st.markdown("##### 📌 أ) صورة السؤال (اختياري):")
        qb_file_up = st.file_uploader("رفع صورة السؤال:", type=["jpg", "png", "jpeg"], key=f"qbc_img_up_{st.session_state.qb_ver}")
        if qb_file_up is not None:
            st.session_state.qb_q_img = base64.b64encode(qb_file_up.read()).decode()

        if st.session_state.qb_q_img:
            st.image(f"data:image/jpeg;base64,{st.session_state.qb_q_img}", width=350)
            if st.form_submit_button("🗑️ مسح صورة السؤال"):
                st.session_state.qb_q_img = ""
                st.session_state.qb_ver += 1
                st.rerun()

        qb_text = st.text_area("نص السؤال المكتوب:", key="qbc_txt", placeholder="أكتب نص السؤال هنا...")

        qb_opt1, qb_opt2, qb_opt3, qb_opt4 = "", "", "", ""
        qb_correct = 1

        if is_qbank_mcq:
            st.markdown("---")
            st.markdown("##### 🎯 ب) خيارات الاختيار من متعدد:")
            qb_opt1 = st.text_input("الخيار (أ):", key="qbc_o1")
            qb_opt2 = st.text_input("الخيار (ب):", key="qbc_o2")
            qb_opt3 = st.text_input("الخيار (ج):", key="qbc_o3")
            qb_opt4 = st.text_input("الخيار (د):", key="qbc_o4")
            qb_correct = st.selectbox("الإجابة الصحيحة هي:*", [1, 2, 3, 4], format_func=lambda x: f"الخيار ({['أ', 'ب', 'ج', 'د'][x-1]})", key="qbc_cor")

        st.write("---")
        qb_points = st.number_input("درجة السؤال:", min_value=0.5, max_value=10.0, value=1.0, step=0.5, key="qbc_pts")

        if st.form_submit_button("💾 حفظ وإدراج السؤال في بنك الأسئلة"):
            if not qb_text.strip() and not st.session_state.qb_q_img:
                st.error("يرجى كتابة نص السؤال أو رفع صورة السؤال على الأقل.")
            else:
                q_payload = {
                    "type": qb_type,
                    "text": qb_text.strip(),
                    "q_img": st.session_state.qb_q_img,
                    "opt1": qb_opt1.strip() if is_qbank_mcq else "",
                    "opt2": qb_opt2.strip() if is_qbank_mcq else "",
                    "opt3": qb_opt3.strip() if is_qbank_mcq else "",
                    "opt4": qb_opt4.strip() if is_qbank_mcq else "",
                    "correct": qb_correct if is_qbank_mcq else 1,
                    "points": qb_points
                }
                new_qbank_row = {
                    "معرف_السؤال": f"QB_{datetime.now().strftime('%Y%m%d%H%M%S')}",
                    "المنهج/الدولة": qb_curr,
                    "المجموعة/الصف": qb_grade,
                    "المادة": qb_subject,
                    "نوع_السؤال": qb_type,
                    "بيانات_السؤال_JSON": json.dumps(q_payload, ensure_ascii=False)
                }
                st.session_state.question_bank_df = pd.concat([st.session_state.question_bank_df, pd.DataFrame([new_qbank_row])], ignore_index=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                st.session_state.qb_q_img = ""
                st.session_state.qb_ver += 1
                st.success("✓ تم حفظ وإدراج السؤال في بنك الأسئلة بنجاح!")

    st.write("---")
    st.markdown("### 📋 الأسئلة الحالية في بنك الأسئلة:")
    qbf_df = st.session_state.question_bank_df
    if qbf_df.empty:
        st.info("لا توجد أسئلة مضافة في بنك الأسئلة بعد.")
    else:
        for qbi, qbr in qbf_df.iterrows():
            raw_json_data = qbr.get("بيانات_السؤال_JSON", "{}")
            try:
                q_data = json.loads(raw_json_data) if pd.notnull(raw_json_data) and str(raw_json_data).strip() else {}
            except Exception:
                q_data = {}

            if not q_data:
                continue

            with st.expander(f"[{qbr['المجموعة/الصف']}] — {q_data.get('type', 'اختيار من متعدد')} (الدرجة: {q_data.get('points', 1.0)})"):
                st.write(f"**نص السؤال:** {q_data.get('text', '')}")
                if q_data.get("q_img"):
                    st.image(f"data:image/jpeg;base64,{q_data['q_img']}", width=300)
                if q_data.get("type") == "اختيار من متعدد":
                    st.write(f"أ) {q_data.get('opt1','')} | ب) {q_data.get('opt2','')} | ج) {q_data.get('opt3','')} | د) {q_data.get('opt4','')}")
                    corr_idx = int(q_data.get('correct', 1)) - 1
                    corr_letter = ['أ', 'ب', 'ج', 'د'][corr_idx] if 0 <= corr_idx < 4 else 'أ'
                    st.write(f"<b>الإجابة الصحيحة:</b> الخيار ({corr_letter})")
                
                if st.button(f"حذف هذا السؤال 🗑️", key=f"del_qb_{qbi}"):
                    st.session_state.question_bank_df = qbf_df.drop(qbi).reset_index(drop=True)
                    save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                    st.warning("تم حذف السؤال.")
                    st.rerun()

elif t_page == "videos":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("🎥 إدارة ورفع الفيديوهات التعليمية للطلاب:")
    with st.form("upload_video_form", clear_on_submit=True):
        vid_title = st.text_input("عنوان الفيديو / الدرس:")
        vid_curr = st.selectbox("المنهج الدراسي / الدولة:", list(CURRICULUM_DATA.keys()), key="vid_c")
        vid_grade = st.selectbox("المرحلة / الصف الدراسي المستهدف:", CURRICULUM_DATA[vid_curr], key="vid_g")
        
        vid_upload_file = st.file_uploader("رفع ملف الفيديو (MP4 أو ما شابه):", type=["mp4", "mov", "avi", "mkv", "webm"])
        vid_link_input = st.text_input("أو ضع رابط فيديو (يوتيوب أو رابط مباشر):", placeholder="https://www.youtube.com/watch?v=...")

        if st.form_submit_button("💾 حفظ ونشر الفيديو للطالب"):
            if not vid_title.strip():
                st.error("يرجى كتابة عنوان الفيديو.")
            else:
                v_bytes_str = ""
                if vid_upload_file is not None:
                    file_bytes = vid_upload_file.read()
                    v_bytes_str = base64.b64encode(file_bytes).decode()

                new_vid = {
                    "معرف_الفيديو": f"VID_{datetime.now().strftime('%Y%m%d%H%M%S')}",
                    "عنوان_الفيديو": vid_title.strip(),
                    "المنهج/الدولة": vid_curr,
                    "المجموعة/الصف": vid_grade,
                    "رابط_الفيديو": vid_link_input.strip() if vid_link_input else "",
                    "فيديو_base64": v_bytes_str,
                    "تاريخ_الرفع": str(date.today())
                }
                st.session_state.videos_df = pd.concat([st.session_state.videos_df, pd.DataFrame([new_vid])], ignore_index=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                st.success(f"✓ تم نشر الفيديو ({vid_title}) بنجاح للمرحلة ({vid_grade})!")

    st.write("---")
    st.markdown("### 📋 الفيديوهات المنشورة مسبقاً:")
    v_df_state = st.session_state.videos_df
    if v_df_state.empty:
        st.info("لا توجد فيديوهات منشورة حالياً.")
    else:
        st.dataframe(v_df_state[["عنوان_الفيديو", "المنهج/الدولة", "المجموعة/الصف", "تاريخ_الرفع"]], use_container_width=True)
        with st.expander("🗑️ حذف فيديو منشور"):
            del_vid_opts = {vi: f"[{vr['المجموعة/الصف']}] {vr['عنوان_الفيديو']}" for vi, vr in v_df_state.iterrows()}
            sel_del_vid = st.selectbox("اختر الفيديو المراد حذفه:", options=list(del_vid_opts.keys()), format_func=lambda x: del_vid_opts[x], key="sel_del_vid")
            if st.button("🚨 تأكيد حذف الفيديو المختار"):
                st.session_state.videos_df = v_df_state.drop(sel_del_vid).reset_index(drop=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                st.success("✓ تم حذف الفيديو بنجاح!")
                st.rerun()

    # --- قسم متابعة تعليقات وأسئلة الطلاب على الفيديوهات ---
    st.write("---")
    st.markdown("### 💬 متابعة تعليقات وأسئلة الطلاب على الفيديوهات:")
    vc_state = st.session_state.video_comments_df
    if vc_state.empty:
        st.info("لا توجد تعليقات أو أسئلة من الطلاب على الفيديوهات حتى الآن.")
    else:
        for c_idx, c_row in vc_state.iterrows():
            st.markdown(f"""
                <div style="background:{card_bg}; border:1px solid {card_border}; border-radius:10px; padding:15px; margin-bottom:12px;">
                    <p style="margin:0; color:#10b981; font-size:16px;"><b>درس: {c_row['عنوان_الفيديو']}</b></p>
                    <p style="margin:4px 0; color:#0284c7; font-size:14px;">الطالب: <b>{c_row['اسم الطالب']}</b> — <span style="font-size:12px; opacity:0.7;">{c_row['التاريخ_والوقت']}</span></p>
                    <p style="margin:8px 0 0 0; font-size:16px;">{c_row['نص_التعليق']}</p>
                </div>
            """, unsafe_allow_html=True)
            if st.button(f"🗑️ حذف هذا التعليق", key=f"del_vcomm_{c_idx}"):
                st.session_state.video_comments_df = vc_state.drop(c_idx).reset_index(drop=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                st.success("تم حذف التعليق.")
                st.rerun()

elif t_page == "abqary":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("💡 إدارة امتحانات ونتائج موقع عبقري:")
    with st.form("upload_abqary_form", clear_on_submit=True):
        ab_title = st.text_input("عنوان امتحان عبقري:")
        ab_curr = st.selectbox("المنهج الدراسي / الدولة:", list(CURRICULUM_DATA.keys()), key="ab_c")
        ab_grade = st.selectbox("المرحلة / الصف الدراسي المستهدف:", CURRICULUM_DATA[ab_curr], key="ab_g")
        ab_link = st.text_input("رابط الامتحان على موقع عبقري (اختياري):", placeholder="https://abqary.com/exam/...")
        ab_html = st.text_area(
            "كود HTML للاختبار الإلكتروني (اختياري):",
            height=280,
            placeholder="الصق هنا كود HTML الكامل للاختبار الذي تريد أن يظهر للطالب داخل المنصة..."
        )
        st.caption("يمكنك استخدام الرابط أو كود HTML أو الاثنين معاً. إذا وضعت كود HTML سيظهر الاختبار تفاعلياً داخل صفحة الطالب.")
        res_link = st.text_input("رابط النتيجة (اختياري):", placeholder="https://abqary.com/result/...")
        secret_pass = st.text_input("الرقم السري لإظهار النتيجة للطالب:", placeholder="مثال: 1234 أو كود خاص")

        if st.form_submit_button("💾 حفظ ونشر امتحان عبقري للطالب"):
            if not ab_title.strip() or (not ab_link.strip() and not ab_html.strip()):
                st.error("يرجى كتابة عنوان الامتحان وإضافة رابط عبقري أو كود HTML للاختبار على الأقل.")
            else:
                new_ab = {
                    "معرف_عبقري": f"ABQ_{datetime.now().strftime('%Y%m%d%H%M%S')}",
                    "عنوان_الإمتحان": ab_title.strip(),
                    "المنهج/الدولة": ab_curr,
                    "المجموعة/الصف": ab_grade,
                    "رابط_الإمتحان": ab_link.strip() if ab_link else "",
                    "كود_HTML": ab_html.strip() if ab_html else "",
                    "رابط_النتيجة": res_link.strip() if res_link else "",
                    "الرقم_السري_للنتيجة": secret_pass.strip(),
                    "تاريخ_النشر": str(date.today())
                }
                st.session_state.abqary_df = pd.concat([st.session_state.abqary_df, pd.DataFrame([new_ab])], ignore_index=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                st.success(f"✓ تم نشر امتحان عبقري ({ab_title}) بنجاح!")

    st.write("---")
    st.markdown("### 📋 امتحانات عبقري المنشورة مسبقاً:")
    ab_df_state = st.session_state.abqary_df
    if ab_df_state.empty:
        st.info("لا توجد امتحانات عبقري منشورة حالياً.")
    else:
        ab_display = ab_df_state.copy()
        ab_display["نوع الاختبار"] = ab_display.apply(
            lambda r: "HTML تفاعلي + رابط" if str(r.get("كود_HTML", "") or "").strip() and str(r.get("رابط_الإمتحان", "") or "").strip() else ("HTML تفاعلي" if str(r.get("كود_HTML", "") or "").strip() else "رابط عبقري"),
            axis=1
        )
        st.dataframe(ab_display[["عنوان_الإمتحان", "المجموعة/الصف", "نوع الاختبار", "تاريخ_النشر"]], use_container_width=True)
        with st.expander("🗑️ حذف امتحان عبقري"):
            del_ab_opts = {ai: f"[{ar['المجموعة/الصف']}] {ar['عنوان_الإمتحان']}" for ai, ar in ab_df_state.iterrows()}
            sel_del_ab = st.selectbox("اختر الامتحان المراد حذفه:", options=list(del_ab_opts.keys()), format_func=lambda x: del_ab_opts[x], key="sel_del_ab")
            if st.button("🚨 تأكيد حذف امتحان عبقري المختار"):
                st.session_state.abqary_df = ab_df_state.drop(sel_del_ab).reset_index(drop=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                st.success("✓ تم حذف الامتحان بنجاح!")
                st.rerun()

elif t_page == "grades":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("📈 سجل درجات ونقاط اختبارات الطلاب:")
    exam_assessments = st.session_state.assessments_df[
        st.session_state.assessments_df["النوع"].astype(str).str.contains("اختبار|كويز", na=False)
    ].copy()

    if exam_assessments.empty:
        st.info("لا توجد نتائج اختبارات مرصودة للطلاب حتى الآن.")
    else:
        st.dataframe(exam_assessments[["التاريخ", "اسم الطالب", "عنوان التكليف", "الدرجة المحصلة", "الدرجة العظمى", "حالة التسليم", "ملاحظات وتوجيهات"]], use_container_width=True)

        with st.expander("🗑️ حذف نتيجة امتحان لطالب محدد"):
            del_exam_opts = {i: f"{r['اسم الطالب']} - {r['عنوان التكليف']} ({r['التاريخ']})" for i, r in exam_assessments.iterrows()}
            sel_del_exam_idx = st.selectbox("اختر السجل المراد حذفه:", options=list(del_exam_opts.keys()), format_func=lambda x: del_exam_opts[x], key="sel_del_exam")
            if st.button("🚨 تأكيد حذف سجل الامتحان المختار"):
                real_idx = exam_assessments.loc[sel_del_exam_idx].name
                st.session_state.assessments_df = st.session_state.assessments_df.drop(real_idx).reset_index(drop=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                st.success("✓ تم حذف النتيجة بنجاح!")
                st.rerun()

        buf_grades = io.BytesIO()
        with pd.ExcelWriter(buf_grades, engine="openpyxl") as writer:
            exam_assessments.to_excel(writer, sheet_name="Exam_Grades", index=False)
        
        c_ex_dl1, c_ex_dl2 = st.columns(2)
        with c_ex_dl1:
            st.download_button(
                label="📥 تصدير درجات الطلاب لملف Excel",
                data=buf_grades.getvalue(),
                file_name="درجات_الاختبارات_للطلاب.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with c_ex_dl2:
            grades_pdf_html = """<!DOCTYPE html>
            <html dir="rtl" lang="ar">
            <head><meta charset="utf-8"><title>سجل درجات الاختبارات</title></head>
            <body style="font-family: Arial; padding: 25px;" onload="window.print()">
                <h2>سجل درجات ونقاط اختبارات الطلاب - م/ محمد غنيم</h2>
                <table border="1" style="width:100%; border-collapse:collapse; text-align:center; margin-top:20px;">
                    <tr style="background:#f1f5f9;">
                        <th style="padding:10px;">التاريخ</th>
                        <th style="padding:10px;">اسم الطالب</th>
                        <th style="padding:10px;">عنوان الاختبار</th>
                        <th style="padding:10px;">الدرجة</th>
                        <th style="padding:10px;">الحالة</th>
                    </tr>
            """
            for _, r in exam_assessments.iterrows():
                grades_pdf_html += f"""
                    <tr>
                        <td style="padding:8px;">{r['التاريخ']}</td>
                        <td style="padding:8px;">{r['اسم الطالب']}</td>
                        <td style="padding:8px;">{r['عنوان التكليف']}</td>
                        <td style="padding:8px;">{r['الدرجة المحصلة']} / {r['الدرجة العظمى']}</td>
                        <td style="padding:8px;">{r['حالة التسليم']}</td>
                    </tr>
                """
            grades_pdf_html += "</table></body></html>"

            st.download_button(
                label="🖨️ طباعة وتصدير درجات الطلاب PDF",
                data=grades_pdf_html.encode("utf-8"),
                file_name="سجل_درجات_الاختبارات.html",
                mime="application/octet-stream"
            )

elif t_page == "essays":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("📝 كنترول تصحيح إجابات الطلاب المقالية وصور الحل:")
    essays = st.session_state.essays_df

    if essays.empty:
        st.info("لا توجد إجابات مقالية مرسلة من الطلاب بعد.")
    else:
        for es_idx, es_row in essays.iterrows():
            st_name = str(es_row.get("اسم الطالب", ""))
            ex_name = str(es_row.get("عنوان الامتحان", ""))
            q_num = es_row.get("رقم السؤال", 1)
            q_text = str(es_row.get("نص السؤال", ""))
            ans_txt = es_row.get("إجابة الطالب النصية", "")
            img_b64_ans = es_row.get("صورة الحل_base64", "")
            q_max = float(es_row.get("درجة السؤال", 1.0))
            status = str(es_row.get("حالة التصحيح", "قيد التصحيح من المعلم"))
            
            raw_fb = es_row.get("ملاحظات المعلم", "")
            t_feedback = str(raw_fb) if pd.notnull(raw_fb) and str(raw_fb) != "nan" else ""

            with st.expander(f"📌 حل الطالب: {st_name} — {ex_name} (س {q_num}) | [{status}]"):
                st.markdown(f"**نص السؤال:** {q_text}")
                st.markdown(f"**إجابة الطالب المكتوبة:**")
                st.write(str(ans_txt) if pd.notnull(ans_txt) and str(ans_txt).strip() and str(ans_txt) != "nan" else "لم يكتب نصاً (أرفق صورة بالأسفل)")

                if pd.notnull(img_b64_ans) and str(img_b64_ans).strip() and str(img_b64_ans) != "nan":
                    st.markdown("**📷 صورة خطوات الحل المرفوعة من الطالب:**")
                    st.image(f"data:image/jpeg;base64,{img_b64_ans}", use_container_width=True)

                with st.form(f"grade_essay_form_{es_idx}"):
                    c_g1, c_g2 = st.columns(2)
                    with c_g1:
                        awarded_score = st.number_input("الدرجة المستحقة:*", min_value=0.0, max_value=q_max, value=float(es_row.get("الدرجة المرصودة", 0.0)) if pd.notnull(es_row.get("الدرجة المرصودة")) else 0.0, step=0.5)
                    with c_g2:
                        teacher_feedback = st.text_input("ملاحظات المعلم وتوجيهه للطالب:", value=t_feedback)

                    c_sub1, c_sub2 = st.columns(2)
                    with c_sub1:
                        submit_grade = st.form_submit_button("💾 اعتماد ورصد الدرجة للطالب")
                    with c_sub2:
                        delete_essay = st.form_submit_button("🗑️ مسح وإلغاء هذه الإجابة المقالية")

                    if submit_grade:
                        essays.at[es_idx, "الدرجة المرصودة"] = awarded_score
                        essays.at[es_idx, "حالة التصحيح"] = "تم التصحيح والاعتماد"
                        essays.at[es_idx, "ملاحظات المعلم"] = str(teacher_feedback).strip()

                        match_ass = st.session_state.assessments_df[
                            (st.session_state.assessments_df["اسم الطالب"].astype(str).str.strip() == st_name.strip())
                            & (st.session_state.assessments_df["عنوان التكليف"].astype(str).str.contains(ex_name, na=False))
                        ]
                        if not match_ass.empty:
                            ass_idx = match_ass.index[-1]
                            current_score = float(st.session_state.assessments_df.at[ass_idx, "الدرجة المحصلة"])
                            st.session_state.assessments_df.at[ass_idx, "الدرجة المحصلة"] = current_score + awarded_score
                            st.session_state.assessments_df.at[ass_idx, "ملاحظات وتوجيهات"] = f"تم تصحيح المقالي: +{awarded_score} درجة. {teacher_feedback}"

                        st.session_state.essays_df = essays
                        save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                        st.success(f"✓ تم رصد درجة الطالب ({st_name}) بنجاح!")
                        st.rerun()

                    if delete_essay:
                        st.session_state.essays_df = essays.drop(es_idx).reset_index(drop=True)
                        save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                        st.warning("⚠️ تم مسح إجابة المقالي بنجاح.")
                        st.rerun()

elif t_page == "chat":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("💬 صندوق محادثات الطلاب والرد على الأسئلة:")
    all_students_with_chat = sorted(list(set(
        [s for s in st.session_state.users_df["اسم الطالب"].dropna().unique() if str(s).strip()]
        + [s for s in st.session_state.messages_df["اسم الطالب"].dropna().unique() if str(s).strip()]
    )))

    if not all_students_with_chat:
        st.info("لا توجد رسائل واردة بعد.")
    else:
        selected_chat_student = st.selectbox("اختر الطالب لمشاهدة محادثته والرد عليه:", all_students_with_chat)
        if selected_chat_student:
            student_thread = st.session_state.messages_df[st.session_state.messages_df["اسم الطالب"].astype(str).str.strip() == selected_chat_student.strip()].copy()
            st.write(f"### سجل المحادثة مع الطالب: **{selected_chat_student}**")

            with st.container():
                if not student_thread.empty:
                    for _, m in student_thread.iterrows():
                        sender = m["المرسل"]
                        t_stamp = m.get("التاريخ_والوقت", "")
                        content = m.get("نص الرسالة", "")
                        img_data = m.get("الصورة_base64", "")
                        if sender == "student":
                            st.markdown(f"<div class='chat-bubble-student'><b>الطالب ({t_stamp}):</b><br>{content}</div>", unsafe_allow_html=True)
                        else:
                            st.markdown(f"<div class='chat-bubble-teacher'><b>أنت (البشمهندس) ({t_stamp}):</b><br>{content}</div>", unsafe_allow_html=True)
                        if pd.notnull(img_data) and str(img_data).strip():
                            st.image(f"data:image/jpeg;base64,{img_data}", width=350)

            with st.form("teacher_reply_form", clear_on_submit=True):
                reply_text = st.text_area("اكتب ردك أو التوجيه للطالب:")
                if st.form_submit_button("📤 إرسال الرد إلى الطالب"):
                    if reply_text.strip():
                        new_rep = {
                            "التاريخ_والوقت": datetime.now().strftime("%Y-%m-%d %H:%M"),
                            "اسم الطالب": selected_chat_student.strip(),
                            "المرسل": "teacher", "نص الرسالة": reply_text.strip(), "الصورة_base64": "",
                        }
                        st.session_state.messages_df = pd.concat([st.session_state.messages_df, pd.DataFrame([new_rep])], ignore_index=True)
                        save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                        st.success("تم إرسال الرد للبشمهندس بنجاح!")
                        st.rerun()

elif t_page == "students":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("👥 بطاقات الطلاب المسجلين والتحكم الكامل:")
    
    bookings_df_state = st.session_state.bookings_df
    if not bookings_df_state.empty:
        st.markdown("#### 📅 طلبات حجز الدروس أونلاين الواردة:")
        st.dataframe(bookings_df_state, use_container_width=True)
        with st.expander("🗑️ حذف أو إدارة طلب حجز"):
            del_book_opts = {bi: f"{br['اسم الطالب']} - {br['المرحلة_الصف']} ({br['رقم_الهاتف']})" for bi, br in bookings_df_state.iterrows()}
            sel_del_book = st.selectbox("اختر الطلب:", options=list(del_book_opts.keys()), format_func=lambda x: del_book_opts[x], key="sel_del_book")
            if st.button("🚨 تأكيد حذف طلب الحجز المختار"):
                st.session_state.bookings_df = bookings_df_state.drop(sel_del_book).reset_index(drop=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                st.success("✓ تم حذف طلب الحجز بنجاح!")
                st.rerun()
        st.write("---")

    bank_req_state = st.session_state.bank_requests_df
    if not bank_req_state.empty:
        st.markdown("#### 💳 طلبات اشتراكات بنك الأسئلة الواردة (تأكيد الدفع):")
        for br_idx, br_row in bank_req_state.iterrows():
            st_req_name = br_row["اسم الطالب"]
            req_phone = br_row["رقم_الهاتف"]
            req_otp = br_row["كود_OTP"]
            req_status = br_row["حالة_الدفع"]
            req_receipt = br_row["إيصال_الدفع_base64"]

            with st.expander(f"طلب اشتراك بنك الأسئلة: الطالب ({st_req_name}) — الهاتف: ({req_phone}) — كود OTP: ({req_otp}) | الحالة: [{req_status}]"):
                if pd.notnull(req_receipt) and str(req_receipt).strip() and str(req_receipt) != "nan":
                    st.markdown("**📷 صورة إيصال التحويل المرفق:**")
                    st.image(f"data:image/jpeg;base64,{req_receipt}", width=350)
                
                c_bk1, c_bk2 = st.columns(2)
                with c_bk1:
                    if st.button(f"✅ تأكيد وتفعيل الاشتراك للطالب {st_req_name}", key=f"confirm_bank_{br_idx}"):
                        bank_req_state.at[br_idx, "حالة_الدفع"] = "مؤكد ومفعل"
                        st.session_state.users_df.loc[st.session_state.users_df["اسم الطالب"].astype(str).str.strip() == str(st_req_name).strip(), "حالة_الاشتراك_البنك"] = "مشترك"
                        save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                        st.success(f"✓ تم تفعيل اشتراك بنك الأسئلة للطالب {st_req_name} بنجاح!")
                        st.rerun()
                with c_bk2:
                    if st.button(f"🗑️ حذف طلب الاشتراك", key=f"del_bank_req_{br_idx}"):
                        st.session_state.bank_requests_df = bank_req_state.drop(br_idx).reset_index(drop=True)
                        save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                        st.warning("تم حذف طلب الاشتراك.")
                        st.rerun()
        st.write("---")

    all_known_students = sorted(list(set(
        [s for s in st.session_state.users_df["اسم الطالب"].dropna().unique() if str(s).strip()]
        + [s for s in st.session_state.sessions_df["اسم الطالب"].dropna().unique() if str(s).strip()]
        + [s for s in st.session_state.assessments_df["اسم الطالب"].dropna().unique() if str(s).strip()]
        + [s for s in st.session_state.weekly_schedule_df["اسم الطالب"].dropna().unique() if str(s).strip()]
        + [s for s in st.session_state.online_schedule_df["اسم الطالب"].dropna().unique() if str(s).strip()]
    )))

    if all_known_students:
        st.markdown("### 📒 كشف الطلاب المسجلين")
        roster_html = build_student_roster_html(all_known_students, "كشف الطلاب المسجلين - م/ محمد غنيم")
        roster_pdf = html_to_pdf_bytes(roster_html)
        rc1, rc2 = st.columns(2)
        with rc1:
            if roster_pdf:
                st.download_button("📄 طباعة كشف المسجلين PDF", roster_pdf, file_name="كشف_الطلاب_المسجلين.pdf", mime="application/pdf", key="students_roster_pdf")
            else:
                st.download_button("🖨️ طباعة كشف المسجلين", roster_html.encode("utf-8"), file_name="كشف_الطلاب_المسجلين.html", mime="text/html", key="students_roster_html")
        with rc2:
            roster_buf = io.BytesIO()
            with pd.ExcelWriter(roster_buf, engine="openpyxl") as writer:
                st.session_state.users_df.to_excel(writer, sheet_name="Students", index=False)
            st.download_button("📥 تصدير كشف المسجلين Excel", roster_buf.getvalue(), file_name="كشف_الطلاب_المسجلين.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="students_roster_excel")
        st.write("---")
        with st.expander("🗑️ حذف طالب محدد نهائياً من كافة السجلات"):
            del_selected_st = st.selectbox("اختر الطالب المراد حذفه نهائياً:", all_known_students, key="del_box_select")
            if st.button("🚨 تأكيد حذف هذا الطالب نهائياً", key="btn_confirm_del_box"):
                delete_student_completely(del_selected_st)
                st.success(f"✓ تم مسح الطالب ({del_selected_st}) نهائياً!")
                st.rerun()

    st.write("---")
    if not all_known_students:
        st.info("لا توجد طلاب مسجلون حالياً.")
    else:
        for idx, st_name in enumerate(all_known_students):
            u_row = st.session_state.users_df[st.session_state.users_df["اسم الطالب"].astype(str).str.strip() == st_name.strip()]
            s_row = st.session_state.sessions_df[st.session_state.sessions_df["اسم الطالب"].astype(str).str.strip() == st_name.strip()]

            st_curr = u_row.iloc[0].get("المنهج/الدولة", "-") if not u_row.empty else (s_row.iloc[-1].get("المنهج/الدولة", "-") if not s_row.empty else "-")
            st_grade = u_row.iloc[0].get("المجموعة/الصف", "-") if not u_row.empty else (s_row.iloc[-1].get("المجموعة/الصف", "-") if not s_row.empty else "-")
            st_date = u_row.iloc[0].get("تاريخ التسجيل", "-") if not u_row.empty else "-"
            st_pass = u_row.iloc[0].get("كلمة المرور", "-") if not u_row.empty else "-"
            st_phone = u_row.iloc[0].get("رقم الهاتف", "-") if not u_row.empty else "-"
            is_banned = u_row.iloc[0].get("الحالة_حظر") == "محظور" if not u_row.empty else False
            bank_sub_state = u_row.iloc[0].get("حالة_الاشتراك_البنك", "غير مشترك") if not u_row.empty else "غير مشترك"

            status_badge = "🚫 محظور" if is_banned else "✅ نشط"
            badge_color = "#dc2626" if is_banned else "#16a34a"
            bank_badge = "📚 مشترك بالبنك" if bank_sub_state == "مشترك" else "🔒 غير مشترك بالبنك"

            st_sessions_card = st.session_state.sessions_df[st.session_state.sessions_df["اسم الطالب"].astype(str).str.strip() == st_name.strip()]
            total_st_sessions = len(st_sessions_card)
            attended_st_sessions = len(st_sessions_card[st_sessions_card["الحالة"] == "حاضر"])
            total_st_cost = st_sessions_card["سعر الحصة"].astype(float, errors="ignore").sum(numeric_only=True) if not st_sessions_card.empty else 0.0

            with st.container():
                col_c1, col_c2, col_c3, col_c4 = st.columns([3, 2, 2, 2])
                with col_c1:
                    st.markdown(f"""
                        <h4 style="margin: 0; color: #0052cc;">{st_name}</h4>
                        <p style="margin: 3px 0; font-size: 14px; font-weight: 900;">{st_curr} — {st_grade}</p>
                        <p style="margin: 0; font-size: 13px; color: #64748b;">الهاتف: {st_phone} | كلمة المرور: <b>{st_pass}</b></p>
                        <p style="margin: 2px 0 0 0; font-size: 13px; color: #0284c7;"><b>{bank_badge}</b></p>
                    """, unsafe_allow_html=True)
                with col_c2:
                    st.markdown(f"""
                        <p style="margin:0; font-size:14px; font-weight:900;">حضور: <b>{attended_st_sessions}/{total_st_sessions}</b></p>
                        <p style="margin:0; font-size:14px; font-weight:900; color:#b91c1c;">إجمالي السعر: <b>{total_st_cost:,.1f}</b></p>
                    """, unsafe_allow_html=True)
                with col_c3:
                    st_sessions_html = f"""<!DOCTYPE html>
                    <html dir="rtl" lang="ar">
                    <head><meta charset="utf-8"><title>تقرير الحضور والأسعار - {st_name}</title></head>
                    <body style="font-family: Arial; padding: 25px;" onload="window.print()">
                        <h2>تقرير الحضور وحساب الحصص - م/ محمد غنيم</h2>
                        <p><b>اسم الطالب:</b> {st_name} | <b>المجموعة/الصف:</b> {st_grade}</p>
                        <p><b>إجمالي الحصص الحضورية:</b> {attended_st_sessions} من {total_st_sessions} | <b>المبلغ الإجمالي المستحق:</b> {total_st_cost:,.1f}</p>
                        <hr>
                        <table border="1" style="width:100%; border-collapse:collapse; text-align:center; margin-top:15px;">
                            <tr style="background:#f1f5f9;">
                                <th style="padding:8px;">التاريخ</th>
                                <th style="padding:8px;">الحالة</th>
                                <th style="padding:8px;">سعر الحصة</th>
                                <th style="padding:8px;">المستوى</th>
                                <th style="padding:8px;">ملاحظات</th>
                            </tr>
                    """
                    for _, s_row in st_sessions_card.iterrows():
                        st_sessions_html += f"""
                            <tr>
                                <td style="padding:6px;">{s_row['التاريخ']}</td>
                                <td style="padding:6px;">{s_row['الحالة']}</td>
                                <td style="padding:6px;">{s_row['سعر الحصة']}</td>
                                <td style="padding:6px;">{s_row['مستوى الطالب']}</td>
                                <td style="padding:6px;">{s_row['ملاحظات']}</td>
                            </tr>
                        """
                    st_sessions_html += "</table></body></html>"

                    st.download_button(
                        label="🖨️ طباعة الحضور والأسعار",
                        data=st_sessions_html.encode("utf-8"),
                        file_name=f"حضور_وأسعار_{st_name}.html",
                        mime="application/octet-stream",
                        key=f"print_att_{idx}"
                    )
                with col_c4:
                    if st.button("📊 درجات الطالب", key=f"btn_grades_{idx}"):
                        st.session_state[f"show_grades_{idx}"] = not st.session_state.get(f"show_grades_{idx}", False)

                # بيانات ولي الأمر قابلة للتعديل وتظهر في التقرير والكشوف
                if not u_row.empty:
                    with st.expander("👨‍👩‍👦 بيانات ولي الأمر", expanded=False):
                        with st.form(f"parent_data_form_{idx}"):
                            p_name = st.text_input("اسم ولي الأمر", value=str(u_row.iloc[0].get("اسم ولي الأمر", "")), key=f"parent_name_{idx}")
                            p_phone = st.text_input("رقم ولي الأمر", value=str(u_row.iloc[0].get("رقم ولي الأمر", "")), key=f"parent_phone_{idx}")
                            if st.form_submit_button("💾 حفظ بيانات ولي الأمر"):
                                ui = u_row.index[0]
                                st.session_state.users_df.at[ui, "اسم ولي الأمر"] = p_name.strip()
                                st.session_state.users_df.at[ui, "رقم ولي الأمر"] = p_phone.strip()
                                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                                st.success("✓ تم حفظ بيانات ولي الأمر")
                                st.rerun()

                if st.session_state.get(f"show_grades_{idx}", False):
                    st.markdown(f"**سجل درجات الطالب: {st_name}**")
                    st_grades_df = st.session_state.assessments_df[st.session_state.assessments_df["اسم الطالب"].astype(str).str.strip() == st_name.strip()]
                    if st_grades_df.empty:
                        st.info("لا توجد درجات مرصودة لهذا الطالب حتى الآن.")
                    else:
                        st.dataframe(st_grades_df[["التاريخ", "النوع", "عنوان التكليف", "الدرجة المحصلة", "الدرجة العظمى", "حالة التسليم", "ملاحظات وتوجيهات"]], use_container_width=True)

                        with st.form(f"clear_st_grades_form_{idx}"):
                            del_grade_choice = st.selectbox("اختر النتيجة المراد مسحها لهذا الطالب:", options=list(st_grades_df.index), format_func=lambda x: f"{st_grades_df.loc[x, 'عنوان التكليف']} ({st_grades_df.loc[x, 'التاريخ']})")
                            if st.form_submit_button("🗑️ مسح هذه الدرجة المحددة للطالب"):
                                st.session_state.assessments_df = st.session_state.assessments_df.drop(del_grade_choice).reset_index(drop=True)
                                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                                st.success("✓ تم مسح الدرجة بنجاح!")
                                st.rerun()

                if is_banned:
                    if st.button("فك الحظر 🔓", key=f"unban_{idx}"):
                        st.session_state.users_df.loc[st.session_state.users_df["اسم الطالب"] == st_name, "الحالة_حظر"] = "نشط"
                        save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                        st.success(f"تم فك حظر {st_name}!")
                        st.rerun()
                else:
                    if st.button("حظر 🚫", key=f"ban_{idx}"):
                        if not u_row.empty:
                            st.session_state.users_df.loc[st.session_state.users_df["اسم الطالب"] == st_name, "الحالة_حظر"] = "محظور"
                            save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                            st.warning(f"تم حظر {st_name}.")
                            st.rerun()
                
                if st.button("حذف نهائي للطالب 🗑️", key=f"del_card_btn_{idx}"):
                    delete_student_completely(st_name)
                    st.success(f"✓ تم مسح الطالب ({st_name}) نهائياً!")
                    st.rerun()
                st.write("---")

elif t_page == "add_session":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("إدخال بيانات الحصة")
    t_curriculum = st.selectbox("اختر المنهج الدراسي:", list(CURRICULUM_DATA.keys()), key="teacher_curr_select")
    t_grades = CURRICULUM_DATA[t_curriculum]

    with st.form("teacher_entry_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            session_date = st.date_input("تاريخ الحصة", value=date.today())
            session_names = sorted(list(set([str(x).strip() for x in st.session_state.users_df["اسم الطالب"].dropna().unique() if str(x).strip()] + [str(x).strip() for x in st.session_state.weekly_schedule_df["اسم الطالب"].dropna().unique() if str(x).strip()] + [str(x).strip() for x in st.session_state.online_schedule_df["اسم الطالب"].dropna().unique() if str(x).strip()])))
            prefill_session = str(st.session_state.get("prefill_student", "")).strip()
            if session_names:
                student_name = st.selectbox("اسم الطالب", session_names, index=(session_names.index(prefill_session) if prefill_session in session_names else 0), key="session_student_select")
            else:
                student_name = st.text_input("اسم الطالب", value=prefill_session, placeholder="مثال: أحمد محمد")
            group_name = st.selectbox("المرحلة / الصف الدراسي:", t_grades)
            status = st.selectbox("حالة الحضور", ["حاضر", "غائب", "متأخر", "بعذر"])
        with col2:
            price = st.number_input("سعر الحصة", min_value=0.0, step=10.0, value=100.0)
            total_sessions = st.number_input("الحصص المنفذة حتى الآن", min_value=1, step=1, value=1)
            payment_type = st.selectbox("نظام الدفع", ["اشتراك شهري", "مقدم", "مؤخر (بعد الحصة)", "مؤجل"])
            student_level = st.selectbox("المستوى الدراسي", ["ممتاز ⭐⭐⭐", "جيد جداً ⭐⭐", "جيد ⭐", "متوسط", "يحتاج متابعة"])

        notes = st.text_area("ملاحظات الأداء والالتزام بالحصة", placeholder="أداء الطالب في الحصة...")
        if st.form_submit_button("💾 رصد وحفظ الحصة"):
            if not student_name.strip():
                st.error("يرجى كتابة اسم الطالب أولاً.")
            else:
                new_row = {
                    "التاريخ": str(session_date), "اسم الطالب": student_name.strip(),
                    "المنهج/الدولة": t_curriculum, "المجموعة/الصف": group_name,
                    "الحالة": status, "سعر الحصة": price, "عدد الحصص الكلي": total_sessions,
                    "نظام الدفع": payment_type, "مستوى الطالب": student_level, "ملاحظات": notes.strip(),
                }
                st.session_state.sessions_df = pd.concat([st.session_state.sessions_df, pd.DataFrame([new_row])], ignore_index=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                st.success(f"✓ تم حفظ سجل الحصة للطالب ({student_name}) بنجاح!")
                st.session_state.pop("prefill_student", None)
                st.session_state.pop("prefill_schedule_idx", None)

elif t_page == "add_hw":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("إضافة درجات الواجبات المنزلية والاختبارات يدوياً")
    all_registered_names = sorted(list(set(
        [s for s in st.session_state.users_df["اسم الطالب"].dropna().unique() if str(s).strip()]
        + [s for s in st.session_state.sessions_df["اسم الطالب"].dropna().unique() if str(s).strip()]
        + [s for s in st.session_state.weekly_schedule_df["اسم الطالب"].dropna().unique() if str(s).strip()]
        + [s for s in st.session_state.online_schedule_df["اسم الطالب"].dropna().unique() if str(s).strip()]
    )))

    with st.form("assessment_form", clear_on_submit=True):
        col_a1, col_a2 = st.columns(2)
        with col_a1:
            if all_registered_names:
                prefill_hw = str(st.session_state.get("prefill_student", ""))
                hw_index = all_registered_names.index(prefill_hw) if prefill_hw in all_registered_names else 0
                ass_student = st.selectbox("اختر الطالب:", all_registered_names, index=hw_index)
            else:
                ass_student = st.text_input("اسم الطالب:")
            ass_type = st.selectbox("نوع التكليف الأكاديمي:", ["واجب منزلي", "اختبار دوري", "كويز سريع", "مهمة إضافية"])
            ass_title = st.text_input("عنوان الدرس / التكليف:", placeholder="مثال: تمارين الهندسة صـ 25")
            ass_date = st.date_input("تاريخ الرصد:", value=date.today())
        with col_a2:
            ass_status = st.selectbox("حالة التسليم والحل:", ["تم التسليم كاملاً وبشكل ممتاز", "تم التسليم مع بعض الأخطاء", "تسليم جزئي / ناقص", "لم يتم التسليم (مقصّر)", "غائب عن الاختبار"])
            ass_score = st.number_input("الدرجة المحصلة:", min_value=0.0, step=0.5, value=10.0)
            ass_max = st.number_input("الدرجة العظمى:", min_value=1.0, step=1.0, value=10.0)
            ass_notes = st.text_input("ملاحظة وتوجيه ولي الأمر:", placeholder="توجيه مباشر يظهر في تقرير ولي الأمر...")

        if st.form_submit_button("💾 رصد التقييم وحفظه"):
            if not str(ass_student).strip() or not ass_title.strip():
                st.error("يرجى التأكد من اختيار اسم الطالب وكتابة عنوان التكليف.")
            else:
                new_ass = {
                    "التاريخ": str(ass_date), "اسم الطالب": str(ass_student).strip(),
                    "النوع": ass_type, "عنوان التكليف": ass_title.strip(),
                    "الدرجة المحصلة": ass_score, "الدرجة العظمى": ass_max,
                    "حالة التسليم": ass_status, "ملاحظات وتوجيهات": ass_notes.strip(),
                }
                st.session_state.assessments_df = pd.concat([st.session_state.assessments_df, pd.DataFrame([new_ass])], ignore_index=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                st.success(f"✓ تم رصد {ass_type} بنجاح للطالب ({ass_student})!")
                st.session_state.pop("prefill_student", None)

    st.write("---")
    st.markdown("### 🗑️ حذف رصد واجب / تقييم")
    assessments_current = st.session_state.assessments_df.copy()
    if assessments_current.empty:
        st.info("لا توجد سجلات واجبات أو تقييمات مسجلة للحذف.")
    else:
        assessment_indices = list(assessments_current.index)
        assessment_choices = []
        for idx, row in assessments_current.iterrows():
            assessment_choices.append(
                f"{idx} — {row.get('اسم الطالب','')} — {row.get('النوع','')} — {row.get('عنوان التكليف','')} — {row.get('التاريخ','')}"
            )
        selected_assessment = st.selectbox(
            "اختر الرصد الذي تريد مسحه:",
            assessment_choices,
            key="delete_assessment_select"
        )
        if st.button("🗑️ مسح الرصد المحدد", key="delete_assessment_button", type="secondary"):
            selected_pos = assessment_choices.index(selected_assessment)
            original_assessment_idx = assessment_indices[selected_pos]
            st.session_state.assessments_df = st.session_state.assessments_df.drop(index=original_assessment_idx).reset_index(drop=True)
            save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
            st.success("✓ تم مسح رصد الواجب / التقييم المحدد بنجاح.")
            st.rerun()

    st.dataframe(st.session_state.assessments_df, use_container_width=True)

elif t_page == "edit_records":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("مراجعة وتعديل بيانات الحصص")
    df = st.session_state.sessions_df

    if df.empty:
        st.info("لا توجد سجلات حصص مسجلة بعد.")
    else:
        record_options = {idx: f"[{row['التاريخ']}] - {row['اسم الطالب']} ({row.get('المنهج/الدولة', '')} | {row['المجموعة/الصف']}) - {row['الحالة']}" for idx, row in df.iterrows()}
        selected_idx = st.selectbox("اختر السجل المراد تعديله:", options=list(record_options.keys()), format_func=lambda x: record_options[x])
        selected_row = df.loc[selected_idx]

        try: curr_date = datetime.strptime(str(selected_row["التاريخ"]), "%Y-%m-%d").date()
        except Exception: curr_date = date.today()

        with st.form("edit_record_form"):
            c1, c2 = st.columns(2)
            with c1:
                edit_name = st.text_input("اسم الطالب:", value=str(selected_row["اسم الطالب"]))
                saved_curr = selected_row.get("المنهج/الدولة", "")
                curr_list = list(CURRICULUM_DATA.keys())
                curr_idx = curr_list.index(saved_curr) if saved_curr in curr_list else 0
                edit_curr = st.selectbox("المنهج / الدولة:", curr_list, index=curr_idx)
                edit_group = st.text_input("المرحلة / الصف الدراسي:", value=str(selected_row["المجموعة/الصف"]))
                edit_date = st.date_input("التاريخ:", value=curr_date)
                s_opts = ["حاضر", "غائب", "متأخر", "بعذر"]
                edit_status = st.selectbox("الحالة:", s_opts, index=(s_opts.index(selected_row["الحالة"]) if selected_row["الحالة"] in s_opts else 0))
            with c2:
                edit_price = st.number_input("سعر الحصة:", min_value=0.0, step=10.0, value=float(selected_row["سعر الحصة"]) if pd.notnull(selected_row["سعر الحصة"]) else 0.0)
                edit_sessions = st.number_input("عدد الحصص الكلي:", min_value=1, step=1, value=int(selected_row["عدد الحصص الكلي"]) if pd.notnull(selected_row["عدد الحصص الكلي"]) else 1)
                p_opts = ["اشتراك شهري", "مقدم", "مؤخر (بعد الحصة)", "مؤجل"]
                edit_pay = st.selectbox("نظام الدفع:", p_opts, index=(p_opts.index(selected_row["نظام الدفع"]) if selected_row["نظام الدفع"] in p_opts else 0))
                l_opts = ["ممتاز ⭐⭐⭐", "جيد جداً ⭐⭐", "جيد ⭐", "متوسط", "يحتاج متابعة", "قيد التقييم"]
                edit_level = st.selectbox("المستوى:", l_opts, index=(l_opts.index(selected_row["مستوى الطالب"]) if selected_row["مستوى الطالب"] in l_opts else 0))

            edit_notes = st.text_area("الملاحظات والتقييم:", value=str(selected_row["ملاحظات"]))
            b1, b2 = st.columns([1, 4])
            with b1: update_btn = st.form_submit_button("💾 تحديث البيانات")
            with b2: delete_btn = st.form_submit_button("🗑️ حذف السجل")

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
                st.session_state.sessions_df = df
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                st.success("✓ تم تحديث بيانات الحصة بنجاح!")
                st.rerun()

            if delete_btn:
                df = df.drop(selected_idx).reset_index(drop=True)
                st.session_state.sessions_df = df
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                st.warning("⚠️ تم حذف السجل.")
                st.rerun()

elif t_page == "all_records":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("📊 نظرة شاملة على السجلات وإحصائيات الطلاب")
    
    all_students_master = sorted(list(set(
        [str(s).strip() for s in st.session_state.users_df["اسم الطالب"].dropna().unique() if str(s).strip()]
        + [str(s).strip() for s in st.session_state.sessions_df["اسم الطالب"].dropna().unique() if str(s).strip()]
        + [str(s).strip() for s in st.session_state.assessments_df["اسم الطالب"].dropna().unique() if str(s).strip()]
        + [str(s).strip() for s in st.session_state.weekly_schedule_df["اسم الطالب"].dropna().unique() if str(s).strip()]
        + [str(s).strip() for s in st.session_state.online_schedule_df["اسم الطالب"].dropna().unique() if str(s).strip()]
    )))

    if not all_students_master:
        st.info("لا توجد بيانات طلاب مسجلة حتى الآن.")
    else:
        selected_master_student = st.selectbox("🔍 اختر طالباً لعرض بياناته التفصيلية:", options=all_students_master, key="master_student_sel")
        
        if selected_master_student:
            u_r = st.session_state.users_df[st.session_state.users_df["اسم الطالب"].astype(str).str.strip() == selected_master_student]
            s_r = st.session_state.sessions_df[st.session_state.sessions_df["اسم الطالب"].astype(str).str.strip() == selected_master_student]
            a_r = st.session_state.assessments_df[st.session_state.assessments_df["اسم الطالب"].astype(str).str.strip() == selected_master_student]

            reg_state = "نعم (مسجل على المنصة)" if not u_r.empty else "لا (مسجل يدويًا)"
            
            grade_val = "-"
            curr_val = "-"
            if not u_r.empty:
                grade_val = str(u_r.iloc[0].get("المجموعة/الصف", "-"))
                curr_val = str(u_r.iloc[0].get("المنهج/الدولة", "-"))
            elif not s_r.empty:
                grade_val = str(s_r.iloc[-1].get("المجموعة/الصف", "-"))
                curr_val = str(s_r.iloc[-1].get("المنهج/الدولة", "-"))

            total_sess = len(s_r)
            attended_sess = len(s_r[s_r["الحالة"] == "حاضر"])
            total_due = s_r["سعر الحصة"].astype(float, errors="ignore").sum(numeric_only=True) if not s_r.empty else 0.0

            st.markdown(f"""
                <div style="background:{card_bg}; border:2px solid #10b981; border-radius:12px; padding:20px; margin-bottom:10px;">
                    <h4 style="color:#059669; margin-top:0;">👤 ملف الطالب: {selected_master_student}</h4>
                    <p style="font-size:16px; margin:5px 0; color:#0f172a;"><b>حالة التسجيل:</b> {reg_state} | <b>المرحلة/الصف:</b> {grade_val} ({curr_val})</p>
                    <p style="font-size:16px; margin:5px 0; color:#0f172a;"><b>إجمالي الحصص:</b> {total_sess} (حاضر: {attended_sess}) | <b>إجمالي المبلغ المستحق:</b> <span style="color:#dc2626;">{total_due:,.1f} جنيه</span></p>
                </div>
            """, unsafe_allow_html=True)

            st_sessions_pdf = st.session_state.sessions_df[st.session_state.sessions_df["اسم الطالب"].astype(str).str.strip() == selected_master_student].copy()
            st_assessments_pdf = st.session_state.assessments_df[st.session_state.assessments_df["اسم الطالب"].astype(str).str.strip() == selected_master_student].copy()

            sess_rows_html = ""
            for _, sr in st_sessions_pdf.iterrows():
                sess_rows_html += f"<tr><td>{sr['التاريخ']}</td><td>{sr['الحالة']}</td><td>{sr['سعر الحصة']}</td><td>{sr['مستوى الطالب']}</td><td>{sr['ملاحظات']}</td></tr>"

            ass_rows_html = ""
            for _, ar in st_assessments_pdf.iterrows():
                ass_rows_html += f"<tr><td>{ar['التاريخ']}</td><td>{ar['النوع']}</td><td>{ar['عنوان التكليف']}</td><td>{ar['الدرجة المحصلة']} / {ar['الدرجة العظمى']}</td><td>{ar['حالة التسليم']}</td></tr>"

            single_student_pdf_html = f"""<!DOCTYPE html>
            <html dir="rtl" lang="ar">
            <head><meta charset="utf-8"><title>تقرير الطالب - {selected_master_student}</title></head>
            <body style="font-family: Arial; padding: 25px;" onload="window.print()">
                <h2>تقرير متابعة الطالب: {selected_master_student}</h2>
                <p><b>المرحلة / الصف:</b> {grade_val} | <b>المنهج:</b> {curr_val}</p>
                <p><b>إجمالي الحصص:</b> {total_sess} | <b>إجمالي الرصيد المستحق:</b> {total_due:,.1f} جنيه</p>
                <hr>
                <h3>سجل الحصص والحضور:</h3>
                <table border="1" style="width:100%; border-collapse:collapse; text-align:center; margin-bottom:20px;">
                    <tr style="background:#f1f5f9;"><th style="padding:8px;">التاريخ</th><th>الحالة</th><th>السعر</th><th>المستوى</th><th>ملاحظات</th></tr>
                    {sess_rows_html if sess_rows_html else "<tr><td colspan='5'>لا توجد حصص مسجلة</td></tr>"}
                </table>
                <h3>سجل الاختبارات والواجبات:</h3>
                <table border="1" style="width:100%; border-collapse:collapse; text-align:center;">
                    <tr style="background:#f1f5f9;"><th style="padding:8px;">التاريخ</th><th>النوع</th><th>العنوان</th><th>الدرجة</th><th>الحالة</th></tr>
                    {ass_rows_html if ass_rows_html else "<tr><td colspan='5'>لا توجد اختبارات مسجلة</td></tr>"}
                </table>
            </body>
            </html>"""

            st.download_button(
                label=f"🖨️ طباعة وتصدير ملف PDF خاص بالطالب ({selected_master_student})",
                data=single_student_pdf_html.encode("utf-8"),
                file_name=f"تقرير_الطالب_{selected_master_student}.html",
                mime="application/octet-stream",
                key=f"dl_single_pdf_{selected_master_student}"
            )

        st.write("---")
        st.markdown("### 📋 جدول ملخص الطلاب الشامل:")
        
        master_data_list = []
        for st_name in all_students_master:
            u_r = st.session_state.users_df[st.session_state.users_df["اسم الطالب"].astype(str).str.strip() == st_name]
            s_r = st.session_state.sessions_df[st.session_state.sessions_df["اسم الطالب"].astype(str).str.strip() == st_name]
            a_r = st.session_state.assessments_df[st.session_state.assessments_df["اسم الطالب"].astype(str).str.strip() == st_name]
            
            is_registered = "نعم (مسجل على المنصة)" if not u_r.empty else "لا (مسجل يدويًا)"
            
            grade_val = "-"
            if not u_r.empty:
                grade_val = str(u_r.iloc[0].get("المجموعة/الصف", "-"))
            elif not s_r.empty:
                grade_val = str(s_r.iloc[-1].get("المجموعة/الصف", "-"))

            total_sess = len(s_r)
            attended_sess = len(s_r[s_r["الحالة"] == "حاضر"])
            avg_price = s_r["سعر الحصة"].astype(float, errors="ignore").mean() if not s_r.empty else 0.0
            total_due = s_r["سعر الحصة"].astype(float, errors="ignore").sum(numeric_only=True) if not s_r.empty else 0.0
            _, total_paid_student, balance_student = get_student_financials(st_name)
            
            exams_count = len(a_r[a_r["النوع"].astype(str).str.contains("اختبار|كويز", na=False)])
            hws_count = len(a_r[a_r["النوع"].astype(str).str.contains("واجب", na=False)])

            master_data_list.append({
                "اسم الطالب": st_name,
                "حالة التسجيل": is_registered,
                "ولي الأمر": str(u_r.iloc[0].get("اسم ولي الأمر", "")) if not u_r.empty else "",
                "المرحلة/الصف": grade_val,
                "إجمالي الحصص": total_sess,
                "الحصص الحاضرة": attended_sess,
                "متوسط سعر الحصة": f"{avg_price:,.1f}",
                "إجمالي الحساب المستحق": f"{total_due:,.1f}",
                "إجمالي المدفوع": f"{total_paid_student:,.1f}",
                "الرصيد المتبقي": f"{max(balance_student,0):,.1f}",
                "عدد الاختبارات": exams_count,
                "عدد الواجبات": hws_count
            })

        master_df = pd.DataFrame(master_data_list)
        st.dataframe(master_df, use_container_width=True)

        all_students_pdf_html = """<!DOCTYPE html>
        <html dir="rtl" lang="ar">
        <head><meta charset="utf-8"><title>تقرير السجلات الشاملة لجميع الطلاب</title></head>
        <body style="font-family: Arial; padding: 25px;" onload="window.print()">
            <h2>تقرير السجلات الشاملة وإحصائيات جميع الطلاب - م/ محمد غنيم</h2>
            <table border="1" style="width:100%; border-collapse:collapse; text-align:center; margin-top:15px;">
                <tr style="background:#f1f5f9;">
                    <th style="padding:8px;">اسم الطالب</th>
                    <th>حالة التسجيل</th><th>ولي الأمر</th>
                    <th>المرحلة/الصف</th>
                    <th>إجمالي الحصص</th>
                    <th>الحصص الحاضرة</th>
                    <th>إجمالي الحساب</th><th>إجمالي المدفوع</th><th>الرصيد المتبقي</th>
                    <th>الاختبارات</th>
                    <th>الواجبات</th>
                </tr>
        """
        for item in master_data_list:
            all_students_pdf_html += f"""
                <tr>
                    <td style="padding:6px;">{item['اسم الطالب']}</td>
                    <td>{item['حالة التسجيل']}</td><td>{item['ولي الأمر']}</td>
                    <td>{item['المرحلة/الصف']}</td>
                    <td>{item['إجمالي الحصص']}</td>
                    <td>{item['الحصص الحاضرة']}</td>
                    <td>{item['إجمالي الحساب المستحق']}</td><td>{item['إجمالي المدفوع']}</td><td>{item['الرصيد المتبقي']}</td>
                    <td>{item['عدد الاختبارات']}</td>
                    <td>{item['عدد الواجبات']}</td>
                </tr>
            """
        all_students_pdf_html += "</table></body></html>"

        c_dl1, c_dl2 = st.columns(2)
        with c_dl1:
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                master_df.to_excel(writer, sheet_name="Master_Students_Report", index=False)
                st.session_state.users_df.to_excel(writer, sheet_name="Users", index=False)
                st.session_state.sessions_df.to_excel(writer, sheet_name="Sessions", index=False)
                st.session_state.assessments_df.to_excel(writer, sheet_name="Assessments", index=False)
                st.session_state.exams_df.to_excel(writer, sheet_name="Exams", index=False)
                st.session_state.videos_df.to_excel(writer, sheet_name="Videos", index=False)

            st.download_button(
                label="📥 تصدير تقرير السجلات الشاملة للطلاب (Excel)",
                data=buf.getvalue(), file_name="تقرير_السجلات_الشاملة_للطلاب.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        with c_dl2:
            st.download_button(
                label="🖨️ طباعة وتصدير تقرير السجلات الشاملة لجميع الطلاب (PDF)",
                data=all_students_pdf_html.encode("utf-8"),
                file_name="تقرير_السجلات_الشاملة_لجميع_الطلاب.html",
                mime="application/octet-stream"
            )

elif t_page == "parent_report":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("📑 إصدار وطباعة تقرير متابعة الطالب لولي الأمر (شامل الأسعار)")
    all_names = sorted(list(set(
        [s for s in st.session_state.sessions_df["اسم الطالب"].dropna().unique() if str(s).strip()]
        + [s for s in st.session_state.assessments_df["اسم الطالب"].dropna().unique() if str(s).strip()]
        + [s for s in st.session_state.users_df["اسم الطالب"].dropna().unique() if str(s).strip()]
        + [s for s in st.session_state.weekly_schedule_df["اسم الطالب"].dropna().unique() if str(s).strip()]
        + [s for s in st.session_state.online_schedule_df["اسم الطالب"].dropna().unique() if str(s).strip()]
    )))

    if not all_names:
        st.info("لا توجد بيانات كافية لإصدار التقرير بعد.")
    else:
        selected_student = st.selectbox("اختر الطالب لإصدار وطباعة تقريره بصيغة PDF:", all_names)

        if selected_student:
            st_sessions = st.session_state.sessions_df[st.session_state.sessions_df["اسم الطالب"] == selected_student].copy().sort_values(by="التاريخ")
            st_assessments = st.session_state.assessments_df[st.session_state.assessments_df["اسم الطالب"] == selected_student].copy().sort_values(by="التاريخ")
            
            u_r_rep = st.session_state.users_df[st.session_state.users_df["اسم الطالب"].astype(str).str.strip() == selected_student]
            parent_name_rep = str(u_r_rep.iloc[0].get("اسم ولي الأمر", "")) if not u_r_rep.empty else ""
            parent_phone_rep = str(u_r_rep.iloc[0].get("رقم ولي الأمر", "")) if not u_r_rep.empty else ""

            curr_val = "-"
            group_val = "-"
            if not u_r_rep.empty:
                group_val = str(u_r_rep.iloc[0].get("المجموعة/الصف", "-"))
                curr_val = str(u_r_rep.iloc[0].get("المنهج/الدولة", "-"))
            elif not st_sessions.empty:
                group_val = str(st_sessions.iloc[-1].get("المجموعة/الصف", "-"))
                curr_val = str(st_sessions.iloc[-1].get("المنهج/الدولة", "-"))

            level_val = st_sessions.iloc[-1].get("مستوى الطالب", "جيد") if not st_sessions.empty else "جيد"
            pay_val = st_sessions.iloc[-1].get("نظام الدفع", "مؤجل") if not st_sessions.empty else "مؤجل"

            # الحساب الشهري يظهر على الموقع فقط، ولا يتم إدخاله داخل ملف تقرير ولي الأمر/PDF
            report_months = set()
            if not st_sessions.empty and "التاريخ" in st_sessions.columns:
                report_months.update([m for m in st_sessions["التاريخ"].apply(_month_from_value) if m])
            student_payments_rep = st.session_state.get("payment_records_df", pd.DataFrame()).copy()
            if not student_payments_rep.empty and "اسم الطالب" in student_payments_rep.columns:
                student_payments_rep = student_payments_rep[student_payments_rep["اسم الطالب"].astype(str).str.strip() == selected_student.strip()]
                if "الشهر" in student_payments_rep.columns:
                    report_months.update([str(m).strip() for m in student_payments_rep["الشهر"].dropna() if str(m).strip()])
                elif "التاريخ" in student_payments_rep.columns:
                    report_months.update([m for m in student_payments_rep["التاريخ"].apply(_month_from_value) if m])

            if report_months:
                report_month_options = sorted(report_months, reverse=True)
                selected_report_month = st.selectbox(
                    "📆 اختر شهر التقرير المالي للطالب:",
                    report_month_options,
                    key="parent_report_month_selector"
                )
                month_due, month_paid, month_remaining = get_student_monthly_financials(selected_student, selected_report_month)
                st.markdown(f"### 📆 الحساب المالي لشهر {selected_report_month}")
                m1, m2, m3 = st.columns(3)
                m1.metric("💳 المستحق خلال الشهر", f"{month_due:,.0f} جنيه")
                m2.metric("✅ المدفوع خلال الشهر", f"{month_paid:,.0f} جنيه")
                m3.metric("💰 المتبقي خلال الشهر", f"{max(month_remaining, 0):,.0f} جنيه")
                if not student_payments_rep.empty:
                    pmonth = student_payments_rep.copy()
                    if "الشهر" in pmonth.columns:
                        pmonth = pmonth[pmonth["الشهر"].astype(str).str.strip() == selected_report_month]
                    elif "التاريخ" in pmonth.columns:
                        pmonth = pmonth[pmonth["التاريخ"].apply(_month_from_value) == selected_report_month]
                    if not pmonth.empty:
                        st.markdown("**عمليات الدفع خلال الشهر المحدد:**")
                        show_cols = [c for c in ["التاريخ", "المبلغ", "طريقة الدفع", "حالة الدفع", "ملاحظات"] if c in pmonth.columns]
                        if show_cols:
                            st.dataframe(pmonth[show_cols].reset_index(drop=True), use_container_width=True)
            else:
                st.info("لا توجد بيانات مالية شهرية مسجلة لهذا الطالب حتى الآن.")

            att_cnt = len(st_sessions[st_sessions["الحالة"] == "حاضر"])
            abs_cnt = len(st_sessions[st_sessions["الحالة"] == "غائب"])
            total_sessions_cnt = len(st_sessions)
            att_percentage = (att_cnt / total_sessions_cnt * 100) if total_sessions_cnt > 0 else 0
            total_cost = st_sessions["سعر الحصة"].astype(float, errors="ignore").sum(numeric_only=True) if not st_sessions.empty else 0.0

            exams = st_assessments[st_assessments["النوع"].str.contains("اختبار", na=False)]
            hws = st_assessments[st_assessments["النوع"].str.contains("واجب", na=False)]
            total_exam_score = exams["الدرجة المحصلة"].astype(float).sum() if not exams.empty else 0
            total_exam_max = exams["الدرجة العظمى"].astype(float).sum() if not exams.empty else 0
            exam_percentage = (total_exam_score / total_exam_max * 100) if total_exam_max > 0 else 100.0

            col_rep1, col_rep2, col_rep3, col_rep4 = st.columns(4)
            col_rep1.metric("نسبة الحضور والالتزام", f"{att_percentage:.1f}%")
            col_rep2.metric("عدد الواجبات المستلمة", len(hws))
            col_rep3.metric("متوسط درجات الاختبارات", f"{exam_percentage:.1f}%")
            col_rep4.metric("إجمالي المبلغ المستحق", f"{total_cost:,.1f}")

            ass_html_rows = ""
            if not st_assessments.empty:
                for _, r in st_assessments.iterrows():
                    ass_html_rows += f"""
                    <tr>
                        <td style="padding: 10px; border: 2px solid #000; font-weight: 900;">{r['التاريخ']}</td>
                        <td style="padding: 10px; border: 2px solid #000; font-weight: 900; color: #0052cc;">{r['النوع']}</td>
                        <td style="padding: 10px; border: 2px solid #000; font-weight: 900;">{r['عنوان التكليف']}</td>
                        <td style="padding: 10px; border: 2px solid #000; font-weight: 900;">{r['الدرجة المحصلة']} / {r['الدرجة العظمى']}</td>
                        <td style="padding: 10px; border: 2px solid #000; font-weight: 900;">{r['حالة التسليم']}</td>
                        <td style="padding: 10px; border: 2px solid #000; font-weight: 900; color: #047857;">{r['ملاحظات وتوجيهات']}</td>
                    </tr>
                    """
            else:
                ass_html_rows = "<tr><td colspan='6' style='padding: 15px; font-weight: 900; border: 2px solid #000;'>لا توجد واجبات أو اختبارات مرصودة حتى الآن.</td></tr>"

            session_html_rows = ""
            if not st_sessions.empty:
                for _, r in st_sessions.iterrows():
                    st_color = "#0f766e" if r["الحالة"] == "حاضر" else ("#b91c1c" if r["الحالة"] == "غائب" else "#b45309")
                    p_curr = float(r.get('سعر الحصة', 0)) if pd.notnull(r.get('سعر الحصة')) else 0.0
                    session_html_rows += f"""
                    <tr>
                        <td style="padding: 10px; border: 2px solid #000; font-weight: 900;">{r['التاريخ']}</td>
                        <td style="padding: 10px; border: 2px solid #000; color:{st_color}; font-weight: 900;">{r['الحالة']}</td>
                        <td style="padding: 10px; border: 2px solid #000; font-weight: 900; color: #16a34a; font-size: 17px;">{p_curr:,.1f}</td>
                        <td style="padding: 10px; border: 2px solid #000; font-weight: 900; color: #0052cc;">{r['مستوى الطالب']}</td>
                        <td style="padding: 10px; border: 2px solid #000; font-weight: 900;">{r['ملاحظات']}</td>
                    </tr>
                    """
                session_html_rows += f"""
                <tr style="background-color: #f1f5f9;">
                    <td colspan="2" style="padding: 12px; border: 2px solid #000; font-weight: 900; font-size: 17px;">الإجمالي الكلي المستحق للحصص:</td>
                    <td style="padding: 12px; border: 2px solid #000; font-weight: 900; font-size: 19px; color: #b91c1c;">{total_cost:,.1f}</td>
                    <td colspan="2" style="padding: 12px; border: 2px solid #000; font-weight: 900; font-size: 16px;">نظام الدفع المعتمد: {pay_val}</td>
                </tr>
                """
            else:
                session_html_rows = "<tr><td colspan='5' style='padding: 15px; font-weight: 900; border: 2px solid #000;'>لا توجد حصص مسجلة بعد.</td></tr>"

            teacher_img_tag = f'<img src="{teacher_image_data_uri(img_b64)}" style="width: 100px; height: 100px; border-radius: 50%; border: 3px solid #0052cc; object-fit: cover;">' if img_b64 else ""

            student_weekly_rep = st.session_state.weekly_schedule_df[st.session_state.weekly_schedule_df["اسم الطالب"].astype(str).str.strip() == selected_student.strip()].copy()
            weekly_report_rows = ""
            for _, wr in student_weekly_rep.iterrows():
                weekly_report_rows += f"<tr><td style='padding:10px;border:2px solid #000;font-weight:900'>{wr.get('اليوم','')}</td><td style='padding:10px;border:2px solid #000;font-weight:900'>{wr.get('الموعد','')}</td><td style='padding:10px;border:2px solid #000;font-weight:900'>{wr.get('اسم الأكاديمية','')}</td><td style='padding:10px;border:2px solid #000;font-weight:900'>{wr.get('المنهج/الدولة','')}</td><td style='padding:10px;border:2px solid #000;font-weight:900'>{wr.get('المجموعة/الصف','')}</td><td style='padding:10px;border:2px solid #000;font-weight:900'>{wr.get('سعر الحصة',0)}</td></tr>"

            parent_report_html = f"""<!DOCTYPE html>
            <html dir="rtl" lang="ar">
            <head>
                <meta charset="utf-8">
                <title>تقرير متابعة ولي الأمر - {selected_student}</title>
                <style>
                    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@800;900&display=swap');
                    body {{ font-family: 'Cairo', Tahoma, Arial, sans-serif; padding: 30px; color: #000000 !important; background-color: #ffffff !important; font-weight: 900; }}
                    .header-box {{ border-bottom: 3px solid #0052cc; padding-bottom: 15px; margin-bottom: 25px; display: flex; justify-content: space-between; align-items: center; }}
                    .brand-name {{ color: #0052cc; margin: 0; font-size: 32px; font-weight: 900; }}
                    .brand-sub {{ color: #000000; margin: 4px 0; font-size: 16px; font-weight: 900; }}
                    .parent-notice {{ background-color: #f0fdf4; border: 2px solid #16a34a; padding: 12px 18px; border-radius: 8px; margin-bottom: 25px; font-size: 17px; font-weight: 900; color: #14532d; }}
                    .stats-box {{ width: 100%; border-collapse: collapse; margin-bottom: 25px; }}
                    .stats-box th, .stats-box td {{ border: 2px solid #000000; padding: 10px; text-align: center; font-size: 16px; font-weight: 900; }}
                    .stats-box th {{ background: #f1f5f9; color: #000000; font-weight: 900; }}
                    .section-title {{ margin-top: 25px; margin-bottom: 10px; color: #0052cc; font-size: 20px; font-weight: 900; border-bottom: 2px solid #0052cc; padding-bottom: 6px; display: inline-block; }}
                    .table-main {{ width: 100%; border-collapse: collapse; text-align: center; margin-top: 10px; margin-bottom: 20px; }}
                    .table-main th {{ background-color: #f8fafc; color: #000000; font-weight: 900; font-size: 16px; padding: 10px; border: 2px solid #000000; }}
                    .footer-note {{ margin-top: 35px; text-align: center; color: #000000; font-size: 16px; font-weight: 900; border-top: 2px solid #000000; padding-top: 15px; }}
                </style>
            </head>
            <body onload="window.print()">
                <div class="header-box">
                    <div style="display: flex; align-items: center; gap: 20px;">
                        {teacher_img_tag}
                        <div>
                            <h2 class="brand-name">م/ محمد غنيم 📐</h2>
                            <p class="brand-sub">تقرير التقييم الدوري لولي الأمر</p>
                            <p style="margin: 2px 0; color: #000000; font-size: 16px; font-weight: 900;"><b>المنهج والمرحلة:</b> {curr_val} — {group_val}</p>
                        </div>
                    </div>
                    <div style="text-align: left;">
                        <h2 style="color: #0052cc; margin: 0; font-size: 26px; font-weight: 900;">الطالب: {selected_student}</h2><p style="margin:3px 0;font-weight:900;">ولي الأمر: {parent_name_rep or "غير مسجل"} | {parent_phone_rep or "غير مسجل"}</p>
                        <p style="margin: 5px 0 0 0; color: #000000; font-size: 15px; font-weight: 900;">تاريخ إصدار التقرير: {date.today()}</p>
                    </div>
                </div>
                <table class="stats-box">
                    <tr>
                        <th>إجمالي الحصص</th><th>مرات الحضور</th><th>مرات الغياب</th><th>نسبة الالتزام بالحضور</th><th>إجمالي الحساب المستحق</th><th>المستوى العام</th>
                    </tr>
                    <tr>
                        <td style="font-weight: 900;">{total_sessions_cnt}</td>
                        <td style="color: #0f766e; font-weight: 900;">{att_cnt}</td>
                        <td style="color: #b91c1c; font-weight: 900;">{abs_cnt}</td>
                        <td style="font-weight: 900; color: #0052cc;">{att_percentage:.1f}%</td>
                        <td style="font-weight: 900; font-size: 18px; color: #b91c1c;">{total_cost:,.1f}</td>
                        <td style="color: #0052cc; font-weight: 900;">{level_val}</td>
                    </tr>
                </table>
                <div class="section-title">1. جدول المواعيد الأسبوعية:</div>
                <table class="table-main">
                    <tr><th>اليوم</th><th>الموعد</th><th>الأكاديمية</th><th>المنهج</th><th>المرحلة</th><th>سعر الحصة</th></tr>
                    {weekly_report_rows if weekly_report_rows else "<tr><td colspan='6'>لا توجد مواعيد أسبوعية مسجلة.</td></tr>"}
                </table>
                <div class="section-title">2. تقرير الواجبات المنزلية والاختبارات الدورية:</div>
                <table class="table-main">
                    <tr><th>التاريخ</th><th>النوع</th><th>عنوان التكليف / الاختبار</th><th>الدرجة المحصلة</th><th>حالة التسليم والالتزام</th><th>ملاحظات وتوجيهات</th></tr>
                    {ass_html_rows}
                </table>
                <div class="section-title">3. سجل الحضور وتفاصيل سعر كل حصة:</div>
                <table class="table-main">
                    <tr><th>التاريخ</th><th>حالة الحضور</th><th>سعر الحصة</th><th>مستوى الطالب بالحصة</th><th>ملاحظات التفاعل والاستيعاب</th></tr>
                    {session_html_rows}
                </table>
                <div class="footer-note">مع تحيات: <b>م/ محمد غنيم | منصة الرياضيات والإحصاء</b> — رقم التواصل المباشر: 01016361440 🌟</div>
            </body>
            </html>"""

            st.download_button(
                label=f"🖨️ تحميل وطباعة تقرير ولي الأمر لـ ({selected_student}) بصيغة PDF",
                data=parent_report_html.encode("utf-8"),
                file_name=f"تقرير_ولي_الأمر_{selected_student}.html",
                mime="application/octet-stream",
            )
