import os
import re
import uuid
import asyncio
import pandas as pd
from flask import Flask, render_template_string, request, jsonify
import edge_tts
import google.generativeai as genai

# ===== التعديل 2: مسار نسبي بدل الثابت =====
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
AUDIO_DIR = os.path.join(BASE_DIR, "static")
DATA_DIR = os.path.join(BASE_DIR, "data_files")

if not os.path.exists(AUDIO_DIR):
    os.makedirs(AUDIO_DIR)
if not os.path.exists(DATA_DIR):
    os.makedirs(DATA_DIR)

# ===== التعديل 3: static_folder صريح =====
app = Flask(__name__, static_folder=AUDIO_DIR, static_url_path='/static')

# إعداد مفتاح Gemini API الخاص بك (يُفضّل ضبطه في متغيرات البيئة على PythonAnywhere)
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "AQ.Ab8RN6IaQy6kyhWHnW3hmshY13Ev5Lt9alhv0pM_K2OgKJIhSw")
genai.configure(api_key=GEMINI_API_KEY)

# إعداد نموذج جيمناي للرد بذكاء واحترافية
generation_config = {"temperature": 0.1, "max_output_tokens": 150}
model = genai.GenerativeModel(model_name="gemini-1.5-flash", generation_config=generation_config)

# تحميل ملف الإكسل (جميع الاقسام.xlsx) عند بدء التشغيل
excel_path = os.path.join(DATA_DIR, "جميع الاقسام.xlsx")
global_df = None
global_df_str = None

try:
    if os.path.exists(excel_path):
        global_df = pd.read_excel(excel_path)
        global_df_str = global_df.astype(str)
        print("تم تحميل ملف الإكسل جميع الاقسام.xlsx بنجاح.")
    else:
        alt_path = os.path.join(DATA_DIR, "جميع الاقسام .xlsx")
        if os.path.exists(alt_path):
            global_df = pd.read_excel(alt_path)
            global_df_str = global_df.astype(str)
            print("تم تحميل ملف الإكسل البديل بنجاح.")
except Exception as e:
    print("Error loading excel file:", e)


def normalize_arabic(text: str) -> str:
    text = text.lower().strip()
    text = re.sub(r'[\u064b-\u065f\u0670]', '', text)
    text = text.replace('أ', 'ا').replace('إ', 'ا').replace('آ', 'ا')
    text = text.replace('ى', 'ي').replace('ة', 'ه')
    text = re.sub(r'\s+', ' ', text)
    return text


def is_explicit_book_search(norm_query: str) -> bool:
    book_keywords = ['كتاب', 'كتب', 'فهرس', 'مؤلف', 'رقم عام', 'الرقم العام', 'سنة النشر', 'طبعة']
    return any(k in norm_query for k in book_keywords)


def match_staff_member(norm_query: str):
    if is_explicit_book_search(norm_query):
        return None

    staff_profiles = [
        {
            "keys": {"عفيفي محمد عوض", "عفيفي محمد", "عفيفي", "مدير المكتبه", "مدير اداره المكتبه"},
            "reply": "مدير إدارة المكتبة هو الأستاذ عفيفي محمد عوض بدرجة مدير عام، ولديه خبرة 35 عاماً في مجال المكتبات الجامعية والإدارية."
        },
        {
            "keys": {"شيرين احمد راشد", "شيرين احمد", "شيرين راشد", "مس شيرين", "دكتوره شيرين", "شيرين", "نائب المدير", "نائب مدير المكتبه"},
            "reply": "نائب مدير المكتبة هي الأستاذة شيرين أحمد راشد، وتشارك في إدارة وتسيير أعمال المكتبة والإشراف على الخدمات المقدمة للمستفيدين."
        },
        {
            "keys": {"رشا ابراهيم احمد", "رشا ابراهيم", "رشا احمد", "مس رشا", "دكتوره رشا", "رشا"},
            "reply": "الأستاذة رشا إبراهيم أحمد هي أخصائي مكتبات بالمكتبة، ومسؤولة عن توجيه المستفسرين وتقديم الخدمات الخاصة بكتب قسم الهندسة الكهربية وقسم هندسة القوى."
        },
        {
            "keys": {"دينا عبد الفتاح ناصف", "دينا عبد الفتاح", "دينا ناصف", "مس دينا", "دكتوره دينا", "دينا"},
            "reply": "الأستاذة دينا عبد الفتاح ناصف أخصائي مكتبات بالمكتبة، ومسؤولة عن توجيه المستفسرين وتقديم الخدمات الخاصة بكتب العلوم الأساسية، والرسائل العلمية، وقسم الهندسة المعمارية."
        },
        {
            "keys": {"اسماء عبد الفتاح عرب", "اسماء عبد الفتاح", "اسماء عرب", "مس اسماء", "دكتوره اسماء", "اسماء"},
            "reply": "الأستاذة أسماء عبد الفتاح عرب أخصائي اجتماعي بالمكتبة، ومسؤولة عن توجيه المستفسرين وتقديم الخدمات الخاصة بكتب قسم الهندسة المدنية وقسم هندسة الإنتاج والتصميم."
        },
    ]

    for profile in staff_profiles:
        if norm_query in profile["keys"]:
            return profile["reply"]

    for profile in staff_profiles:
        for key in profile["keys"]:
            if len(key.split()) >= 2 and key in norm_query:
                return profile["reply"]

    return None


# دالة البحث في فهرس الكتب
def search_excel_books(query):
    global global_df, global_df_str
    if global_df is None or global_df_str is None:
        return None, None
    try:
        search_query = query.strip()
        q_check = normalize_arabic(search_query)

        forbidden_exact = [
            'عفيفي محمد عوض', 'عفيفي محمد', 'عفيفي',
            'رشا ابراهيم', 'شيرين احمد', 'اسماء عبد الفتاح', 'دينا عبد الفتاح',
            'فريق العمل', 'العاملين بالمكتبه', 'الموظفين بالمكتبه',
            'تدوين صوتي', 'بودكاست', 'soundcloud', 'ساوند', 'موقع المكتبه',
            'مواعيد العمل', 'ساعات العمل', 'قاعات المكتبه', 'مجله الكليه', 'بنك المعرفه', 'استعاره عن بعد',
            'ازاي اوصل لكتاب', 'كيف اصل لكتاب', 'كيف اوصل لكتاب',
            'زهقان', 'ملل', 'زعلان', 'تعبان', 'زهقت'
        ]
        if any(f_name in q_check for f_name in forbidden_exact) and not is_explicit_book_search(q_check):
            return None, None

        results = pd.DataFrame()

        clean_query = re.sub(r'^(عايز|اريد|أريد|محتاج|ابحث عن|أبحث عن|كتاب|بحث عن|عن كتاب)\s+', '', search_query, flags=re.IGNORECASE).strip()
        if not clean_query:
            clean_query = search_query

        q_lower = clean_query.lower()

        master_library_dict = {
            'ثيرمو': ['thermodynamics', 'thermodynamic', 'thermal'],
            'ثيرموديناميك': ['thermodynamics', 'thermodynamic', 'thermal'],
            'ترموديناميك': ['thermodynamics', 'thermodynamic', 'thermal'],
            'ديناميك': ['dynamics', 'dynamic', 'thermodynamics'],
            'حرارة': ['thermodynamics', 'heat', 'thermal', 'heating'],
            'فايبريشن': ['vibration', 'vibrations'],
            'اهتزازات': ['vibration', 'vibrations'],
            'فلويد': ['fluid', 'fluids'],
            'موائع': ['fluid', 'fluids'],
            'ميكانيكا': ['mechanics', 'mechanic', 'mechanical', 'machinery'],
            'ميكانيكي': ['mechanics', 'mechanic', 'mechanical'],
            'فيزياء': ['physics', 'physical'],
            'رياضيات': ['mathematics', 'mathematical', 'math'],
            'بايثون': ['python'],
            'باور': ['power', 'energy'],
            'طاقة': ['power', 'energy'],
            'بامب': ['pumps', 'pump'],
            'مضخات': ['pumps', 'pump'],
            'كونكريت': ['concrete'],
            'خرسانة': ['concrete'],
            'ديزاين': ['design', 'designs'],
            'تصميم': ['design', 'designs'],
            'إلكتريك': ['electric', 'electrical', 'electricity'],
            'كهرباء': ['electric', 'electrical', 'electricity'],
            'إلكترونيك': ['electronic', 'electronics'],
            'إلكترونكس': ['electronic', 'electronics'],
            'الكترونيك': ['electronic', 'electronics'],
            'الكترونيكس': ['electronic', 'electronics'],
            'ديفايس': ['device', 'devices'],
            'ديفايسز': ['device', 'devices'],
            'أجهزة': ['device', 'devices', 'instrumentation', 'apparatus'],
            'كنترول': ['control'],
            'تحكم': ['control'],
            'سيستم': ['system', 'systems'],
            'أنظمة': ['system', 'systems'],
            'ستراكشر': ['structure', 'structures', 'structural'],
            'إنشاءات': ['structure', 'structures', 'structural', 'construction', 'building'],
            'هندسة': ['engineering', 'engineer'],
            'أنجنيرنج': ['engineering', 'engineer'],
            'تحليل': ['analysis', 'analytical'],
            'مشاكل': ['problems', 'problem'],
            'تمارين': ['problems', 'solved'],
            'تكنولوجي': ['technology', 'technological', 'technical'],
            'مواد': ['materials', 'material'],
            'هاندبوك': ['handbook', 'manual', 'guide'],
            'مرجع': ['handbook', 'manual', 'guide'],
            'مقدمة': ['introduction', 'introductory', 'basic', 'basics'],
            'أساسيات': ['fundamentals', 'fundamental', 'principles', 'basic', 'basics'],
            'آلات': ['machines', 'machine', 'machinery'],
            'دائرة': ['circuits', 'circuit'],
            'دوائر': ['circuits', 'circuit']
        }

        search_terms = [q_lower]
        for ar_term, en_list in master_library_dict.items():
            if ar_term in q_lower:
                search_terms.extend(en_list)

        for word in q_lower.split():
            if len(word) > 2 and word not in search_terms:
                search_terms.append(word)

        for term in search_terms:
            if not term:
                continue
            for col in global_df_str.columns:
                col_l = str(col).lower()
                if 'عنوان' in col_l or 'name' in col_l or 'title' in col_l:
                    mask = global_df_str[col].str.contains(re.escape(term), case=False, na=False)
                    results = global_df[mask]
                    if not results.empty:
                        break
            if not results.empty:
                break

        if results.empty:
            for term in search_terms:
                if not term:
                    continue
                mask_all = global_df_str.apply(lambda row: row.str.contains(re.escape(term), case=False, na=False).any(), axis=1)
                results = global_df[mask_all]
                if not results.empty:
                    break

        if not results.empty:
            limited_results = results.head(3)
            response_text = f"وجدت {len(results)} نتيجة مطابقة (عرض أول 3 نتائج):\n\n"
            speech_summary = "وجدت نتائج مطابقة في الكتب."

            for _, row in limited_results.iterrows():
                response_text += "────────────────────────\n"
                book_title = ""
                book_no = ""

                for col, val in row.items():
                    col_str = str(col).strip()
                    if any(k in col_str.lower() for k in ['phonetic', 'arabic_phonetic', 'النطق']):
                        continue

                    if pd.notna(val) and str(val).strip() != "":
                        response_text += f"• **{col_str}**: {val}\n"
                        col_lower = col_str.lower()
                        if any(k in col_lower for k in ['عنوان', 'اسم', 'book', 'title', 'name']):
                            if not book_title:
                                book_title = str(val)
                        if any(k in col_lower for k in ['رقم', 'عام', 'number', 'code', 'id']):
                            if not book_no:
                                book_no = str(val)

                if book_title:
                    speech_summary += f" الكتاب هو {book_title}"
                    if book_no:
                        speech_summary += f" والرقم العام هو {book_no} "

            response_text += "────────────────────────"
            return response_text.strip(), speech_summary
    except Exception as e:
        print("Excel Search Error:", e)
    return None, None


LIBRARY_DATA = """نبذة تعريفية عن المكتبة:
بدأت المكتبة نشاطها منذ أن كانت مكتبة تابعة للمعهد العالي الصناعي عام 1958، ثم تحولت إلى مكتبة لكلية الهندسة عام 1985، وهي مستمرة في أداء رسالتها حتى الآن.

مواعيد العمل وقاعات المكتبة:
تبدأ ساعات العمل اليومية من الساعة 8:45 صباحاً إلى الساعة 1:45 ظهراً. وتقدم الخدمات الإلكترونية طوال أيام الأسبوع على مدار 24 ساعة.
تحتوي المكتبة على قاعتين رئيسيتين: القاعة الأولى بالدور الأرضي بمبنى القوى تخص قسم هندسة القوى الميكانيكية، العلوم الأساسية، الكتب الثقافية، المعمارية، والرسائل الجامعية. والقاعة الرئيسية بالدور الأول علوي تحتوي على قسم الهندسة المدنية، الإنتاج والتصميم، والهندسة الكهربية.

فريق العمل بالمكتبة وتخصصاتهم:
1. الأستاذ عفيفي محمد عوض: مدير إدارة المكتبة بدرجة مدير عام، ولديه خبرة 35 عاماً في مجال المكتبات الجامعية والإدارية.
2. الأستاذة شيرين أحمد راشد: نائب مدير المكتبة، وتشارك في إدارة وتسيير أعمال المكتبة والإشراف على الخدمات المقدمة للمستفيدين.
3. الأستاذة رشا إبراهيم أحمد: أخصائي مكتبات بالمكتبة، ومسؤولة عن توجيه المستفسرين وتقديم الخدمات الخاصة بكتب قسم الهندسة الكهربية وقسم هندسة القوى.
4. الأستاذة دينا عبد الفتاح ناصف: أخصائي مكتبات بالمكتبة، ومسؤولة عن توجيه المستفسرين وتقديم الخدمات الخاصة بكتب العلوم الأساسية، والرسائل العلمية، وقسم الهندسة المعمارية.
5. الأستاذة أسماء عبد الفتاح عرب: أخصائي اجتماعي بالمكتبة، ومسؤولة عن توجيه المستفسرين وتقديم الخدمات الخاصة بكتب قسم الهندسة المدنية وقسم هندسة الإنتاج والتصميم.

كيفية التسجيل وعمل حساب على بنك المعرفة المصري:
1. الدخول على الرابط: https://www.ekb.eg/ar/home
2. الضغط على زر "سجل".
3. ابدأ في تحديد الفئة الخاصة بك (طالب جامعي أو باحث).
4. ابدأ في ملء بياناتك المطلوبة.
- ملاحظة: في حالة عدم الاستطاعة، يمكنك ملء الاستمارة والمكتبة تقوم بعمل حساب خاص بك، أو التوجه إلى المكتبة للمساعدة.

مجلة الكلية:
- تستطيع الوصول إلى مجلة الكلية من خلال الرابط التالي: https://erjm.journals.ekb.eg
- وتستطيع الوصول إلى مجلة الكلية التفاعلية من خلال الروابط التالية: 
  1. https://canva.link/63hkb3cg54jy0rf
  2. https://canva.link/ouxmps9fhubijh9

خدمات المكتبة (التقليدية والإلكترونية):
1. الخدمات التقليدية للمكتبة: الاطلاع الداخلي، الإرشاد والتوجيه، الاستعارة الخارجية الورقية والإلكترونية، خدمة الإحاطة الجارية، وخدمة البث الانتقائي.
2. الخدمات الإلكترونية للمكتبة (تقدم طوال 24 ساعة): 
   - موقع المكتبة: https://englib2010.wixsite.com
   - تطبيق الهاتف الخاص بالمكتبة من خلال موقع المكتبة.
   - الاستعارة عن بعد.
   - التدوين الصوتي (خدمة إلكترونية): تعد راديو خاص بالمكتبة، وتستطيع الوصول إليه مباشرة عبر الرابط التالي: https://on.soundcloud.com/vhZCzabhbuFJjncVnU
   - مجلة الكلية.
   - المنصة التعليمية.
   - الذكاء الاصطناعي.
   - الاختبارات الجامعية.

كيفية الاستعارة عن بعد:
1. الدخول على الرابط: https://englib2010.wixsite.com
2. الضغط علي E-Book borrowing واختيار القسم العلمي.
3. الضغط علي Search now للبحث عن الكتاب ثم اضغط علي الأيقونة الخضراء Click here.
4. ملء استمارة الطلب وإرسالها ليقوم الأخصائي بإرسال الكتاب عبر بريدك الإلكتروني. (مدة الاستعارة أسبوع واحد قابل للتجديد لكتاب واحد أسبوعياً للقراءة فقط).

كيفية البحث عن كتاب فى فهرس المكتبة:
1. الدخول على الرابط: www.eulc.edu.eg
2. اختيار جامعة المنوفية - كلية الهندسة وتحديد مقتنيات المكتبة وكتابة مصطلح البحث ثم اضغط بحث.

الامتحانات والمحاضرات الرقمية:
- الامتحانات السابقة متوفرة عبر الرابط: https://englib2010.wixsite.com/englibrary/exam
- المحاضرات الإلكترونية لأعضاء هيئة التدريس متاحة عبر الرابط: https://englib2010.wixsite.com/englibrary/electronic-platform

لائحـة وتعليمـات الإعارة الخارجية:
أعضاء هيئة التدريس 5 كتب، المدرسون المساعدون والمعيدون 4 كتب, العاملون وطلبة الدراسات العليا 3 كتب، طلبة المرحلة الجامعية الأولى كتابان.

الجوائز:
حصلت المكتبة على المركز الأول أعوام 2018 و2019 و2021 و2023 و2024، وحصلت على المستوى الثاني عام 2025."""

with open(os.path.join(DATA_DIR, "library_qa.txt"), "w", encoding="utf-8") as f:
    f.write(LIBRARY_DATA)


def clean_text_for_speech(text):
    if not text:
        return ""
    text_clean = re.sub(r'https?://\S+|www\.\S+', '', text)
    if 'مجلة الكلية' in text or 'canva.link' in text:
        text_clean = re.sub(r'^\s*\d{1,2}\s*[\.\-]\s*', '', text_clean, flags=re.MULTILINE)
        text_clean = re.sub(r'\n\s*\d{1,2}\s*[\.\-]\s*', ' ', text_clean)
    text_clean = re.sub(r'(\d+)', r' \1 ', text_clean)
    text_clean = re.sub(r'[*#_`–—•\[\]\(\)\-]', ' ', text_clean)
    text_clean = text_clean.replace('الكليّة', 'الكُلّيّة').replace('الكلية', 'الكُلّيّة')
    text_clean = re.sub(r'([،.؟!])', r'\1 ', text_clean)
    text_clean = re.sub(r'\s+', ' ', text_clean).strip()
    return text_clean


# ===== التعديل 4: edge-tts بدل gTTS =====
async def _save_tts_audio(text, path):
    communicate = edge_tts.Communicate(text, voice="ar-EG-SalmaNeural")
    await communicate.save(path)


def synthesize_speech(speech_text: str, prefix: str):
    unique_filename = f"{prefix}_{uuid.uuid4().hex}.mp3"
    audio_path = os.path.join(AUDIO_DIR, unique_filename)
    try:
        asyncio.run(_save_tts_audio(speech_text, audio_path))
        if os.path.exists(audio_path) and os.path.getsize(audio_path) > 0:
            print(f"[TTS OK] {unique_filename} ({os.path.getsize(audio_path)} bytes)")
            return f"/static/{unique_filename}", None
        return None, "ملف الصوت فارغ أو لم يُنشأ."
    except Exception as e:
        print("edge-tts Error:", e)
        return None, str(e)


def ask_gemini_library_assistant(user_query):
    q_lower = user_query.lower()
    q_norm = normalize_arabic(user_query)

    # 1. التدوين الصوتي بأولوية مطلقة
    if any(k in q_lower for k in ['تدوين', 'بودكاست', 'soundcloud', 'ساوند']) or ('صوتي' in q_lower and 'كتب' not in q_lower and 'مشروع' not in q_lower):
        reply = "خدمة التدوين الصوتي تعد راديو خاص بالمكتبة، وتستطيع الوصول إليه من خلال الرابط التالي: https://on.soundcloud.com/vhZCzabhbuFJjncVnU"
        return reply, reply

    # 2. مواعيد العمل وقاعات المكتبة بأولوية مطلقة
    if any(k in q_lower for k in ['مواعيد', 'ساعات العمل', 'أوقات العمل', 'قاعات', 'القاعات']):
        reply = "تبدأ ساعات العمل اليومية من الساعة 8:45 صباحاً إلى الساعة 1:45 ظهراً. وتقدم الخدمات الإلكترونية طوال أيام الأسبوع على مدار 24 ساعة.\n\nتحتوي المكتبة على قاعتين رئيسيتين:\n• القاعة الأولى: بالدور الأرضي بمبنى القوى (تخص قسم هندسة القوى الميكانيكية، العلوم الأساسية، الكتب الثقافية، المعمارية، والرسائل الجامعية).\n• القاعة الرئيسية: بالدور الأول علوي (تخص قسم الهندسة المدنية، الإنتاج والتصميم، والهندسة الكهربية)."
        return reply, reply

    # 3. معالجة المشاعر (زهقان، زعلان، ملل، تعبان) بشكل ودود واحترافي
    if any(k in q_lower for k in ['زهقان', 'ملل', 'زعلان', 'تعبان', 'زهقت']):
        reply = "لا داعي للضيق أو الملل أبداً! أنا هنا دائماً لمساعدتك وإسعادك. ما رأيك أن نقوم بجولة ممتعة معاً؟ يمكنني مساعدتك في البحث عن كتاب ممتع وشيق يغير جو يومك، أو يمكنك الاستماع إلى البودكاست وراديو المكتبة الخاص عبر الرابط: https://on.soundcloud.com/vhZCzabhbuFJjncVnU\n\nكيف يمكنني مساعدتك الآن لتشعر بتحسن؟"
        return reply, reply

    # 4. معالجة الأسماء وفريق العمل
    if q_norm in {'فريق العمل', 'العاملين بالمكتبه', 'الموظفين بالمكتبه'}:
        reply = "فريق العمل بالمكتبة يضم:\n• الأستاذ عفيفي محمد عوض (مدير إدارة المكتبة)\n• الأستاذة شيرين أحمد راشد (نائب المدير)\n• الأستاذة رشا إبراهيم أحمد (أخصائي مكتبات - كتب الكهربية والقوى)\n• الأستاذة دينا عبد الفتاح ناصف (أخصائي مكتبات - العلوم الأساسية والمعمارية)\n• الأستاذة أسماء عبد الفتاح عرب (أخصائي اجتماعي - المدني والإنتاج)."
        return reply, reply

    staff_reply = match_staff_member(q_norm)
    if staff_reply:
        return staff_reply, staff_reply

    # الخدمات والروابط الأخرى
    if 'تطبيق' in q_lower or 'موقع المكتبة' in q_lower or 'الموقع' in q_lower:
        reply = "يمكنك الوصول إلى موقع المكتبة وتطبيق الهاتف الخاص بها عبر الرابط الرئيسي: https://englib2010.wixsite.com"
        return reply, reply
    elif 'محاضرات' in q_lower:
        reply = "تتاح المحاضرات الإلكترونية لأعضاء هيئة التدريس عبر الرابط: https://englib2010.wixsite.com/englibrary/electronic-platform"
        return reply, reply
    elif 'امتحانات' in q_lower or 'امتحان' in q_lower:
        reply = "يمكنك الاطلاع على الامتحانات السابقة عبر الرابط: https://englib2010.wixsite.com/englibrary/exam"
        return reply, reply
    elif q_lower in ['ازاى اوصل لكتاب', 'إزاى أوصل لكتاب', 'ازاي اوصل لكتاب', 'كيف أصل لكتاب', 'كيف اوصل لكتاب']:
        reply = "للبحث عن كتاب أو الوصول إليه في فهرس المكتبة، ادخل على الرابط www.eulc.edu.eg، اختر جامعة المنوفية كلية الهندسة، حدد مقتنيات المكتبة وكتابة مصطلح البحث ثم اضغط بحث."
        return reply, reply
    elif any(k in q_lower for k in ['فهرس', 'البحث عن كتاب']):
        reply = "للبحث عن كتاب أو الوصول إليه في فهرس المكتبة، ادخل على الرابط www.eulc.edu.eg، اختر جامعة المنوفية كلية الهندسة، حدد مقتنيات المكتبة وكتابة مصطلح البحث ثم اضغط بحث."
        return reply, reply
    elif 'تسجيل' in q_lower or 'عمل حساب' in q_lower or 'بنك المعرفه' in q_lower or 'بنك المعرفة' in q_lower:
        reply = "لعمل حساب وتسجيل الدخول على بنك المعرفة المصري، يرجى الدخول على الرابط https://www.ekb.eg/ar/home، ثم اضغط على سجل، وحدد فئتك كطالب جامعي أو باحث ثم املأ بياناتك."
        return reply, reply
    elif 'مجلة الكلية' in q_lower or 'مجله الكليه' in q_lower or 'المجلة' in q_lower:
        reply = "تستطيع الوصول إلى مجلة الكلية من خلال الرابط التالي: https://erjm.journals.ekb.eg\n\nوتستطيع الوصول إلى مجلة الكلية التفاعلية من خلال الروابط التالية:\n1. https://canva.link/63hkb3cg54jy0rf\n2. https://canva.link/ouxmps9fhubijh9"
        return reply, reply
    elif 'خدمات' in q_lower:
        reply = "تقدم المكتبة خدمات تقليدية (الاطلاع، الإرشاد، الاستعارة) وخدمات إلكترونية على مدار 24 ساعة تشمل موقع المكتبة (https://englib2010.wixsite.com), وتطبيق الهاتف الخاص بالمكتبة, الاستعارة عن بعد, خدمة التدوين الصوتي (تعد راديو خاص بالمكتبة عبر الرابط: https://on.soundcloud.com/vhZCzabhbuFJjncVnU), مجلة الكلية, المنصة التعليمية, الذكاء الاصطناعي, والاختبارات الجامعية."
        return reply, reply
    elif 'استعاره' in q_lower or 'استعارة' in q_lower or 'عن بعد' in q_lower:
        reply = "للإستعارة عن بعد، يرجى الدخول على الرابط https://englib2010.wixsite.com ثم اختر E-Book borrowing، حدد القسم وابحث عن الكتاب ثم اضغط Click here لملء وإرسال استمارة الطلب."
        return reply, reply
    elif 'جائزة' in q_lower or 'جوائز' in q_lower or 'مركز' in q_lower:
        reply = "حصلت المكتبة على المركز الأول أعوام 2018 و2019 و2021 و2023 و2024، وحصلت على المستوى الثاني عام 2025."
        return reply, reply

    # البحث في الكتب (الإكسل)
    excel_result, speech_summary = search_excel_books(user_query)
    if excel_result:
        return excel_result, speech_summary

    # التحيات والدردشة العامة
    if any(word in q_lower for word in ['شكرا', 'تسلم', 'ممتاز', 'رائع', 'يعطيك العافية', 'كفوا', 'تسلم ايدك']):
        reply = "مرحباً بك! يسعدني تقديم المساعدة دائماً لخدمة رواد مكتبة كلية الهندسة. هل تحتاج لأي مساعدة أخرى؟"
        return reply, reply

    if any(word in q_lower for word in ['غبي', 'حمار', 'زفت', 'احم', 'أحمق', 'لعن', 'كلب'] or re.search(r'(شتم|اوسخ|تبا)', q_lower)):
        reply = "أنا موظف استقبال رسمي لمكتبة كلية الهندسة، وهدفي هو مساعدتك بكل احترام وتقدير. تفضل بسؤالك العلمي أو الإداري وسأكون سعيداً بخدمتك."
        return reply, reply

    try:
        prompt = f"""أنت موظف استقبال رسمي واحترافي لمكتبة كلية الهندسة بشبين الكوم (جامعة المنوفية).
أجب بدقة تامة واحترافية وبشكل منظم عن استفسار الزائر بناءً على معلومات المكتبة التالية:

معلومات المكتبة:
{LIBRARY_DATA}

سؤال الزائر: {user_query}"""

        response = model.generate_content(prompt)
        if response and response.text:
            ans = response.text.strip()
            return ans, ans
    except Exception as e:
        print("Gemini API Error details:", e)

    fallback = "مرحباً بك في مكتبة كلية الهندسة. يمكنك السؤال عن فريق العمل، الخدمات، التدوين الصوتي، قواعد الإعارة، أو البحث عن الكتب."
    return fallback, fallback


@app.route("/")
def home():
    # الواجهة التفاعلية العصرية الديناميكية
    return render_template_string('''
        <!DOCTYPE html>
        <html lang="ar" dir="rtl">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
            <title>المساعد الصوتي الرقمي - مكتبة كلية الهندسة</title>
            <link href="https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700;800&display=swap" rel="stylesheet">
            <style>
                * { box-sizing: border-box; margin: 0; padding: 0; }
                
                body { 
                    font-family: 'Tajawal', 'Segoe UI', Tahoma, sans-serif; 
                    background: linear-gradient(135deg, #0f172a, #1e293b, #1e3a8a);
                    background-size: 400% 400%;
                    animation: gradientBG 15s ease infinite;
                    color: #0f172a; 
                    height: 100vh; 
                    display: flex; 
                    justify-content: center; 
                    align-items: center; 
                    overflow: hidden;
                    position: relative;
                }

                @keyframes gradientBG {
                    0% { background-position: 0% 50%; }
                    50% { background-position: 100% 50%; }
                    100% { background-position: 0% 50%; }
                }

                .bg-glow-1, .bg-glow-2 {
                    position: absolute;
                    border-radius: 50%;
                    filter: blur(80px);
                    opacity: 0.4;
                    z-index: 0;
                    animation: floatGlow 10s ease-in-out infinite alternate;
                }
                .bg-glow-1 { width: 300px; height: 300px; background: #3b82f6; top: 10%; right: 10%; }
                .bg-glow-2 { width: 250px; height: 250px; background: #6366f1; bottom: 10%; left: 10%; animation-delay: -5s; }

                @keyframes floatGlow {
                    0% { transform: translate(0, 0) scale(1); }
                    100% { transform: translate(30px, -40px) scale(1.15); }
                }

                .chat-container { 
                    width: 95%; 
                    max-width: 460px; 
                    height: 92vh; 
                    max-height: 780px; 
                    display: flex; 
                    flex-direction: column; 
                    overflow: hidden; 
                    background: rgba(255, 255, 255, 0.88);
                    backdrop-filter: blur(16px);
                    -webkit-backdrop-filter: blur(16px);
                    border-radius: 28px; 
                    box-shadow: 0 20px 50px rgba(0, 0, 0, 0.35), 0 0 0 1px rgba(255, 255, 255, 0.3) inset; 
                    z-index: 10;
                    transition: all 0.3s ease;
                }
                
                .chat-header { 
                    background: linear-gradient(135deg, #1e3a8a, #2563eb, #3b82f6); 
                    color: #ffffff; 
                    padding: 18px 16px; 
                    text-align: center; 
                    border-bottom: 1px solid rgba(255, 255, 255, 0.2); 
                    display: flex; 
                    flex-direction: column; 
                    gap: 6px; 
                    align-items: center; 
                    box-shadow: 0 4px 15px rgba(0, 0, 0, 0.12);
                    position: relative;
                }
                
                .uni-badge { 
                    font-size: 0.75rem; 
                    background: rgba(255, 255, 255, 0.18); 
                    padding: 3px 14px; 
                    border-radius: 20px; 
                    font-weight: 700;
                    letter-spacing: 0.5px;
                    border: 1px solid rgba(255, 255, 255, 0.25);
                }

                .header-title { 
                    font-size: 1.15rem; 
                    font-weight: 800; 
                    display: flex; 
                    align-items: center; 
                    gap: 8px; 
                }

                .header-subtitle { 
                    font-size: 0.82rem; 
                    color: #e0f2fe; 
                    font-weight: 500; 
                    opacity: 0.9;
                }

                .live-indicator { 
                    width: 9px; 
                    height: 9px; 
                    background: #22c55e; 
                    border-radius: 50%; 
                    display: inline-block; 
                    box-shadow: 0 0 10px #22c55e; 
                    animation: pulseGlow 2s infinite;
                }

                @keyframes pulseGlow {
                    0% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(34, 197, 94, 0.7); }
                    70% { transform: scale(1.1); box-shadow: 0 0 0 8px rgba(34, 197, 94, 0); }
                    100% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(34, 197, 94, 0); }
                }

                .quick-chips { 
                    display: flex; 
                    gap: 8px; 
                    padding: 10px 14px; 
                    background: rgba(248, 250, 252, 0.6); 
                    overflow-x: auto; 
                    white-space: nowrap; 
                    border-bottom: 1px solid rgba(226, 232, 240, 0.8); 
                    scrollbar-width: none; 
                }
                .quick-chips::-webkit-scrollbar { display: none; }
                
                .chip { 
                    background: #ffffff; 
                    color: #1e3a8a; 
                    border: 1px solid #cbd5e1; 
                    padding: 6px 14px; 
                    border-radius: 18px; 
                    font-size: 0.8rem; 
                    cursor: pointer; 
                    transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1); 
                    font-weight: 700; 
                    flex-shrink: 0; 
                    box-shadow: 0 2px 5px rgba(0, 0, 0, 0.03);
                    display: flex;
                    align-items: center;
                    gap: 4px;
                }
                .chip:hover { 
                    background: linear-gradient(135deg, #2563eb, #1d4ed8); 
                    color: #ffffff; 
                    border-color: #2563eb;
                    transform: translateY(-2px);
                    box-shadow: 0 4px 12px rgba(37, 99, 235, 0.3);
                }

                .chat-box { 
                    flex: 1; 
                    overflow-y: auto; 
                    padding: 16px; 
                    display: flex; 
                    flex-direction: column; 
                    gap: 14px; 
                    scroll-behavior: smooth; 
                    background: transparent;
                }

                .message { 
                    padding: 12px 16px; 
                    border-radius: 18px; 
                    line-height: 1.6; 
                    max-width: 88%; 
                    font-size: 0.92rem; 
                    position: relative; 
                    white-space: pre-line; 
                    word-wrap: break-word; 
                    animation: messageIn 0.35s cubic-bezier(0.18, 0.89, 0.32, 1.28) forwards;
                    box-shadow: 0 3px 10px rgba(0, 0, 0, 0.04);
                }

                @keyframes messageIn {
                    from { opacity: 0; transform: translateY(15px) scale(0.96); }
                    to { opacity: 1; transform: translateY(0) scale(1); }
                }

                .user { 
                    background: linear-gradient(135deg, #2563eb, #1d4ed8); 
                    color: #ffffff; 
                    margin-left: auto; 
                    border-bottom-left-radius: 4px; 
                    font-weight: 500;
                }

                .bot { 
                    background: #ffffff; 
                    color: #1e293b; 
                    margin-right: auto; 
                    border-bottom-right-radius: 4px; 
                    border: 1px solid #e2e8f0; 
                    font-weight: 500;
                }

                .audio-container {
                    margin-top: 10px;
                    background: #f1f5f9;
                    padding: 6px;
                    border-radius: 12px;
                    border: 1px solid #e2e8f0;
                }

                .audio-notice {
                    font-size: 0.78rem;
                    color: #b91c1c;
                    margin-top: 8px;
                }
                
                .controls-wrapper { 
                    padding: 14px; 
                    background: rgba(255, 255, 255, 0.95); 
                    border-top: 1px solid rgba(226, 232, 240, 0.8); 
                }

                .controls { 
                    display: flex; 
                    gap: 8px; 
                    background: #f8fafc; 
                    padding: 6px 8px 6px 14px; 
                    border-radius: 30px; 
                    border: 1.5px solid #cbd5e1; 
                    align-items: center; 
                    transition: all 0.3s ease;
                }
                .controls:focus-within {
                    border-color: #2563eb;
                    box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.15);
                    background: #ffffff;
                }

                input[type="text"] { 
                    flex: 1; 
                    background: transparent; 
                    border: none; 
                    outline: none; 
                    color: #0f172a; 
                    font-size: 0.93rem; 
                    font-family: 'Tajawal', sans-serif;
                    padding: 6px 0; 
                    min-width: 0; 
                    font-weight: 500;
                }

                button { 
                    background: linear-gradient(135deg, #2563eb, #1d4ed8); 
                    color: #ffffff; 
                    border: none; 
                    border-radius: 50%; 
                    width: 40px; 
                    height: 40px; 
                    cursor: pointer; 
                    font-weight: bold; 
                    display: flex; 
                    align-items: center; 
                    justify-content: center; 
                    flex-shrink: 0; 
                    transition: all 0.25s ease; 
                    box-shadow: 0 3px 8px rgba(37, 99, 235, 0.3);
                }
                button:hover { 
                    transform: scale(1.06); 
                    box-shadow: 0 5px 14px rgba(37, 99, 235, 0.45);
                }

                .mic-btn { 
                    background: linear-gradient(135deg, #10b981, #059669); 
                    color: #ffffff; 
                    box-shadow: 0 3px 8px rgba(16, 185, 129, 0.3);
                }
                .mic-btn:hover {
                    box-shadow: 0 5px 14px rgba(16, 185, 129, 0.45);
                }

                .mic-btn.listening { 
                    background: linear-gradient(135deg, #ef4444, #dc2626); 
                    animation: pulseMic 1.2s infinite alternate; 
                    box-shadow: 0 0 15px rgba(239, 68, 68, 0.6);
                }

                @keyframes pulseMic { 
                    0% { transform: scale(1); } 
                    100% { transform: scale(1.15); } 
                }

                audio { 
                    margin-top: 8px; 
                    width: 100%; 
                    height: 32px; 
                    border-radius: 8px; 
                }

                /* ===== إضافة صغيرة: ستايل زر التشغيل اليدوي ===== */
                .play-btn {
                    margin-bottom: 8px;
                    padding: 6px 14px;
                    border-radius: 8px;
                    background: linear-gradient(135deg, #2563eb, #1d4ed8);
                    color: #ffffff;
                    border: none;
                    cursor: pointer;
                    font-size: 0.85rem;
                    font-family: 'Tajawal', sans-serif;
                    font-weight: 700;
                    width: auto;
                    height: auto;
                }
            </style>
        </head>
        <body>
            <div class="bg-glow-1"></div>
            <div class="bg-glow-2"></div>

            <div class="chat-container">
                <div class="chat-header">
                    <span class="uni-badge">جامعة المنوفية 🏛️</span>
                    <div class="header-title">
                        <span class="live-indicator"></span> مكتبة كلية الهندسة
                    </div>
                    <div class="header-subtitle">المساعد الصوتي الرقمي الذكي</div>
                </div>

                <div class="quick-chips">
                    <div class="chip" onclick="sendPreset('التسجيل على بنك المعرفة')">🌐 بنك المعرفة</div>
                    <div class="chip" onclick="sendPreset('فريق العمل')">👥 فريق العمل</div>
                    <div class="chip" onclick="sendPreset('تدوين صوتي')">🎧 البودكاست</div>
                    <div class="chip" onclick="sendPreset('مجلة الكلية')">📚 المجلة</div>
                </div>

                <div class="chat-box" id="chatBox"></div>

                <div class="controls-wrapper">
                    <div class="controls">
                        <input type="text" id="userInput" placeholder="اكتب سؤالك هنا أو اختر من الأعلى..." onkeypress="if(event.key === 'Enter') sendMessage()">
                        <button id="micBtn" class="mic-btn" onclick="toggleVoiceInput()" title="تحدث صوتياً">🎤</button>
                        <button onclick="sendMessage()" title="إرسال">⬆</button>
                    </div>
                </div>
            </div>

            <script>
                let recognition = null;
                let isListening = false;

                window.onload = function() {
                    loadWelcomeMessage();
                };

                function appendBotMessage(data) {
                    const chatBox = document.getElementById('chatBox');
                    const botMessageDiv = document.createElement('div');
                    botMessageDiv.className = 'message bot';

                    const textDiv = document.createElement('div');
                    textDiv.innerText = data.reply;
                    botMessageDiv.appendChild(textDiv);

                    if (data.audio_url) {
                        const audioContainer = document.createElement('div');
                        audioContainer.className = 'audio-container';

                        // ===== التعديل 5: زر تشغيل يدوي بدل autoplay =====
                        const playBtn = document.createElement('button');
                        playBtn.className = 'play-btn';
                        playBtn.innerText = '🔊 تشغيل الرد الصوتي';

                        const audioElem = document.createElement('audio');
                        audioElem.controls = true;
                        // audioElem.autoplay = true;  // ← اتشال عشان Chrome بيمنع
                        audioElem.src = data.audio_url;

                        playBtn.onclick = function() {
                            audioElem.play().catch(err => {
                                console.error('فشل التشغيل:', err);
                                alert('فشل التشغيل، جرّب مرة أخرى');
                            });
                        };

                        audioContainer.appendChild(playBtn);
                        audioContainer.appendChild(audioElem);
                        botMessageDiv.appendChild(audioContainer);
                    } else {
                        const noticeDiv = document.createElement('div');
                        noticeDiv.className = 'audio-notice';
                        noticeDiv.innerText = 'تعذر إنشاء ملف الصوت حالياً.';
                        botMessageDiv.appendChild(noticeDiv);
                    }

                    chatBox.appendChild(botMessageDiv);
                    chatBox.scrollTop = chatBox.scrollHeight;
                }

                async function loadWelcomeMessage() {
                    try {
                        const response = await fetch('/welcome');
                        const data = await response.json();
                        appendBotMessage(data);
                    } catch (error) {
                        console.error("Welcome Error:", error);
                    }
                }

                function sendPreset(text) {
                    document.getElementById('userInput').value = text;
                    sendMessage();
                }

                function toggleVoiceInput() {
                    const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
                    if (!SpeechRecognition) { alert("متصفحك لا يدعم الإدخال الصوتي."); return; }
                    if (isListening) { stopVoiceInput(); return; }
                    recognition = new SpeechRecognition();
                    recognition.lang = 'ar-EG';
                    recognition.onstart = function() { isListening = true; document.getElementById('micBtn').classList.add('listening'); };
                    recognition.onresult = function(event) { document.getElementById('userInput').value = event.results[0][0].transcript; sendMessage(); };
                    recognition.onend = function() { stopVoiceInput(); };
                    recognition.start();
                }

                function stopVoiceInput() {
                    isListening = false;
                    const micBtn = document.getElementById('micBtn');
                    if (micBtn) micBtn.classList.remove('listening');
                    if (recognition) { try { recognition.stop(); } catch(e){} }
                }

                async function sendMessage() {
                    stopVoiceInput();
                    const input = document.getElementById('userInput');
                    const chatBox = document.getElementById('chatBox');
                    const userText = input.value.trim();
                    if (!userText) return;

                    chatBox.innerHTML += `<div class="message user">${userText}</div>`;
                    input.value = '';
                    chatBox.scrollTop = chatBox.scrollHeight;

                    try {
                        document.querySelectorAll('audio').forEach(audio => {
                            audio.pause();
                            audio.currentTime = 0;
                            audio.remove();
                        });

                        const response = await fetch('/chat', {
                            method: 'POST',
                            headers: { 'Content-Type': 'application/json' },
                            body: JSON.stringify({ message: userText })
                        });
                        const data = await response.json();
                        appendBotMessage(data);
                    } catch (error) {
                        console.error("Error:", error);
                    }
                }
            </script>
        </body>
        </html>
    ''')


@app.route("/welcome", methods=["GET"])
def welcome():
    reply_text = "مرحباً بكم، أنا موظف استقبال صوتي لمكتبة كلية الهندسة بجامعة المنوفية. كيف يمكنني مساعدتك اليوم؟"
    speech_text = clean_text_for_speech(reply_text)
    audio_url, audio_error = synthesize_speech(speech_text, "welcome")

    return jsonify({
        "reply": reply_text,
        "audio_url": audio_url,
        "audio_error": audio_error
    })


@app.route("/chat", methods=["POST"])
def chat():
    data = request.get_json()
    user_msg = data.get("message", "")

    reply_text, speech_summary = ask_gemini_library_assistant(user_msg)

    speech_text = clean_text_for_speech(speech_summary if speech_summary else reply_text)
    if not speech_text:
        speech_text = "عذراً، لم أتمكن من قراءة الإجابة."

    audio_url, audio_error = synthesize_speech(speech_text, "response")

    return jsonify({
        "reply": reply_text,
        "audio_url": audio_url,
        "audio_error": audio_error
    })


@app.errorhandler(404)
def page_not_found(e):
    return "الصفحة غير موجودة، تأكد من صحة الرابط.", 404


if __name__ == '__main__':
    app.run(debug=True)