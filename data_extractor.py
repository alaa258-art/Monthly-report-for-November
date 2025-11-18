import pandas as pd
import json
import numpy as np
import os
import sys

# ==============================================================================
# الإعدادات: يجب تعديل هذه المتغيرات عند تغيير الملف
# ==============================================================================

# اسم ملف الإكسيل الذي سيتم معالجته. يجب أن يكون الملف في نفس مجلد هذا السكريبت.
EXCEL_FILE_NAME = 'التقريرالشهري.xlsx' 

# اسم ملف JSON الناتج الذي سيتم استخدامه مع صفحة الويب الآمنة.
JSON_OUTPUT_FILE = 'StudentDatamonth.json' # تم تغيير الاسم بناءً على متطلبات المستخدم

# أسماء أوراق العمل (Sheets) التي تحتوي على بيانات الطلاب.
# تم التأكد من الأسماء من ملف الإكسيل المرفق.
SHEET_NAMES = ['GRADE 5', 'GRADE 6', 'GRADE 7', 'GRADE 8']

# ==============================================================================
# هيكل البيانات: تم تعديل هذا الجزء بناءً على تحليل ملف الإكسيل المرفق
# ==============================================================================

# خريطة المواد والترجمة الإنجليزية (للاستخدام في كود الويب)
# تم تبسيطها لتتطابق مع الأعمدة الفعلية في ملف الإكسيل
SUBJECT_COLUMNS_MAP = [
    ('التربية الإسلامية', 'islamic_education'),
    ('اللغة العربية', 'arabic_language'),
    ('ENGLISH', 'english_language'), # تم تغيير الاسم ليتطابق مع الإكسيل
    ('الدراسات الاجتماعية', 'social_studies'),
    ('Math', 'mathematics'), # تم تغيير الاسم ليتطابق مع الإكسيل
    ('العلوم', 'science'), # سيتم التعامل مع 'العلوم العامة' و 'علوم' في الكود
]

# عدد الأعمدة الفرعية لكل مادة (5 أعمدة)
# بناءً على تحليل الإكسيل، كل مادة تأخذ 5 أعمدة (درجة، مستوى، سلوك، إلخ)
SUB_COLUMNS_COUNT = 5

# مؤشرات الأعمدة الرئيسية (0-based index)
# تم التأكد من هذه المؤشرات من تحليل ملف الإكسيل
NATIONAL_ID_COL_IDX = 0  # رقم الهوية الوطنية
STUDENT_ID_COL_IDX = 1   # رقم الطالب
STUDENT_NAME_COL_IDX = 2 # اسم الطالب (العمود الذي يحمل العنوان الطويل)
CLASS_COL_IDX = 3        # الصف (العمود الرابع)
BEHAVIOR_COL_IDX = 4     # السلوك (العمود الخامس)

# مؤشر بداية أعمدة الدرجات (العمود السادس)
GRADES_START_COL_IDX = 5

# ==============================================================================
# الدالة الرئيسية لاستخراج البيانات
# ==============================================================================

def extract_data_to_json(file_path):
    """
    يقرأ ملف الإكسيل ويستخرج بيانات الطلاب إلى ملف JSON بهدف البحث المزدوج.
    """
    all_students_data = {}
    total_students = 0
    
    print(f"بدء معالجة ملف الإكسيل: {file_path}")

    for sheet_name in SHEET_NAMES:
        try:
            print(f"  - معالجة ورقة العمل: {sheet_name}...")
            
            # قراءة ورقة العمل، مع تحديد الصف الرابع (skiprows=2) كبداية للبيانات الفعلية
            # هذا يضمن أن يتم تضمين الصف الذي يحتوي على أول طالب (الصف الرابع في الإكسيل)
            df = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=2, header=None, engine='openpyxl')
            
            # إعادة تسمية الأعمدة لتسهيل الوصول إليها
            # نستخدم الأعمدة التي تم تحديدها مسبقاً
            df.rename(columns={
                NATIONAL_ID_COL_IDX: 'National_ID',
                STUDENT_ID_COL_IDX: 'Student_ID',
                STUDENT_NAME_COL_IDX: 'Student_Name',
                CLASS_COL_IDX: 'Class_Name',
                BEHAVIOR_COL_IDX: 'General_Behavior'
            }, inplace=True)
            
            # التأكد من وجود الأعمدة الأساسية
            if len(df.columns) < GRADES_START_COL_IDX:
                print(f"    [خطأ] عدد الأعمدة في ورقة {sheet_name} غير كافٍ. تم تخطي الورقة.")
                continue

            # تصفية الصفوف التي لا تحتوي على رقم طالب
            # نستخدم العمود الذي تم إعادة تسميته 'Student_ID' للتصفية
            df.dropna(subset=['Student_ID'], inplace=True)
            
            # التكرار على صفوف البيانات (الطلاب)
            for index, row in df.iterrows():
                try:
                    # استخراج البيانات الأساسية باستخدام المؤشرات
                    student_id_raw = row['Student_ID']
                    national_id_raw = row['National_ID']
                    
                    # تنظيف رقم الطالب (تحويله إلى نص وإزالة .0 إذا وجدت)
                    student_id = str(student_id_raw).strip()
                    if student_id.endswith('.0'):
                        student_id = student_id[:-2]
                    
                    # تنظيف رقم الهوية الوطنية (تحويله إلى نص وإزالة .0 إذا وجدت)
                    national_id = str(national_id_raw).strip()
                    if national_id.endswith('.0'):
                        national_id = national_id[:-2]
                    
                    # التحقق من صحة البيانات الأساسية
                    if not student_id or not national_id or national_id == 'nan' or student_id == 'nan':
                        continue
                        
                    # المفتاح المزدوج للبحث: "رقم الطالب_رقم الهوية الوطنية"
                    combined_key = f"{student_id}_{national_id}"
                        
                    # اسم الطالب (العمود الذي تم إعادة تسميته)
                    student_name = str(row['Student_Name']).strip()
                    
                    # إذا كان اسم الطالب هو العنوان الطويل، فهذا يعني أن الصف غير صحيح، نتخطاه
                    if student_name == 'التقرير الشهري للتقويم الثاني للفصل الاول 2025/2026':
                        continue

                    student_data = {
                        'student_name': student_name,
                        'class_name': str(row['Class_Name']).strip(),
                        'general_behavior': str(row['General_Behavior']).strip(),
                        'grades': {}
                    }
                    
                    # استخراج الدرجات
                    current_col_idx = GRADES_START_COL_IDX
                    
                    # قائمة بأسماء الأعمدة الفرعية (التي تظهر في الويب)
                    sub_column_keys = ['formative_exam', 'academic_level', 'participation', 'doing_tasks', 'attending_books']
                    
                    for ar_subject, en_subject in SUBJECT_COLUMNS_MAP:
                        subject_grades = {}
                        
                        # التحقق من حدود الأعمدة
                        if current_col_idx + SUB_COLUMNS_COUNT > len(row):
                            break 
                            
                        # استخراج الأعمدة الفرعية الخمسة للمادة الحالية
                        for i in range(SUB_COLUMNS_COUNT):
                            # نستخدم مؤشر العمود الحالي (لأننا قرأنا بدون رأس، فالمؤشرات هي الأرقام)
                            value = str(row[current_col_idx + i]).strip() if pd.notna(row[current_col_idx + i]) else 'N/A'
                            subject_grades[sub_column_keys[i]] = value
                        
                        # إضافة المادة فقط إذا كانت تحتوي على درجات فعلية
                        if any(v != 'N/A' for v in subject_grades.values()):
                            student_data['grades'][en_subject] = subject_grades
                        
                        # الانتقال إلى بداية المادة التالية
                        current_col_idx += SUB_COLUMNS_COUNT

                    all_students_data[combined_key] = student_data
                    total_students += 1
                
                except Exception as e:
                    # طباعة خطأ في صف معين للمساعدة في التصحيح
                    print(f"    [خطأ في الصف] حدث خطأ أثناء معالجة الصف رقم {index + 3} في ورقة {sheet_name}: {e}")
                    continue

        except Exception as e:
            print(f"  [خطأ عام] حدث خطأ أثناء معالجة ورقة {sheet_name}: {e}")

    # حفظ البيانات في ملف JSON
    with open(JSON_OUTPUT_FILE, 'w', encoding='utf-8') as f:
        json.dump(all_students_data, f, ensure_ascii=False, indent=4)

    print("\n" + "="*50)
    print(f"✓ تم الانتهاء من معالجة البيانات بنجاح.")
    print(f"✓ إجمالي عدد الطلاب المستخرجين: {total_students}")
    print(f"✓ تم حفظ ملف البيانات الجديد في: {JSON_OUTPUT_FILE}")
    print("="*50)

# تنفيذ الدالة
if __name__ == "__main__":
    # التحقق من وجود ملف الإكسيل
    if not os.path.exists(EXCEL_FILE_NAME):
        print(f"!!! خطأ: لم يتم العثور على ملف الإكسيل باسم '{EXCEL_FILE_NAME}' في هذا المجلد.")
        print("!!! يرجى التأكد من وضع ملف الإكسيل في نفس مجلد هذا السكريبت وتسميته بشكل صحيح.")
        sys.exit(1)
    else:
        extract_data_to_json(EXCEL_FILE_NAME)
