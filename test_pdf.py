import requests
import json

url = "http://127.0.0.1:8000/generate-pdf"
payload = {
    "SERVICE_CODE": "TEST-001",
    "ID_NUMBER": "1234567890",
    "NAME_AR": "تجربة مستخدم",
    "NAME_EN": "Test User",
    "DAYS_COUNT": 3,
    "ENTRY_DATE_GREGORIAN": "2023-10-01",
    "EXIT_DATE_GREGORIAN": "2023-10-04",
    "ENTRY_DATE_HIJRI": "16/03/1445",
    "EXIT_DATE_HIJRI": "19/03/1445",
    "REPORT_ISSUE_DATE": "2023-10-01",
    "NATIONALITY_AR": "سعودي",
    "NATIONALITY_EN": "Saudi",
    "DOCTOR_NAME_AR": "د. أحمد",
    "DOCTOR_NAME_EN": "Dr. Ahmed",
    "JOB_TITLE_AR": "طبيب",
    "JOB_TITLE_EN": "Doctor",
    "HOSPITAL_NAME_AR": "مستشفى الأمل",
    "HOSPITAL_NAME_EN": "Hope Hospital",
    "PRINT_DATE": "Tuesday, 22 April 2025",
    "PRINT_TIME": "12:32 PM"
}

print(f"إرسال طلب تجريبي إلى: {url}...")
try:
    response = requests.post(url, json=payload)
    if response.status_code == 200:
        print("✅ نجاح! تم استلام ملف PDF.")
        with open("test_output.pdf", "wb") as f:
            f.write(response.content)
        print("📁 تم حفظ ملف الاختبار باسم 'test_output.pdf' في مجلد الباك اند.")
    else:
        print(f"❌ فشل الاستجابة: {response.status_code}")
        print(response.text)
except Exception as e:
    print(f"❌ خطأ أثناء الاتصال: {str(e)}")
