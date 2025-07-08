import streamlit as st
import pandas as pd
import re
import io
# ==============================
# Hàm chuẩn hóa số điện thoại
import re
import pandas as pd
import phonenumbers
from phonenumbers import geocoder




# Danh sách mã quốc gia phổ biến để tự động thêm dấu +
COUNTRY_CODES = {
    '886': 'Taiwan',
    '1': 'USA/Canada',
    '81': 'Japan',
    '82': 'South Korea',
    '85': 'Hong Kong',
    '86': 'China',
    '855': 'Cambodia',
    '856': 'Laos',
    '95': 'Myanmar',
    '44': 'UK',
    '61': 'Australia',
    '65': 'Singapore',
    '66': 'Thailand',
}

def normalize_phone(phone):
    if pd.isna(phone):
        return None

    # Làm sạch ký tự đặc biệt, chỉ giữ số và dấu +
    phone = str(phone).strip()
    phone = re.sub(r'[^\d+]', '', phone)

    # 1️⃣ Xử lý số Việt Nam bắt đầu bằng +84 hoặc 84
    if phone.startswith('+84'):
        phone = '0' + phone[3:]
    elif phone.startswith('84') and len(phone) in [10, 11]:
        phone = '0' + phone[2:]

    # 2️⃣ Nếu giờ là số Việt Nam:
    # - Di động: 10 số, bắt đầu từ 03-09
    # - Bàn: 11 số, bắt đầu từ 02
    if (phone.startswith('02') and len(phone) == 11) or \
       (phone.startswith(('03', '04', '05', '06', '07', '08', '09')) and len(phone) == 10):
        return phone

    # 3️⃣ Nếu có 9 số và bắt đầu từ 3–9 → thêm 0 rồi kiểm tra lại
    if len(phone) == 9 and phone[0] in '3456789':
        phone = '0' + phone
        if (phone.startswith('02') and len(phone) == 11) or \
           (phone.startswith(('03', '04', '05', '06', '07', '08', '09')) and len(phone) == 10):
            return phone

    # 4️⃣ Nếu có dấu + → xử lý bằng thư viện phonenumbers
    if phone.startswith('+'):
        try:
            parsed = phonenumbers.parse(phone, None)
            if phonenumbers.is_valid_number(parsed):
                country = geocoder.description_for_number(parsed, 'en')
                if parsed.country_code == 84:
                    return None  # Không trả về số Việt Nam dạng +84 nữa
                return f"{phone} / {country}"
        except:
            return None

    # 5️⃣ Nếu không có dấu + nhưng bắt đầu bằng mã quốc gia → thêm +
    for code in sorted(COUNTRY_CODES.keys(), key=lambda x: -len(x)):
        if phone.startswith(code) and len(phone) >= len(code) + 7:
            fake_plus = '+' + phone
            try:
                parsed = phonenumbers.parse(fake_plus, None)
                if phonenumbers.is_valid_number(parsed):
                    country = geocoder.description_for_number(parsed, 'en')
                    if parsed.country_code == 84:
                        return None
                    return f"{fake_plus} / {country}"
            except:
                continue

    # ❌ Không hợp lệ
    return None



# ==============================
# Hàm chuẩn hóa tên khách
def clean_name(name):
    if pd.isna(name): return ''
    name = str(name).strip()
    name = re.sub(r'\s+', ' ', name)
    return name.title()

# ==============================
# Giao diện Streamlit
st.set_page_config(page_title="Thống kê khách hàng siêng học", layout="wide")
st.title("📊 Thống Kê Khách Hàng Siêng Năng Nhất Theo Số Lớp Học Offline")

uploaded_file = st.file_uploader("📥 Kéo thả file Excel có nhiều sheets (mỗi sheet là một lớp học):", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    customer_dict = {}
    all_cleaned_rows = []

    for sheet in xls.sheet_names:
        try:
            df = xls.parse(sheet, skiprows=2, usecols="D:E", names=["Tên", "SĐT"])
            
            # Chuẩn hóa dữ liệu
            df["Tên"] = df["Tên"].apply(clean_name)
            df["SĐT"] = df["SĐT"].apply(normalize_phone)

            # Xóa dòng không có SĐT
            df = df.dropna(subset=["SĐT"])

            for _, row in df.iterrows():
                name = row["Tên"]
                phone = row["SĐT"]

                if not phone:
                    continue  # Bỏ dòng nếu không có SĐT

                name = name or "Không rõ"  # Gán tên mặc định nếu rỗng

                # ✅ Ghi lại khách đã chuẩn hóa
                all_cleaned_rows.append({
                    "Tên lớp": sheet,
                    "Tên khách": name,
                    "SĐT": phone
                })

                # ✅ Ghi nhận cho thống kê lớp
                key = (name, phone)
                if key not in customer_dict:
                    customer_dict[key] = set()
                customer_dict[key].add(sheet)
        except Exception as e:
            st.warning(f"⚠️ Sheet '{sheet}' lỗi: {e}")
            continue

    if customer_dict:
        result_df = pd.DataFrame([
            {
                "Tên khách": name,
                "SĐT": phone,
                "Số lớp đã tham gia": len(classes),
                "Tên lớp đã tham gia": ', '.join(sorted(classes))
            }
            for (name, phone), classes in customer_dict.items()
        ])

        result_df = result_df.sort_values(by="Số lớp đã tham gia", ascending=False)
        st.success("✅ Đã xử lý xong dữ liệu!")

        st.dataframe(result_df, use_container_width=True)

        # ✅ Tải file thống kê
        buffer = io.BytesIO()
        result_df.to_excel(buffer, index=False)
        st.download_button(
            "📤 Tải kết quả về Excel",
            buffer.getvalue(),
            file_name="khach_sieng_nang.xlsx",
            key="download_summary"
        )

        # ✅ Tải file toàn bộ khách đã chuẩn hóa
        if all_cleaned_rows:
            cleaned_df = pd.DataFrame(all_cleaned_rows)
            buffer_all = io.BytesIO()
            cleaned_df.to_excel(buffer_all, index=False)
            st.download_button(
                "🧾 Tải file tất cả khách đã chuẩn hóa",
                buffer_all.getvalue(),
                file_name="tat_ca_khach_sach.xlsx",
                key="download_all_cleaned"
            )
        else:
            st.info("Không tìm thấy khách hàng hợp lệ trong file.")
    else:
        st.warning("Không tìm thấy khách hàng hợp lệ nào để thống kê.")
else:
    st.info("📂 Vui lòng upload file Excel để bắt đầu.")
