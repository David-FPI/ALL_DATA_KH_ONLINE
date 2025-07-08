import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="📊 Thống kê lớp học ZOOM", layout="wide")
st.title("📊 Tổng hợp số lần tham gia & điểm danh theo nhân viên")

# ======================= HÀM XUẤT EXCEL =======================
def to_excel_file(df, sheet_name="Sheet1"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)
    return output

# ======================= HÀM ĐỌC TÊN NHÂN VIÊN =======================
def extract_all_names(files, selected_col, additional_cols):
    all_data = []
    for file in files:
        xls = pd.ExcelFile(file)
        for sheet in xls.sheet_names:
            try:
                df = pd.read_excel(xls, sheet_name=sheet, skiprows=2, engine="openpyxl", dtype=str)
                if selected_col not in df.columns:
                    continue
                sub_df = df[[selected_col] + [col for col in additional_cols if col in df.columns]].copy()
                sub_df = sub_df.dropna(subset=[selected_col])
                sub_df["Sheet"] = sheet
                all_data.append(sub_df)
            except Exception as e:
                st.warning(f"⚠️ Lỗi sheet `{sheet}` trong file `{file.name}`: {e}")
    if all_data:
        return pd.concat(all_data, ignore_index=True)
    return pd.DataFrame()


# ======================= GIAO DIỆN UPLOAD =======================
uploaded_files = st.file_uploader("📤 Tải lên các file Excel (.xlsx)", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    try:
        # Đọc file đầu tiên, sheet đầu tiên để lấy sample
        sample_df = pd.read_excel(uploaded_files[0], sheet_name=0, skiprows=2, engine="openpyxl", nrows=5, dtype=str)
        st.subheader("📃 Data mẫu (5 dòng đầu tiên trong sheet đầu tiên)")
        st.dataframe(sample_df)

        # Chọn cột chính và các cột phụ
        selected_col = st.selectbox("🔍 Chọn cột để thống kê tên nhân viên", options=sample_df.columns.tolist())
        additional_cols = st.multiselect("➕ Chọn các cột bổ sung đi kèm", options=[col for col in sample_df.columns if col != selected_col])
    except Exception as e:
        st.error(f"❌ Không thể đọc dữ liệu mẫu: {e}")
        st.stop()

    all_names = extract_all_names(uploaded_files, selected_col, additional_cols)

    if not all_names.empty:

        df_raw = all_names.rename(columns={selected_col: "Tên Nhân Viên", "Sheet": "Buổi Học (Sheet)"})

        # ======================= BẢNG 1 =======================
        with st.expander("📋 Danh sách tên nhân viên và buổi học (không gộp)", expanded=True):
            show_cols_1 = st.multiselect("🧩 Chọn cột hiển thị", df_raw.columns.tolist(), default=df_raw.columns.tolist())


            st.dataframe(df_raw[show_cols_1], use_container_width=True)

            st.download_button("📥 Tải Excel bảng này", data=to_excel_file(df_raw, "DanhSach"), file_name="bang_1_danh_sach.xlsx")

        # ======================= BẢNG 2 =======================
        with st.expander("📊 Tổng hợp tham gia lớp học theo nhân viên", expanded=True):
            df_summary = (
                df_raw.groupby("Tên Nhân Viên")
                .agg({
                    "Buổi Học (Sheet)": [
                        "count",
                        lambda x: len(set(x)),
                        lambda x: ", ".join(sorted(set(x)))
                    ]
                })
            )
            df_summary.columns = ["Tổng số lần xuất hiện", "Số buổi học khác nhau", "Danh sách buổi học"]
            df_summary = df_summary.reset_index().sort_values(by="Tổng số lần xuất hiện", ascending=False)

            # 👉 Gộp thêm các cột bổ sung từ df_raw (lấy giá trị đầu tiên)
            if additional_cols:
                add_info = df_raw.groupby("Tên Nhân Viên")[additional_cols].first().reset_index()
                df_summary = pd.merge(df_summary, add_info, on="Tên Nhân Viên", how="left")

            show_cols_2 = st.multiselect("🧩 Chọn cột hiển thị", df_summary.columns.tolist(), default=df_summary.columns.tolist())
            st.dataframe(df_summary[show_cols_2], use_container_width=True)
            st.download_button("📥 Tải Excel tổng hợp", data=to_excel_file(df_summary, "TongHop"), file_name="bang_2_tong_hop.xlsx")


        # ======================= BẢNG 3 =======================
        with st.expander("✅ Bảng chấm công (Nhân viên x Buổi học)", expanded=True):
            df_pivot = pd.pivot_table(
                df_raw,
                index="Tên Nhân Viên",
                columns="Buổi Học (Sheet)",
                aggfunc="size",
                fill_value=0
            )
            df_pivot = df_pivot.applymap(lambda x: "✅" if x > 0 else "")
            df_pivot["Tổng buổi tham gia"] = (df_pivot == "✅").sum(axis=1)
            df_pivot = df_pivot.reset_index()

            # 👉 Gộp thêm các cột bổ sung từ df_raw
            if additional_cols:
                add_info = df_raw.groupby("Tên Nhân Viên")[additional_cols].first().reset_index()
                df_pivot = pd.merge(df_pivot, add_info, on="Tên Nhân Viên", how="left")

            show_cols_3 = st.multiselect("🧩 Chọn cột hiển thị", df_pivot.columns.tolist(), default=df_pivot.columns.tolist())
            st.dataframe(df_pivot[show_cols_3], use_container_width=True)
            st.download_button("📥 Tải Excel bảng chấm công", data=to_excel_file(df_pivot, "ChamCong"), file_name="bang_3_cham_cong.xlsx")


    else:
        st.warning("⚠️ Không tìm thấy dữ liệu hợp lệ trong cột đã chọn.")
