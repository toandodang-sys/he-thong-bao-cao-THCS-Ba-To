import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from copy import copy
import io
import os
import re

st.set_page_config(page_title="Hệ thống Báo cáo Tiết dạy", layout="centered")

st.title("Hệ thống Nộp và Tổng hợp Báo cáo")
st.write("Phiên bản Tối ưu (Trích xuất theo đúng tọa độ Cột và Dòng chuẩn của biểu mẫu).")


# --- HÀM COPY ĐỊNH DẠNG ---
def copy_sheet(source_sheet, target_sheet):
    for col, dim in source_sheet.column_dimensions.items():
        target_sheet.column_dimensions[col].width = dim.width
    for row, dim in source_sheet.row_dimensions.items():
        target_sheet.row_dimensions[row].height = dim.height

    for row in source_sheet.iter_rows():
        for cell in row:
            if type(cell).__name__ == 'MergedCell':
                continue
            target_cell = target_sheet.cell(row=cell.row, column=cell.column)
            target_cell.value = cell.value

            if cell.has_style:
                try:
                    if cell.border: target_cell.border = copy(cell.border)
                except:
                    pass
                try:
                    if cell.fill: target_cell.fill = copy(cell.fill)
                except:
                    pass
                try:
                    if cell.number_format: target_cell.number_format = copy(cell.number_format)
                except:
                    pass
                try:
                    if cell.alignment: target_cell.alignment = copy(cell.alignment)
                except:
                    pass

            try:
                if cell.font:
                    new_color = copy(cell.font.color) if cell.font.color else None
                    target_cell.font = Font(
                        name='Times New Roman',
                        size=cell.font.size,
                        bold=cell.font.bold,
                        italic=cell.font.italic,
                        vertAlign=cell.font.vertAlign,
                        underline=cell.font.underline,
                        strike=cell.font.strike,
                        color=new_color
                    )
                else:
                    target_cell.font = Font(name='Times New Roman', size=11)
            except:
                target_cell.font = Font(name='Times New Roman', size=11)

    for merge_range in source_sheet.merged_cells.ranges:
        try:
            target_sheet.merge_cells(str(merge_range))
        except:
            pass

        # --- HÀM TRÍCH XUẤT SỐ ---


def get_num(sheet, row, col):
    """Hàm lấy giá trị số từ 1 ô cụ thể, xử lý cả trường hợp lẫn chữ"""
    v = sheet.cell(row=row, column=col).value
    if v is None:
        return 0
    s = str(v).strip()
    match = re.search(r'\d+(\.\d+)?', s)
    if match:
        return float(match.group()) if '.' in match.group() else int(match.group())
    return 0


# --- HÀM TẠO SHEET TỔNG HỢP ---
def create_summary_sheet(wb_merged, list_of_sheets, nam_hoc, hoc_ky, tuan):
    ws_th = wb_merged.create_sheet(title="Tổng hợp")

    th_font = Font(name='Times New Roman', size=11)
    bold_font = Font(name='Times New Roman', size=11, bold=True)
    center_aligned = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'),
                         bottom=Side(style='thin'))

    ws_th['A1'] = "UBND XÃ BA TƠ"
    ws_th['D1'] = "CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM"
    ws_th['A2'] = "TRƯỜNG THCS BA TƠ"
    ws_th['D2'] = "Độc lập – Tự do – Hạnh phúc"

    for row in range(1, 3):
        ws_th.cell(row=row, column=1).font = bold_font
        ws_th.cell(row=row, column=4).font = bold_font
        ws_th.cell(row=row, column=1).alignment = center_aligned
        ws_th.cell(row=row, column=4).alignment = center_aligned

    ws_th.merge_cells('A1:B1')
    ws_th.merge_cells('D1:G1')
    ws_th.merge_cells('A2:B2')
    ws_th.merge_cells('D2:G2')

    title_text = f"BẢNG TỔNG HỢP THEO DÕI BÁO CÁO THỰC HIỆN CHƯƠNG TRÌNH, TIẾT DẠY HÀNG TUẦN - NĂM HỌC {nam_hoc}"
    ws_th['A4'] = title_text
    ws_th['A4'].font = bold_font
    ws_th['A4'].alignment = center_aligned
    ws_th.merge_cells('A4:J4')

    ws_th['A5'] = f"({hoc_ky})"
    ws_th['A5'].font = th_font
    ws_th['A5'].alignment = center_aligned
    ws_th.merge_cells('A5:J5')

    headers = [
        "TT",
        "Họ và tên CB, giáo viên",
        "Tổng số tiết thực dạy",
        "Số tiết kiêm nhiệm",
        "Số tiết đi công tác",
        "Số tiết dạy thay",
        "Số tiết lấp giờ, tăng tiết",
        "Số tiết coi KT, dự giờ",
        "Tổng số tiết thực hiện",
        "Ghi chú"
    ]

    ws_th.row_dimensions[7].height = 40

    for col, header in enumerate(headers, 1):
        cell = ws_th.cell(row=7, column=col)
        cell.value = header
        cell.font = bold_font
        cell.alignment = center_aligned
        cell.border = thin_border

    ws_th.column_dimensions['A'].width = 5
    ws_th.column_dimensions['B'].width = 25
    for col in ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']:
        ws_th.column_dimensions[col].width = 15

    row_idx = 8
    tt = 1

    for sheet_name in list_of_sheets:
        source_ws = wb_merged[sheet_name]

        # --- LẤY DỮ LIỆU CHÍNH XÁC THEO CỘT VÀ DÒNG BÁC ĐÃ CHỈ ---
        # 1. Cột H (8): Tiết thực dạy (Cộng từ dòng 13 đến 78)
        t_day = sum(get_num(source_ws, r, 8) for r in range(13, 79))

        # 2. Số tiết kiêm nhiệm: Lấy từ Cột H (8), dòng 80 và 81
        t_kiem_nhiem = get_num(source_ws, 80, 8) + get_num(source_ws, 81, 8)

        # 3. Cột I (9): Đi công tác
        t_cong_tac = sum(get_num(source_ws, r, 9) for r in range(13, 79))

        # 4. Cột J (10): Dạy thay
        t_day_thay = sum(get_num(source_ws, r, 10) for r in range(13, 79))

        # 5. Cột K (11): Lấp giờ, tăng tiết
        t_tang_tiet = sum(get_num(source_ws, r, 11) for r in range(13, 79))

        # 6. Cột L (12): Coi KT, dự giờ
        t_coi_kt = sum(get_num(source_ws, r, 12) for r in range(13, 79))

        tong_cong = t_day + t_kiem_nhiem + t_cong_tac + t_day_thay + t_tang_tiet + t_coi_kt

        data_row = [
            tt,
            sheet_name,
            t_day,
            t_kiem_nhiem,
            t_cong_tac,
            t_day_thay,
            t_tang_tiet,
            t_coi_kt,
            tong_cong,
            ""
        ]

        for col, val in enumerate(data_row, 1):
            cell = ws_th.cell(row=row_idx, column=col)
            cell.value = val
            cell.font = th_font
            cell.alignment = center_aligned
            cell.border = thin_border

        row_idx += 1
        tt += 1


# --- TẠO THƯ MỤC LƯU TRỮ ---
SAVE_DIR = "Du_Lieu_Bao_Cao"
if not os.path.exists(SAVE_DIR):
    os.makedirs(SAVE_DIR)

# --- DANH SÁCH GIÁO VIÊN ---
DANH_SACH_GV = [
    "Nguyễn Văn Lộc", "Đỗ Văn Linh", "Huỳnh Thị Hạ Quyên", "Phạm Thị Mỹ Thuận",
    "Huỳnh Thị Huyên", "Đỗ Đặng Toàn", "Bùi Thị Xuân Nhựt", "Nguyễn Thị Như Ái",
    "Phạm Thị Lai Tình", "Trần Văn Hoàng", "Ngô Hữu Hoá", "Phạm Thuỵ Thuỳ Nghi",
    "Đỗ Thanh Vũ", "Phạm Bá Quyết", "Võ Quang Tuyên", "Nguyễn Văn Thân",
    "Nguyễn Minh Văn", "Lê Thị Tuyết Lệ", "Trần Đình Thảo", "Bùi Thanh Tâm",
    "Bùi Thị Bích Vân", "K Mah Ri Lan", "Nguyễn Mẫn Thu", "Lê Thị Kim Tuyết",
    "Lê Thị Tường Vy", "Nguyễn Thị Thúy Hằng", "Trần Thị Kim Anh", "Nguyễn Thị Hoan",
    "Đinh Thị Xuân Trâm"
]

# Sidebar Cấu Hình Chung
st.sidebar.header("Cài đặt thông số")
nam_hoc = st.sidebar.text_input("Năm học:", "2025 - 2026")
hoc_ky = st.sidebar.selectbox("Học kỳ:", ["HỌC KỲ I", "HỌC KỲ II"], index=1)

danh_sach_tuan = [f"Tuần {i}" for i in range(1, 36)]

tab1, tab2 = st.tabs(["📤 Khu vực Giáo viên nộp bài", "⚙️ Khu vực Quản lý tổng hợp"])

with tab1:
    st.header(f"Nộp báo cáo hàng tuần ({hoc_ky} - {nam_hoc})")

    tuan_nop = st.selectbox("Chọn tuần báo cáo:", danh_sach_tuan, index=24)
    ten_gv = st.selectbox("Chọn tên của bạn:", DANH_SACH_GV)
    uploaded_file = st.file_uploader("Chọn file Excel báo cáo (.xlsx) của bạn", type=['xlsx'])

    if st.button("📤 Nộp Báo Cáo"):
        if not uploaded_file:
            st.warning("⚠️ Vui lòng chọn file báo cáo!")
        else:
            ky_nam_dir = os.path.join(SAVE_DIR, nam_hoc.replace(" ", ""), hoc_ky.replace(" ", "_"))
            tuan_dir = os.path.join(ky_nam_dir, tuan_nop.replace(" ", "_"))

            if not os.path.exists(tuan_dir):
                os.makedirs(tuan_dir)

            file_path = os.path.join(tuan_dir, f"{ten_gv}.xlsx")
            with open(file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            st.success(
                f"🎉 Đã lưu thành công! Giáo viên **{ten_gv}** đã nộp báo cáo cho **{tuan_nop}**, {hoc_ky}, Năm học {nam_hoc}.")
            st.info(
                "💡 Nếu bạn nộp nhầm file, chỉ cần chọn lại tên và nộp file mới. Hệ thống sẽ tự động thay thế file cũ.")

with tab2:
    st.header("Danh sách và Tổng hợp")
    tuan_tong_hop = st.selectbox("Chọn tuần cần kiểm tra và tổng hợp:", danh_sach_tuan, index=24, key="tuan_th")

    ky_nam_dir_th = os.path.join(SAVE_DIR, nam_hoc.replace(" ", ""), hoc_ky.replace(" ", "_"))
    tuan_dir_th = os.path.join(ky_nam_dir_th, tuan_tong_hop.replace(" ", "_"))

    danh_sach_file = []
    if os.path.exists(tuan_dir_th):
        danh_sach_file = [f for f in os.listdir(tuan_dir_th) if f.endswith('.xlsx')]

    st.subheader(f"Danh sách đã nộp ({len(danh_sach_file)}/{len(DANH_SACH_GV)} giáo viên):")
    if len(danh_sach_file) > 0:
        for f_name in danh_sach_file:
            st.write(f"✅ {f_name.replace('.xlsx', '')}")

        st.write("---")
        if st.button(f"⚙️ Tiến hành tổng hợp {tuan_tong_hop}"):
            wb_merged = openpyxl.Workbook()
            for style in wb_merged._named_styles:
                if style.name == 'Normal':
                    style.font = Font(name='Times New Roman', size=11)
            wb_merged.remove(wb_merged.active)

            list_of_sheets = []

            with st.spinner('Đang xử lý, trích xuất dữ liệu và đồng bộ Font...'):
                for f_name in danh_sach_file:
                    file_path = os.path.join(tuan_dir_th, f_name)
                    try:
                        wb_source = openpyxl.load_workbook(file_path, data_only=True)
                        source_sheet = wb_source.active

                        sheet_name = f_name.replace('.xlsx', '')[:31]
                        target_sheet = wb_merged.create_sheet(title=sheet_name)
                        copy_sheet(source_sheet, target_sheet)
                        list_of_sheets.append(sheet_name)

                    except Exception as e:
                        st.error(f"❌ Lỗi ở file '{f_name}': {e}")

                if len(list_of_sheets) > 0:
                    create_summary_sheet(wb_merged, list_of_sheets, nam_hoc, hoc_ky, tuan_tong_hop)

            if len(wb_merged.sheetnames) > 0:
                output = io.BytesIO()
                wb_merged.save(output)
                output.seek(0)

                file_name_out = f"Tong_hop_BC_{nam_hoc.replace(' ', '')}_{hoc_ky.replace(' ', '_')}_{tuan_tong_hop.replace(' ', '_')}.xlsx"
                st.download_button(
                    label=f"📥 Tải file tổng hợp {tuan_tong_hop} về máy",
                    data=output,
                    file_name=file_name_out,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    else:
        st.info(f"Chưa có giáo viên nào nộp báo cáo trong {tuan_tong_hop} ({hoc_ky} - {nam_hoc}).")