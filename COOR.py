import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document
from datetime import datetime, timezone, timedelta
from tab_bao_cao_su_vu import tab_bao_cao_su_vu
import re
import json
import os

st.set_page_config(page_title="COOR TOOL VJ DAD", layout="wide")

# Lấy giờ Việt Nam
now_vn = datetime.now(timezone(timedelta(hours=7)))

tab1, tab2 = st.tabs(["✈️ Kế hoạch Kéo tàu", "📋 Báo cáo Sự cố CAAV"])

with tab1:
    st.title("✈️ Trình tạo mail kéo tàu Vietjet DAD")
    st.caption(f"Ngày tạo báo cáo: {now_vn.strftime('%d/%m/%Y')}")
    # ---HÀM JSON---
    DATA_FILE = "plans_data.json"

    def save_plans(plans):
        with open(DATA_FILE, "w", encoding="utf-8") as f:
            json.dump(plans, f, ensure_ascii=False, indent=2)

    def load_plans():
        if os.path.exists(DATA_FILE):
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        return []
    # --- KHỞI TẠO SESSION STATE ---
    if 'plans' not in st.session_state:
        st.session_state.plans = load_plans()
    if 'editing_index' not in st.session_state:
        st.session_state.editing_index = None

    # --- HÀM HỖ TRỢ ---
    def create_word_document(content):
        doc = Document()
        # Loại bỏ các tag HTML highlight và ký hiệu == nếu có trước khi xuất Word
        clean_content = re.sub(r'<span.*?>|<\/span>', '', content)
        clean_content = clean_content.replace('==', '')
        bold_pattern = re.compile(r'\*\*(.*?)\*\*')
        for line in clean_content.split('\n'):
            p = doc.add_paragraph()
            p.style.font.name = 'Arial'
            parts = re.split(r'(\*\*.*?\*\*)', line)
            for part in parts:
                if bold_pattern.match(part):
                    p.add_run(part.strip('*')).bold = True
                else:
                    p.add_run(part)
        bio = BytesIO()
        doc.save(bio)
        bio.seek(0)
        return bio.getvalue()

    def generate_report_content(plans, highlight=False, kinh_gui=""):
        today_str = now_vn.strftime('%d/%m/%Y')
        
        # Header 
        report_lines = []
        # Thêm tiêu đề mail
        report_lines.append(f"**KẾ HOẠCH KÉO TÀU VJ DAD NGÀY {today_str}**")
        report_lines.append("")  # Dòng trống sau tiêu đề
        if kinh_gui:
            report_lines.extend(kinh_gui.split('\n'))
        
        report_lines.append("") # Dòng trống sau Kính gửi
        report_lines.append(f"VJ DAD gửi kế hoạch kéo/đẩy tàu bay ngày {today_str} như sau:")
        report_lines.append("")

        # Body
        for i, plan in enumerate(plans):
            changed = plan.get('changed_fields', [])
            
            # Tiêu đề mục (Ưu tiên hiển thị Đang bãi)
            tau_str = f"VN-{plan['Tàu']}"
            if 'Tàu' in changed and highlight: tau_str = f"=={tau_str}=="
            
            if plan.get('Đang bãi'):
                db_str = f"Đang bãi {plan['Đang bãi']}"
                if 'Đang bãi' in changed and highlight: db_str = f"=={db_str}=="
                title = f"**{i+1}. {tau_str}/ {db_str}**"
            else:
                ch_str = plan['Chuyến']
                if 'Chuyến' in changed and highlight: ch_str = f"=={ch_str}=="
                title = f"**{i+1}. {tau_str}/{ch_str}**"
                if plan['STA']:
                    sta_str = f"STA {plan['STA']}"
                    if 'STA' in changed and highlight: sta_str = f"=={sta_str}=="
                    title += f" **{sta_str}**"
            
            if plan['Ghi chú']:
                gc_upper = str(plan['Ghi chú']).upper()
                if "CNX" in gc_upper:
                    gc_display = "HUỶ KẾ HOẠCH KÉO"
                elif "DONE" in gc_upper:
                    gc_display = "ĐÃ HOÀN THÀNH"
                else:
                    gc_display = plan['Ghi chú']
                gc_str = f"({gc_display})"
                if 'Ghi chú' in changed and highlight: gc_str = f"=={gc_str}=="
                title += f" **{gc_str}**"
            report_lines.append(title)
            # Kiểm tra điều kiện ẩn chi tiết
            gc_upper = str(plan['Ghi chú']).upper()
            if "CNX" in gc_upper or "DONE" in gc_upper:
                report_lines.append("")
                continue

            # Kéo tới (Luôn hiện nếu có thông tin)
            if plan['Kéo tới']:
                kt_val = plan['Kéo tới']
                if 'Kéo tới' in changed and highlight: kt_val = f"=={kt_val}=="
                line = f"    - Kéo về **{kt_val}**"
                
                if plan['Thời gian kéo']:
                    tgk_val = plan['Thời gian kéo']
                    if 'Thời gian kéo' in changed and highlight: tgk_val = f"=={tgk_val}=="
                    line += f" dự kiến vào lúc: **{tgk_val}**"
                report_lines.append(line)

            if plan['Kéo ga lớn'] == "CÓ":
                line = f"    - Kéo ra ga lớn: **CÓ**"
                if plan['Thời gian kéo ga lớn']:
                    tggl_val = plan['Thời gian kéo ga lớn']
                    if 'Thời gian kéo ga lớn' in changed and highlight: tggl_val = f"=={tggl_val}=="
                    line += f". Thời gian dự kiến: **{tggl_val}**"
                report_lines.append(line)
            
            if plan['Kéo khai thác'] == "CÓ":
                if not plan['Thời gian kéo khai thác'] or plan['Thời gian kéo khai thác'] == "THÔNG BÁO SAU":
                    line = "    - Kéo ra bãi khai thác: **CÓ**. Thời gian dự kiến: **THÔNG BÁO SAU**"
                else:
                    ktch_val = plan['Khai thác chuyến']
                    if 'Khai thác chuyến' in changed and highlight: ktch_val = f"=={ktch_val}=="
                    tgkt_val = plan['Thời gian kéo khai thác']
                    if 'Thời gian kéo khai thác' in changed and highlight: tgkt_val = f"=={tgkt_val}=="
                    line = f"    - Kéo ra bãi khai thác chuyến: **{ktch_val}**. Thời gian dự kiến: **{tgkt_val}**"
                report_lines.append(line)

            dv_val = plan['Đơn vị kéo']
            if 'Đơn vị kéo' in changed and highlight: dv_val = f"=={dv_val}=="
            report_lines.append(f"    - Đơn vị kéo đẩy: **{dv_val}**")
            
            asu_val = plan['ASU-GPU']
            if 'ASU-GPU' in changed and highlight: asu_val = f"=={asu_val}=="
            report_lines.append(f"    - Cần ASU, GPU: **{asu_val}**")
            report_lines.append("") # Dòng trống giữa các mục

        # Footer cố định
        report_lines.append("Kính mong TBT, ĐHSĐ sắp xếp tàu về bến thuận tiện cho việc kéo đẩy.")
        report_lines.append("VJ sẽ cập nhật thông tin khi kế hoạch thay đổi.")
        
        return "\n".join(report_lines)

    def convert_to_html(markdown_text):
        """Chuyển đổi Markdown sang HTML để hỗ trợ copy paste bôi đen giữ định dạng."""
        html = markdown_text.replace('\n', '<br>')
        # Xử lý highlight ==text== -> <span style="background-color: yellow">text</span>
        html = re.sub(r'==(.*?)==', r'<span style="background-color: #FFFF00; color: black; padding: 0 2px; border-radius: 2px;">\1</span>', html)
        # Xử lý in đậm **text** -> <b>text</b>
        html = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', html)
        # Xử lý thụt lề (4 dấu cách) -> &nbsp;
        html = html.replace('    ', '&nbsp;&nbsp;&nbsp;&nbsp;')
        return f'<div style="font-family: Arial; font-size: 14px; line-height: 1.5; color: black;">{html}</div>'

    # --- SIDEBAR - Cấu hình ---
    st.sidebar.header("⚙️ Cấu hình Mail mẫu")
    default_kinh_gui = """Kính Gửi:
        -Trực Ban Trưởng
        -Điều Hành Sân Đỗ
        -Đài kiểm soát mặt đất"""

    kinh_gui_input = st.sidebar.text_area(
        "Danh sách Kính gửi (Edit tại đây):", 
        value=default_kinh_gui,
        height=150
    )

    st.sidebar.divider()
    st.sidebar.info("Footer đã được cố định theo mẫu.")

    # --- FORM NHẬP LIỆU ---
    st.subheader("📋 1. Nhập Kế hoạch Kéo tàu")

    # Lấy dữ liệu nếu đang ở chế độ chỉnh sửa
    edit_idx = st.session_state.editing_index
    edit_data = st.session_state.plans[edit_idx] if edit_idx is not None else {}

    with st.form("plan_form", clear_on_submit=True):
        if edit_idx is not None:
            st.warning(f"🔄 Đang chỉnh sửa kế hoạch số {edit_idx + 1}")
        
        st.write("Điền thông tin cho một tàu bay và nhấn nút Thêm/Cập nhật.")
        c1, c2, c3, c4, c5 = st.columns(5)
        tau = c1.text_input("Tàu (VN-)", value=edit_data.get("Tàu", ""), placeholder="A662")
        chuyen = c2.text_input("Chuyến", value=edit_data.get("Chuyến", ""), placeholder="VJ703")
        sta = c3.text_input("STA / Ghi chú", value=edit_data.get("STA", ""), placeholder="12:30 hoặc PHASE CHECK")
        dang_bai = c4.text_input("Đang bãi", value=edit_data.get("Đang bãi", ""), placeholder="VJ01 hoặc 3M")
        ghi_chu = c5.text_input("Ghi chú thêm (nếu có)", value=edit_data.get("Ghi chú", ""))

        st.write("Kéo về bãi:")
        c1, c2 = st.columns([1, 2])
        keo_toi = c1.text_input("Kéo về bãi", value=edit_data.get("Kéo tới", ""), placeholder="VJ01")
        tg_keo = c2.text_input("Thời gian kéo về bãi", value=edit_data.get("Thời gian kéo", ""), placeholder="11:00")

        st.write("Kéo ra ga lớn & khai thác:")
        c1, c2, c3, c4, c5 = st.columns(5)
        
        kg_idx = ["KHÔNG", "CÓ"].index(edit_data.get("Kéo ga lớn", "KHÔNG"))
        keo_ga_lon = c1.selectbox("Kéo ga lớn?", ["KHÔNG", "CÓ"], index=kg_idx)
        tg_ga_lon = c2.text_input("Giờ kéo ga lớn", value=edit_data.get("Thời gian kéo ga lớn", "THÔNG BÁO SAU"))
        
        kk_idx = ["CÓ", "KHÔNG"].index(edit_data.get("Kéo khai thác", "CÓ"))
        keo_kt = c3.selectbox("Kéo khai thác?", ["CÓ", "KHÔNG"], index=kk_idx)
        kt_chuyen = c4.text_input("Chuyến khai thác", value=edit_data.get("Khai thác chuyến", ""), placeholder="VJ703")
        tg_kt = c5.text_input("Giờ kéo khai thác", value=edit_data.get("Thời gian kéo khai thác", "THÔNG BÁO SAU"))

        st.write("Thông tin khác:")
        c1, c2 = st.columns(2)
        don_vi = c1.text_input("Đơn vị kéo", value=edit_data.get("Đơn vị kéo", "VJ"))
        
        asu_idx = ["KHÔNG", "CÓ"].index(edit_data.get("ASU-GPU", "KHÔNG"))
        asu_gpu = c2.selectbox("Cần ASU/GPU?", ["KHÔNG", "CÓ"], index=asu_idx)

        btn_label = "➕ Thêm vào danh sách" if edit_idx is None else "💾 Cập nhật kế hoạch"
        submitted = st.form_submit_button(btn_label, use_container_width=True)
        
        if submitted and tau:
            new_plan = {
                "Tàu": tau, "Chuyến": chuyen, "STA": sta, "Đang bãi": dang_bai, "Ghi chú": ghi_chu,
                "Kéo tới": keo_toi, "Thời gian kéo": tg_keo,
                "Kéo ga lớn": keo_ga_lon, "Thời gian kéo ga lớn": tg_ga_lon,
                "Kéo khai thác": keo_kt, "Khai thác chuyến": kt_chuyen, "Thời gian kéo khai thác": tg_kt,
                "Đơn vị kéo": don_vi, "ASU-GPU": asu_gpu
            }
            
            if edit_idx is not None:
                # So sánh để tìm các trường thay đổi
                old_plan = st.session_state.plans[edit_idx]
                changed_fields = []
                for k, v in new_plan.items():
                    if v != old_plan.get(k):
                        changed_fields.append(k)
                
                new_plan['changed_fields'] = changed_fields
                st.session_state.plans[edit_idx] = new_plan
                save_plans(st.session_state.plans)
                st.session_state.editing_index = None
                st.success("Đã cập nhật kế hoạch!")
            else:
                new_plan['changed_fields'] = []
                st.session_state.plans.append(new_plan)
                save_plans(st.session_state.plans)
                st.success("Đã thêm kế hoạch mới!")
            st.rerun()

    if st.session_state.editing_index is not None:
        if st.button("❌ Hủy chỉnh sửa", use_container_width=True):
            st.session_state.editing_index = None
            st.rerun()

    # --- DANH SÁCH KẾ HOẠCH ĐÃ THÊM ---
    st.subheader("📝 2. Danh sách Kế hoạch")
    if not st.session_state.plans:
        st.info("Chưa có kế hoạch nào được thêm.")
    else:
        for i, plan in enumerate(st.session_state.plans):
            with st.container(border=True):
                c1, c2, c3, c4, c5 = st.columns([6, 1, 1, 1, 1])
                
                # Hiển thị thông tin tiêu đề trong danh sách
                if plan.get('Đang bãi'):
                    info = f"**{i+1}. VN-{plan['Tàu']}/ Đang bãi {plan['Đang bãi']}**"
                else:
                    info = f"**{i+1}. VN-{plan['Tàu']}/{plan['Chuyến']}** - STA: {plan['STA']}"
                
                if plan['Ghi chú']: info += f" ({plan['Ghi chú']})"
                
                details = []
                if plan['Kéo tới']: 
                    details.append(f"Kéo về: {plan['Kéo tới']} ({plan['Thời gian kéo']})")
                if plan['Kéo ga lớn'] == "CÓ": details.append(f"Ga lớn: CÓ ({plan['Thời gian kéo ga lớn']})")
                if plan['Kéo khai thác'] == "CÓ": 
                    if not plan['Thời gian kéo khai thác'] or plan['Thời gian kéo khai thác'] == "THÔNG BÁO SAU":
                        details.append("Khai thác: CÓ (THÔNG BÁO SAU)")
                    else: details.append(f"Khai thác: {plan['Khai thác chuyến']} ({plan['Thời gian kéo khai thác']})")
                
                c1.markdown(info)
                if details:
                    c1.caption(" | ".join(details))
                
                # Nút di chuyển lên
                if c2.button("🔼", key=f"up_{i}", help="Di chuyển lên"):
                    if i > 0:
                        st.session_state.plans[i], st.session_state.plans[i-1] = st.session_state.plans[i-1], st.session_state.plans[i]
                        save_plans(st.session_state.plans)
                        st.rerun()
                
                # Nút di chuyển xuống
                if c3.button("🔽", key=f"down_{i}", help="Di chuyển xuống"):
                    if i < len(st.session_state.plans) - 1:
                        st.session_state.plans[i], st.session_state.plans[i+1] = st.session_state.plans[i+1], st.session_state.plans[i]
                        st.rerun()

                # Nút sửa
                if c4.button("📝", key=f"edit_{i}", help="Chỉnh sửa"):
                    st.session_state.editing_index = i
                    st.rerun()
                
                # Nút xóa
                if c5.button("❌", key=f"del_{i}", help="Xóa"):
                    st.session_state.plans.pop(i)
                    save_plans(st.session_state.plans)
                    if st.session_state.editing_index == i:
                        st.session_state.editing_index = None
                    st.rerun()

    # --- TẠO MAIL ---
    st.divider()
    col_btn1, col_btn2 = st.columns(2)
    show_highlight = col_btn1.toggle("Highlight thông tin thay đổi", value=True)

    if st.button("🚀 3. Tạo Mail mẫu", use_container_width=True, type="primary"):
        if not st.session_state.plans:
            st.warning("Vui lòng thêm ít nhất một kế hoạch trước khi tạo mail.")
        else:
            st.subheader("📧 4. Kết quả Mail mẫu")
            report_content = generate_report_content(
                st.session_state.plans, 
                highlight=show_highlight, 
                kinh_gui=kinh_gui_input
            )
            
            # Hiển thị HTML để người dùng có thể bôi đen và copy paste giữ định dạng
            st.caption("👇 Bạn có thể bôi đen (Highlight) nội dung dưới đây để Copy & Paste vào Mail/Zalo (giữ nguyên định dạng in đậm):")
            st.write(convert_to_html(report_content), unsafe_allow_html=True)
            
            st.divider()
            st.caption("Hoặc copy text thô tại đây (Không bao gồm Highlight):")
            # Text thô không nên có highlight markdown ==text==
            raw_content = re.sub(r'==(.*?)==', r'\1', report_content)
            st.code(raw_content, language="text")
            
            st.download_button(
                label="📄 Tải xuống file Word (.docx)",
                data=create_word_document(report_content),
                file_name=f"Ke_hoach_keo_tau_{now_vn.strftime('%d%m%Y')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
with tab2:
    tab_bao_cao_su_vu()
