"""Ứng dụng Streamlit xuất sổ đo thủy chuẩn hạng 4."""

from __future__ import annotations

import pandas as pd
import streamlit as st

from excel_builder import DuAn, TongHop, build_workbook, workbook_to_bytes
from kiem_tra_so_do import DauVao, tinh_ket_qua
from phan_bo_tram import (
    kiem_tra_hang4,
    phan_bo_tram,
    tinh_ket_qua_tram,
    tong_kiem_tra,
)

st.set_page_config(page_title="Sổ đo thủy chuẩn", page_icon="📐", layout="wide")

st.title("Sổ đo thủy chuẩn hạng 4")
st.caption("Phân bổ số đo theo tỷ lệ mẫu So do.xlsx — co giãn theo tổng hợp dự án.")

col_left, col_right = st.columns(2)

with col_left:
    st.subheader("Thông tin dự án")
    do_tu = st.text_input("Đo từ", value="(tên mốc mua)")
    den = st.text_input("Đến", value="(mốc GPSIV- gần nhất)")
    bat_dau = st.text_input("Bắt đầu lúc", value="")
    ket_thuc = st.text_input("Kết thúc lúc", value="")
    ngay_do = st.text_input("Ngày đo", value="(theo nhật ký)")
    thoi_tiet = st.text_input("Thời tiết", value="")
    nguoi_do = st.text_input("Người đo", value="")
    nguoi_ghi = st.text_input("Người ghi", value="")
    nguoi_kiem_tra = st.text_input("Người kiểm tra", value="")

with col_right:
    st.subheader("Thông số kỹ thuật")
    n_tram = st.number_input("Số trạm đo", min_value=1, max_value=50, value=19, step=1)
    k1 = st.number_input("Hằng số K1 (mm)", value=4575.0, step=1.0)
    k2 = st.number_input("Hằng số K2 (mm)", value=4475.0, step=1.0)
    delta_k = st.number_input("Chênh hằng số ΔK (mm)", value=100.0, step=1.0)
    st.divider()
    st.markdown("**Số tổng hợp**")
    l_sau = st.number_input("Tổng khoảng cách nhìn sau L_sau (m)", value=1794.9, format="%.1f")
    l_truoc = st.number_input("Tổng khoảng cách nhìn trước L_truoc (m)", value=1798.1, format="%.1f")
    h_tong = st.number_input("Tổng chênh mặt đen H (mm)", value=537.0, step=1.0)
    i_tong = st.number_input("Tổng chênh mặt đỏ I (mm)", value=638.0, step=1.0)

du_an = DuAn(
    do_tu=do_tu,
    den=den,
    bat_dau=bat_dau,
    ket_thuc=ket_thuc,
    ngay_do=ngay_do,
    thoi_tiet=thoi_tiet,
    nguoi_do=nguoi_do,
    nguoi_ghi=nguoi_ghi,
    nguoi_kiem_tra=nguoi_kiem_tra,
)
tong_hop = TongHop(l_sau=l_sau, l_truoc=l_truoc, h_tong=h_tong, i_tong=i_tong)

tram_list, mau_sel, chenh_cao_list = phan_bo_tram(
    int(n_tram), k1, k2, l_sau, l_truoc, h_tong, i_tong, delta_k
)
kiem_tra = tong_kiem_tra(tram_list, mau_sel, chenh_cao_list, delta_k)
hang4 = kiem_tra_hang4(tram_list, mau_sel, chenh_cao_list, delta_k)

dau_vao = DauVao(
    so_tram=int(n_tram),
    k1=k1,
    k2=k2,
    delta_k=delta_k,
    l_sau=l_sau,
    l_truoc=l_truoc,
    h_tong=h_tong,
    i_tong=i_tong,
)
kq = tinh_ket_qua(dau_vao)

st.divider()
st.subheader("Số đo đã phân bổ từng trạm")

preview_rows = []
for i, doc in enumerate(tram_list, start=1):
    mt = mau_sel[i - 1]
    kq_tram = tinh_ket_qua_tram(
        doc, mt.n_s, mt.n_t, delta_k, i, chenh_cao_list[i - 1]
    )
    preview_rows.append(
        {
            "Trạm": i,
            "KC sau (m)": round(kq_tram.kc_sau, 2),
            "KC trước (m)": round(kq_tram.kc_truoc, 2),
            "Chênh H (mm)": kq_tram.chenh_h,
            "Chênh cao K (mm)": kq_tram.chenh_cao,
            "|H-mid| sau (mm)": kq_tram.sai_so_h_s,
            "|H-mid| trước (mm)": kq_tram.sai_so_h_t,
            "J sau": kq_tram.j_s,
            "J trước": kq_tram.j_t,
        }
    )

st.dataframe(pd.DataFrame(preview_rows), use_container_width=True, hide_index=True)

st.caption(
    f"Tổng tính lại: L_sau={kiem_tra['l_sau']} m, L_truoc={kiem_tra['l_truoc']} m, "
    f"H={kiem_tra['h_tong']} mm"
)

st.subheader("Kiểm tra hạng IV (từng trạm)")
st.dataframe(pd.DataFrame(hang4), use_container_width=True, hide_index=True)

vi_pham = [
    h for h in hang4
    if not (h["dat_h_s_5mm"] and h["dat_h_t_5mm"] and h["dat_j_s"] and h["dat_j_t"])
]
if vi_pham:
    st.error(f"Có {len(vi_pham)} trạm vi phạm: |H−mid| > 5 mm hoặc |J| > 2 mm.")
else:
    st.success("Tất cả trạm đạt: |H−mid stadia| ≤ 5 mm, |J| ≤ 2 mm.")

st.subheader("Xem trước kiểm tra tổng hợp")

c1, c2, c3, c4 = st.columns(4)
c1.metric("Chênh khoảng cách", f"{kq.delta_l:.1f} m")
c2.metric("Chiều dài tuyến L", f"{kq.l_km:.3f} km")
c3.metric("Dung sai hạng 4", f"±{kq.mp:.1f} mm")
c4.metric("Chênh I − H", f"{kq.delta_hi:g} mm")

st.markdown(
    f"""
| Tiêu chí | Kết quả |
|----------|---------|
| Chênh khoảng cách (\\|ΔL\\| ≤ {kq.gioi_han_chenh_kc:g} m) | **{"ĐẠT" if kq.dat_chenh_kc else "KHÔNG ĐẠT"}** |
| Khép vòng cao độ (\\|H\\| ≤ {kq.mp:.1f} mm) | **{"ĐẠT" if kq.dat_khep_vong else "KHÔNG ĐẠT"}** |
| Kiểm tra K | **{"ĐẠT" if kq.dat_kiem_tra_k else "KHÔNG ĐẠT"}** |
"""
)

wb = build_workbook(
    du_an, int(n_tram), k1, k2, delta_k, tong_hop, tram_list, mau_sel, chenh_cao_list
)
excel_bytes = workbook_to_bytes(wb)

ten_file = f"So_do_{do_tu.replace(' ', '_')[:30] or 'du_an'}.xlsx"
st.download_button(
    label="Tải file Excel",
    data=excel_bytes,
    file_name=ten_file,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.info(
    f"File Excel gồm **{int(n_tram)}** trạm — số đo phân bổ theo tỷ lệ mẫu "
    f"(trạm 1: K≈−11, trạm 2: K≈48 mm với tổng mặc định)."
)
