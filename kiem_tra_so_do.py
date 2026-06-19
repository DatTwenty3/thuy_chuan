"""Kiểm tra tổng hợp sổ đo thủy chuẩn hạng 4."""

from __future__ import annotations

import math
import sys
from dataclasses import dataclass

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stdin.reconfigure(encoding="utf-8")


DEFAULTS = {
    "so_tram": 19,
    "k1": 4575,
    "k2": 4475,
    "delta_k": 100,
    "l_sau": 1794.9,
    "l_truoc": 1798.1,
    "h_tong": 537,
    "i_tong": 638,
}


@dataclass
class DauVao:
    so_tram: int
    k1: float
    k2: float
    delta_k: float
    l_sau: float
    l_truoc: float
    h_tong: float
    i_tong: float


@dataclass
class KetQua:
    delta_l: float
    l_km: float
    mp: float
    delta_hi: float
    delta_k_ky_vong: float
    gioi_han_chenh_kc: float
    dat_chenh_kc: bool
    dat_khep_vong: bool
    dat_kiem_tra_k: bool


def tinh_ket_qua(data: DauVao) -> KetQua:
    delta_l = data.l_sau - data.l_truoc
    l_km = (data.l_sau + data.l_truoc) / 2 / 1000
    mp = 20 * math.sqrt(l_km)
    delta_hi = data.i_tong - data.h_tong
    delta_k_ky_vong = data.delta_k * (data.so_tram / 2)
    gioi_han_chenh_kc = 0.5 * data.so_tram

    lech_k_luan_phien = abs(delta_hi - delta_k_ky_vong) <= data.delta_k
    lech_k_tong_hop = abs(delta_hi - data.delta_k) <= data.delta_k

    return KetQua(
        delta_l=delta_l,
        l_km=l_km,
        mp=mp,
        delta_hi=delta_hi,
        delta_k_ky_vong=delta_k_ky_vong,
        gioi_han_chenh_kc=gioi_han_chenh_kc,
        dat_chenh_kc=abs(delta_l) <= gioi_han_chenh_kc,
        dat_khep_vong=abs(data.h_tong) <= mp,
        dat_kiem_tra_k=lech_k_luan_phien or lech_k_tong_hop,
    )


def _nhap_so(prompt: str, default: float, kieu: type = float):
    raw = input(f"{prompt} [{default}]: ").strip()
    if not raw:
        return default
    return kieu(raw)


def nhap_lieu() -> DauVao:
    print("Nhập thông số (Enter = dùng giá trị mặc định):\n")
    return DauVao(
        so_tram=_nhap_so("Số trạm đo", DEFAULTS["so_tram"], int),
        k1=_nhap_so("Hằng số K1 (mm)", DEFAULTS["k1"]),
        k2=_nhap_so("Hằng số K2 (mm)", DEFAULTS["k2"]),
        delta_k=_nhap_so("Chênh hằng số ΔK (mm)", DEFAULTS["delta_k"]),
        l_sau=_nhap_so("Tổng khoảng cách nhìn sau L_sau (m)", DEFAULTS["l_sau"]),
        l_truoc=_nhap_so("Tổng khoảng cách nhìn trước L_truoc (m)", DEFAULTS["l_truoc"]),
        h_tong=_nhap_so("Tổng chênh mặt đen H (mm)", DEFAULTS["h_tong"]),
        i_tong=_nhap_so("Tổng chênh mặt đỏ I (mm)", DEFAULTS["i_tong"]),
    )


def _trang_thai(dat: bool) -> str:
    return "ĐẠT" if dat else "KHÔNG ĐẠT"


def in_bao_cao(data: DauVao, kq: KetQua) -> None:
    print("\n=== KẾT QUẢ KIỂM TRA TỔNG HỢP ===")
    print(f"Số trạm đo          : {data.so_tram}")
    print(f"Hằng số K1          : {data.k1:g} mm")
    print(f"Hằng số K2          : {data.k2:g} mm")
    print(f"Chênh hằng số ΔK     : {data.delta_k:g} mm")
    print(f"Tổng nhìn sau       : {data.l_sau:g} m")
    print(f"Tổng nhìn trước     : {data.l_truoc:g} m")
    print(f"Chênh khoảng cách   : {kq.delta_l:.1f} m")
    print(f"Chiều dài tuyến L   : {kq.l_km:.3f} km")
    print(f"Dung sai hạng 4     : ±{kq.mp:.1f} mm")
    print(f"Tổng chênh H (đen)  : {data.h_tong:g} mm")
    print(f"Tổng chênh I (đỏ)   : {data.i_tong:g} mm")
    print(f"Chênh I - H         : {kq.delta_hi:g} mm")
    print(f"ΔK kỳ vọng (luân phiên): {kq.delta_k_ky_vong:g} mm")
    print()
    print("--- Đánh giá ---")
    print(
        f"Chênh khoảng cách (|ΔL| ≤ {kq.gioi_han_chenh_kc:g} m): "
        f"{_trang_thai(kq.dat_chenh_kc)}"
    )
    print(
        f"Khép vòng cao độ (|H| ≤ {kq.mp:.1f} mm): "
        f"{_trang_thai(kq.dat_khep_vong)}"
    )
    print(
        f"Kiểm tra K (|I-H - ΔK×n/2| ≤ ΔK): "
        f"{_trang_thai(kq.dat_kiem_tra_k)}"
    )
    tong_dat = kq.dat_chenh_kc and kq.dat_khep_vong and kq.dat_kiem_tra_k
    print(f"\nKết luận chung      : {_trang_thai(tong_dat)}")


def main() -> None:
    print("=== KIỂM TRA TỔNG HỢP SỔ ĐO THỦY CHUẨN HẠNG 4 ===\n")
    while True:
        data = nhap_lieu()
        in_bao_cao(data, tinh_ket_qua(data))
        tiep = input("\nNhập tiếp? (y/n) [n]: ").strip().lower()
        if tiep not in ("y", "yes", "c", "co"):
            break
        print()


if __name__ == "__main__":
    main()
