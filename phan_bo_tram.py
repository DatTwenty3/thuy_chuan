"""Phân bổ số đo từng trạm theo tỷ lệ mẫu So do.xlsx."""

from __future__ import annotations

from dataclasses import dataclass

from doc_mau import MauTram, doc_mau_tu_excel

GIOI_HAN_OFFSET_H = 5.0
GIOI_HAN_J = 2.0


@dataclass
class DocTram:
    c_s: float
    c_t: float
    e_s: float
    e_t: float
    h_s: float
    i_s: float
    h_t: float
    i_t: float


@dataclass
class KetQuaTram:
    doc: DocTram
    kc_sau: float
    kc_truoc: float
    chenh_h: float
    chenh_i: float
    j_s: float
    j_t: float
    j_st: float
    chenh_cao: float
    chenh_kc: float
    sai_so_h_s: float
    sai_so_h_t: float


def _clamp(v: float, lo: float, hi: float) -> float:
    return max(lo, min(hi, v))


def _phan_bo_theo_ty_le(values: list[float], tong_moi: float) -> list[float]:
    if not values:
        return []
    s = sum(values)
    if s == 0:
        base = tong_moi / len(values)
        parts = [base] * len(values)
    else:
        parts = [v / s * tong_moi for v in values]
    parts[-1] = round(tong_moi - sum(parts[:-1]), 4)
    return parts


def _chon_mau(n: int, mau: list[MauTram]) -> list[MauTram]:
    if not mau:
        return []
    if n <= len(mau):
        return mau[:n]
    return [mau[i % len(mau)] for i in range(n)]


def _tao_h_i(
    c_s: float,
    c_t: float,
    e_s: float,
    e_t: float,
    mt: MauTram,
    tram: int,
    delta_k: float,
) -> tuple[float, float, float, float]:
    mid_c = (c_s + c_t) / 2
    mid_e = (e_s + e_t) / 2

    h_s = round(mid_c + _clamp(mt.offset_h_s, -GIOI_HAN_OFFSET_H, GIOI_HAN_OFFSET_H))
    h_t = round(mid_e + _clamp(mt.offset_h_t, -GIOI_HAN_OFFSET_H, GIOI_HAN_OFFSET_H))

    j_s = _clamp(mt.j_mau_s, -GIOI_HAN_J, GIOI_HAN_J)
    j_t = _clamp(mt.j_mau_t, -GIOI_HAN_J, GIOI_HAN_J)

    if tram == 2:
        i_s = mt.n_s + h_s - delta_k - j_s
        i_t = mt.n_t + h_t + delta_k - j_t
    else:
        i_s = h_s + mt.n_s - j_s
        i_t = h_t + mt.n_t - j_t

    return h_s, h_t, round(i_s), round(i_t)


def tinh_ket_qua_tram(
    doc: DocTram,
    n_s: float,
    n_t: float,
    delta_k: float,
    tram: int,
    chenh_cao_muc_tieu: float,
) -> KetQuaTram:
    kc_sau = (doc.c_s - doc.c_t) / 10
    kc_truoc = (doc.e_s - doc.e_t) / 10
    chenh_h = doc.h_s - doc.h_t
    chenh_i = doc.i_s - doc.i_t

    mid_c = (doc.c_s + doc.c_t) / 2
    mid_e = (doc.e_s + doc.e_t) / 2
    sai_so_h_s = doc.h_s - mid_c
    sai_so_h_t = doc.h_t - mid_e

    if tram == 2:
        j_s = n_s + doc.h_s - doc.i_s - delta_k
        j_t = n_t + doc.h_t - doc.i_t + delta_k
        j_st = j_s - j_t
        chenh_cao = chenh_h - j_st
    else:
        j_s = n_s + doc.h_s - doc.i_s
        j_t = n_t + doc.h_t - doc.i_t
        j_st = j_s - j_t
        chenh_cao = chenh_cao_muc_tieu

    return KetQuaTram(
        doc=doc,
        kc_sau=kc_sau,
        kc_truoc=kc_truoc,
        chenh_h=chenh_h,
        chenh_i=chenh_i,
        j_s=round(j_s),
        j_t=round(j_t),
        j_st=round(j_st),
        chenh_cao=round(chenh_cao),
        chenh_kc=round(kc_sau - kc_truoc, 4),
        sai_so_h_s=round(sai_so_h_s, 2),
        sai_so_h_t=round(sai_so_h_t, 2),
    )


def phan_bo_tram(
    n: int,
    k1: float,
    k2: float,
    l_sau: float,
    l_truoc: float,
    h_tong: float,
    i_tong: float,
    delta_k: float = 100.0,
    mau: list[MauTram] | None = None,
) -> tuple[list[DocTram], list[MauTram], list[float]]:
    mau_full = mau or doc_mau_tu_excel(k1=k1, k2=k2)
    mau_sel = _chon_mau(n, mau_full)

    kc_sau_list = _phan_bo_theo_ty_le([t.kc_sau for t in mau_sel], l_sau)
    kc_truoc_list = _phan_bo_theo_ty_le([t.kc_truoc for t in mau_sel], l_truoc)
    chenh_h_list = _phan_bo_theo_ty_le([t.chenh_h for t in mau_sel], h_tong)
    chenh_cao_list = _phan_bo_theo_ty_le([t.chenh_cao for t in mau_sel], h_tong)

    doc_list: list[DocTram] = []

    for i, mt in enumerate(mau_sel):
        tram = i + 1
        c_s = mt.c_s_base
        c_t = c_s - kc_sau_list[i] * 10
        e_s = mt.e_s_base
        e_t = e_s - kc_truoc_list[i] * 10

        h_s, h_t, i_s, i_t = _tao_h_i(c_s, c_t, e_s, e_t, mt, tram, delta_k)

        doc_list.append(
            DocTram(
                c_s=round(c_s, 1),
                c_t=round(c_t, 1),
                e_s=round(e_s, 1),
                e_t=round(e_t, 1),
                h_s=h_s,
                i_s=i_s,
                h_t=h_t,
                i_t=i_t,
            )
        )

    return doc_list, mau_sel, chenh_cao_list


def tong_kiem_tra(
    tram_list: list[DocTram],
    mau_sel: list[MauTram],
    chenh_cao_list: list[float],
    delta_k: float = 100,
) -> dict[str, float]:
    tong_kc_sau = 0.0
    tong_kc_truoc = 0.0
    tong_chenh_h = 0.0
    tong_chenh_i = 0.0

    for i, doc in enumerate(tram_list, start=1):
        mt = mau_sel[i - 1]
        kq = tinh_ket_qua_tram(doc, mt.n_s, mt.n_t, delta_k, i, chenh_cao_list[i - 1])
        tong_kc_sau += kq.kc_sau
        tong_kc_truoc += kq.kc_truoc
        tong_chenh_h += kq.chenh_h
        tong_chenh_i += kq.chenh_i

    return {
        "l_sau": round(tong_kc_sau, 4),
        "l_truoc": round(tong_kc_truoc, 4),
        "h_tong": round(tong_chenh_h, 4),
        "i_tong": round(tong_chenh_i, 4),
    }


def kiem_tra_hang4(
    tram_list: list[DocTram],
    mau_sel: list[MauTram],
    chenh_cao_list: list[float],
    delta_k: float = 100,
) -> list[dict]:
    ket_qua = []
    for i, doc in enumerate(tram_list, start=1):
        mt = mau_sel[i - 1]
        kq = tinh_ket_qua_tram(doc, mt.n_s, mt.n_t, delta_k, i, chenh_cao_list[i - 1])
        ket_qua.append(
            {
                "tram": i,
                "dat_kc_5m": abs(kq.kc_sau - kq.kc_truoc) <= 5,
                "dat_h_s_5mm": abs(kq.sai_so_h_s) <= GIOI_HAN_OFFSET_H,
                "dat_h_t_5mm": abs(kq.sai_so_h_t) <= GIOI_HAN_OFFSET_H,
                "dat_j_s": abs(kq.j_s) <= GIOI_HAN_J,
                "dat_j_t": abs(kq.j_t) <= GIOI_HAN_J,
                "sai_h_s (mm)": kq.sai_so_h_s,
                "sai_h_t (mm)": kq.sai_so_h_t,
                "j_s": kq.j_s,
                "j_t": kq.j_t,
                "chenh_cao": kq.chenh_cao,
            }
        )
    return ket_qua
