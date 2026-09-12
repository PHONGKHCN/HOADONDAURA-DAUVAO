# -*- coding: utf-8 -*-
"""Đọc số tiền VNĐ thành chữ tiếng Việt."""

CHU_SO = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"]
DON_VI = ["", "nghìn", "triệu", "tỷ"]


def _doc_ba_so(so, day_du=False):
    """Đọc một nhóm 3 chữ số (000-999) thành chữ."""
    tram = so // 100
    chuc = (so % 100) // 10
    donvi = so % 10
    parts = []

    if tram > 0 or day_du:
        parts.append(CHU_SO[tram] + " trăm")
    elif tram == 0 and (chuc > 0 or donvi > 0) and day_du:
        parts.append("không trăm")

    if chuc == 0:
        if donvi > 0 and (tram > 0 or day_du):
            parts.append("lẻ")
    elif chuc == 1:
        parts.append("mười")
    else:
        parts.append(CHU_SO[chuc] + " mươi")

    if donvi > 0:
        if chuc >= 2 and donvi == 1:
            parts.append("mốt")
        elif chuc >= 1 and donvi == 5:
            parts.append("lăm")
        else:
            parts.append(CHU_SO[donvi])

    return " ".join(parts)


def so_thanh_chu(so_tien):
    """Chuyển số tiền (int/float, đơn vị đồng) thành chữ tiếng Việt, viết hoa chữ đầu."""
    so_tien = int(round(so_tien))
    if so_tien == 0:
        return "Không đồng"

    am = so_tien < 0
    so_tien = abs(so_tien)

    # Tách thành các nhóm 3 chữ số từ phải sang trái
    groups = []
    n = so_tien
    if n == 0:
        groups = [0]
    while n > 0:
        groups.append(n % 1000)
        n //= 1000
    groups.reverse()  # nhóm cao nhất trước

    num_groups = len(groups)
    words = []
    for idx, g in enumerate(groups):
        if g == 0:
            continue
        pos_from_right = num_groups - 1 - idx
        # tỷ đơn vị lặp lại theo bội số của 3 nhóm (tỷ, nghìn tỷ, triệu tỷ...)
        don_vi_idx = pos_from_right % 4 if pos_from_right < 4 else None
        if pos_from_right >= 4:
            # số quá lớn (nghìn tỷ trở lên) - lặp lại "tỷ" theo nhóm 4
            level = pos_from_right // 4
            suffix = " " + " ".join(["tỷ"] * level) if level else ""
            unit_label = DON_VI[pos_from_right % 4] + suffix
        else:
            unit_label = DON_VI[pos_from_right]

        day_du = idx > 0  # nhóm không phải đầu tiên thì đọc đủ "trăm/không trăm"
        chu = _doc_ba_so(g, day_du=day_du)
        if unit_label:
            words.append(chu + " " + unit_label)
        else:
            words.append(chu)

    result = " ".join(words).strip()
    result = " ".join(result.split())  # gộp khoảng trắng thừa
    result = result[0].upper() + result[1:] + " đồng"
    if am:
        result = "Âm " + result
    return result


if __name__ == "__main__":
    tests = [0, 5, 10, 15, 21, 100, 105, 110, 121, 1000, 1001, 1050,
             87050000, 2912000, 1000000000, 123456789]
    for t in tests:
        print(f"{t:>15,} -> {so_thanh_chu(t)}")
