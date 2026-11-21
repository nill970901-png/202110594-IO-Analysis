"""산업연관분석 계산 모듈 (브라우저 Pyodide 및 일반 Python 공용).

모든 금액 단위는 백만원(생산자가격, 2023년), 계수는 단위 없음.
분석 수준(level): basic(기본부문 380) / small(소분류 165) / medium(중분류 83) / large(대분류 33)
"""
import json, io
import numpy as np
import openpyxl

LEVEL_LABEL = {"basic": "기본부문", "small": "소분류", "medium": "중분류", "large": "대분류"}
CODE_LEN = {"basic": 4, "small": 3, "medium": 2}
_RANGES = [("01", 1, 5), ("02", 6, 7), ("03", 8, 10), ("04", 11, 12), ("05", 13, 15), ("06", 16, 16),
           ("07", 17, 24), ("08", 25, 26), ("09", 27, 29), ("10", 30, 30), ("11", 31, 36), ("12", 37, 37),
           ("13", 38, 39), ("14", 40, 42), ("15", 43, 43), ("16", 44, 44), ("17", 45, 46), ("18", 47, 49),
           ("19", 50, 51), ("20", 52, 52), ("21", 53, 57), ("22", 58, 58), ("23", 59, 64), ("24", 65, 67),
           ("25", 68, 69), ("26", 70, 72), ("27", 73, 74), ("28", 75, 75), ("29", 76, 76), ("30", 77, 78),
           ("31", 79, 80), ("32", 81, 82), ("33", 83, 83)]
LARGE_OF_MEDIUM = {f"{m:02d}": l for l, a, b in _RANGES for m in range(a, b + 1)}
FD = ["9111", "9112", "9121", "9122", "9131", "9132", "9140"]
FD_NAME = {"9111": "민간소비지출", "9112": "정부소비지출", "9121": "민간고정자본형성", "9122": "정부고정자본형성",
           "9131": "재고증감", "9132": "귀중품순취득", "9140": "수출"}
VA = {"9610": "피용자보수", "9621": "영업잉여", "9622": "고정자본소모", "9630": "생산세(보조금공제)"}

LEVELS = {}  # level -> dict


def emit(x):
    return json.dumps(x, ensure_ascii=False, allow_nan=False)


def to_level_code(basic_code, level):
    """기본부문 코드 → 해당 분류의 코드."""
    if level == "large":
        return LARGE_OF_MEDIUM[basic_code[:2]]
    return basic_code[:CODE_LEN[level]]


CODE_WIDTH = {"basic": 4, "small": 3, "medium": 2, "large": 2}


def _norm(c, level):
    """코드 정규화: Excel에 숫자로 저장돼 앞자리 0이 빠진 경우(예: 1 → '01') 복원."""
    if c is None:
        return None
    s = str(c).strip()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    if s.isdigit() and int(s) < 9000:
        s = s.zfill(CODE_WIDTH[level])
    return s


def _is_sector(c):
    return c is not None and str(c).isdigit() and int(str(c)) < 9000


def _clean(v):
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        f = float(v)
        return f if np.isfinite(f) else None
    return None if v is None else str(v)


# ------------------------------------------------------------------ 불러오기
def load_level(level, raw):
    if level not in LEVEL_LABEL:
        raise ValueError("알 수 없는 분류 수준: " + str(level))
    wb = openpyxl.load_workbook(io.BytesIO(bytes(raw)), read_only=True, data_only=True)
    book = {s: [tuple(r) for r in wb[s].values] for s in wb.sheetnames}
    wb.close()
    need = ["A표_총거래표(생산자)", "A표_국산거래표(생산자)", "A표_수입거래표(생산자)", "생산유발계수"]
    for s in need:
        if s not in book:
            raise ValueError("필수 시트가 없습니다: " + s + " (시트 목록: " + ", ".join(book) + ")")

    def frame(sheet):
        rows = book[sheet]
        cols = [_norm(c, level) for c in rows[4]]
        ci = {c: i for i, c in enumerate(cols) if c}
        ri = {_norm(r[0], level): r for r in rows[6:] if r and r[0] is not None}
        return ri, ci, cols

    tri, tci, tcols = frame("A표_총거래표(생산자)")
    codes = [c for c in tcols[2:] if _is_sector(c)]
    names = [str(tri[c][1]) for c in codes]
    n = len(codes)

    def grab(sheet, rcodes, ccodes):
        ri, ci, _ = frame(sheet)
        out = np.array([[ri[r][ci[c]] if ci[c] < len(ri[r]) else None for c in ccodes] for r in rcodes], dtype=float)
        if not np.isfinite(out).all():
            raise ValueError(sheet + ": 비수치 또는 결측 데이터")
        return out

    T = "A표_총거래표(생산자)"
    x = grab(T, ["9790"], codes)[0]
    Z = grab(T, codes, codes)
    Zd = grab("A표_국산거래표(생산자)", codes, codes)
    Zm = grab("A표_수입거래표(생산자)", codes, codes)
    va = grab(T, list(VA) + ["9690"], codes)
    Fd = grab("A표_국산거래표(생산자)", codes, FD)
    Fm = grab("A표_수입거래표(생산자)", codes, FD)
    sup = grab(T, codes, ["9310", "9329", "9390", "9140", "9190", "9090"])
    safe = np.where(x > 0, x, 1.0)
    A, Ad, Am = Z / safe, Zd / safe, Zm / safe
    vrow = va[-1] / safe
    L = np.linalg.inv(np.eye(n) - Ad)
    M = Am @ L
    V = vrow[:, None] * L
    official = grab("생산유발계수", codes, codes)
    resid = {
        "생산유발계수(공표표 대비)": float(np.max(np.abs(official - L))),
        "(I-Ad)L-I": float(np.max(np.abs((np.eye(n) - Ad) @ L - np.eye(n)))),
        "총산출-총투입 상대오차": float(np.max(np.abs(sup[:, 0] - x) / np.maximum(x, 1))),
    }
    if max(resid.values()) > 1e-6:
        raise ValueError("계수 검증 실패: " + json.dumps(resid, ensure_ascii=False))
    cs, rs = L.sum(0), L.sum(1)
    LEVELS[level] = dict(book=book, codes=codes, names=names, idx={c: i for i, c in enumerate(codes)}, x=x,
                         A=A, Ad=Ad, Am=Am, Zd=Zd, Zm=Zm, L=L, M=M, V=V, v=vrow, va=va, Fd=Fd, Fm=Fm, sup=sup,
                         BL=cs / cs.mean(), FL=rs / rs.mean())
    return emit({"level": level, "label": LEVEL_LABEL[level], "sheets": list(book),
                 "sectors": [{"code": c, "name": nm} for c, nm in zip(codes, names)],
                 "residuals": resid, "totalOutput": float(x.sum())})


def _lv(level):
    if level not in LEVELS:
        raise ValueError(LEVEL_LABEL.get(level, level) + " 자료를 먼저 불러오세요.")
    return LEVELS[level]


# ------------------------------------------------------------------ 1. 원자료 탐색
def explore(level, sheet, rowcode):
    d = _lv(level)
    rows = d["book"][sheet]
    body = [r for r in rows[6:] if r and len(r) > 1 and r[0] is not None]
    choices = [{"code": str(r[0]), "name": str(r[1]) if r[1] is not None else ""} for r in body]
    chosen = next((r for r in body if str(r[0]) == rowcode), body[0])
    head, sub = rows[4], rows[5] if len(rows) > 5 else ()
    cells = []
    for i in range(2, len(head)):
        if head[i] is None:
            continue
        val = chosen[i] if i < len(chosen) else None
        cells.append({"code": str(head[i]), "name": str(sub[i]).replace("\n", " ") if i < len(sub) and sub[i] else "",
                      "value": _clean(val)})
    unit = str(rows[3][0]) if rows[3] and rows[3][0] else "단위 : 계수(단위 없음, 산출 1단위당 비율)"
    return emit({"rows": choices, "selected": str(chosen[0]), "name": str(chosen[1]), "unit": unit,
                 "title": str(rows[0][0] or sheet), "cells": cells})


# ------------------------------------------------------------------ 2. 산업구조 지표
def overview(level):
    d = _lv(level)
    x, va, sup = d["x"], d["va"], d["sup"]
    safe = np.where(x > 0, x, 1.0)
    supply = np.where(sup[:, 2] > 0, sup[:, 2], 1.0)
    out = []
    for i, (c, nm) in enumerate(zip(d["codes"], d["names"])):
        out.append({
            "code": c, "name": nm,
            "output": float(x[i]),                                  # 백만원
            "valueAdded": float(va[-1][i]),                         # 백만원
            "vaRate": float(va[-1][i] / safe[i] * 100),             # %
            "laborShare": float(va[0][i] / va[-1][i] * 100) if va[-1][i] > 0 else 0.0,  # 피용자보수/부가가치 %
            "importInputRate": float(d["Am"][:, i].sum() * 100),    # 수입중간투입/총투입 %
            "exportRate": float(sup[i, 3] / safe[i] * 100),         # 수출/총산출 %
            "importDependence": float(sup[i, 1] / supply[i] * 100), # 수입/총공급 %
            "exports": float(sup[i, 3]), "imports": float(sup[i, 1]),
            "prodMultiplier": float(d["L"][:, i].sum()),            # 생산유발계수 열합
            "vaMultiplier": float(d["V"][:, i].sum()),
            "importMultiplier": float(d["M"][:, i].sum()),
            "BL": float(d["BL"][i]), "FL": float(d["FL"][i]),
        })
    tot = {"output": float(x.sum()), "valueAdded": float(va[-1].sum()),
           "vaRate": float(va[-1].sum() / x.sum() * 100), "exports": float(sup[:, 3].sum()),
           "imports": float(sup[:, 1].sum()), "laborShare": float(va[0].sum() / va[-1].sum() * 100)}
    return emit({"rows": out, "total": tot})


# ------------------------------------------------------------------ 3. 수요충격 시뮬레이션
def _demand_vector(d, level, shocks):
    f = np.zeros(len(d["codes"]))
    direct_imp = 0.0
    for s in shocks:
        if s.get("type") == "item":
            item = s["item"]
            if item not in FD:
                raise ValueError("알 수 없는 최종수요 항목: " + item)
            pct = float(s["pct"]) / 100
            j = FD.index(item)
            if not np.isfinite(pct):
                raise ValueError("변화율은 유한한 수여야 합니다.")
            mask = np.ones(len(f), bool)
            if s.get("code") not in (None, "", "ALL"):
                mask[:] = False
                mask[d["idx"][s["code"]]] = True
            f[mask] += d["Fd"][mask, j] * pct
            direct_imp += float((d["Fm"][mask, j] * pct).sum())
        else:
            a = float(s["amount"])
            if not np.isfinite(a):
                raise ValueError("수요 변화액은 유한한 수여야 합니다.")
            if s["code"] not in d["idx"]:
                raise ValueError("현재 분류에 없는 산업 코드: " + s["code"])
            f[d["idx"][s["code"]]] += a
    return f, direct_imp


def simulate(level, payload):
    d = _lv(level)
    shocks = json.loads(payload)
    f, direct_imp = _demand_vector(d, level, shocks)
    if not np.any(f != 0):
        raise ValueError("0이 아닌 수요 변화가 필요합니다.")
    xx, m, v = d["L"] @ f, d["M"] @ f, d["V"] @ f
    emp = d["va"][0] / np.where(d["x"] > 0, d["x"], 1.0) * xx      # 피용자보수 유발액
    rows = [{"code": c, "name": n, "demand": float(a), "production": float(p), "pct": float(p / x * 100) if x > 0 else 0.0,
             "imports": float(b), "valueAdded": float(e), "wages": float(w)}
            for c, n, a, p, x, b, e, w in zip(d["codes"], d["names"], f, xx, d["x"], m, v, emp)]
    rows.sort(key=lambda r: abs(r["production"]), reverse=True)
    fs = float(f.sum())
    return emit({"demand": fs, "production": float(xx.sum()), "imports": float(m.sum()), "directImports": direct_imp,
                 "valueAdded": float(v.sum()), "wages": float(emp.sum()),
                 "multiplier": float(xx.sum() / fs) if fs else None,
                 "gdpShare": float(v.sum() / d["va"][-1].sum() * 100), "rows": rows})


# ------------------------------------------------------------------ 4. 전후방 연관효과
def linkage(level):
    d = _lv(level)
    q = {"I": "핵심(전후방 모두 높음)", "II": "후방연관 높음(원자재 수요형)", "III": "독립형(둘 다 낮음)",
         "IV": "전방연관 높음(중간재 공급형)"}
    rows = []
    for c, n, bl, fl, x in zip(d["codes"], d["names"], d["BL"], d["FL"], d["x"]):
        k = "I" if bl >= 1 and fl >= 1 else "II" if bl >= 1 else "IV" if fl >= 1 else "III"
        rows.append({"code": c, "name": n, "BL": float(bl), "FL": float(fl), "output": float(x), "quad": k, "quadName": q[k]})
    return emit({"rows": rows, "quadNames": q})


# ------------------------------------------------------------------ 5. 가격파급효과 (비용인상 모형)
def price(level, payload):
    """Δp = (I - Ad')^-1 Δc,  Δc_j = Σ_k r_k·Am_kj (수입가격) + r_w·w_j (임금 등 부가가치 항목)."""
    d = _lv(level)
    shocks = json.loads(payload)
    n = len(d["codes"])
    dc = np.zeros(n)
    rimp = np.zeros(n)  # 수입품 가격 상승률
    for s in shocks:
        r = float(s["pct"]) / 100
        if not np.isfinite(r):
            raise ValueError("상승률은 유한한 수여야 합니다.")
        if s["kind"] == "import":
            rimp[d["idx"][s["code"]]] += r
        else:  # 부가가치 항목 (예: 9610 피용자보수)
            k = list(VA).index(s["component"])
            share = d["va"][k] / np.where(d["x"] > 0, d["x"], 1.0)
            if s.get("code") in (None, "", "ALL"):
                dc += r * share
            else:
                j = d["idx"][s["code"]]
                dc[j] += r * share[j]
    dc += rimp @ d["Am"]
    dp = d["L"].T @ dc
    # 소비자물가 근사: 민간소비(국산+수입) 가중평균
    cd, cm = d["Fd"][:, 0], d["Fm"][:, 0]
    cpi = float((cd @ dp + cm @ rimp) / (cd.sum() + cm.sum()) * 100)
    ppi = float(d["x"] @ dp / d["x"].sum() * 100)
    rows = [{"code": c, "name": nm, "direct": float(a * 100), "total": float(b * 100)}
            for c, nm, a, b in zip(d["codes"], d["names"], dc, dp)]
    rows.sort(key=lambda r: r["total"], reverse=True)
    return emit({"cpi": cpi, "ppi": ppi, "rows": rows})


# ------------------------------------------------------------------ 산업 공급망 리포트
def profile(level, code, topn=10):
    """한 산업의 공급처(후방)·판매처(전방)·수입 원자재·유발효과를 한 번에 정리."""
    d = _lv(level)
    if code not in d["idx"]:
        raise ValueError("현재 분류에 없는 산업 코드: " + str(code))
    j = d["idx"][code]
    names, codes, x = d["names"], d["codes"], d["x"]
    topn = int(topn)

    def top(vec, amt, k=topn, skip_self=False):
        order = np.argsort(-vec)
        out = []
        for i in order:
            if vec[i] <= 0 or len(out) >= k:
                break
            out.append({"code": codes[i], "name": names[i], "coef": float(vec[i]), "amount": float(amt[i]),
                        "self": bool(i == j)})
        return out

    # 후방: 산출 1원 생산에 필요한 국산·수입 중간재
    suppliers = top(d["Ad"][:, j], d["Zd"][:, j])
    imported = top(d["Am"][:, j], d["Zm"][:, j])
    # 전방: 이 산업의 국산품이 어디에 팔리는가 (중간수요 + 최종수요 항목)
    row_int = d["Zd"][j]
    row_fd = d["Fd"][j]
    total_sales = float(row_int.sum() + row_fd.sum())
    safe = total_sales if total_sales > 0 else 1.0
    buyers = [{"code": codes[i], "name": names[i], "amount": float(row_int[i]), "share": float(row_int[i] / safe * 100)}
              for i in np.argsort(-row_int)[:topn] if row_int[i] > 0]
    final = [{"code": c, "name": FD_NAME[c], "amount": float(v), "share": float(v / safe * 100)}
             for c, v in zip(FD, row_fd) if abs(v) > 0]
    # 유발효과: 이 산업 최종수요 1조원(1,000,000백만원) 증가 시
    shock = 1_000_000.0
    induced = d["L"][:, j] * shock
    spill = [{"code": codes[i], "name": names[i], "amount": float(induced[i])}
             for i in np.argsort(-induced)[:topn + 1] if i != j][:topn]
    va = d["va"]
    rank = lambda arr: int((arr > arr[j]).sum() + 1)
    pm = d["L"].sum(0)
    return emit({
        "code": code, "name": names[j], "level": LEVEL_LABEL[level], "n": len(codes),
        "output": float(x[j]), "outputShare": float(x[j] / x.sum() * 100), "outputRank": rank(x),
        "valueAdded": float(va[-1][j]), "vaRate": float(va[-1][j] / x[j] * 100) if x[j] > 0 else 0.0,
        "wages": float(va[0][j]), "profit": float(va[1][j]),
        "exportRate": float(d["sup"][j, 3] / x[j] * 100) if x[j] > 0 else 0.0,
        "importDependence": float(d["sup"][j, 1] / d["sup"][j, 2] * 100) if d["sup"][j, 2] > 0 else 0.0,
        "domInputRate": float(d["Ad"][:, j].sum() * 100), "impInputRate": float(d["Am"][:, j].sum() * 100),
        "prodMultiplier": float(pm[j]), "prodMultiplierAvg": float(pm.mean()), "prodRank": rank(pm),
        "vaMultiplier": float(d["V"][:, j].sum()), "importMultiplier": float(d["M"][:, j].sum()),
        "BL": float(d["BL"][j]), "FL": float(d["FL"][j]),
        "suppliers": suppliers, "imported": imported, "buyers": buyers, "final": final, "totalSales": total_sales,
        "spill": spill, "selfInduced": float(induced[j]), "totalInduced": float(induced.sum()),
    })


# ------------------------------------------------------------------ 6. 분류 수준별 비교
def compare(basic_code, amount):
    a = float(amount)
    if not np.isfinite(a) or a == 0:
        raise ValueError("0이 아닌 금액을 입력하세요.")
    out = []
    for level in ["basic", "small", "medium", "large"]:
        if level not in LEVELS:
            continue
        d = LEVELS[level]
        code = to_level_code(basic_code, level)
        if code not in d["idx"]:
            continue
        j = d["idx"][code]
        out.append({"level": level, "label": LEVEL_LABEL[level], "sectors": len(d["codes"]), "code": code,
                    "name": d["names"][j], "production": float(d["L"][:, j].sum() * a),
                    "valueAdded": float(d["V"][:, j].sum() * a), "imports": float(d["M"][:, j].sum() * a),
                    "prodMultiplier": float(d["L"][:, j].sum()), "BL": float(d["BL"][j]), "FL": float(d["FL"][j])})
    if not out:
        raise ValueError("불러온 분류 자료가 없습니다.")
    return emit({"rows": out})
