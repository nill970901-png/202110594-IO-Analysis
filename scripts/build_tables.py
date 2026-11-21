"""기본부문(380) 투입산출표를 소분류(165)·중분류(83)·대분류(33)로 통합하는 스크립트.

사용법:  python scripts/build_tables.py data/io-2023-basic.xlsx data/

1) 거래표(총·국산·수입, 백만원)를 부문분류에 따라 행·열 합산
2) 합산된 거래표에서 투입계수, 생산·수입·부가가치 유발계수, 최종수요 항목별 유발액 재계산
3) 같은 계산을 기본부문에 적용해 한국은행 공표 계수와의 오차를 검증
"""
import sys, json
from pathlib import Path
import numpy as np
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill

sys.path.insert(0, str(Path(__file__).parent))
from classification import SMALL, MEDIUM, LARGE, parents

FD = ["9111", "9112", "9121", "9122", "9131", "9132", "9140"]
TRADE_SHEETS = ["A표_총거래표(생산자)", "A표_국산거래표(생산자)", "A표_수입거래표(생산자)"]


def read_sheet(ws):
    rows = list(ws.values)
    cols = [str(c) if c is not None else None for c in rows[4]]
    names = list(rows[5])
    body = [r for r in rows[6:] if r and r[0] is not None]
    return rows[:4], cols, names, body


def is_sector(code):
    return code is not None and code.isdigit() and len(code) == 4 and int(code) < 9000


def aggregate_trade(ws, groups, gnames, key):
    """거래표 1장을 groups 기준으로 합산. key(code)->group code."""
    head, cols, names, body = read_sheet(ws)
    col_idx = {c: i for i, c in enumerate(cols) if c}
    sec_cols = [c for c in cols if is_sector(c)]
    extra_cols = [c for c in cols[2:] if c and not is_sector(c)]
    extra_names = {c: names[col_idx[c]] for c in extra_cols}
    out_rows = {}  # code -> {colcode: value}
    order_rows = []
    for r in body:
        rc = str(r[0])
        tgt = key(rc) if is_sector(rc) else rc
        if tgt not in out_rows:
            out_rows[tgt] = {}
            order_rows.append((tgt, gnames.get(tgt, r[1])))
        d = out_rows[tgt]
        for c in cols[2:]:
            if not c:
                continue
            v = r[col_idx[c]]
            v = float(v) if isinstance(v, (int, float)) else 0.0
            cc = key(c) if is_sector(c) else c
            d[cc] = d.get(cc, 0.0) + v
    new_cols = list(groups) + extra_cols
    new_names = [gnames[g] for g in groups] + [extra_names[c] for c in extra_cols]
    matrix = [[code, name] + [out_rows[code].get(c, 0.0) for c in new_cols] for code, name in order_rows]
    return head, new_cols, new_names, matrix


def to_np(matrix, cols, row_codes, col_codes):
    ri = {r[0]: i for i, r in enumerate(matrix)}
    ci = {c: i + 2 for i, c in enumerate(cols)}
    return np.array([[matrix[ri[r]][ci[c]] for c in col_codes] for r in row_codes], dtype=float)


def coefficients(total, dom, imp, codes):
    """거래표(총·국산·수입)에서 계수표 계산. 반환: dict of numpy arrays."""
    tm, tc = total[3], total[1]
    x = to_np(tm, tc, ["9790"], codes)[0]                       # 총투입 = 총산출
    Z = to_np(tm, tc, codes, codes)
    Zd = to_np(dom[3], dom[1], codes, codes)
    Zm = to_np(imp[3], imp[1], codes, codes)
    va_codes = ["9590", "9610", "9621", "9622", "9630", "9690", "9790"]
    VArows = to_np(tm, tc, va_codes, codes)
    with np.errstate(divide="ignore", invalid="ignore"):
        div = np.where(x > 0, x, np.nan)
        A = np.nan_to_num(Z / div)
        Ad = np.nan_to_num(Zd / div)
        Am = np.nan_to_num(Zm / div)
        Arows = np.nan_to_num(VArows / div)
    n = len(codes)
    L = np.linalg.inv(np.eye(n) - Ad)
    v = Arows[va_codes.index("9690")]
    M = Am @ L
    V = v[:, None] * L
    Fd = to_np(dom[3], dom[1], codes, FD)                      # 국산 최종수요
    Fm = to_np(imp[3], imp[1], codes, FD)                      # 수입 최종수요
    return dict(x=x, A=A, Ad=Ad, Am=Am, Arows=Arows, va_codes=va_codes, L=L, M=M, V=V, v=v, Fd=Fd, Fm=Fm)


# ---------------------------------------------------------------- 엑셀 쓰기
BOLD = Font(bold=True)
HEAD_FILL = PatternFill("solid", fgColor="DCE6F0")


def write_header(ws, title, unit, cols, names, corner="상품\n상품"):
    ws.append([title])
    ws.append(["2023년"])
    ws.append(["생산자가격"])
    ws.append([unit])
    ws.append([corner, None] + cols)
    ws.append([None, None] + names)
    ws["A1"].font = Font(bold=True, size=12)
    for row in ws.iter_rows(min_row=5, max_row=6):
        for c in row:
            c.font = BOLD
            c.fill = HEAD_FILL
            c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.freeze_panes = "C7"
    ws.column_dimensions["B"].width = 26


def write_trade(wb, title_sheet, head, cols, names, matrix, label):
    ws = wb.create_sheet(title_sheet)
    write_header(ws, str(head[0][0]).replace("(2023년 기준)", f"(2023년 기준, {label})"), "단위 : 백만원", cols, names)
    for r in matrix:
        ws.append([r[0], r[1]] + [round(v, 6) for v in r[2:]])


def write_coef(wb, sheet, title, codes, names, rows, row_labels, extra_cols=None, unit=None):
    """rows: list of (code, name, values[n] [+extra])"""
    ws = wb.create_sheet(sheet)
    ec = extra_cols or []
    write_header(ws, title, unit, codes + [c for c, _ in ec], names + [n for _, n in ec])
    for code, name, vals in rows:
        ws.append([code, name] + [float(v) for v in vals])


def build_workbook(c, codes, names, trade, label, members, out_path):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for sheet, (head, cols, cnames, matrix) in zip(TRADE_SHEETS, trade):
        write_trade(wb, sheet, head, cols, cnames, matrix, label)
    n = len(codes)
    va_names = {"9590": "중간투입계", "9610": "피용자보수", "9621": "영업잉여", "9622": "고정자본소모",
                "9630": "생산세(보조금공제)", "9690": "부가가치계", "9790": "총투입계"}
    t = f"(2023년 가격, {label})"
    # 총투입계수
    rows = [(codes[i], names[i], list(c["A"][i])) for i in range(n)]
    rows += [(vc, va_names[vc], list(c["Arows"][k])) for k, vc in enumerate(c["va_codes"])]
    write_coef(wb, "총투입계수(A)", "투입산출표_총투입계수표" + t, codes, names, rows, None)
    for key, sheet, nm in [("Am", "수입투입계수(Am)", "수입투입계수표"), ("Ad", "국산투입계수(Ad)", "국산투입계수표")]:
        m = c[key]
        rows = [(codes[i], names[i], list(m[i])) for i in range(n)] + [("9590", "소계", list(m.sum(0)))]
        write_coef(wb, sheet, "투입산출표_" + nm + t, codes, names, rows, None)
    L = c["L"]
    cs, rs = L.sum(0), L.sum(1)
    rows = [(codes[i], names[i], list(L[i]) + [rs[i], rs[i] / rs.mean()]) for i in range(n)]
    rows += [("9590", "열합", list(cs) + [cs.sum(), None]), ("9801", "영향력계수", list(cs / cs.mean()) + [None, None])]
    ws_rows = [(a, b, [0 if v is None else v for v in vals]) for a, b, vals in rows]
    write_coef(wb, "생산유발계수", "투입산출표_생산유발계수표" + t, codes, names, ws_rows, None,
               extra_cols=[("9090", "행합"), ("9802", "감응도계수")])
    ws = wb["생산유발계수"]
    ws.cell(row=ws.max_row - 1, column=n + 4).value = None
    ws.cell(row=ws.max_row, column=n + 3).value = None
    ws.cell(row=ws.max_row, column=n + 4).value = None
    for key, sheet, nm in [("M", "수입유발계수", "수입유발계수표"), ("V", "부가가치유발계수", "부가가치유발계수표")]:
        m = c[key]
        rows = [(codes[i], names[i], list(m[i])) for i in range(n)] + [("9590", "열합", list(m.sum(0)))]
        write_coef(wb, sheet, "투입산출표_" + nm + t, codes, names, rows, None)
    # 최종수요 항목별 유발액
    fd_names = ["민간소비지출\n(가계 및 가계봉사\n비영리단체)", "정부소비지출", "민간고정\n자본형성", "정부고정\n자본형성",
                "재고증감", "귀중품\n순취득", "수출"]
    Fd, Fm = c["Fd"], c["Fm"]
    ftot = (Fd + Fm).sum(0)
    for kind, sheet, nm in [("X", "최종수요 항목별 생산유발액", "생산유발액"), ("M", "최종수요 항목별 수입유발액", "수입유발액"),
                            ("V", "최종수요 항목별 부가가치유발액", "부가가치유발액")]:
        if kind == "X":
            E = L @ Fd
        elif kind == "M":
            E = c["M"] @ Fd
            E = E + Fm  # 최종수요의 직접 수입 포함
        else:
            E = c["V"] @ Fd
        E = np.column_stack([E, E.sum(1)])
        tot = E.sum(0)
        with np.errstate(divide="ignore", invalid="ignore"):
            share = np.nan_to_num(tot / tot[-1] * 100)
            coef = np.nan_to_num(tot / np.append(ftot, ftot.sum()))
        rows = [(codes[i], names[i], list(E[i])) for i in range(n)]
        rows += [("9819", "열합", list(tot)), ("9812", "구성비(%)", list(share)), ("9811", "유발계수", list(coef))]
        ws = wb.create_sheet(sheet)
        write_header(ws, "투입산출표_최종수요 항목별 " + nm + t, "단위 : 백만원", FD + ["9190"], fd_names + ["계"],
                     corner="최종수요 항목\n상품")
        for code, name, vals in rows:
            ws.append([code, name] + [float(v) for v in vals])
    # 부문분류 대응표
    ws = wb.create_sheet("부문분류")
    ws.append([f"부문분류 대응표 ({label})"])
    ws.append(["2023년"])
    ws.append(["코드 체계: 기본부문 4자리 → 소분류 앞 3자리 → 중분류 앞 2자리 → 대분류"])
    ws.append(["단위 : 없음"])
    ws.append(["코드", "부문명", "포함 기본부문 수", "포함 기본부문"])
    ws.append([None, None, None, None])
    for code, name in zip(codes, names):
        mem = members[code]
        ws.append([code, name, len(mem), ", ".join(f"{a} {b}" for a, b in mem)])
    for row in ws.iter_rows(min_row=5, max_row=5):
        for cell in row:
            cell.font = BOLD
    wb.save(out_path)


def main(src, outdir):
    outdir = Path(outdir)
    wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
    _, cols, names, body = read_sheet(wb["A표_총거래표(생산자)"])
    basic = [(str(r[0]), str(r[1])) for r in body if is_sector(str(r[0]))]
    assert len(basic) == 380
    bnames = dict(basic)

    levels = {
        "basic": ("기본부문", lambda c: c, bnames),
        "small": ("소분류", lambda c: parents(c)[0], SMALL),
        "medium": ("중분류", lambda c: parents(c)[1], MEDIUM),
        "large": ("대분류", lambda c: parents(c)[2], LARGE),
    }
    report = {}
    official = {}
    for key, (label, fn, gnames) in levels.items():
        groups = []
        members = {}
        for bc, bn in basic:
            g = fn(bc)
            if g not in members:
                groups.append(g)
                members[g] = []
            members[g].append((bc, bn))
        trade = [aggregate_trade(wb[s], groups, gnames, fn) for s in TRADE_SHEETS]
        c = coefficients(trade[0], trade[1], trade[2], groups)
        # 균형 점검: 행합(총산출) = 열합(총투입)
        tot = to_np(trade[0][3], trade[0][1], groups, ["9310"])[:, 0]
        bal = float(np.max(np.abs(tot - c["x"]) / np.maximum(c["x"], 1)))
        info = {"부문수": len(groups), "총산출_조원": round(float(c["x"].sum()) / 1e6, 3),
                "행렬균형_최대상대오차": bal,
                "생산유발승수_평균": float(c["L"].sum(0).mean())}
        if key == "basic":
            # 한국은행 공표 계수와 비교 (계산 방법 검증)
            def off(sheet):
                _, oc, _, ob = read_sheet(wb[sheet])
                ci = {cc: i for i, cc in enumerate(oc)}
                ri = {str(r[0]): r for r in ob}
                return np.array([[ri[a][ci[b]] for b in groups] for a in groups], dtype=float)
            info["한은계수와_최대절대오차"] = {
                "국산투입계수": float(np.max(np.abs(off("국산투입계수(Ad)") - c["Ad"]))),
                "생산유발계수": float(np.max(np.abs(off("생산유발계수") - c["L"]))),
                "수입유발계수": float(np.max(np.abs(off("수입유발계수") - c["M"]))),
                "부가가치유발계수": float(np.max(np.abs(off("부가가치유발계수") - c["V"]))),
            }
            for s, kind in [("최종수요 항목별 생산유발액", "X"), ("최종수요 항목별 수입유발액", "M"), ("최종수요 항목별 부가가치유발액", "V")]:
                _, oc, _, ob = read_sheet(wb[s])
                ci = {cc: i for i, cc in enumerate(oc)}
                ri = {str(r[0]): r for r in ob}
                O = np.array([[ri[a][ci[b]] for b in FD] for a in groups], dtype=float)
                E = {"X": c["L"] @ c["Fd"], "M": c["M"] @ c["Fd"] + c["Fm"], "V": c["V"] @ c["Fd"]}[kind]
                info["한은계수와_최대절대오차"][s + "(백만원)"] = float(np.max(np.abs(O - E)))
        else:
            build_workbook(c, groups, [gnames[g] for g in groups], trade, label, members, outdir / f"io-2023-{key}.xlsx")
        report[label] = info
        print(label, json.dumps(info, ensure_ascii=False))
    (outdir / "build-report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    # 분류 대응 CSV
    lines = ["기본부문코드,기본부문명,소분류코드,소분류명,중분류코드,중분류명,대분류코드,대분류명"]
    for bc, bn in basic:
        s, m, l = parents(bc)
        q = lambda t: '"' + t.replace('"', '""') + '"'
        lines.append(",".join([bc, q(bn), s, q(SMALL[s]), m, q(MEDIUM[m]), l, q(LARGE[l])]))
    (outdir / "classification.csv").write_text("﻿" + "\n".join(lines) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
