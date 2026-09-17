"""LBO 杠杆收购模型 Excel 构建器。

Sheet 结构（7 sheets）:
1. Assumptions       - 交易/资本结构/利率/偿债/运营/退出参数
2. Sources & Uses     - 资金来源与用途（Revolver 为平衡项）
3. Debt Schedule     - Revolver/TLA/TLB/SrNotes/Sub 滚动 + 现金
4. Income Statement  - 预测利润表（各层级利息）
5. Cash Flow (CFADS)  - 偿债可用现金流
6. Exit & Returns     - 退出价值/IRR/MoIC
7. Sensitivity       - 入场×退出倍数 IRR 敏感性

跨表链接:
  S&U Revolver → Debt Schedule 期初
  Debt Schedule 期初余额 → IS 利息费用
  IS 净利润 → CFADS → Debt Schedule 偿债现金
  Debt Schedule 期末债务/现金 → Exit 净债务 → 股权价值 → IRR
"""
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from backend import styles as S
from backend.styles import (
    style_input, style_formula, style_total, style_header,
    style_subheader, style_label, style_title, set_col_widths,
    write_header_row, NUM_FMT, PCT_FMT, NUM_FMT_INT, MULT_FMT, PRICE_FMT,
)
from backend.schemas import LBOParams
from backend.custom_items import expression_to_excel


def ycol(year_index):
    return get_column_letter(2 + year_index)


def _write_assumptions(wb, p: LBOParams):
    ws = wb.create_sheet("Assumptions")
    ws.cell(row=1, column=1, value="Assumptions 假设参数").font = S.TITLE_FONT
    ws.cell(row=1, column=1).alignment = S.LEFT_ALIGN
    n = 5  # LBO 固定 5 年预测
    headers = ["参数 Parameter", "Value"] + [f"Year {i}" for i in range(1, n + 1)]
    write_header_row(ws, 3, headers)

    def section(row, title):
        c = ws.cell(row=row, column=1, value=title)
        style_subheader(c)
        for col in range(2, 2 + n + 1):
            style_subheader(ws.cell(row=row, column=col))

    def inp(row, label, value, fmt=NUM_FMT, year_col=None):
        c = ws.cell(row=row, column=1, value=label)
        style_label(c)
        if year_col is None:
            cell = ws.cell(row=row, column=2, value=value)
            style_input(cell, fmt)
        else:
            cell = ws.cell(row=row, column=2 + year_col, value=value)
            style_input(cell, fmt)

    # 交易假设
    section(4, "交易假设 Transaction Assumptions")
    inp(5, "LTM Revenue", p.ltm_revenue)
    inp(6, "LTM EBITDA", p.ltm_ebitda)
    inp(7, "Entry EV/EBITDA", p.entry_ev_ebitda, MULT_FMT)
    # B8 = Entry EV (formula)
    c = ws.cell(row=8, column=1, value="Entry EV 入场企业价值")
    style_label(c)
    cell = ws.cell(row=8, column=2, value="=B6*B7")
    style_formula(cell, NUM_FMT)
    inp(9, "Existing Net Debt 现有净债务", p.existing_net_debt)
    inp(10, "Existing Cash 现有现金", p.existing_cash)
    inp(11, "Transaction Fees 交易费用", p.transaction_fees)
    inp(12, "Financing Fees 融资费用", p.financing_fees)

    # 资本结构
    section(14, "资本结构 Capital Structure")
    inp(14, "Sponsor Equity", p.sponsor_equity) if False else None
    ws.cell(row=14, column=1, value="Sponsor Equity").font = S.LABEL_FONT
    # fix: proper label
    c14 = ws.cell(row=14, column=1, value="Sponsor Equity")
    style_label(c14)
    inp(14, "Sponsor Equity", p.sponsor_equity) if False else None
    # 直接写
    style_label(ws.cell(row=14, column=1, value="Sponsor Equity"))
    cell = ws.cell(row=14, column=2, value=p.sponsor_equity)
    style_input(cell)
    inp(15, "Revolver Capacity", p.revolver_capacity)
    inp(16, "Term Loan A", p.term_loan_a)
    inp(17, "Term Loan B", p.term_loan_b)
    inp(18, "Senior Notes", p.senior_notes)
    inp(19, "Subordinated Debt", p.subordinated_debt)

    # 利率
    section(21, "利率 Interest Rates")
    inp(21, "Revolver Rate", p.revolver_rate, PCT_FMT)
    inp(22, "Term Loan A Rate", p.tla_rate, PCT_FMT)
    inp(23, "Term Loan B Rate", p.tlb_rate, PCT_FMT)
    inp(24, "Senior Notes Rate", p.senior_notes_rate, PCT_FMT)
    inp(25, "Subordinated Debt Rate", p.sub_rate, PCT_FMT)

    # 偿债
    section(27, "偿债 Debt Repayment")
    inp(27, "TLA Mandatory Amort %", p.tla_mandatory_amort, PCT_FMT)
    inp(28, "TLB Mandatory Amort %", p.tlb_mandatory_amort, PCT_FMT)
    inp(29, "Cash Sweep %", p.cash_sweep_pct, PCT_FMT)
    inp(30, "Min Cash Balance", p.min_cash_balance)

    # 运营
    section(32, "运营假设 Operating Assumptions")
    style_label(ws.cell(row=32, column=1, value="Revenue Growth %"))
    for i, g in enumerate([p.rev_growth_y1, p.rev_growth_y2, p.rev_growth_y3,
                           p.rev_growth_y4, p.rev_growth_y5][:n]):
        inp(32, "Revenue Growth %", g, PCT_FMT, year_col=i + 1)
    style_label(ws.cell(row=33, column=1, value="EBITDA Margin %"))
    for i, m in enumerate([p.ebitda_margin] * n):
        inp(33, "EBITDA Margin %", m, PCT_FMT, year_col=i + 1)
    inp(34, "D&A % of Revenue", p.da_pct, PCT_FMT)
    inp(35, "CapEx % of Revenue", p.capex_pct, PCT_FMT)
    inp(36, "ΔNWC % of Revenue", p.nwc_pct, PCT_FMT)
    inp(37, "Tax Rate", p.tax_rate, PCT_FMT)

    # 退出
    section(39, "退出 Exit")
    inp(39, "Exit EV/EBITDA", p.exit_ev_ebitda, MULT_FMT)
    inp(40, "Exit Year", p.exit_year, NUM_FMT_INT)
    inp(41, "Cash Interest Rate", p.cash_interest_rate, PCT_FMT)

    set_col_widths(ws, {"A": 44, "B": 14})
    for i in range(n):
        ws.column_dimensions[get_column_letter(3 + i)].width = 14
    ws.sheet_view.showGridLines = False


def _write_sources_uses(wb, p: LBOParams):
    ws = wb.create_sheet("Sources & Uses")
    ws.cell(row=1, column=1, value="Sources & Uses 资金来源与用途").font = S.TITLE_FONT
    write_header_row(ws, 3, ["项目 Item", "Amount"])

    A = "Assumptions"

    def line(row, label, formula, fmt=NUM_FMT, total=False, input_val=None):
        c = ws.cell(row=row, column=1, value=label)
        (style_total if total else style_label)(c)
        cell = ws.cell(row=row, column=2)
        if input_val is not None:
            cell.value = input_val
            style_input(cell, fmt)
        else:
            cell.value = formula
            style_formula(cell, fmt, link="!" in (formula or ""))

    # Uses
    sec = ws.cell(row=4, column=1, value="Uses 用途")
    style_subheader(sec)
    style_subheader(ws.cell(row=4, column=2))

    line(6, "Purchase of Equity 购买股权", f"={A}!B8-{A}!B9")
    line(7, "Refinance Existing Debt 再融资债务", f"={A}!B9+{A}!B10")
    line(8, "Transaction Fees 交易费用", f"={A}!B11")
    line(9, "Financing Fees 融资费用", f"={A}!B12")
    line(10, "Total Uses 用途合计", "=SUM(B6:B9)", total=True)

    # Sources
    sec2 = ws.cell(row=12, column=1, value="Sources 来源")
    style_subheader(sec2)
    style_subheader(ws.cell(row=12, column=2))

    line(13, "Sponsor Equity", f"={A}!B14")
    line(14, "Revolver Draw (plug) 循环额度提取", "=B10-B13-B15-B16-B17-B18-B19")
    line(15, "Term Loan A", f"={A}!B16")
    line(16, "Term Loan B", f"={A}!B17")
    line(17, "Senior Notes", f"={A}!B18")
    line(18, "Subordinated Debt", f"={A}!B19")
    line(19, "Existing Cash 现有现金", f"={A}!B10")
    line(20, "Total Sources 来源合计", "=SUM(B13:B19)", total=True)

    # Balance check
    ws.cell(row=22, column=1, value="平衡校验 Balance Check").font = S.SUBHEADER_FONT
    style_label(ws.cell(row=22, column=1, value="平衡校验 Balance Check"))
    cell = ws.cell(row=22, column=2,
                   value='=IF(ROUND(B20-B10,2)=0,"BALANCED","IMBALANCED")')
    cell.font = S.TOTAL_FONT
    cell.alignment = S.CENTER_ALIGN
    cell.border = S.BORDER

    set_col_widths(ws, {"A": 40, "B": 16})
    ws.sheet_view.showGridLines = False


def _write_debt_schedule(wb, p: LBOParams):
    ws = wb.create_sheet("Debt Schedule")
    ws.cell(row=1, column=1, value="Debt Schedule 债务滚动表").font = S.TITLE_FONT
    n = 5
    write_header_row(ws, 3, ["项目 Item", "Close"] + [f"Year {i}" for i in range(1, n + 1)])

    A = "Assumptions"
    SU = "'Sources & Uses'"
    CF = "'Cash Flow (CFADS)'"

    def row(r, label, formula_func, fmt=NUM_FMT, total=False, indent=0):
        c = ws.cell(row=r, column=1, value=label)
        style_total(c) if total else style_label(c, indent)
        for yi in range(n + 1):
            col = ycol(yi)
            f = formula_func(col, yi)
            cell = ws.cell(row=r, column=2 + yi)
            if f is not None:
                cell.value = f
            if total:
                style_total(cell, fmt)
            else:
                style_formula(cell, fmt, link="!" in (f or ""))

    # === Cash & CFADS ===
    sec = ws.cell(row=4, column=1, value="现金与偿债可用 Cash & CFADS")
    style_subheader(sec)
    for col in range(2, 2 + n + 1):
        style_subheader(ws.cell(row=4, column=col))

    def beg_cash(col, yi):
        if yi == 0:
            return f"={A}!B10"
        return f"={ycol(yi-1)}35"
    row(5, "Beginning Cash 期初现金", beg_cash)

    def cfads(col, yi):
        return f"=0" if yi == 0 else f"={CF}!{col}13"
    row(6, "CFADS 偿债可用现金流", cfads)

    row(7, "Min Cash 最低现金", lambda c, yi: f"={A}!$B$30")
    row(8, "Cash Available 可用现金", lambda c, yi: f"=MAX(0,{c}5+{c}6-{c}7)")

    # === Revolver ===
    sec2 = ws.cell(row=10, column=1, value="Revolver 循环额度")
    style_subheader(sec2)
    for col in range(2, 2 + n + 1):
        style_subheader(ws.cell(row=10, column=col))

    def rev_beg(col, yi):
        if yi == 0:
            return f"={SU}!B14"
        return f"={ycol(yi-1)}13"
    row(11, "Beginning 期初", rev_beg)

    row(12, "Repayment 偿还", lambda c, yi: f"=MIN({c}11,{c}8)")

    def rev_end(col, yi):
        return f"={col}11-{col}12"
    row(13, "Ending 期末", rev_end)

    # === Term Loan A ===
    sec3 = ws.cell(row=15, column=1, value="Term Loan A")
    style_subheader(sec3)
    for col in range(2, 2 + n + 1):
        style_subheader(ws.cell(row=15, column=col))

    def tla_beg(col, yi):
        if yi == 0:
            return f"={A}!B16"
        return f"={ycol(yi-1)}18"
    row(16, "Beginning 期初", tla_beg)

    row(17, "Mandatory Amort 强制摊销", lambda c, yi: f"=MIN({c}16,{A}!$B$16*{A}!$B$27)")
    row(18, "Ending 期末", lambda c, yi: f"={c}16-{c}17")

    # === Term Loan B ===
    sec4 = ws.cell(row=20, column=1, value="Term Loan B")
    style_subheader(sec4)
    for col in range(2, 2 + n + 1):
        style_subheader(ws.cell(row=20, column=col))

    def tlb_beg(col, yi):
        if yi == 0:
            return f"={A}!B17"
        return f"={ycol(yi-1)}24"
    row(21, "Beginning 期初", tlb_beg)

    row(22, "Mandatory Amort 强制摊销", lambda c, yi: f"=MIN({c}21,{A}!$B$17*{A}!$B$28)")
    def tlb_sweep(col, yi):
        return f"=MIN({col}21-{col}22,MAX(0,{col}8-{col}12-{col}17-{col}22)*{A}!$B$29)"
    row(23, "Cash Sweep 现金清偿", tlb_sweep)
    row(24, "Ending 期末", lambda c, yi: f"={c}21-{c}22-{c}23")

    # === Senior Notes (bullet) ===
    sec5 = ws.cell(row=26, column=1, value="Senior Notes 次优先票据")
    style_subheader(sec5)
    for col in range(2, 2 + n + 1):
        style_subheader(ws.cell(row=26, column=col))

    def sn_beg(col, yi):
        if yi == 0:
            return f"={A}!B18"
        return f"={ycol(yi-1)}28"
    row(27, "Beginning 期初", sn_beg)
    row(28, "Ending 期末", lambda c, yi: f"={c}27")

    # === Subordinated (bullet) ===
    sec6 = ws.cell(row=30, column=1, value="Subordinated Debt 次级债务")
    style_subheader(sec6)
    for col in range(2, 2 + n + 1):
        style_subheader(ws.cell(row=30, column=col))

    def sub_beg(col, yi):
        if yi == 0:
            return f"={A}!B19"
        return f"={ycol(yi-1)}32"
    row(31, "Beginning 期初", sub_beg)
    row(32, "Ending 期末", lambda c, yi: f"={c}31")

    # === Totals ===
    row(34, "Total Debt 总债务", lambda c, yi: f"={c}13+{c}18+{c}24+{c}28+{c}32", total=True)

    def cash_end(col, yi):
        return f"=MAX({col}7,{col}5+{col}6-{col}12-{col}17-{col}22-{col}23)"
    row(35, "Cash Ending 期末现金", cash_end, total=True)

    set_col_widths(ws, {"A": 36, "B": 14})
    for i in range(n):
        ws.column_dimensions[get_column_letter(3 + i)].width = 14
    ws.sheet_view.showGridLines = False


def _write_income_statement(wb, p: LBOParams):
    ws = wb.create_sheet("Income Statement")
    ws.cell(row=1, column=1, value="Income Statement 利润表").font = S.TITLE_FONT
    n = 5
    write_header_row(ws, 3, ["项目 Item", "LTM"] + [f"Year {i}" for i in range(1, n + 1)])

    A = "Assumptions"
    DS = "'Debt Schedule'"

    def row(r, label, formula_func, fmt=NUM_FMT, total=False, indent=0):
        c = ws.cell(row=r, column=1, value=label)
        style_total(c) if total else style_label(c, indent)
        for yi in range(n + 1):
            col = ycol(yi)
            f = formula_func(col, yi)
            cell = ws.cell(row=r, column=2 + yi)
            if f is not None:
                cell.value = f
            if total:
                style_total(cell, fmt)
            else:
                style_formula(cell, fmt, link="!" in (f or ""))

    def rev(col, yi):
        if yi == 0:
            return f"={A}!B5"
        return f"={ycol(yi-1)}5*(1+{A}!{col}32)"
    row(5, "Revenue 营业收入", rev)

    def growth(col, yi):
        return None if yi == 0 else f"={col}5/{ycol(yi-1)}5-1"
    row(6, "Revenue Growth %", growth, PCT_FMT)

    def ebitda(col, yi):
        if yi == 0:
            return f"={A}!B6"
        return f"={col}5*{A}!{col}33"
    row(7, "EBITDA", ebitda, total=True)

    def ebitda_m(col, yi):
        return None if yi == 0 else f"={col}7/{col}5"
    row(8, "EBITDA Margin %", ebitda_m, PCT_FMT)

    row(9, "D&A 折旧摊销", lambda c, yi: f"={c}5*{A}!$B$34")
    row(10, "EBIT 营业利润", lambda c, yi: f"={c}7-{c}9", total=True)

    row(12, "Interest - Revolver 利息(循环)", lambda c, yi: f"={A}!$B$21*{DS}!{c}11")
    row(13, "Interest - TL A 利息(TLA)", lambda c, yi: f"={A}!$B$22*{DS}!{c}16")
    row(14, "Interest - TL B 利息(TLB)", lambda c, yi: f"={A}!$B$23*{DS}!{c}21")
    row(15, "Interest - Sr Notes 利息(优先票据)", lambda c, yi: f"={A}!$B$24*{DS}!{c}27")
    row(16, "Interest - Sub 利息(次级)", lambda c, yi: f"={A}!$B$25*{DS}!{c}31")
    row(17, "Total Interest 利息合计", lambda c, yi: f"=SUM({c}12:{c}16)", total=True)
    row(18, "Interest Income 利息收入", lambda c, yi: f"={A}!$B$41*{DS}!{c}5")
    row(19, "EBT 税前利润", lambda c, yi: f"={c}10-{c}17+{c}18", total=True)
    row(20, "Taxes 所得税", lambda c, yi: f"=MAX(0,{c}19*{A}!$B$37)")
    row(21, "Net Income 净利润", lambda c, yi: f"={c}19-{c}20", total=True)

    # ========== 自定义行项 Custom Line Items ==========
    if p.custom_items:
        # Section header
        sec = ws.cell(row=23, column=1, value="自定义行项 Custom Line Items")
        style_subheader(sec)
        for col_idx in range(2, 2 + n + 1):
            style_subheader(ws.cell(row=23, column=col_idx))

        DS = "'Debt Schedule'"

        def build_cell_map(col):
            """构建变量名 → Excel 单元格引用 的映射。"""
            return {
                "revenue": f"{col}5",
                "ebitda": f"{col}7",
                "da": f"{col}9",
                "ebit": f"{col}10",
                "interest_revolver": f"{col}12",
                "interest_tla": f"{col}13",
                "interest_tlb": f"{col}14",
                "interest_sn": f"{col}15",
                "interest_sub": f"{col}16",
                "total_interest": f"{col}17",
                "interest_income": f"{col}18",
                "ebt": f"{col}19",
                "taxes": f"{col}20",
                "net_income": f"{col}21",
                "total_debt": f"{DS}!{col}34",
                "cash": f"{DS}!{col}35",
                "revolver": f"{DS}!{col}13",
                "tla": f"{DS}!{col}18",
                "tlb": f"{DS}!{col}24",
                "senior_notes": f"{DS}!{col}28",
                "subordinated_debt": f"{DS}!{col}32",
                "cfads": f"'Cash Flow (CFADS)'!{col}13",
            }

        for idx, ci in enumerate(p.custom_items):
            r = 24 + idx
            c = ws.cell(row=r, column=1, value=ci.name)
            style_label(c, 0)
            fmt = PCT_FMT if ci.format == "percent" else NUM_FMT
            for yi in range(n + 1):
                col = ycol(yi)
                cell_map = build_cell_map(col)
                try:
                    formula = expression_to_excel(ci.formula, cell_map)
                except Exception as e:
                    formula = f'="#ERR: {str(e)[:30]}"'
                cell = ws.cell(row=r, column=2 + yi, value=formula)
                style_formula(cell, fmt, link=True)

    set_col_widths(ws, {"A": 38, "B": 14})
    for i in range(n):
        ws.column_dimensions[get_column_letter(3 + i)].width = 14
    ws.sheet_view.showGridLines = False


def _write_cfads(wb, p: LBOParams):
    ws = wb.create_sheet("Cash Flow (CFADS)")
    ws.cell(row=1, column=1, value="Cash Flow (CFADS) 偿债可用现金流").font = S.TITLE_FONT
    n = 5
    write_header_row(ws, 3, ["项目 Item", "LTM"] + [f"Year {i}" for i in range(1, n + 1)])

    A = "Assumptions"
    IS = "'Income Statement'"

    def row(r, label, formula_func, fmt=NUM_FMT, total=False, indent=0):
        c = ws.cell(row=r, column=1, value=label)
        style_total(c) if total else style_label(c, indent)
        for yi in range(n + 1):
            col = ycol(yi)
            f = formula_func(col, yi)
            cell = ws.cell(row=r, column=2 + yi)
            if f is not None:
                cell.value = f
            if total:
                style_total(cell, fmt)
            else:
                style_formula(cell, fmt, link="!" in (f or ""))

    row(5, "Net Income 净利润", lambda c, yi: None if yi == 0 else f"={IS}!{c}21")
    row(7, "D&A 折旧摊销", lambda c, yi: None if yi == 0 else f"={IS}!{c}9")
    def dnwc(col, yi):
        return None if yi == 0 else f"=-{IS}!{col}5*{A}!$B$36"
    row(8, "Δ NWC 净营运资本变动", dnwc)
    row(9, "Cash from Operations 经营现金流", lambda c, yi: None if yi == 0 else f"={c}5+{c}7+{c}8", total=True)
    def capex(col, yi):
        return None if yi == 0 else f"=-{IS}!{col}5*{A}!$B$35"
    row(11, "CapEx 资本支出", capex)
    row(13, "CFADS 偿债可用现金流", lambda c, yi: None if yi == 0 else f"={c}9+{c}11", total=True)

    set_col_widths(ws, {"A": 38, "B": 14})
    for i in range(n):
        ws.column_dimensions[get_column_letter(3 + i)].width = 14
    ws.sheet_view.showGridLines = False


def _write_exit_returns(wb, p: LBOParams):
    ws = wb.create_sheet("Exit & Returns")
    ws.cell(row=1, column=1, value="Exit & Returns 退出与回报").font = S.TITLE_FONT
    write_header_row(ws, 3, ["项目 Item", "Value"])

    A = "Assumptions"
    IS = "'Income Statement'"
    DS = "'Debt Schedule'"
    exit_year = p.exit_year
    exit_col = ycol(exit_year)  # e.g. exit_year=5 -> G

    def line(row, label, formula, fmt=NUM_FMT, total=False):
        c = ws.cell(row=row, column=1, value=label)
        (style_total if total else style_label)(c)
        cell = ws.cell(row=row, column=2)
        cell.value = formula
        style_formula(cell, fmt, link="!" in (formula or ""))

    line(5, f"Exit Year ({exit_year}) EBITDA 退出年EBITDA", f"={IS}!{exit_col}7")
    line(6, "Exit EV/EBITDA 退出倍数", f"={A}!B39", MULT_FMT)
    line(7, "Exit Enterprise Value 退出企业价值", f"=B5*B6", total=True)
    line(8, "Less: Total Debt 减:总债务", f"=-{DS}!{exit_col}34")
    line(9, "Plus: Cash 加:现金", f"={DS}!{exit_col}35")
    line(10, "Equity Value at Exit 退出股权价值", f"=B7+B8+B9", total=True)
    line(12, "Sponsor Equity Invested 投入股本", f"={A}!B14")
    line(13, "MoIC (x) 投资倍数", f"=B10/B12", MULT_FMT, total=True)

    # Cash flow timeline for IRR
    sec = ws.cell(row=15, column=1, value="现金流时间轴 Cash Flow Timeline")
    style_subheader(sec)
    style_subheader(ws.cell(row=15, column=2))
    for yi in range(6):
        c = ws.cell(row=15, column=2 + yi, value=f"Year {yi}" if yi > 0 else "Year 0")
        style_subheader(c)

    # Row 16: cash flows
    # Year 0 (B16) = -Sponsor Equity
    c = ws.cell(row=16, column=1, value="Sponsor Cash Flows")
    style_label(c)
    for yi in range(6):
        col = ycol(yi)
        cell = ws.cell(row=16, column=2 + yi)
        if yi == 0:
            cell.value = f"=-{A}!B14"
        elif yi == exit_year:
            cell.value = "=B10"
        elif yi < exit_year:
            cell.value = "=0"
        else:
            cell.value = None  # after exit, blank
        style_formula(cell, NUM_FMT, link="!" in (cell.value or ""))

    # IRR
    c = ws.cell(row=18, column=1, value="IRR 内部收益率")
    style_total(c)
    irr_range = f"B16:{exit_col}16"
    cell = ws.cell(row=18, column=2, value=f"=IRR({irr_range})")
    style_total(cell, PCT_FMT)

    set_col_widths(ws, {"A": 40, "B": 16})
    for i in range(6):
        ws.column_dimensions[get_column_letter(2 + i)].width = 14
    ws.sheet_view.showGridLines = False


def _write_sensitivity(wb, p: LBOParams):
    ws = wb.create_sheet("Sensitivity")
    ws.cell(row=1, column=1, value="Sensitivity 敏感性分析").font = S.TITLE_FONT
    ws.cell(row=2, column=1, value="行=入场倍数，列=退出倍数；单元格=MoIC").font = S.SUBHEADER_FONT

    A = "Assumptions"
    IS = "'Income Statement'"
    DS = "'Debt Schedule'"
    exit_col = ycol(p.exit_year)

    entry_multiples = [p.entry_ev_ebitda - 1, p.entry_ev_ebitda - 0.5,
                       p.entry_ev_ebitda, p.entry_ev_ebitda + 0.5, p.entry_ev_ebitda + 1]
    exit_multiples = [p.exit_ev_ebitda - 1, p.exit_ev_ebitda - 0.5,
                      p.exit_ev_ebitda, p.exit_ev_ebitda + 0.5, p.exit_ev_ebitda + 1]

    # header
    ws.cell(row=4, column=1, value="Entry \\ Exit").font = S.HEADER_FONT
    ws.cell(row=4, column=1).fill = S.HEADER_FILL
    ws.cell(row=4, column=1).alignment = S.CENTER_ALIGN
    ws.cell(row=4, column=1).border = S.BORDER
    for j, m in enumerate(exit_multiples):
        c = ws.cell(row=4, column=2 + j, value=m)
        style_header(c)
        c.number_format = MULT_FMT

    for i, em in enumerate(entry_multiples):
        rc = ws.cell(row=5 + i, column=1, value=em)
        rc.font = S.HEADER_FONT
        rc.fill = S.HEADER_FILL
        rc.alignment = S.CENTER_ALIGN
        rc.border = S.BORDER
        rc.number_format = MULT_FMT

        for j, xm in enumerate(exit_multiples):
            # Entry EV = LTM EBITDA * entry multiple
            # Equity at entry = Entry EV - Net Debt + Cash = em*EBITDA - B9 + B10
            # Exit EV = exit_year EBITDA * exit multiple
            # Equity at exit = Exit EV - Debt + Cash (from Debt Schedule)
            # MoIC = Equity_exit / Equity_entry
            entry_equity = f"({A}!$B$6*{em}-{A}!$B$9+{A}!$B$10)"
            exit_ev = f"{IS}!{exit_col}7*{xm}"
            exit_equity = f"({exit_ev}-{DS}!{exit_col}34+{DS}!{exit_col}35)"
            moic = f"({exit_equity})/({entry_equity})"
            cell = ws.cell(row=5 + i, column=2 + j, value=f"={moic}")
            style_formula(cell, MULT_FMT)

    set_col_widths(ws, {"A": 12})
    for j in range(len(exit_multiples)):
        ws.column_dimensions[get_column_letter(2 + j)].width = 12
    ws.sheet_view.showGridLines = False


def build(p: LBOParams) -> Workbook:
    """构建 LBO 模型 workbook。"""
    wb = Workbook()
    wb.remove(wb.active)
    _write_assumptions(wb, p)
    _write_sources_uses(wb, p)
    _write_debt_schedule(wb, p)
    _write_income_statement(wb, p)
    _write_cfads(wb, p)
    _write_exit_returns(wb, p)
    _write_sensitivity(wb, p)
    wb.active = wb.sheetnames.index("Assumptions")
    return wb
