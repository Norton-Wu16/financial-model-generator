"""三表联动模型 Excel 构建器。

Sheet 结构（6 sheets）:
1. Assumptions          - 全部输入参数
2. Income Statement     - 利润表（收入→净利润）
3. Balance Sheet        - 资产负债表（资产/负债/权益，平衡）
4. Cash Flow Statement  - 现金流量表（经营/投资/筹资）
5. Supporting Schedules - PP&E 滚动 + 债务滚动
6. Dashboard            - 关键指标摘要 + 平衡校验

跨表链接闭环:
  IS 净利润 → BS 留存收益
  Schedules D&A → IS & CFS
  BS 营运资本变动 → CFS
  BS PP&E 变动 → CFS 投资
  CFS 净现金变动 → BS 期末现金（平衡项）
"""
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from backend import styles as S
from backend.styles import (
    style_input, style_formula, style_total, style_header,
    style_subheader, style_label, style_title, set_col_widths,
    write_header_row, NUM_FMT, PCT_FMT, NUM_FMT_INT,
)
from backend.schemas import ThreeStatementParams
from backend.custom_items import expression_to_excel


# 列映射: Year 0 = B, Year 1..N = C, D, ...
def ycol(year_index):
    """year_index: 0=Year0(B), 1=Year1(C)..."""
    return get_column_letter(2 + year_index)


def _write_assumptions(wb, p: ThreeStatementParams):
    ws = wb.create_sheet("Assumptions")
    ws.cell(row=1, column=1, value="Assumptions 假设参数").font = S.TITLE_FONT
    ws.cell(row=1, column=1).alignment = S.LEFT_ALIGN

    n = p.projection_years
    headers = ["参数 Parameter", "Year 0"] + [f"Year {i}" for i in range(1, n + 1)]
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

    # 运营假设
    section(4, "运营假设 Operating Assumptions")
    inp(5, "Revenue (base) 基期收入", p.revenue_y0)
    style_label(ws.cell(row=6, column=1, value="Revenue Growth % 收入增长率"))
    for i, g in enumerate([p.rev_growth_y1, p.rev_growth_y2, p.rev_growth_y3,
                          p.rev_growth_y4, p.rev_growth_y5][:n]):
        inp(6, "Revenue Growth % 收入增长率", g, PCT_FMT, year_col=i + 1)
    inp(7, "COGS % of Revenue 成本率", p.cogs_pct, PCT_FMT)
    inp(8, "SG&A % of Revenue 销管费用率", p.sga_pct, PCT_FMT)
    inp(9, "R&D % of Revenue 研发费用率", p.rd_pct, PCT_FMT)
    inp(10, "D&A % of Revenue 折旧摊销率", p.da_pct, PCT_FMT)
    inp(11, "Interest Rate on Debt 债务利率", p.interest_rate, PCT_FMT)
    inp(12, "Tax Rate 税率", p.tax_rate, PCT_FMT)
    inp(13, "Dividend Payout % 股利分配率", p.dividend_pct, PCT_FMT)
    inp(14, "CapEx % of Revenue 资本支出率", p.capex_pct, PCT_FMT)

    # 营运资本天数
    section(16, "营运资本天数 Working Capital Days")
    inp(17, "DSO 应收天数", p.dso, NUM_FMT_INT)
    inp(18, "DIO 库存天数", p.dio, NUM_FMT_INT)
    inp(19, "DPO 应付天数", p.dpo, NUM_FMT_INT)
    inp(20, "Accrued Days 应计天数", p.accrued_days, NUM_FMT_INT)

    # 期初资产负债表
    section(22, "期初资产负债表 Beginning Balance Sheet")
    inp(23, "Beginning Cash 期初现金", p.beg_cash)
    inp(24, "Beginning AR 期初应收", p.beg_ar)
    inp(25, "Beginning Inventory 期初库存", p.beg_inventory)
    inp(26, "Beginning PP&E 期初固定资产", p.beg_ppe)
    inp(27, "Beginning AP 期初应付", p.beg_ap)
    inp(28, "Beginning Accrued 期初应计", p.beg_accrued)
    inp(29, "Beginning Debt 期初债务", p.beg_debt)
    inp(30, "Common Stock 普通股", p.common_stock)
    inp(31, "Beginning Retained Earnings 期初留存收益", p.beg_retn_earn)

    # 其他
    section(33, "其他 Other")
    style_label(ws.cell(row=34, column=1, value="New Debt Issuance (annual) 新增债务"))
    for i in range(n):
        cell = ws.cell(row=34, column=3 + i, value=p.new_debt_issuance)
        style_input(cell, NUM_FMT)
    style_label(ws.cell(row=34, column=2, value=""))

    set_col_widths(ws, {"A": 42, "B": 14})
    for i in range(n):
        ws.column_dimensions[get_column_letter(3 + i)].width = 14
    ws.sheet_view.showGridLines = False


def _write_income_statement(wb, p: ThreeStatementParams):
    ws = wb.create_sheet("Income Statement")
    ws.cell(row=1, column=1, value="Income Statement 利润表").font = S.TITLE_FONT
    n = p.projection_years
    write_header_row(ws, 3, ["项目 Item", "Year 0"] + [f"Year {i}" for i in range(1, n + 1)])

    A = "Assumptions"
    SCH = "'Supporting Schedules'"

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
        prev = ycol(yi - 1)
        return f"={prev}5*(1+{A}!{col}6)"
    row(5, "Revenue 营业收入", rev)

    def growth(col, yi):
        if yi == 0:
            return None
        return f"={col}5/{ycol(yi-1)}5-1"
    row(6, "Revenue Growth % 增长率", growth, PCT_FMT)

    row(7, "COGS 营业成本", lambda c, yi: f"={c}5*{A}!$B$7")
    row(8, "Gross Profit 毛利", lambda c, yi: f"={c}5-{c}7", total=True)
    row(9, "SG&A 销管费用", lambda c, yi: f"={c}5*{A}!$B$8")
    row(10, "R&D 研发费用", lambda c, yi: f"={c}5*{A}!$B$9")
    row(11, "EBITDA", lambda c, yi: f"={c}8-{c}9-{c}10", total=True)

    def da(col, yi):
        return f"={col}5*{A}!$B$10" if yi == 0 else f"={SCH}!{col}6"
    row(12, "D&A 折旧摊销", da)

    row(13, "EBIT 营业利润", lambda c, yi: f"={c}11-{c}12", total=True)

    def interest(col, yi):
        if yi == 0:
            return f"={A}!$B$11*{A}!B29"
        return f"={SCH}!{col}17"
    row(14, "Interest Expense 利息费用", interest)

    row(15, "EBT 税前利润", lambda c, yi: f"={c}13-{c}14", total=True)
    row(16, "Taxes 所得税", lambda c, yi: f"=MAX(0,{c}15*{A}!$B$12)")
    row(17, "Net Income 净利润", lambda c, yi: f"={c}15-{c}16", total=True)
    row(18, "Dividends 股利", lambda c, yi: f"={c}17*{A}!$B$13")

    # ========== 自定义行项 Custom Line Items ==========
    if p.custom_items:
        # Section header
        sec = ws.cell(row=20, column=1, value="自定义行项 Custom Line Items")
        style_subheader(sec)
        for col_idx in range(2, 2 + n + 1):
            style_subheader(ws.cell(row=20, column=col_idx))

        BS = "'Balance Sheet'"
        SCH = "'Supporting Schedules'"

        def build_cell_map(col):
            """构建变量名 → Excel 单元格引用 的映射。"""
            return {
                "revenue": f"{col}5",
                "cogs": f"{col}7",
                "gross_profit": f"{col}8",
                "sga": f"{col}9",
                "rd": f"{col}10",
                "ebitda": f"{col}11",
                "da": f"{col}12",
                "ebit": f"{col}13",
                "interest": f"{col}14",
                "ebt": f"{col}15",
                "taxes": f"{col}16",
                "net_income": f"{col}17",
                "dividends": f"{col}18",
                "cash": f"{BS}!{col}5",
                "ar": f"{BS}!{col}6",
                "inventory": f"{BS}!{col}7",
                "ppe": f"{BS}!{col}8",
                "total_assets": f"{BS}!{col}9",
                "ap": f"{BS}!{col}13",
                "accrued": f"{BS}!{col}14",
                "debt": f"{BS}!{col}15",
                "total_liabilities": f"{BS}!{col}16",
                "common_stock": f"{BS}!{col}20",
                "retained_earnings": f"{BS}!{col}21",
                "total_equity": f"{BS}!{col}22",
                "total_le": f"{BS}!{col}25",
                "capex": f"{SCH}!{col}8",
                "new_debt": f"{SCH}!{col}19",
            }

        for idx, ci in enumerate(p.custom_items):
            r = 21 + idx
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

    set_col_widths(ws, {"A": 34, "B": 14})
    for i in range(n):
        ws.column_dimensions[get_column_letter(3 + i)].width = 14
    ws.sheet_view.showGridLines = False


def _write_supporting_schedules(wb, p: ThreeStatementParams):
    ws = wb.create_sheet("Supporting Schedules")
    ws.cell(row=1, column=1, value="Supporting Schedules 辅助表").font = S.TITLE_FONT
    n = p.projection_years
    write_header_row(ws, 3, ["项目 Item", "Year 0"] + [f"Year {i}" for i in range(1, n + 1)])

    A = "Assumptions"
    BS = "'Balance Sheet'"
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

    # PP&E Roll-forward
    sec = ws.cell(row=4, column=1, value="PP&E 滚动 PP&E Roll-forward")
    style_subheader(sec)
    for col in range(2, 2 + n + 1):
        style_subheader(ws.cell(row=4, column=col))

    row(6, "D&A 折旧摊销", lambda c, yi: f"={IS}!{c}5*{A}!$B$10")

    def begppe(col, yi):
        if yi == 0:
            return f"={A}!B26"
        return f"={ycol(yi-1)}10"
    row(7, "Beginning PP&E 期初固定资产", begppe)

    row(8, "CapEx 资本支出", lambda c, yi: f"={IS}!{c}5*{A}!$B$14")
    row(9, "Less: D&A 减:折旧", lambda c, yi: f"=-{c}6")

    def endppe(col, yi):
        return f"={col}7" if yi == 0 else f"={col}7+{col}8+{col}9"
    row(10, "Ending PP&E 期末固定资产", endppe, total=True)

    # Debt Roll-forward
    sec2 = ws.cell(row=15, column=1, value="债务滚动 Debt Roll-forward")
    style_subheader(sec2)
    for col in range(2, 2 + n + 1):
        style_subheader(ws.cell(row=15, column=col))

    def interest(col, yi):
        if yi == 0:
            return f"={A}!$B$11*{A}!B29"
        return f"={A}!$B$11*{BS}!{ycol(yi-1)}15"
    row(17, "Interest Expense 利息费用", interest)

    def begdebt(col, yi):
        if yi == 0:
            return f"={A}!B29"
        return f"={ycol(yi-1)}20"
    row(18, "Beginning Debt 期初债务", begdebt)

    def newdebt(col, yi):
        return f"=0" if yi == 0 else f"={A}!{col}34"
    row(19, "New Debt Issuance 新增债务", newdebt)

    def enddebt(col, yi):
        return f"={col}18" if yi == 0 else f"={col}18+{col}19"
    row(20, "Ending Debt 期末债务", enddebt, total=True)

    set_col_widths(ws, {"A": 34, "B": 14})
    for i in range(n):
        ws.column_dimensions[get_column_letter(3 + i)].width = 14
    ws.sheet_view.showGridLines = False


def _write_balance_sheet(wb, p: ThreeStatementParams):
    ws = wb.create_sheet("Balance Sheet")
    ws.cell(row=1, column=1, value="Balance Sheet 资产负债表").font = S.TITLE_FONT
    n = p.projection_years
    write_header_row(ws, 3, ["项目 Item", "Year 0"] + [f"Year {i}" for i in range(1, n + 1)])

    A = "Assumptions"
    IS = "'Income Statement'"
    CFS = "'Cash Flow Statement'"
    SCH = "'Supporting Schedules'"

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

    # Assets
    sec = ws.cell(row=4, column=1, value="资产 Assets")
    style_subheader(sec)
    for col in range(2, 2 + n + 1):
        style_subheader(ws.cell(row=4, column=col))

    def cash(col, yi):
        if yi == 0:
            return f"={A}!B23"
        return f"={ycol(yi-1)}5+{CFS}!{col}23"
    row(5, "Cash 现金", cash)

    def ar(col, yi):
        if yi == 0:
            return f"={A}!B24"
        return f"={IS}!{col}5/365*{A}!$B$17"
    row(6, "Accounts Receivable 应收账款", ar)

    def inv(col, yi):
        if yi == 0:
            return f"={A}!B25"
        return f"={IS}!{col}7/365*{A}!$B$18"
    row(7, "Inventory 存货", inv)

    def ppe(col, yi):
        if yi == 0:
            return f"={A}!B26"
        return f"={SCH}!{col}10"
    row(8, "PP&E, net 固定资产净额", ppe)

    row(9, "Total Assets 总资产", lambda c, yi: f"=SUM({c}5:{c}8)", total=True)

    # Liabilities
    sec2 = ws.cell(row=12, column=1, value="负债 Liabilities")
    style_subheader(sec2)
    for col in range(2, 2 + n + 1):
        style_subheader(ws.cell(row=12, column=col))

    def ap(col, yi):
        if yi == 0:
            return f"={A}!B27"
        return f"={IS}!{col}7/365*{A}!$B$19"
    row(13, "Accounts Payable 应付账款", ap)

    def acc(col, yi):
        if yi == 0:
            return f"={A}!B28"
        return f"={IS}!{col}5/365*{A}!$B$20"
    row(14, "Accrued Expenses 应计费用", acc)

    def debt(col, yi):
        if yi == 0:
            return f"={A}!B29"
        return f"={SCH}!{col}20"
    row(15, "Debt 债务", debt)

    row(16, "Total Liabilities 总负债", lambda c, yi: f"=SUM({c}13:{c}15)", total=True)

    # Equity
    sec3 = ws.cell(row=19, column=1, value="权益 Equity")
    style_subheader(sec3)
    for col in range(2, 2 + n + 1):
        style_subheader(ws.cell(row=19, column=col))

    row(20, "Common Stock 普通股", lambda c, yi: f"={A}!$B$30")

    def re(col, yi):
        if yi == 0:
            return f"={A}!B31"
        return f"={ycol(yi-1)}21+{IS}!{col}17-{IS}!{col}18"
    row(21, "Retained Earnings 留存收益", re)

    row(22, "Total Equity 总权益", lambda c, yi: f"={c}20+{c}21", total=True)
    row(25, "Total Liabilities + Equity 总负债权益", lambda c, yi: f"={c}16+{c}22", total=True)

    # Balance check
    ws.cell(row=27, column=1, value="平衡校验 Balance Check").font = S.SUBHEADER_FONT
    style_label(ws.cell(row=27, column=1, value="平衡校验 Balance Check"))
    for yi in range(n + 1):
        col = ycol(yi)
        cell = ws.cell(row=27, column=2 + yi,
                       value=f'=IF(ROUND({col}9-{col}25,2)=0,"BALANCED","OUT OF BALANCE")')
        cell.font = S.TOTAL_FONT
        cell.alignment = S.CENTER_ALIGN
        cell.border = S.BORDER

    set_col_widths(ws, {"A": 34, "B": 14})
    for i in range(n):
        ws.column_dimensions[get_column_letter(3 + i)].width = 14
    ws.sheet_view.showGridLines = False


def _write_cash_flow(wb, p: ThreeStatementParams):
    ws = wb.create_sheet("Cash Flow Statement")
    ws.cell(row=1, column=1, value="Cash Flow Statement 现金流量表").font = S.TITLE_FONT
    n = p.projection_years
    write_header_row(ws, 3, ["项目 Item", "Year 0"] + [f"Year {i}" for i in range(1, n + 1)])

    A = "Assumptions"
    IS = "'Income Statement'"
    BS = "'Balance Sheet'"
    SCH = "'Supporting Schedules'"

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

    # Operating
    sec = ws.cell(row=4, column=1, value="经营活动 Operating Activities")
    style_subheader(sec)
    for col in range(2, 2 + n + 1):
        style_subheader(ws.cell(row=4, column=col))

    def zero(col, yi):
        return None if yi == 0 else f"={IS}!{col}17"
    row(6, "Net Income 净利润", zero)

    def da(col, yi):
        return None if yi == 0 else f"={SCH}!{col}6"
    row(7, "D&A 折旧摊销", da)

    def dar(col, yi):
        if yi == 0:
            return None
        return f"=-({BS}!{col}6-{BS}!{ycol(yi-1)}6)"
    row(8, "Change in AR 应收变动", dar)

    def dinv(col, yi):
        if yi == 0:
            return None
        return f"=-({BS}!{col}7-{BS}!{ycol(yi-1)}7)"
    row(9, "Change in Inventory 存货变动", dinv)

    def dap(col, yi):
        if yi == 0:
            return None
        return f"={BS}!{col}13-{BS}!{ycol(yi-1)}13"
    row(10, "Change in AP 应付变动", dap)

    def dacc(col, yi):
        if yi == 0:
            return None
        return f"={BS}!{col}14-{BS}!{ycol(yi-1)}14"
    row(11, "Change in Accrued 应计变动", dacc)

    def cfo(col, yi):
        return None if yi == 0 else f"=SUM({col}6:{col}11)"
    row(12, "Cash from Operations 经营现金流", cfo, total=True)

    # Investing
    sec2 = ws.cell(row=14, column=1, value="投资活动 Investing Activities")
    style_subheader(sec2)
    for col in range(2, 2 + n + 1):
        style_subheader(ws.cell(row=14, column=col))

    def capex(col, yi):
        return None if yi == 0 else f"=-{IS}!{col}5*{A}!$B$14"
    row(15, "CapEx 资本支出", capex)
    row(16, "Cash from Investing 投资现金流", lambda c, yi: None if yi == 0 else f"={c}15", total=True)

    # Financing
    sec3 = ws.cell(row=18, column=1, value="筹资活动 Financing Activities")
    style_subheader(sec3)
    for col in range(2, 2 + n + 1):
        style_subheader(ws.cell(row=18, column=col))

    def ndebt(col, yi):
        if yi == 0:
            return None
        return f"={BS}!{col}15-{BS}!{ycol(yi-1)}15"
    row(19, "Net Debt Change 债务净变动", ndebt)

    def divpaid(col, yi):
        return None if yi == 0 else f"=-{IS}!{col}18"
    row(20, "Dividends Paid 支付股利", divpaid)

    row(21, "Cash from Financing 筹资现金流", lambda c, yi: None if yi == 0 else f"=SUM({c}19:{c}20)", total=True)
    row(23, "Net Change in Cash 现金净变动", lambda c, yi: None if yi == 0 else f"={c}12+{c}16+{c}21", total=True)

    set_col_widths(ws, {"A": 34, "B": 14})
    for i in range(n):
        ws.column_dimensions[get_column_letter(3 + i)].width = 14
    ws.sheet_view.showGridLines = False


def _write_dashboard(wb, p: ThreeStatementParams):
    ws = wb.create_sheet("Dashboard")
    ws.cell(row=1, column=1, value="Dashboard 关键指标摘要").font = S.TITLE_FONT
    n = p.projection_years
    write_header_row(ws, 3, ["指标 Metric", "Year 0"] + [f"Year {i}" for i in range(1, n + 1)])

    IS = "'Income Statement'"
    BS = "'Balance Sheet'"

    def row(r, label, formula_func, fmt=NUM_FMT, total=False):
        c = ws.cell(row=r, column=1, value=label)
        (style_total if total else style_label)(c)
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

    row(5, "Revenue 营业收入", lambda c, yi: f"={IS}!{c}5")
    row(6, "EBITDA", lambda c, yi: f"={IS}!{c}11", total=True)
    row(7, "EBIT 营业利润", lambda c, yi: f"={IS}!{c}13")
    row(8, "Net Income 净利润", lambda c, yi: f"={IS}!{c}17", total=True)
    row(9, "Total Assets 总资产", lambda c, yi: f"={BS}!{c}9")
    row(10, "Total Debt 总债务", lambda c, yi: f"={BS}!{c}15")
    row(11, "Cash 现金", lambda c, yi: f"={BS}!{c}5")
    row(12, "Retained Earnings 留存收益", lambda c, yi: f"={BS}!{c}21")

    def ebitda_margin(col, yi):
        return None if yi == 0 else f"={IS}!{col}11/{IS}!{col}5"
    row(14, "EBITDA Margin %", ebitda_margin, PCT_FMT)
    def ni_margin(col, yi):
        return None if yi == 0 else f"={IS}!{col}17/{IS}!{col}5"
    row(15, "Net Margin %", ni_margin, PCT_FMT)

    set_col_widths(ws, {"A": 34, "B": 14})
    for i in range(n):
        ws.column_dimensions[get_column_letter(3 + i)].width = 14
    ws.sheet_view.showGridLines = False


def build(p: ThreeStatementParams) -> Workbook:
    """构建三表联动模型 workbook。"""
    wb = Workbook()
    wb.remove(wb.active)
    _write_assumptions(wb, p)
    _write_income_statement(wb, p)
    _write_supporting_schedules(wb, p)
    _write_balance_sheet(wb, p)
    _write_cash_flow(wb, p)
    _write_dashboard(wb, p)
    wb.active = wb.sheetnames.index("Assumptions")
    return wb
