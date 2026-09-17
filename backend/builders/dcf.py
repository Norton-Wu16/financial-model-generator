"""DCF 现金流折现模型 Excel 构建器。

Sheet 结构（6 sheets）:
1. Assumptions       - 运营/WACC/终值/EV桥参数
2. Operating Model   - 收入→EBITDA→EBIT→NOPAT→Unlevered FCF
3. WACC              - CAPM 成本 + 加权
4. Terminal Value    - Gordon 增长 & 退出倍数
5. DCF Valuation      - PV 汇总 + EV→Equity 桥 + 每股价值
6. Sensitivity       - WACC × g 敏感性矩阵

跨表链接:
  Assumptions → OpModel（收入/利润率驱动）
  Assumptions → WACC（CAPM 参数）
  OpModel + WACC + Assumptions → Terminal Value
  OpModel + WACC + TV → DCF Valuation（PV 汇总）
"""
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from backend import styles as S
from backend.styles import (
    style_input, style_formula, style_total, style_header,
    style_subheader, style_label, style_title, set_col_widths,
    write_header_row, NUM_FMT, PCT_FMT, NUM_FMT_INT, MULT_FMT, PRICE_FMT,
)
from backend.schemas import DCFParams
from backend.custom_items import expression_to_excel


def ycol(year_index):
    return get_column_letter(2 + year_index)


def _write_assumptions(wb, p: DCFParams):
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

    section(4, "运营假设 Operating Assumptions")
    inp(5, "Revenue (base) 基期收入", p.revenue_y0)
    style_label(ws.cell(row=6, column=1, value="Revenue Growth % 收入增长率"))
    for i, g in enumerate([p.rev_growth_y1, p.rev_growth_y2, p.rev_growth_y3,
                          p.rev_growth_y4, p.rev_growth_y5][:n]):
        inp(6, "Revenue Growth %", g, PCT_FMT, year_col=i + 1)
    style_label(ws.cell(row=7, column=1, value="EBITDA Margin %"))
    for i, m in enumerate([p.ebitda_margin_y1, p.ebitda_margin_y2, p.ebitda_margin_y3,
                           p.ebitda_margin_y4, p.ebitda_margin_y5][:n]):
        inp(7, "EBITDA Margin %", m, PCT_FMT, year_col=i + 1)
    inp(8, "D&A % of Revenue 折旧摊销率", p.da_pct, PCT_FMT)
    inp(9, "CapEx % of Revenue 资本支出率", p.capex_pct, PCT_FMT)
    inp(10, "NWC % of Revenue 净营运资本率", p.nwc_pct, PCT_FMT)
    inp(11, "Tax Rate 税率", p.tax_rate, PCT_FMT)

    section(13, "WACC 资本成本")
    inp(14, "Risk-free Rate 无风险利率", p.risk_free_rate, PCT_FMT)
    inp(15, "Equity Risk Premium 股权风险溢价", p.equity_risk_premium, PCT_FMT)
    inp(16, "Beta", p.beta)
    inp(17, "Pre-tax Cost of Debt 税前债务成本", p.pre_tax_cost_of_debt, PCT_FMT)
    inp(18, "Target Debt Weight 目标债务权重", p.debt_weight, PCT_FMT)

    section(20, "终值 Terminal Value")
    inp(21, "Terminal Growth Rate g 终值增长率", p.terminal_growth, PCT_FMT)
    inp(22, "Exit EV/EBITDA 退出倍数", p.exit_multiple, MULT_FMT)
    inp(23, "TV Method 终值方法", p.tv_method)
    ws.cell(row=23, column=2).alignment = S.CENTER_ALIGN

    section(25, "EV→Equity 桥")
    inp(26, "Total Debt 总债务", p.net_debt)
    inp(27, "Cash 现金", p.cash)
    inp(28, "Minority Interest 少数股东权益", p.minority_interest)
    inp(29, "Preferred Stock 优先股", p.preferred_stock)
    inp(30, "Stock-Based Compensation 股权激励", p.stock_based_comp)
    inp(31, "Pension Deficit 养老金缺口", p.pension_deficit)
    inp(32, "Shares Outstanding (M) 流通股本", p.shares_outstanding)
    inp(33, "Current Share Price 当前股价", p.current_price, PRICE_FMT)
    inp(34, "Valuation Timing 估值时点 (mid/end)", p.valuation_timing)
    ws.cell(row=34, column=2).alignment = S.CENTER_ALIGN

    set_col_widths(ws, {"A": 44, "B": 14})
    for i in range(n):
        ws.column_dimensions[get_column_letter(3 + i)].width = 14
    ws.sheet_view.showGridLines = False


def _write_operating_model(wb, p: DCFParams):
    ws = wb.create_sheet("Operating Model")
    ws.cell(row=1, column=1, value="Operating Model 运营模型").font = S.TITLE_FONT
    n = p.projection_years
    write_header_row(ws, 3, ["项目 Item", "Year 0"] + [f"Year {i}" for i in range(1, n + 1)])

    A = "Assumptions"

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
        return f"={ycol(yi-1)}5*(1+{A}!{col}6)"
    row(5, "Revenue 营业收入", rev)

    def growth(col, yi):
        return None if yi == 0 else f"={col}5/{ycol(yi-1)}5-1"
    row(6, "Revenue Growth % 增长率", growth, PCT_FMT)

    def ebitda(col, yi):
        return f"={col}5*{A}!{col}7" if yi > 0 else f"={col}5*{A}!C7"
    row(7, "EBITDA", ebitda, total=True)

    def ebitda_margin(col, yi):
        return None if yi == 0 else f"={col}7/{col}5"
    row(8, "EBITDA Margin %", ebitda_margin, PCT_FMT)

    row(9, "D&A 折旧摊销", lambda c, yi: f"={c}5*{A}!$B$8")
    row(10, "EBIT 营业利润", lambda c, yi: f"={c}7-{c}9", total=True)

    def ebit_margin(col, yi):
        return None if yi == 0 else f"={col}10/{col}5"
    row(11, "EBIT Margin %", ebit_margin, PCT_FMT)

    row(12, "NOPAT 税后营业利润", lambda c, yi: f"={c}10*(1-{A}!$B$11)", total=True)

    def dnwc(col, yi):
        return None if yi == 0 else f"=({col}5-{ycol(yi-1)}5)*{A}!$B$10"
    row(13, "Δ NWC 净营运资本变动", dnwc)

    row(14, "CapEx 资本支出", lambda c, yi: f"={c}5*{A}!$B$9")

    def ufcf(col, yi):
        return None if yi == 0 else f"={col}12+{col}9-{col}13-{col}14"
    row(15, "Unlevered FCF 无杠杆自由现金流", ufcf, total=True)

    # ========== 自定义行项 Custom Line Items ==========
    if p.custom_items:
        # Section header
        sec = ws.cell(row=17, column=1, value="自定义行项 Custom Line Items")
        style_subheader(sec)
        for col_idx in range(2, 2 + n + 1):
            style_subheader(ws.cell(row=17, column=col_idx))

        W = "WACC"
        TV = "'Terminal Value'"
        DV = "'DCF Valuation'"
        A2 = "Assumptions"

        def build_cell_map(col):
            """构建变量名 → Excel 单元格引用 的映射。"""
            return {
                "revenue": f"{col}5",
                "ebitda": f"{col}7",
                "da": f"{col}9",
                "ebit": f"{col}10",
                "nopat": f"{col}12",
                "delta_nwc": f"{col}13",
                "capex": f"{col}14",
                "ufcf": f"{col}15",
                "wacc": f"{W}!$B$14",
                "cost_of_equity": f"{W}!$B$8",
                "after_tax_kd": f"{W}!$B$11",
                "tv_gordon": f"{TV}!$B$5",
                "tv_exit": f"{TV}!$B$6",
                "selected_tv": f"{TV}!$B$7",
                "pv_tv": f"{TV}!$B$9",
                "sum_pv_fcf": f"{DV}!$B$10",
                "enterprise_value": f"{DV}!$B$12",
                "equity_value": f"{DV}!$B$22",
                "implied_price": f"{DV}!$B$24",
                "current_price": f"{A2}!$B$33",
                "shares_outstanding": f"{A2}!$B$32",
                "net_debt": f"{A2}!$B$26",
                "cash": f"{A2}!$B$27",
            }

        for idx, ci in enumerate(p.custom_items):
            r = 18 + idx
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


def _write_wacc(wb, p: DCFParams):
    ws = wb.create_sheet("WACC")
    ws.cell(row=1, column=1, value="WACC 加权平均资本成本").font = S.TITLE_FONT
    write_header_row(ws, 3, ["项目 Item", "Value"])

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

    line(5, "Risk-free Rate (Rf) 无风险利率", f"={A}!B14", PCT_FMT)
    line(6, "Equity Risk Premium (ERP) 股权风险溢价", f"={A}!B15", PCT_FMT)
    line(7, "Beta", f"={A}!B16")
    line(8, "Cost of Equity (CAPM) 股权成本", "=B5+B6*B7", PCT_FMT, total=True)
    line(9, "Pre-tax Cost of Debt 税前债务成本", f"={A}!B17", PCT_FMT)
    line(10, "Tax Rate 税率", f"={A}!B11", PCT_FMT)
    line(11, "After-tax Cost of Debt 税后债务成本", "=B9*(1-B10)", PCT_FMT)
    line(12, "Equity Weight 股权重", f"=1-{A}!B18", PCT_FMT)
    line(13, "Debt Weight 债权重", f"={A}!B18", PCT_FMT)
    line(14, "WACC 加权平均资本成本", "=B8*B12+B11*B13", PCT_FMT, total=True)

    set_col_widths(ws, {"A": 40, "B": 14})
    ws.sheet_view.showGridLines = False


def _write_terminal_value(wb, p: DCFParams):
    ws = wb.create_sheet("Terminal Value")
    ws.cell(row=1, column=1, value="Terminal Value 终值").font = S.TITLE_FONT
    write_header_row(ws, 3, ["项目 Item", "Value"])

    A = "Assumptions"
    OP = "'Operating Model'"
    W = "WACC"
    n = p.projection_years
    last_col = ycol(n)  # Year N 列

    def line(row, label, formula, fmt=NUM_FMT, total=False):
        c = ws.cell(row=row, column=1, value=label)
        (style_total if total else style_label)(c)
        cell = ws.cell(row=row, column=2)
        cell.value = formula
        style_formula(cell, fmt, link="!" in (formula or ""))

    line(5, "Gordon Growth TV 戈登增长终值",
         f"={OP}!{last_col}15*(1+{A}!$B$21)/({W}!$B$14-{A}!$B$21)", total=True)
    line(6, "Exit Multiple TV 退出倍数终值",
         f"={OP}!{last_col}7*{A}!$B$22", MULT_FMT)
    line(7, "Selected TV 选定终值",
         f'=IF({A}!$B$23="gordon",B5,B6)', total=True)
    line(8, "Discount Period (Years) 折现期数", f"={n}")
    line(9, "PV of Terminal Value 终值现值",
         f"=B7/(1+{W}!$B$14)^B8", total=True)

    set_col_widths(ws, {"A": 42, "B": 16})
    ws.sheet_view.showGridLines = False


def _write_dcf_valuation(wb, p: DCFParams):
    ws = wb.create_sheet("DCF Valuation")
    ws.cell(row=1, column=1, value="DCF Valuation 估值汇总").font = S.TITLE_FONT
    n = p.projection_years
    headers = ["项目 Item"] + [f"Year {i}" for i in range(1, n + 1)]
    write_header_row(ws, 3, headers)

    A = "Assumptions"
    OP = "'Operating Model'"
    W = "WACC"
    TV = "'Terminal Value'"

    def row(r, label, formula_func, fmt=NUM_FMT, total=False):
        c = ws.cell(row=r, column=1, value=label)
        (style_total if total else style_label)(c)
        for yi in range(n):
            col = ycol(yi + 1)  # Year 1..N = C..G
            f = formula_func(col, yi)
            cell = ws.cell(row=r, column=2 + yi)
            if f is not None:
                cell.value = f
            if total:
                style_total(cell, fmt)
            else:
                style_formula(cell, fmt, link="!" in (f or ""))

    # Year 1..N = columns C..G in OpModel
    row(5, "Unlevered FCF 自由现金流", lambda c, yi: f"={OP}!{c}15")

    # Discount period: mid-year = 0.5,1.5,...; end-year = 1,2,...
    is_mid = p.valuation_timing == "mid"
    def disc_period(col, yi):
        period = yi + 0.5 if is_mid else yi + 1
        return f"={period}"
    row(6, "Discount Period 折现期数", disc_period)

    row(7, "Discount Factor 折现因子", lambda c, yi: f"=1/(1+{W}!$B$14)^{c}6", fmt='0.0000')
    row(8, "PV of FCF FCF现值", lambda c, yi: f"={c}5*{c}7", total=True)

    # 汇总区
    def line(row, label, formula, fmt=NUM_FMT, total=False):
        c = ws.cell(row=row, column=1, value=label)
        (style_total if total else style_label)(c)
        cell = ws.cell(row=row, column=2)
        cell.value = formula
        if total:
            style_total(cell, fmt)
        else:
            style_formula(cell, fmt, link="!" in (formula or ""))

    # Sum of PV FCF (in column B, but the FCF PVs are in C..G)
    first_pv = ycol(1)  # C
    last_pv = ycol(n)   # G
    line(10, "Sum of PV of FCF FCF现值合计", f"=SUM({first_pv}8:{last_pv}8)")
    line(11, "PV of Terminal Value 终值现值", f"={TV}!B9")
    line(12, "Enterprise Value 企业价值", f"=B10+B11", total=True)

    # Equity bridge
    sec = ws.cell(row=14, column=1, value="EV → Equity 桥")
    style_subheader(sec)
    style_subheader(ws.cell(row=14, column=2))

    line(15, "Enterprise Value 企业价值", f"=B12", total=True)
    line(16, "Less: Total Debt 减:总债务", f"=-{A}!B26")
    line(17, "Plus: Cash 加:现金", f"={A}!B27")
    line(18, "Less: Minority Interest 减:少数股东权益", f"=-{A}!B28")
    line(19, "Less: Preferred Stock 减:优先股", f"=-{A}!B29")
    line(20, "Less: Stock-Based Comp 减:股权激励", f"=-{A}!B30")
    line(21, "Less: Pension Deficit 减:养老金缺口", f"=-{A}!B31")
    line(22, "Equity Value 股权价值", f"=SUM(B15:B21)", total=True)
    line(23, "Shares Outstanding (M) 流通股本", f"={A}!B32")
    line(24, "Implied Share Price 隐含股价", f"=B22/B23", PRICE_FMT, total=True)
    line(25, "Current Share Price 当前股价", f"={A}!B33", PRICE_FMT)
    line(26, "Upside/(Downside) % 涨跌幅", f"=B24/B25-1", PCT_FMT, total=True)

    set_col_widths(ws, {"A": 40})
    for i in range(n):
        ws.column_dimensions[get_column_letter(2 + i)].width = 14
    ws.sheet_view.showGridLines = False


def _write_sensitivity(wb, p: DCFParams):
    ws = wb.create_sheet("Sensitivity")
    ws.cell(row=1, column=1, value="Sensitivity 敏感性分析").font = S.TITLE_FONT
    ws.cell(row=2, column=1, value="行=终值增长率 g，列=WACC；单元格=隐含股价").font = S.SUBHEADER_FONT

    A = "Assumptions"
    OP = "'Operating Model'"
    n = p.projection_years

    # g 和 WACC 的取值范围
    g_values = [p.terminal_growth - 0.01, p.terminal_growth - 0.005,
                p.terminal_growth, p.terminal_growth + 0.005, p.terminal_growth + 0.01]
    wacc_values = [0.08, 0.09, 0.10, 0.11, 0.12]

    # 表头
    ws.cell(row=4, column=1, value="g \\ WACC").font = S.HEADER_FONT
    ws.cell(row=4, column=1).fill = S.HEADER_FILL
    ws.cell(row=4, column=1).alignment = S.CENTER_ALIGN
    ws.cell(row=4, column=1).border = S.BORDER
    for j, w in enumerate(wacc_values):
        c = ws.cell(row=4, column=2 + j, value=w)
        style_header(c)
        c.number_format = PCT_FMT

    last_col = ycol(n)  # OpModel Year N

    # 每个单元格: EV = Σ FCF_t/(1+w)^t + TV/(1+w)^N ; TV = FCF_N*(1+g)/(w-g)
    # Equity = EV - Debt + Cash - Minority - Preferred - SBC - Pension
    # Price = Equity / Shares
    for i, g in enumerate(g_values):
        rc = ws.cell(row=5 + i, column=1, value=g)
        rc.font = S.HEADER_FONT
        rc.fill = S.HEADER_FILL
        rc.alignment = S.CENTER_ALIGN
        rc.border = S.BORDER
        rc.number_format = PCT_FMT

        for j, w in enumerate(wacc_values):
            # 用命名变量引用 WACC 列头（B4..F4 对应列2..6）
            wcell = f"{get_column_letter(2 + j)}$4"
            gcell = f"$A{5 + i}"
            # PV of FCFs: Σ OpModel!{c}15 / (1+w)^t  for t=1..N
            pv_terms = "+".join(
                f"{OP}!{ycol(t)}15/(1+{wcell})^{t}" for t in range(1, n + 1)
            )
            # Terminal Value (Gordon): OpModel!{last}15*(1+g)/(w-g) , then PV
            tv = f"{OP}!{last_col}15*(1+{gcell})/({wcell}-{gcell})"
            pv_tv = f"{tv}/(1+{wcell})^{n}"
            ev = f"({pv_terms})+{pv_tv}"
            equity = f"({ev})-{A}!$B$26+{A}!$B$27-{A}!$B$28-{A}!$B$29-{A}!$B$30-{A}!$B$31"
            price = f"({equity})/{A}!$B$32"
            cell = ws.cell(row=5 + i, column=2 + j, value=f"={price}")
            style_formula(cell, PRICE_FMT)

    set_col_widths(ws, {"A": 12})
    for j in range(len(wacc_values)):
        ws.column_dimensions[get_column_letter(2 + j)].width = 12
    ws.sheet_view.showGridLines = False


def build(p: DCFParams) -> Workbook:
    """构建 DCF 模型 workbook。"""
    wb = Workbook()
    wb.remove(wb.active)
    _write_assumptions(wb, p)
    _write_operating_model(wb, p)
    _write_wacc(wb, p)
    _write_terminal_value(wb, p)
    _write_dcf_valuation(wb, p)
    _write_sensitivity(wb, p)
    wb.active = wb.sheetnames.index("Assumptions")
    return wb
