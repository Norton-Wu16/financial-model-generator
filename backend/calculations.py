"""纯 Python 计算引擎 - 复现三个模型的 Excel 公式逻辑。

每个 calculate_xxx 函数严格对照 backend/builders/xxx.py 的 Excel 公式，
返回结构化 JSON，供 /api/preview 端点使用。

计算顺序严格遵循跨表依赖链，确保与 Excel 结果一致（误差 < 0.01）。
"""
from typing import Dict, List, Any
from backend.schemas import ThreeStatementParams, DCFParams, LBOParams
from backend.custom_items import CustomItem, evaluate_expression


# ==================== 辅助函数 ====================

def _round2(x: float) -> float:
    """两位小数四舍五入，对齐 Excel 默认精度。"""
    if x is None:
        return 0.0
    return round(float(x), 2)


def _safe_div(a: float, b: float) -> float:
    """安全除法，避免除零。"""
    if b == 0:
        return 0.0
    return a / b


def _eval_custom_items(
    custom_items: List[CustomItem],
    years: List[int],
    context_per_year: List[Dict[str, float]],
) -> List[Dict[str, Any]]:
    """对每个自定义行项，按年求值，返回 JSON 结构。

    返回: [{name, format, values: [v0, v1, ...], error: "" 或 错误信息}]
    求值失败的年份填 None，整行解析失败则 error 字段填错误信息且 values 为空。
    """
    if not custom_items:
        return []
    n_years = len(years)
    result: List[Dict[str, Any]] = []
    for ci in custom_items:
        entry: Dict[str, Any] = {
            "name": ci.name,
            "format": ci.format,
            "formula": ci.formula,
            "values": [],
            "error": "",
        }
        # 先整体校验语法
        try:
            from backend.custom_items import tokenize, parse
            tokens = tokenize(ci.formula)
            ast = parse(tokens)
        except Exception as e:
            entry["error"] = f"公式语法错误: {e}"
            result.append(entry)
            continue
        # 按年求值
        values: List[Any] = []
        for i in range(n_years):
            ctx = context_per_year[i] if i < len(context_per_year) else {}
            try:
                from backend.custom_items import evaluate
                v = evaluate(ast, ctx)
                values.append(_round2(v))
            except Exception as e:
                values.append(None)  # 求值失败的年份填 None
        entry["values"] = values
        result.append(entry)
    return result


def calculate_irr(cash_flows: List[float], max_iter: int = 100, tol: float = 1e-8) -> float:
    """用二分法求 IRR，与 Excel IRR 函数对齐。

    Excel IRR 假设 rate 在 [-1, 1]（即 -100% 到 100%）范围内。
    二分法需要 NPV(rate) 在区间两端异号。
    """
    n = len(cash_flows)
    if n < 2:
        return 0.0

    def npv(rate: float) -> float:
        total = 0.0
        for i, cf in enumerate(cash_flows):
            total += cf / ((1 + rate) ** i)
        return total

    # 找符号变化的区间
    lo, hi = -0.9999, 10.0  # IRR 上限放宽到 1000%，实际项目不会超过
    npv_lo = npv(lo)
    npv_hi = npv(hi)

    # 若两端同号，尝试扩大范围
    if npv_lo * npv_hi > 0:
        # 可能全正或全负（无解），返回 0 作为兜底
        # 尝试在更小区间找
        for test_rate in [x * 0.01 for x in range(-99, 1001)]:
            if npv(test_rate) * npv_lo < 0:
                hi = test_rate
                npv_hi = npv(test_rate)
                break
        else:
            return 0.0

    # 二分迭代
    for _ in range(max_iter):
        mid = (lo + hi) / 2
        npv_mid = npv(mid)
        if abs(npv_mid) < tol:
            return mid
        if npv_mid * npv_lo < 0:
            hi = mid
            npv_hi = npv_mid
        else:
            lo = mid
            npv_lo = npv_mid

    return (lo + hi) / 2


# ==================== 三表联动模型 ====================

def calculate_three_statement(p: ThreeStatementParams) -> Dict[str, Any]:
    """复现 backend/builders/three_statement.py 的 Excel 公式。

    计算依赖链:
      Revenue → COGS/SG&A/R&D/D&A → EBITDA → EBIT
      Schedules: PP&E 滚动、Debt 滚动 → Interest
      → EBT → Taxes → Net Income → Dividends
      BS: AR/Inv/AP/Accrued (营运资本) → Total Assets/Liab
      BS: RE = prev_RE + NI - Dividends
      CFS: ΔWC + NI + D&A + CapEx + ΔDebt → Net Change in Cash
      BS Cash = prev_cash + Net Change (平衡项)
      Balance Check: |Total Assets - Total L+E| < 0.01
    """
    n = p.projection_years
    years = list(range(n + 1))

    # ========== 1. Income Statement ==========
    revenue = [0.0] * (n + 1)
    revenue[0] = p.revenue_y0  # IS!B5 = Assumptions!B5
    growths = [p.rev_growth_y1, p.rev_growth_y2, p.rev_growth_y3,
               p.rev_growth_y4, p.rev_growth_y5]
    for i in range(1, n + 1):
        g = growths[i - 1] if i <= len(growths) else 0.0
        revenue[i] = revenue[i - 1] * (1 + g)  # IS!Ci5 = prev*(1+g)

    cogs = [r * p.cogs_pct for r in revenue]          # IS!B7 = revenue * cogs_pct
    gross_profit = [r - c for r, c in zip(revenue, cogs)]  # IS!B8
    sga = [r * p.sga_pct for r in revenue]             # IS!B9
    rd = [r * p.rd_pct for r in revenue]               # IS!B10
    ebitda = [gp - s - r for gp, s, r in zip(gross_profit, sga, rd)]  # IS!B11
    da_is = [r * p.da_pct for r in revenue]            # IS!B12 / SCH!C6 (D&A 始终 = revenue * da_pct)
    ebit = [e - d for e, d in zip(ebitda, da_is)]      # IS!B13

    # ========== 2. Supporting Schedules ==========
    # PP&E Roll-forward
    ppe_beg = [0.0] * (n + 1)
    ppe_beg[0] = p.beg_ppe  # SCH!B7 = Assumptions!B26
    capex = [r * p.capex_pct for r in revenue]  # SCH!B8 = IS!B5 * capex_pct
    # SCH!B9 = -SCH!B6 = -da
    ppe_end = [0.0] * (n + 1)
    ppe_end[0] = p.beg_ppe  # SCH!B10 (Y0) = SCH!B7 (只 beg, 不加 CapEx/D&A)
    for i in range(1, n + 1):
        ppe_beg[i] = ppe_end[i - 1]  # SCH!Ci7 = prev end
        ppe_end[i] = ppe_beg[i] + capex[i] - da_is[i]  # SCH!Ci10 = beg + capex + (-da)

    # Debt Roll-forward
    debt_beg = [0.0] * (n + 1)
    debt_beg[0] = p.beg_debt  # SCH!B18 = Assumptions!B29
    new_debt = [0.0] * (n + 1)
    new_debt[0] = 0.0  # SCH!B19 (Y0) = 0
    for i in range(1, n + 1):
        new_debt[i] = p.new_debt_issuance  # SCH!Ci19 = Assumptions!Ci34
    debt_end = [0.0] * (n + 1)
    debt_end[0] = p.beg_debt  # SCH!B20 (Y0) = SCH!B18
    for i in range(1, n + 1):
        debt_beg[i] = debt_end[i - 1]  # SCH!Ci18 = prev end
        debt_end[i] = debt_beg[i] + new_debt[i]  # SCH!Ci20 = beg + new

    # Interest Expense: Y0 = interest_rate * beg_debt; Yi = interest_rate * prev_debt
    interest = [0.0] * (n + 1)
    interest[0] = p.interest_rate * p.beg_debt  # IS!B14 / SCH!B17 (Y0)
    for i in range(1, n + 1):
        # SCH!Ci17 = interest_rate * BS!prev_col15 (prev period debt = debt_end[i-1])
        interest[i] = p.interest_rate * debt_end[i - 1]

    # ========== 3. 回到 Income Statement ==========
    ebt = [e - intr for e, intr in zip(ebit, interest)]  # IS!B15
    taxes = [max(0, e * p.tax_rate) for e in ebt]        # IS!B16 = MAX(0, EBT * tax_rate)
    net_income = [e - t for e, t in zip(ebt, taxes)]     # IS!B17
    dividends = [ni * p.dividend_pct for ni in net_income]  # IS!B18

    # ========== 4. Balance Sheet ==========
    # Assets
    cash = [0.0] * (n + 1)
    cash[0] = p.beg_cash  # BS!B5 = Assumptions!B23
    ar = [0.0] * (n + 1)
    ar[0] = p.beg_ar      # BS!B6
    inventory = [0.0] * (n + 1)
    inventory[0] = p.beg_inventory  # BS!B7
    ppe_bs = [0.0] * (n + 1)
    ppe_bs[0] = p.beg_ppe  # BS!B8 = Assumptions!B26
    for i in range(1, n + 1):
        ar[i] = revenue[i] / 365 * p.dso            # BS!Ci6 = IS!Ci5 / 365 * dso
        inventory[i] = cogs[i] / 365 * p.dio        # BS!Ci7 = IS!Ci7 (COGS) / 365 * dio
        ppe_bs[i] = ppe_end[i]                       # BS!Ci8 = SCH!Ci10

    # 现金流表先算 Net Change，再回填 BS Cash
    # Liabilities
    ap = [0.0] * (n + 1)
    ap[0] = p.beg_ap  # BS!B13
    accrued = [0.0] * (n + 1)
    accrued[0] = p.beg_accrued  # BS!B14
    debt_bs = [0.0] * (n + 1)
    debt_bs[0] = p.beg_debt  # BS!B15
    for i in range(1, n + 1):
        ap[i] = cogs[i] / 365 * p.dpo               # BS!Ci13 = IS!Ci7 (COGS) / 365 * dpo
        accrued[i] = revenue[i] / 365 * p.accrued_days  # BS!Ci14 = IS!Ci5 (Revenue) / 365 * accrued_days
        debt_bs[i] = debt_end[i]                    # BS!Ci15 = SCH!Ci20

    # Equity
    common_stock = [p.common_stock] * (n + 1)  # BS!B20 = Assumptions!$B$30 (constant)
    re = [0.0] * (n + 1)
    re[0] = p.beg_retn_earn  # BS!B21 = Assumptions!B31
    for i in range(1, n + 1):
        re[i] = re[i - 1] + net_income[i] - dividends[i]  # BS!Ci21 = prev_RE + NI - Div

    # ========== 5. Cash Flow Statement ==========
    # Y0 全部为 0/None (Excel 里是空)
    delta_ar = [0.0] * (n + 1)
    delta_inventory = [0.0] * (n + 1)
    delta_ap = [0.0] * (n + 1)
    delta_accrued = [0.0] * (n + 1)
    cfo = [0.0] * (n + 1)
    cfi = [0.0] * (n + 1)
    cff = [0.0] * (n + 1)
    net_change_cash = [0.0] * (n + 1)
    capex_cfs = [0.0] * (n + 1)  # 负值
    net_debt_change = [0.0] * (n + 1)
    dividends_paid = [0.0] * (n + 1)  # 负值

    for i in range(1, n + 1):
        delta_ar[i] = -(ar[i] - ar[i - 1])             # CFS!Ci8 = -(BS_AR - prev)
        delta_inventory[i] = -(inventory[i] - inventory[i - 1])
        delta_ap[i] = ap[i] - ap[i - 1]                # CFS!Ci10 = BS_AP - prev
        delta_accrued[i] = accrued[i] - accrued[i - 1]
        cfo[i] = net_income[i] + da_is[i] + delta_ar[i] + delta_inventory[i] + delta_ap[i] + delta_accrued[i]
        capex_cfs[i] = -capex[i]                       # CFS!Ci15 = -IS!Ci5 * capex_pct
        cfi[i] = capex_cfs[i]                          # CFS!Ci16
        net_debt_change[i] = debt_bs[i] - debt_bs[i - 1]  # CFS!Ci19
        dividends_paid[i] = -dividends[i]              # CFS!Ci20
        cff[i] = net_debt_change[i] + dividends_paid[i]
        net_change_cash[i] = cfo[i] + cfi[i] + cff[i]

    # 回填 BS Cash: Yi = prev + net_change
    for i in range(1, n + 1):
        cash[i] = cash[i - 1] + net_change_cash[i]  # BS!Ci5 = prev + CFS!Ci23

    # ========== 6. 汇总与平衡校验 ==========
    total_assets = [cash[i] + ar[i] + inventory[i] + ppe_bs[i] for i in range(n + 1)]
    total_liab = [ap[i] + accrued[i] + debt_bs[i] for i in range(n + 1)]
    total_equity = [common_stock[i] + re[i] for i in range(n + 1)]
    total_le = [total_liab[i] + total_equity[i] for i in range(n + 1)]
    balance_check = [
        "BALANCED" if abs(total_assets[i] - total_le[i]) < 0.01 else "OUT OF BALANCE"
        for i in range(n + 1)
    ]

    # Key metrics
    ebitda_margin = [0.0] + [_safe_div(ebitda[i], revenue[i]) for i in range(1, n + 1)]
    net_margin = [0.0] + [_safe_div(net_income[i], revenue[i]) for i in range(1, n + 1)]

    # ========== 自定义行项求值 ==========
    def _build_ts_context(i):
        return {
            "revenue": revenue[i], "cogs": cogs[i], "gross_profit": gross_profit[i],
            "sga": sga[i], "rd": rd[i], "ebitda": ebitda[i], "da": da_is[i],
            "ebit": ebit[i], "interest": interest[i], "ebt": ebt[i],
            "taxes": taxes[i], "net_income": net_income[i], "dividends": dividends[i],
            "cash": cash[i], "ar": ar[i], "inventory": inventory[i],
            "ppe": ppe_bs[i], "total_assets": total_assets[i],
            "ap": ap[i], "accrued": accrued[i], "debt": debt_bs[i],
            "total_liabilities": total_liab[i], "common_stock": common_stock[i],
            "retained_earnings": re[i], "total_equity": total_equity[i],
            "total_le": total_le[i], "capex": capex[i],
            "new_debt": new_debt[i],
        }
    contexts = [_build_ts_context(i) for i in range(n + 1)]
    custom_results = _eval_custom_items(p.custom_items, years, contexts)

    return {
        "model_type": "three_statement",
        "company_name": p.company_name,
        "years": years,
        "income_statement": {
            "revenue": [_round2(x) for x in revenue],
            "cogs": [_round2(x) for x in cogs],
            "gross_profit": [_round2(x) for x in gross_profit],
            "sga": [_round2(x) for x in sga],
            "rd": [_round2(x) for x in rd],
            "ebitda": [_round2(x) for x in ebitda],
            "da": [_round2(x) for x in da_is],
            "ebit": [_round2(x) for x in ebit],
            "interest": [_round2(x) for x in interest],
            "ebt": [_round2(x) for x in ebt],
            "taxes": [_round2(x) for x in taxes],
            "net_income": [_round2(x) for x in net_income],
            "dividends": [_round2(x) for x in dividends],
        },
        "balance_sheet": {
            "cash": [_round2(x) for x in cash],
            "ar": [_round2(x) for x in ar],
            "inventory": [_round2(x) for x in inventory],
            "ppe": [_round2(x) for x in ppe_bs],
            "total_assets": [_round2(x) for x in total_assets],
            "ap": [_round2(x) for x in ap],
            "accrued": [_round2(x) for x in accrued],
            "debt": [_round2(x) for x in debt_bs],
            "total_liabilities": [_round2(x) for x in total_liab],
            "common_stock": [_round2(x) for x in common_stock],
            "retained_earnings": [_round2(x) for x in re],
            "total_equity": [_round2(x) for x in total_equity],
            "total_le": [_round2(x) for x in total_le],
            "balance_check": balance_check,
        },
        "cash_flow": {
            "net_income": [_round2(x) for x in net_income],
            "da": [_round2(x) for x in da_is],
            "delta_ar": [_round2(x) for x in delta_ar],
            "delta_inventory": [_round2(x) for x in delta_inventory],
            "delta_ap": [_round2(x) for x in delta_ap],
            "delta_accrued": [_round2(x) for x in delta_accrued],
            "cfo": [_round2(x) for x in cfo],
            "capex": [_round2(x) for x in capex_cfs],
            "cfi": [_round2(x) for x in cfi],
            "net_debt_change": [_round2(x) for x in net_debt_change],
            "dividends_paid": [_round2(x) for x in dividends_paid],
            "cff": [_round2(x) for x in cff],
            "net_change_in_cash": [_round2(x) for x in net_change_cash],
        },
        "key_metrics": {
            "ebitda_margin": [_round2(x) for x in ebitda_margin],
            "net_margin": [_round2(x) for x in net_margin],
        },
        "custom_items": custom_results,
    }


# ==================== DCF 模型 ====================

def calculate_dcf(p: DCFParams) -> Dict[str, Any]:
    """复现 backend/builders/dcf.py 的 Excel 公式。

    计算链:
      Revenue → EBITDA → D&A → EBIT → NOPAT
      ΔNWC = (Rev - prev_Rev) * nwc_pct
      UFCF = NOPAT + D&A - ΔNWC - CapEx
      WACC = CAPM + 税后债务成本加权
      TV = Gordon (UFCF_N*(1+g)/(wacc-g)) 或 Exit (EBITDA_N * exit_multiple)
      PV汇总 = Σ UFCF/(1+wacc)^t + TV/(1+wacc)^N
      EV → Equity 桥 → Implied Price
    """
    n = p.projection_years
    years = list(range(n + 1))

    # ========== Operating Model ==========
    revenue = [0.0] * (n + 1)
    revenue[0] = p.revenue_y0
    growths = [p.rev_growth_y1, p.rev_growth_y2, p.rev_growth_y3,
               p.rev_growth_y4, p.rev_growth_y5]
    for i in range(1, n + 1):
        g = growths[i - 1] if i <= len(growths) else 0.0
        revenue[i] = revenue[i - 1] * (1 + g)

    # EBITDA: Y0 = Rev * ebitda_margin_y1; Yi = Rev * ebitda_margin_yi
    margins = [p.ebitda_margin_y1, p.ebitda_margin_y2, p.ebitda_margin_y3,
               p.ebitda_margin_y4, p.ebitda_margin_y5]
    ebitda = [0.0] * (n + 1)
    ebitda[0] = revenue[0] * margins[0]  # Y0 uses C7 (Year 1 margin) per Excel formula
    for i in range(1, n + 1):
        m = margins[i - 1] if i <= len(margins) else margins[-1]
        ebitda[i] = revenue[i] * m

    da = [r * p.da_pct for r in revenue]  # D&A = Rev * da_pct
    ebit = [e - d for e, d in zip(ebitda, da)]
    nopat = [e * (1 - p.tax_rate) for e in ebit]

    # ΔNWC = (Rev - prev_Rev) * nwc_pct (Y0 = 0)
    delta_nwc = [0.0] * (n + 1)
    for i in range(1, n + 1):
        delta_nwc[i] = (revenue[i] - revenue[i - 1]) * p.nwc_pct

    capex = [r * p.capex_pct for r in revenue]

    # UFCF = NOPAT + D&A - ΔNWC - CapEx (Y0 = 0)
    ufcf = [0.0] * (n + 1)
    for i in range(1, n + 1):
        ufcf[i] = nopat[i] + da[i] - delta_nwc[i] - capex[i]

    # ========== WACC ==========
    cost_of_equity = p.risk_free_rate + p.equity_risk_premium * p.beta
    after_tax_kd = p.pre_tax_cost_of_debt * (1 - p.tax_rate)
    equity_weight = 1 - p.debt_weight
    wacc = cost_of_equity * equity_weight + after_tax_kd * p.debt_weight

    # ========== Terminal Value ==========
    last_ufcf = ufcf[n]
    last_ebitda = ebitda[n]
    tv_gordon = last_ufcf * (1 + p.terminal_growth) / (wacc - p.terminal_growth) if wacc != p.terminal_growth else 0.0
    tv_exit = last_ebitda * p.exit_multiple
    selected_tv = tv_gordon if p.tv_method == "gordon" else tv_exit
    pv_tv = selected_tv / ((1 + wacc) ** n)

    # ========== DCF Valuation ==========
    # 折现期: mid = yi + 0.5; end = yi + 1 (yi 从 1 开始)
    is_mid = p.valuation_timing == "mid"
    pv_fcfs = [0.0] * (n + 1)
    for i in range(1, n + 1):
        period = i - 0.5 if is_mid else i  # 对应 Excel: yi + 0.5 or yi + 1, yi 从 1 起
        # 注意: Excel 中 yi 从 0 起 (Year 1 对应 yi=0), 所以 period = yi + 0.5
        # 这里 i 从 1 起 (Year 1), 所以 period = i - 0.5 (mid) 或 i (end)
        pv_fcfs[i] = ufcf[i] / ((1 + wacc) ** period)

    sum_pv_fcf = sum(pv_fcfs[1:])
    ev = sum_pv_fcf + pv_tv

    # Equity 桥
    equity_value = ev - p.net_debt + p.cash - p.minority_interest - p.preferred_stock - p.stock_based_comp - p.pension_deficit
    implied_price = _safe_div(equity_value, p.shares_outstanding)
    upside = _safe_div(implied_price, p.current_price) - 1 if p.current_price else 0.0

    # ========== 汇总 ==========
    ebitda_margin_calc = [0.0] + [_safe_div(ebitda[i], revenue[i]) for i in range(1, n + 1)]
    ebit_margin = [0.0] + [_safe_div(ebit[i], revenue[i]) for i in range(1, n + 1)]

    # ========== 自定义行项求值 ==========
    def _build_dcf_context(i):
        return {
            "revenue": revenue[i], "ebitda": ebitda[i], "da": da[i], "ebit": ebit[i],
            "nopat": nopat[i], "delta_nwc": delta_nwc[i], "capex": capex[i],
            "ufcf": ufcf[i], "wacc": wacc, "cost_of_equity": cost_of_equity,
            "after_tax_kd": after_tax_kd, "tv_gordon": tv_gordon,
            "tv_exit": tv_exit, "selected_tv": selected_tv, "pv_tv": pv_tv,
            "sum_pv_fcf": sum_pv_fcf, "enterprise_value": ev,
            "equity_value": equity_value, "implied_price": implied_price,
            "current_price": p.current_price, "shares_outstanding": p.shares_outstanding,
            "net_debt": p.net_debt, "cash": p.cash,
        }
    contexts = [_build_dcf_context(i) for i in range(n + 1)]
    custom_results = _eval_custom_items(p.custom_items, years, contexts)

    return {
        "model_type": "dcf",
        "company_name": p.company_name,
        "years": years,
        "operating_model": {
            "revenue": [_round2(x) for x in revenue],
            "ebitda": [_round2(x) for x in ebitda],
            "da": [_round2(x) for x in da],
            "ebit": [_round2(x) for x in ebit],
            "nopat": [_round2(x) for x in nopat],
            "delta_nwc": [_round2(x) for x in delta_nwc],
            "capex": [_round2(x) for x in capex],
            "ufcf": [_round2(x) for x in ufcf],
        },
        "wacc": {
            "cost_of_equity": _round2(cost_of_equity),
            "after_tax_cost_of_debt": _round2(after_tax_kd),
            "equity_weight": _round2(equity_weight),
            "debt_weight": _round2(p.debt_weight),
            "wacc": _round2(wacc),
        },
        "terminal_value": {
            "tv_method": p.tv_method,
            "tv_gordon": _round2(tv_gordon),
            "tv_exit": _round2(tv_exit),
            "selected_tv": _round2(selected_tv),
            "pv_tv": _round2(pv_tv),
        },
        "valuation": {
            "sum_pv_fcf": _round2(sum_pv_fcf),
            "pv_tv": _round2(pv_tv),
            "enterprise_value": _round2(ev),
            "equity_value": _round2(equity_value),
            "shares_outstanding": p.shares_outstanding,
            "implied_price": _round2(implied_price),
            "current_price": p.current_price,
            "upside": _round2(upside),
        },
        "key_metrics": {
            "ebitda_margin": [_round2(x) for x in ebitda_margin_calc],
            "ebit_margin": [_round2(x) for x in ebit_margin],
        },
        "custom_items": custom_results,
    }


# ==================== LBO 模型 ====================

def calculate_lbo(p: LBOParams) -> Dict[str, Any]:
    """复现 backend/builders/lbo.py 的 Excel 公式。

    计算链:
      Entry EV = LTM EBITDA * entry_multiple
      S&U: Uses = Purchase Equity + Refi Debt + Fees; Sources = Sponsor + Revolver(plug) + 各债务
      Debt Schedule: 现金/CFADS → Revolver repayment → TLA/TLB mandatory+sweep → SrNotes/Sub bullet
      IS: 各层级 Interest = rate * debt_beg; EBT → NI
      CFADS: NI + D&A - ΔNWC - CapEx
      Exit: Exit EV = EBITDA_exit * exit_mult; Equity = EV - Debt + Cash; MoIC, IRR
    """
    n = 5  # LBO 固定 5 年预测
    years = list(range(n + 1))

    # ========== Sources & Uses ==========
    entry_ev = p.ltm_ebitda * p.entry_ev_ebitda
    purchase_equity = entry_ev - p.existing_net_debt  # SU!B6
    refi_debt = p.existing_net_debt + p.existing_cash  # SU!B7
    trans_fees = p.transaction_fees
    fin_fees = p.financing_fees
    total_uses = purchase_equity + refi_debt + trans_fees + fin_fees  # SU!B10

    # Sources: Sponsor + Revolver(plug) + TLA + TLB + SrNotes + Sub + Existing Cash
    # Revolver Draw = Total Uses - Sponsor - TLA - TLB - SrNotes - Sub - Existing Cash
    revolver_draw = total_uses - p.sponsor_equity - p.term_loan_a - p.term_loan_b - p.sponsor_equity * 0 \
        - p.senior_notes - p.subordinated_debt - p.existing_cash
    # 修正: 去掉错误项
    revolver_draw = total_uses - p.sponsor_equity - p.term_loan_a - p.term_loan_b - p.senior_notes - p.subordinated_debt - p.existing_cash
    total_sources = p.sponsor_equity + revolver_draw + p.term_loan_a + p.term_loan_b + p.senior_notes + p.subordinated_debt + p.existing_cash
    su_balance = abs(total_sources - total_uses) < 0.01

    # ========== Operating: Revenue / EBITDA ==========
    revenue = [0.0] * (n + 1)
    revenue[0] = p.ltm_revenue
    growths = [p.rev_growth_y1, p.rev_growth_y2, p.rev_growth_y3,
               p.rev_growth_y4, p.rev_growth_y5]
    for i in range(1, n + 1):
        g = growths[i - 1] if i <= len(growths) else 0.0
        revenue[i] = revenue[i - 1] * (1 + g)

    ebitda = [0.0] * (n + 1)
    ebitda[0] = p.ltm_ebitda
    for i in range(1, n + 1):
        ebitda[i] = revenue[i] * p.ebitda_margin

    da = [r * p.da_pct for r in revenue]
    ebit = [e - d for e, d in zip(ebitda, da)]
    capex = [r * p.capex_pct for r in revenue]
    delta_nwc = [0.0] * (n + 1)
    for i in range(1, n + 1):
        delta_nwc[i] = -revenue[i] * p.nwc_pct  # ΔNWC = -Rev * nwc_pct (Excel: =-IS*col5*A!B36)

    # ========== Debt Schedule ==========
    # Cash & CFADS
    beg_cash = [0.0] * (n + 1)
    beg_cash[0] = p.existing_cash
    cfads = [0.0] * (n + 1)  # Y0 = 0
    cash_avail = [0.0] * (n + 1)

    # Revolver
    rev_beg = [0.0] * (n + 1)
    rev_beg[0] = revolver_draw
    rev_repay = [0.0] * (n + 1)
    rev_end = [0.0] * (n + 1)
    rev_end[0] = revolver_draw

    # TLA
    tla_beg = [0.0] * (n + 1)
    tla_beg[0] = p.term_loan_a
    tla_mandatory = [0.0] * (n + 1)
    tla_end = [0.0] * (n + 1)
    tla_end[0] = p.term_loan_a

    # TLB
    tlb_beg = [0.0] * (n + 1)
    tlb_beg[0] = p.term_loan_b
    tlb_mandatory = [0.0] * (n + 1)
    tlb_sweep = [0.0] * (n + 1)
    tlb_end = [0.0] * (n + 1)
    tlb_end[0] = p.term_loan_b

    # Sr Notes (bullet)
    sn_beg = [0.0] * (n + 1)
    sn_beg[0] = p.senior_notes
    sn_end = [0.0] * (n + 1)
    sn_end[0] = p.senior_notes

    # Sub (bullet)
    sub_beg = [0.0] * (n + 1)
    sub_beg[0] = p.subordinated_debt
    sub_end = [0.0] * (n + 1)
    sub_end[0] = p.subordinated_debt

    total_debt = [0.0] * (n + 1)
    cash_end = [0.0] * (n + 1)
    total_debt[0] = rev_end[0] + tla_end[0] + tlb_end[0] + sn_end[0] + sub_end[0]
    cash_end[0] = beg_cash[0]  # Y0 = beg_cash

    # 需要先算 IS 的 NI 和 CFADS 的 CFADS 才能滚动 Debt Schedule
    # 但 IS Interest 依赖 Debt Schedule 的 beg balances
    # 这是循环依赖, Excel 里通过跨表引用解决, Python 里需要迭代或同步计算

    # 同步计算: 对每一年, 先用 prev 期末债务算 Interest, 再算 IS, 再算 CFADS, 再滚动 Debt
    interest_revolver = [0.0] * (n + 1)
    interest_tla = [0.0] * (n + 1)
    interest_tlb = [0.0] * (n + 1)
    interest_sn = [0.0] * (n + 1)
    interest_sub = [0.0] * (n + 1)
    interest_income = [0.0] * (n + 1)
    total_interest = [0.0] * (n + 1)
    ebt = [0.0] * (n + 1)
    taxes = [0.0] * (n + 1)
    net_income = [0.0] * (n + 1)

    for i in range(1, n + 1):
        # 1. Interest (基于期初债务, 已在上一轮算好)
        interest_revolver[i] = p.revolver_rate * rev_beg[i]
        interest_tla[i] = p.tla_rate * tla_beg[i]
        interest_tlb[i] = p.tlb_rate * tlb_beg[i]
        interest_sn[i] = p.senior_notes_rate * sn_beg[i]
        interest_sub[i] = p.sub_rate * sub_beg[i]
        total_interest[i] = interest_revolver[i] + interest_tla[i] + interest_tlb[i] + interest_sn[i] + interest_sub[i]
        interest_income[i] = p.cash_interest_rate * beg_cash[i]

        # 2. IS
        ebt[i] = ebit[i] - total_interest[i] + interest_income[i]
        taxes[i] = max(0, ebt[i] * p.tax_rate)
        net_income[i] = ebt[i] - taxes[i]

        # 3. CFADS
        cfo = net_income[i] + da[i] + delta_nwc[i]
        cfads[i] = cfo - capex[i]  # CFADS = NI + D&A + ΔNWC - CapEx (注意 ΔNWC 已是负数)

        # 4. Debt Schedule 滚动
        cash_avail[i] = max(0, beg_cash[i] + cfads[i] - p.min_cash_balance)

        # Revolver repayment
        rev_repay[i] = min(rev_beg[i], cash_avail[i])
        rev_end[i] = rev_beg[i] - rev_repay[i]

        # TLA mandatory amort
        tla_mandatory[i] = min(tla_beg[i], p.term_loan_a * p.tla_mandatory_amort)
        tla_end[i] = tla_beg[i] - tla_mandatory[i]

        # TLB mandatory + sweep
        tlb_mandatory[i] = min(tlb_beg[i], p.term_loan_b * p.tlb_mandatory_amort)
        # Cash Sweep = MIN(beg - mandatory, MAX(0, cash_avail - rev_repay - tla_mandatory - tlb_mandatory) * sweep_pct)
        remaining_cash = max(0, cash_avail[i] - rev_repay[i] - tla_mandatory[i] - tlb_mandatory[i])
        tlb_sweep[i] = min(tlb_beg[i] - tlb_mandatory[i], remaining_cash * p.cash_sweep_pct)
        tlb_end[i] = tlb_beg[i] - tlb_mandatory[i] - tlb_sweep[i]

        # SrNotes & Sub: bullet (不变)
        sn_end[i] = sn_beg[i]
        sub_end[i] = sub_beg[i]

        # Totals
        total_debt[i] = rev_end[i] + tla_end[i] + tlb_end[i] + sn_end[i] + sub_end[i]
        # Cash End = MAX(min_cash, beg_cash + cfads - rev_repay - tla_mandatory - tlb_mandatory - tlb_sweep)
        cash_end[i] = max(p.min_cash_balance,
                          beg_cash[i] + cfads[i] - rev_repay[i] - tla_mandatory[i] - tlb_mandatory[i] - tlb_sweep[i])

        # 准备下一年期初
        if i + 1 <= n:
            beg_cash[i + 1] = cash_end[i]
            rev_beg[i + 1] = rev_end[i]
            tla_beg[i + 1] = tla_end[i]
            tlb_beg[i + 1] = tlb_end[i]
            sn_beg[i + 1] = sn_end[i]
            sub_beg[i + 1] = sub_end[i]

    # ========== Exit & Returns ==========
    exit_year = min(p.exit_year, n)
    exit_ebitda = ebitda[exit_year]
    exit_ev = exit_ebitda * p.exit_ev_ebitda
    exit_equity = exit_ev - total_debt[exit_year] + cash_end[exit_year]
    moic = _safe_div(exit_equity, p.sponsor_equity)

    # IRR: Year 0 = -Sponsor; Year 1..exit-1 = 0; Year exit = exit_equity; after = 0
    # Excel IRR range = B16:{exit_col}16
    irr_cfs = [0.0] * (n + 1)
    irr_cfs[0] = -p.sponsor_equity
    for i in range(1, n + 1):
        if i == exit_year:
            irr_cfs[i] = exit_equity
        elif i < exit_year:
            irr_cfs[i] = 0.0
        else:
            irr_cfs[i] = 0.0  # after exit, blank in Excel but we use 0 for IRR calc
    irr = calculate_irr(irr_cfs[:exit_year + 1])

    # ========== 汇总 ==========
    ebitda_margin = [0.0] + [_safe_div(ebitda[i], revenue[i]) for i in range(1, n + 1)]

    # ========== 自定义行项求值 ==========
    def _build_lbo_context(i):
        return {
            "revenue": revenue[i], "ebitda": ebitda[i], "da": da[i], "ebit": ebit[i],
            "interest_revolver": interest_revolver[i], "interest_tla": interest_tla[i],
            "interest_tlb": interest_tlb[i], "interest_sn": interest_sn[i],
            "interest_sub": interest_sub[i], "total_interest": total_interest[i],
            "interest_income": interest_income[i], "ebt": ebt[i], "taxes": taxes[i],
            "net_income": net_income[i], "total_debt": total_debt[i],
            "cash": cash_end[i], "revolver": rev_end[i], "tla": tla_end[i],
            "tlb": tlb_end[i], "senior_notes": sn_end[i],
            "subordinated_debt": sub_end[i], "cfads": cfads[i],
        }
    contexts = [_build_lbo_context(i) for i in range(n + 1)]
    custom_results = _eval_custom_items(p.custom_items, years, contexts)

    return {
        "model_type": "lbo",
        "company_name": p.company_name,
        "years": years,
        "sources_uses": {
            "entry_ev": _round2(entry_ev),
            "purchase_equity": _round2(purchase_equity),
            "refinance_debt": _round2(refi_debt),
            "transaction_fees": _round2(trans_fees),
            "financing_fees": _round2(fin_fees),
            "total_uses": _round2(total_uses),
            "sponsor_equity": _round2(p.sponsor_equity),
            "revolver_draw": _round2(revolver_draw),
            "term_loan_a": _round2(p.term_loan_a),
            "term_loan_b": _round2(p.term_loan_b),
            "senior_notes": _round2(p.senior_notes),
            "subordinated_debt": _round2(p.subordinated_debt),
            "existing_cash": _round2(p.existing_cash),
            "total_sources": _round2(total_sources),
            "balance_check": "BALANCED" if su_balance else "IMBALANCED",
        },
        "income_statement": {
            "revenue": [_round2(x) for x in revenue],
            "ebitda": [_round2(x) for x in ebitda],
            "da": [_round2(x) for x in da],
            "ebit": [_round2(x) for x in ebit],
            "interest_revolver": [_round2(x) for x in interest_revolver],
            "interest_tla": [_round2(x) for x in interest_tla],
            "interest_tlb": [_round2(x) for x in interest_tlb],
            "interest_sn": [_round2(x) for x in interest_sn],
            "interest_sub": [_round2(x) for x in interest_sub],
            "total_interest": [_round2(x) for x in total_interest],
            "interest_income": [_round2(x) for x in interest_income],
            "ebt": [_round2(x) for x in ebt],
            "taxes": [_round2(x) for x in taxes],
            "net_income": [_round2(x) for x in net_income],
        },
        "debt_schedule": {
            "beg_cash": [_round2(x) for x in beg_cash],
            "cfads": [_round2(x) for x in cfads],
            "cash_available": [_round2(x) for x in cash_avail],
            "revolver_beg": [_round2(x) for x in rev_beg],
            "revolver_repay": [_round2(x) for x in rev_repay],
            "revolver_end": [_round2(x) for x in rev_end],
            "tla_beg": [_round2(x) for x in tla_beg],
            "tla_mandatory": [_round2(x) for x in tla_mandatory],
            "tla_end": [_round2(x) for x in tla_end],
            "tlb_beg": [_round2(x) for x in tlb_beg],
            "tlb_mandatory": [_round2(x) for x in tlb_mandatory],
            "tlb_sweep": [_round2(x) for x in tlb_sweep],
            "tlb_end": [_round2(x) for x in tlb_end],
            "sn_beg": [_round2(x) for x in sn_beg],
            "sn_end": [_round2(x) for x in sn_end],
            "sub_beg": [_round2(x) for x in sub_beg],
            "sub_end": [_round2(x) for x in sub_end],
            "total_debt": [_round2(x) for x in total_debt],
            "cash_end": [_round2(x) for x in cash_end],
        },
        "exit_returns": {
            "exit_year": exit_year,
            "exit_ebitda": _round2(exit_ebitda),
            "exit_ev_ebitda": p.exit_ev_ebitda,
            "exit_ev": _round2(exit_ev),
            "total_debt_at_exit": _round2(total_debt[exit_year]),
            "cash_at_exit": _round2(cash_end[exit_year]),
            "equity_value": _round2(exit_equity),
            "sponsor_equity_invested": _round2(p.sponsor_equity),
            "moic": _round2(moic),
            "irr": _round2(irr),
            "irr_display": f"{irr * 100:.1f}%",
        },
        "key_metrics": {
            "ebitda_margin": [_round2(x) for x in ebitda_margin],
        },
        "custom_items": custom_results,
    }


# ==================== 分发函数 ====================

def calculate(model_type: str, params: dict) -> Dict[str, Any]:
    """根据模型类型分发到对应计算函数。"""
    if model_type == "three_statement":
        p = ThreeStatementParams(**params)
        return calculate_three_statement(p)
    elif model_type == "dcf":
        p = DCFParams(**params)
        return calculate_dcf(p)
    elif model_type == "lbo":
        p = LBOParams(**params)
        return calculate_lbo(p)
    else:
        raise ValueError(f"Unknown model type: {model_type}")
