"""财报上传解析器。

支持 CSV / Excel (.xlsx) 格式，通过中英文关键词模糊匹配 schema 字段。
返回 {matched: {...}, unmatched: [...], warnings: [...]} 让用户确认。

匹配策略:
1. 读取所有单元格，构建 (label, value) 对列表
2. 对每个 schema 字段，用 FIELD_KEYWORDS 的关键词列表做包含匹配
3. 百分比字段自动识别 "25%" / "0.25" / "25" 三种格式（>1 当作百分比）
4. 多个候选时选置信度最高的，其余放入 warnings
"""
import io
import csv
import re
from typing import Dict, Any, List, Tuple

from openpyxl import load_workbook

from backend.schemas import ThreeStatementParams, DCFParams, LBOParams


# ==================== 字段关键词映射 ====================
# key = schema 字段名, value = 关键词列表（按优先级排序）
# 匹配时用小写包含判断
FIELD_KEYWORDS: Dict[str, List[str]] = {
    # 公司信息
    "company_name": ["company", "公司名称", "公司", "entity name", "企业名称"],
    "projection_years": ["projection years", "预测年数", "预测年限", "years"],
    # 收入
    "revenue_y0": ["revenue y0", "base revenue", "基期收入", "year 0 revenue", "revenue (base)", "revenue-0", "initial revenue"],
    "ltm_revenue": ["ltm revenue", "ltm 收入", "last twelve months revenue", "滚动12月收入"],
    # 增长率 (5 年)
    "rev_growth_y1": ["revenue growth y1", "收入增长率 y1", "revenue growth year 1", "rev growth y1", "增长 y1", "y1 growth"],
    "rev_growth_y2": ["revenue growth y2", "收入增长率 y2", "revenue growth year 2", "rev growth y2", "增长 y2", "y2 growth"],
    "rev_growth_y3": ["revenue growth y3", "收入增长率 y3", "revenue growth year 3", "rev growth y3", "增长 y3", "y3 growth"],
    "rev_growth_y4": ["revenue growth y4", "收入增长率 y4", "revenue growth year 4", "rev growth y4", "增长 y4", "y4 growth"],
    "rev_growth_y5": ["revenue growth y5", "收入增长率 y5", "revenue growth year 5", "rev growth y5", "增长 y5", "y5 growth"],
    # 利润率
    "cogs_pct": ["cogs", "营业成本", "cost of revenue", "cost of sales", "成本率", "主营业务成本"],
    "sga_pct": ["sg&a", "销管费用", "销售管理费用", "selling general", "销售费用"],
    "rd_pct": ["r&d", "研发费用", "research and development", "研发投入", "研发"],
    "da_pct": ["d&a", "折旧摊销", "depreciation", "amortization", "折旧", "摊销"],
    "capex_pct": ["capex", "资本支出", "capital expenditure", "资本性支出"],
    "nwc_pct": ["nwc", "净营运资本", "networking capital", "营运资金"],
    "ebitda_margin": ["ebitda margin", "ebitda 利润率", "ebitda率", "ebitda margin %"],
    "ebitda_margin_y1": ["ebitda margin y1", "ebitda 利润率 y1", "ebitda率 y1"],
    "ebitda_margin_y2": ["ebitda margin y2", "ebitda 利润率 y2", "ebitda率 y2"],
    "ebitda_margin_y3": ["ebitda margin y3", "ebitda 利润率 y3", "ebitda率 y3"],
    "ebitda_margin_y4": ["ebitda margin y4", "ebitda 利润率 y4", "ebitda率 y4"],
    "ebitda_margin_y5": ["ebitda margin y5", "ebitda 利润率 y5", "ebitda率 y5"],
    "ebitda": ["ebitda", "息税折旧摊销前利润"],
    "ltm_ebitda": ["ltm ebitda", "ltm 息税折旧摊销前利润", "last twelve months ebitda"],
    # 利率/税率
    "interest_rate": ["interest rate", "债务利率", "借款利率", "贷款利率"],
    "tax_rate": ["tax rate", "税率", "所得税率", "所得税", "effective tax"],
    "dividend_pct": ["dividend", "股利", "分红", "payout", "分配率"],
    "cash_interest_rate": ["cash interest", "现金存款利率", "存款利率"],
    # 营运资本天数
    "dso": ["dso", "应收天数", "days sales outstanding", "days sales", "应收账款天数"],
    "dio": ["dio", "库存天数", "days inventory", "days inventory outstanding", "存货天数"],
    "dpo": ["dpo", "应付天数", "days payable", "days payable outstanding", "应付账款天数"],
    "accrued_days": ["accrued days", "应计费用天数", "accrued"],
    # 期初资产负债表
    "beg_cash": ["beginning cash", "期初现金", "cash (beg)", "期初货币资金"],
    "beg_ar": ["beginning ar", "期初应收", "accounts receivable (beg)", "期初应收账款"],
    "beg_inventory": ["beginning inventory", "期初库存", "期初存货", "inventory (beg)"],
    "beg_ppe": ["beginning ppe", "期初固定资产", "ppe (beg)", "期初pp&e"],
    "beg_ap": ["beginning ap", "期初应付", "accounts payable (beg)", "期初应付账款"],
    "beg_accrued": ["beginning accrued", "期初应计", "accrued (beg)", "期初应计费用"],
    "beg_debt": ["beginning debt", "期初债务", "debt (beg)", "期初负债"],
    "common_stock": ["common stock", "普通股", "股本"],
    "beg_retn_earn": ["retained earnings", "留存收益", "re (beg)", "期初留存收益", "期初未分配利润"],
    "new_debt_issuance": ["new debt", "新增债务", "新增借款", "debt issuance"],
    # WACC
    "risk_free_rate": ["risk-free rate", "无风险利率", "risk free rate", "rf", "无风险"],
    "equity_risk_premium": ["equity risk premium", "股权风险溢价", "erp", "风险溢价"],
    "beta": ["beta", "贝塔", "β"],
    "pre_tax_cost_of_debt": ["pre-tax cost of debt", "税前债务成本", "cost of debt", "债务成本"],
    "debt_weight": ["debt weight", "债务权重", "target debt", "目标债务"],
    # 终值
    "terminal_growth": ["terminal growth", "终值增长率", "gordon growth", "永续增长率"],
    "exit_multiple": ["exit multiple", "退出倍数", "exit ev/ebitda", "退出 ev/ebitda"],
    "tv_method": ["tv method", "终值方法", "terminal value method"],
    # EV 桥
    "net_debt": ["net debt", "净债务", "总债务"],
    "cash": ["cash", "现金", "货币资金"],
    "minority_interest": ["minority interest", "少数股东权益", "少数股东"],
    "preferred_stock": ["preferred stock", "优先股"],
    "stock_based_comp": ["stock-based comp", "股权激励", "sbc", "股份支付", "股权支付"],
    "pension_deficit": ["pension deficit", "养老金缺口", "养老金"],
    "shares_outstanding": ["shares outstanding", "流通股本", "总股本", "股本数"],
    "current_price": ["current price", "当前股价", "现价", "股价"],
    "valuation_timing": ["valuation timing", "估值时点", "timing"],
    # LBO 交易假设
    "entry_ev_ebitda": ["entry ev/ebitda", "入场倍数", "入场 ev/ebitda", "entry multiple"],
    "existing_net_debt": ["existing net debt", "现有净债务", "现有债务"],
    "existing_cash": ["existing cash", "现有现金"],
    "transaction_fees": ["transaction fees", "交易费用", "交易费"],
    "financing_fees": ["financing fees", "融资费用", "融资费"],
    # LBO 资本结构
    "sponsor_equity": ["sponsor equity", "sponsor 股权", "基金股权", "股本投入"],
    "revolver_capacity": ["revolver capacity", "revolver 额度", "循环额度"],
    "term_loan_a": ["term loan a", "tla", "term loan a"],
    "term_loan_b": ["term loan b", "tlb", "term loan b"],
    "senior_notes": ["senior notes", "高级票据", "优先票据"],
    "subordinated_debt": ["subordinated debt", "次级债务", "次级债"],
    # LBO 利率
    "revolver_rate": ["revolver rate", "revolver 利率"],
    "tla_rate": ["tla rate", "term loan a 利率", "tla 利率"],
    "tlb_rate": ["tlb rate", "term loan b 利率", "tlb 利率"],
    "senior_notes_rate": ["senior notes rate", "senior notes 利率", "优先票据利率"],
    "sub_rate": ["sub rate", "subordinated 利率", "次级债务利率"],
    # LBO 偿债
    "tla_mandatory_amort": ["tla mandatory amort", "tla 强制摊销"],
    "tlb_mandatory_amort": ["tlb mandatory amort", "tlb 强制摊销"],
    "cash_sweep_pct": ["cash sweep", "现金清偿", "清偿率"],
    "min_cash_balance": ["min cash", "最低现金", "最低现金余额"],
    # LBO 退出
    "exit_ev_ebitda": ["exit ev/ebitda", "退出倍数", "exit multiple"],
    "exit_year": ["exit year", "退出年"],
}

# 百分比类字段（值需在 0-1 之间或自动转换）
PERCENT_FIELDS = {
    "cogs_pct", "sga_pct", "rd_pct", "da_pct", "capex_pct", "nwc_pct",
    "interest_rate", "tax_rate", "dividend_pct", "cash_interest_rate",
    "rev_growth_y1", "rev_growth_y2", "rev_growth_y3", "rev_growth_y4", "rev_growth_y5",
    "ebitda_margin", "ebitda_margin_y1", "ebitda_margin_y2", "ebitda_margin_y3",
    "ebitda_margin_y4", "ebitda_margin_y5",
    "risk_free_rate", "equity_risk_premium", "pre_tax_cost_of_debt", "debt_weight",
    "terminal_growth", "tla_mandatory_amort", "tlb_mandatory_amort", "cash_sweep_pct",
    "revolver_rate", "tla_rate", "tlb_rate", "senior_notes_rate", "sub_rate",
}


# ==================== Schema 字段列表 ====================

def get_schema_fields(model_type: str) -> Dict[str, Any]:
    """返回某模型的 schema 字段名 → 默认值/类型信息。"""
    if model_type == "three_statement":
        schema = ThreeStatementParams
    elif model_type == "dcf":
        schema = DCFParams
    elif model_type == "lbo":
        schema = LBOParams
    else:
        raise ValueError(f"Unknown model type: {model_type}")
    fields = {}
    for name, info in schema.model_fields.items():
        fields[name] = info.default
    return fields


# ==================== 文件读取 ====================

def _read_csv(content: bytes) -> List[Tuple[str, Any]]:
    """读取 CSV，返回 (label, value) 对列表。"""
    pairs = []
    try:
        text = content.decode("utf-8-sig")  # 处理 BOM
    except UnicodeDecodeError:
        text = content.decode("gbk", errors="ignore")
    reader = csv.reader(text.splitlines())
    for row in reader:
        if not row or len(row) < 2:
            continue
        label = str(row[0]).strip()
        # 找第一个非空值
        value = None
        for cell in row[1:]:
            if cell and cell.strip():
                value = cell.strip()
                break
        if label and value is not None:
            pairs.append((label, value))
    return pairs


def _read_excel(content: bytes) -> List[Tuple[str, Any]]:
    """读取 Excel，返回 (label, value) 对列表。

    支持两种布局:
    1. 标签在 A 列，值在 B 列（同 schema 布局）
    2. 标签在某行，值在下一行（横向布局）
    """
    pairs = []
    wb = load_workbook(io.BytesIO(content), data_only=True, read_only=True)
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            if not row or len(row) < 2:
                continue
            label = row[0]
            if label is None:
                continue
            label = str(label).strip()
            # 找第一个非空值
            value = None
            for cell in row[1:]:
                if cell is not None and str(cell).strip():
                    value = str(cell).strip()
                    break
            if label and value is not None:
                pairs.append((label, value))
    wb.close()
    return pairs


# ==================== 值解析 ====================

def _parse_number(value: str, is_percent: bool) -> Tuple[float, str]:
    """解析数值。返回 (value, format_note)。

    百分比字段:
    - "25%" → 0.25
    - "0.25" → 0.25 (已在小数范围)
    - "25" → 0.25 (大于 1，视为百分比)
    - "2.5" → 0.025 (大于 1，视为百分比)
    """
    s = value.strip().replace(",", "").replace("，", "")
    # 去除货币符号
    s = re.sub(r"[¥$€£]", "", s).strip()

    if is_percent:
        if "%" in s:
            s = s.replace("%", "").strip()
            try:
                v = float(s) / 100
                return v, "百分比(含%)"
            except ValueError:
                return (0.0, "解析失败")
        try:
            v = float(s)
            if abs(v) > 1:  # 大于1视为百分比
                return v / 100, "百分比(>1转换)"
            return v, "小数"
        except ValueError:
            return 0.0, "解析失败"
    else:
        try:
            return float(s), "数值"
        except ValueError:
            return 0.0, "解析失败"


# ==================== 匹配逻辑 ====================

def _match_field(field_name: str, pairs: List[Tuple[str, Any]]) -> Tuple[Any, str, List[str]]:
    """对单个字段做关键词匹配。

    返回 (value, matched_label, warnings)
    """
    keywords = FIELD_KEYWORDS.get(field_name, [])
    if not keywords:
        return (None, "", [])

    candidates = []  # [(label, value, keyword, score)]
    for label, value in pairs:
        label_lower = label.lower()
        for kw in keywords:
            kw_lower = kw.lower()
            if kw_lower in label_lower:
                # 完全匹配优先级最高；包含匹配次之
                score = 100 if label_lower.strip() == kw_lower else 50 + len(kw_lower)
                candidates.append((label, value, kw, score))
                break

    if not candidates:
        return (None, "", [])

    # 按分数降序
    candidates.sort(key=lambda x: -x[3])
    best = candidates[0]
    warnings = []
    if len(candidates) > 1:
        other_labels = [c[0] for c in candidates[1:3]]
        warnings.append(f"字段 '{field_name}' 找到 {len(candidates)} 个候选，已选 '{best[0]}'，其他: {other_labels}")

    return (best[1], best[0], warnings)


# ==================== 主入口 ====================

def parse_financial_report(file_bytes: bytes, filename: str, model_type: str) -> Dict[str, Any]:
    """解析上传的财报，返回映射结果。

    Returns:
        {
            "matched": {field_name: value, ...},  # 成功映射的字段
            "unmatched": ["field1", ...],         # 未找到的字段
            "warnings": ["...", ...],             # 多候选等警告
            "model_type": "three_statement",
            "filename": "report.csv",
        }
    """
    # 1. 读取文件
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    if ext == "csv":
        pairs = _read_csv(file_bytes)
    elif ext in ("xlsx", "xls"):
        pairs = _read_excel(file_bytes)
    else:
        raise ValueError(f"Unsupported file extension: {ext}")

    if not pairs:
        raise ValueError("文件为空或无法解析，请检查格式（需要标签列 + 值列）")

    # 2. 遍历 schema 字段做匹配
    schema_fields = get_schema_fields(model_type)
    matched = {}
    unmatched = []
    all_warnings = []

    for field_name, default_val in schema_fields.items():
        # 文本字段（company_name, tv_method, valuation_timing）特殊处理
        is_text = isinstance(default_val, str)
        is_int = isinstance(default_val, int) and not isinstance(default_val, bool)

        raw_value, matched_label, warnings = _match_field(field_name, pairs)
        all_warnings.extend(warnings)

        if raw_value is None:
            unmatched.append(field_name)
            continue

        if is_text:
            matched[field_name] = raw_value
        elif is_int:
            try:
                matched[field_name] = int(float(str(raw_value).replace(",", "")))
            except ValueError:
                unmatched.append(field_name)
                all_warnings.append(f"字段 '{field_name}' 值 '{raw_value}' 无法转整数")
        else:
            is_percent = field_name in PERCENT_FIELDS
            num_val, note = _parse_number(str(raw_value), is_percent)
            if note == "解析失败":
                unmatched.append(field_name)
                all_warnings.append(f"字段 '{field_name}' 值 '{raw_value}' 解析数值失败")
            else:
                matched[field_name] = round(num_val, 4)

    # 3. 返回结果
    return {
        "matched": matched,
        "unmatched": unmatched,
        "warnings": all_warnings,
        "model_type": model_type,
        "filename": filename,
        "total_fields": len(schema_fields),
        "matched_count": len(matched),
    }
