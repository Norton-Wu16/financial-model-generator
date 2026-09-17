/* app.js - 纯前端实现, 1:1 翻译后端 Python 逻辑
 * 所有后端 API 端点的 JavaScript 等价实现
 * 暴露到 window 全局供 index.html 调用
 */

// ==================== 1. 工具函数 ====================

function _round2(x) {
  if (x === null || x === undefined || isNaN(x)) return 0.0;
  return Math.round(parseFloat(x) * 100) / 100;
}

function _safeDiv(a, b) {
  if (b === 0) return 0.0;
  return a / b;
}

// 1-based 列号 → Excel 列字母 (1=A, 2=B, ..., 27=AA)
function ycol(yearIndex) {
  var col = 2 + yearIndex; // yearIndex 0 → B, 1 → C, ...
  var result = "";
  while (col > 0) {
    var mod = (col - 1) % 26;
    result = String.fromCharCode(65 + mod) + result;
    col = Math.floor((col - 1) / 26);
  }
  return result;
}

// 将 aoa 转为 SheetJS worksheet, 公式字符串(=开头)转为公式单元格
function aoaToSheet(aoa) {
  var ws = XLSX.utils.aoa_to_sheet(aoa);
  if (!ws['!ref']) return ws;
  var range = XLSX.utils.decode_range(ws['!ref']);
  for (var r = range.s.r; r <= range.e.r; r++) {
    for (var c = range.s.c; c <= range.e.c; c++) {
      var addr = XLSX.utils.encode_cell({ r: r, c: c });
      var cell = ws[addr];
      if (cell && typeof cell.v === 'string' && cell.v.charAt(0) === '=') {
        ws[addr] = { t: 'f', f: cell.v.substring(1) };
      }
    }
  }
  return ws;
}


// ==================== 1b. 参数归一化与年度取值 helper ====================

function _num(v, d) { var n = parseFloat(v); return isNaN(n) ? (d === undefined ? 0 : d) : n; }

// 合并默认值，保证新增字段始终存在；年数限制 1~10
function normalizeParams(modelType, p) {
  var base = jsDefaults[modelType] || {};
  var m = Object.assign({}, base, p || {});
  if (modelType === "lbo") {
    m.exit_year = Math.max(1, Math.min(10, parseInt(m.exit_year, 10) || 5));
  } else {
    m.projection_years = Math.max(1, Math.min(10, parseInt(m.projection_years, 10) || 5));
  }
  // 税率口径：仅接受 etr / mtr，未提供时沿用模型默认
  if (m.tax_basis !== "etr" && m.tax_basis !== "mtr") m.tax_basis = base.tax_basis || "etr";
  if (modelType === "dcf") {
    m.debt_weight = _num(m.debt_weight, base.debt_weight);
    m.preferred_weight = _num(m.preferred_weight, base.preferred_weight);
    m.cost_of_preferred = _num(m.cost_of_preferred, base.cost_of_preferred);
  }
  return m;
}

// 年度收入增长率 rev_growth_y{i}；未填返回 0
function growthYear(p, i) {
  var v = parseFloat(p["rev_growth_y" + i]);
  return isNaN(v) ? 0 : v;
}

// 年度比率：优先 <baseKey>_y<i>（按年细化值），否则用统一值 baseKey，再否则 0
function rateYear(p, baseKey, i) {
  var yv = parseFloat(p[baseKey + "_y" + i]);
  if (!isNaN(yv)) return yv;
  var uv = parseFloat(p[baseKey]);
  return isNaN(uv) ? 0 : uv;
}

// DCF EBITDA 利润率 ebitda_margin_y{i}（始终按年）；越界沿用 Y5
function marginYear(p, i) {
  var v = parseFloat(p["ebitda_margin_y" + i]);
  if (!isNaN(v)) return v;
  var last = parseFloat(p.ebitda_margin_y5);
  return isNaN(last) ? 0 : last;
}

// DCF 三层资本结构：普通股 We = 1 - Wd - Wp
function dcfWeights(p) {
  var wd = _num(p.debt_weight, 0);
  var wp = _num(p.preferred_weight, 0);
  return { wd: wd, wp: wp, we: 1 - wd - wp };
}

// 权重非法（普通股为负）
function waccInvalid(p) {
  return dcfWeights(p).we < -1e-6;
}


// ==================== 2. jsDefaults (schemas.py 默认值) ====================

var jsDefaults = {
  three_statement: {
    company_name: "示例公司 Sample Co.",
    projection_years: 5,
    revenue_y0: 100.0,
    rev_growth_y1: 0.10,
    rev_growth_y2: 0.08,
    rev_growth_y3: 0.06,
    rev_growth_y4: 0.05,
    rev_growth_y5: 0.04,
    cogs_pct: 0.60,
    sga_pct: 0.15,
    rd_pct: 0.05,
    da_pct: 0.05,
    interest_rate: 0.06,
    tax_rate: 0.25,
    dividend_pct: 0.30,
    capex_pct: 0.08,
    dso: 45.0,
    dio: 60.0,
    dpo: 30.0,
    accrued_days: 20.0,
    beg_cash: 20.0,
    beg_ar: 12.0,
    beg_inventory: 16.0,
    beg_ppe: 80.0,
    beg_ap: 8.0,
    beg_accrued: 5.0,
    beg_debt: 50.0,
    common_stock: 30.0,
    beg_retn_earn: 35.0,
    new_debt_issuance: 0.0,
    tax_basis: "etr",
  },
  dcf: {
    company_name: "示例公司 Sample Co.",
    projection_years: 5,
    revenue_y0: 100.0,
    rev_growth_y1: 0.15,
    rev_growth_y2: 0.12,
    rev_growth_y3: 0.10,
    rev_growth_y4: 0.08,
    rev_growth_y5: 0.06,
    ebitda_margin_y1: 0.25,
    ebitda_margin_y2: 0.25,
    ebitda_margin_y3: 0.25,
    ebitda_margin_y4: 0.25,
    ebitda_margin_y5: 0.25,
    da_pct: 0.05,
    capex_pct: 0.06,
    nwc_pct: 0.10,
    tax_rate: 0.25,
    risk_free_rate: 0.04,
    equity_risk_premium: 0.05,
    beta: 1.20,
    pre_tax_cost_of_debt: 0.06,
    debt_weight: 0.30,
    preferred_weight: 0.05,
    cost_of_preferred: 0.08,
    terminal_growth: 0.03,
    exit_multiple: 10.0,
    tv_method: "gordon",
    net_debt: 20.0,
    cash: 30.0,
    minority_interest: 0.0,
    preferred_stock: 0.0,
    stock_based_comp: 5.0,
    pension_deficit: 0.0,
    shares_outstanding: 50.0,
    current_price: 20.0,
    valuation_timing: "mid",
    tax_basis: "mtr",
  },
  lbo: {
    company_name: "示例公司 Sample Co.",
    ltm_revenue: 250.0,
    ltm_ebitda: 50.0,
    entry_ev_ebitda: 10.0,
    existing_net_debt: 100.0,
    existing_cash: 5.0,
    transaction_fees: 10.0,
    financing_fees: 15.0,
    sponsor_equity: 150.0,
    revolver_capacity: 50.0,
    term_loan_a: 100.0,
    term_loan_b: 150.0,
    senior_notes: 75.0,
    subordinated_debt: 25.0,
    revolver_rate: 0.04,
    tla_rate: 0.04,
    tlb_rate: 0.05,
    senior_notes_rate: 0.07,
    sub_rate: 0.10,
    tla_mandatory_amort: 0.05,
    tlb_mandatory_amort: 0.02,
    cash_sweep_pct: 0.50,
    min_cash_balance: 10.0,
    rev_growth_y1: 0.08,
    rev_growth_y2: 0.07,
    rev_growth_y3: 0.06,
    rev_growth_y4: 0.05,
    rev_growth_y5: 0.04,
    ebitda_margin: 0.20,
    da_pct: 0.04,
    capex_pct: 0.05,
    nwc_pct: 0.08,
    tax_rate: 0.25,
    exit_ev_ebitda: 10.0,
    exit_year: 5,
    cash_interest_rate: 0.02,
    tax_basis: "mtr",
  },
};


// ==================== 3. jsPresets (presets.py) ====================

var jsPresets = {
  tech_saas: {
    name: "科技/SaaS", name_en: "Tech / SaaS", icon: "💻",
    description: "高增长(20-30%)、高毛利(70%+)、轻资产、高研发投入、高估值倍数。代表: 软件公司",
    params: {
      three_statement: {
        company_name: "TechCo 科技公司", projection_years: 5, revenue_y0: 200.0,
        rev_growth_y1: 0.30, rev_growth_y2: 0.25, rev_growth_y3: 0.20, rev_growth_y4: 0.15, rev_growth_y5: 0.12,
        cogs_pct: 0.30, sga_pct: 0.25, rd_pct: 0.15, da_pct: 0.03, interest_rate: 0.04, tax_rate: 0.20,
        dividend_pct: 0.0, capex_pct: 0.03, dso: 45.0, dio: 15.0, dpo: 30.0, accrued_days: 25.0,
        beg_cash: 80.0, beg_ar: 25.0, beg_inventory: 8.0, beg_ppe: 30.0, beg_ap: 15.0, beg_accrued: 20.0,
        beg_debt: 20.0, common_stock: 100.0, beg_retn_earn: 50.0, new_debt_issuance: 0.0,
      },
      dcf: {
        company_name: "TechCo 科技公司", projection_years: 5, revenue_y0: 200.0,
        rev_growth_y1: 0.30, rev_growth_y2: 0.25, rev_growth_y3: 0.20, rev_growth_y4: 0.15, rev_growth_y5: 0.12,
        ebitda_margin_y1: 0.35, ebitda_margin_y2: 0.38, ebitda_margin_y3: 0.40, ebitda_margin_y4: 0.42, ebitda_margin_y5: 0.45,
        da_pct: 0.03, capex_pct: 0.03, nwc_pct: 0.05, tax_rate: 0.20,
        risk_free_rate: 0.04, equity_risk_premium: 0.05, beta: 1.30, pre_tax_cost_of_debt: 0.05, debt_weight: 0.10,
        terminal_growth: 0.03, exit_multiple: 18.0, tv_method: "gordon",
        net_debt: 10.0, cash: 80.0, minority_interest: 0.0, preferred_stock: 0.0,
        stock_based_comp: 8.0, pension_deficit: 0.0, shares_outstanding: 50.0, current_price: 25.0, valuation_timing: "mid",
      },
      lbo: {
        company_name: "TechCo 科技公司", ltm_revenue: 200.0, ltm_ebitda: 70.0, entry_ev_ebitda: 12.0,
        existing_net_debt: 20.0, existing_cash: 80.0, transaction_fees: 8.0, financing_fees: 12.0,
        sponsor_equity: 300.0, revolver_capacity: 50.0, term_loan_a: 150.0, term_loan_b: 200.0,
        senior_notes: 100.0, subordinated_debt: 50.0,
        revolver_rate: 0.045, tla_rate: 0.045, tlb_rate: 0.055, senior_notes_rate: 0.07, sub_rate: 0.10,
        tla_mandatory_amort: 0.05, tlb_mandatory_amort: 0.02, cash_sweep_pct: 0.50, min_cash_balance: 15.0,
        rev_growth_y1: 0.15, rev_growth_y2: 0.12, rev_growth_y3: 0.10, rev_growth_y4: 0.08, rev_growth_y5: 0.06,
        ebitda_margin: 0.35, da_pct: 0.03, capex_pct: 0.03, nwc_pct: 0.05, tax_rate: 0.20,
        exit_ev_ebitda: 12.0, exit_year: 5, cash_interest_rate: 0.02,
      },
    },
  },
  retail: {
    name: "零售/消费品", name_en: "Retail / Consumer", icon: "🛒",
    description: "中低增长(5-8%)、中等毛利(35-40%净利)、高营运资本需求、重渠道。代表: 实体零售、快消",
    params: {
      three_statement: {
        company_name: "RetailCo 零售公司", projection_years: 5, revenue_y0: 500.0,
        rev_growth_y1: 0.08, rev_growth_y2: 0.07, rev_growth_y3: 0.06, rev_growth_y4: 0.05, rev_growth_y5: 0.04,
        cogs_pct: 0.65, sga_pct: 0.22, rd_pct: 0.01, da_pct: 0.03, interest_rate: 0.05, tax_rate: 0.25,
        dividend_pct: 0.40, capex_pct: 0.04, dso: 7.0, dio: 60.0, dpo: 45.0, accrued_days: 15.0,
        beg_cash: 50.0, beg_ar: 10.0, beg_inventory: 85.0, beg_ppe: 200.0, beg_ap: 40.0, beg_accrued: 25.0,
        beg_debt: 100.0, common_stock: 150.0, beg_retn_earn: 120.0, new_debt_issuance: 0.0,
      },
      dcf: {
        company_name: "RetailCo 零售公司", projection_years: 5, revenue_y0: 500.0,
        rev_growth_y1: 0.08, rev_growth_y2: 0.07, rev_growth_y3: 0.06, rev_growth_y4: 0.05, rev_growth_y5: 0.04,
        ebitda_margin_y1: 0.10, ebitda_margin_y2: 0.10, ebitda_margin_y3: 0.11, ebitda_margin_y4: 0.11, ebitda_margin_y5: 0.12,
        da_pct: 0.03, capex_pct: 0.04, nwc_pct: 0.08, tax_rate: 0.25,
        risk_free_rate: 0.04, equity_risk_premium: 0.05, beta: 0.90, pre_tax_cost_of_debt: 0.05, debt_weight: 0.30,
        terminal_growth: 0.02, exit_multiple: 8.0, tv_method: "exit",
        net_debt: 100.0, cash: 50.0, minority_interest: 0.0, preferred_stock: 0.0,
        stock_based_comp: 2.0, pension_deficit: 0.0, shares_outstanding: 80.0, current_price: 12.0, valuation_timing: "mid",
      },
      lbo: {
        company_name: "RetailCo 零售公司", ltm_revenue: 500.0, ltm_ebitda: 50.0, entry_ev_ebitda: 8.0,
        existing_net_debt: 100.0, existing_cash: 50.0, transaction_fees: 15.0, financing_fees: 20.0,
        sponsor_equity: 150.0, revolver_capacity: 60.0, term_loan_a: 120.0, term_loan_b: 150.0,
        senior_notes: 80.0, subordinated_debt: 30.0,
        revolver_rate: 0.04, tla_rate: 0.04, tlb_rate: 0.05, senior_notes_rate: 0.07, sub_rate: 0.10,
        tla_mandatory_amort: 0.05, tlb_mandatory_amort: 0.02, cash_sweep_pct: 0.50, min_cash_balance: 10.0,
        rev_growth_y1: 0.06, rev_growth_y2: 0.05, rev_growth_y3: 0.04, rev_growth_y4: 0.04, rev_growth_y5: 0.03,
        ebitda_margin: 0.10, da_pct: 0.03, capex_pct: 0.04, nwc_pct: 0.08, tax_rate: 0.25,
        exit_ev_ebitda: 9.0, exit_year: 5, cash_interest_rate: 0.02,
      },
    },
  },
  manufacturing: {
    name: "制造业", name_en: "Manufacturing", icon: "🏭",
    description: "低增长(3-5%)、低毛利(20-25%净利)、高CapEx、重资产、长营运资本周期。代表: 工业制造",
    params: {
      three_statement: {
        company_name: "ManufCo 制造公司", projection_years: 5, revenue_y0: 300.0,
        rev_growth_y1: 0.06, rev_growth_y2: 0.05, rev_growth_y3: 0.04, rev_growth_y4: 0.04, rev_growth_y5: 0.03,
        cogs_pct: 0.70, sga_pct: 0.12, rd_pct: 0.04, da_pct: 0.08, interest_rate: 0.055, tax_rate: 0.25,
        dividend_pct: 0.30, capex_pct: 0.12, dso: 55.0, dio: 90.0, dpo: 40.0, accrued_days: 20.0,
        beg_cash: 30.0, beg_ar: 45.0, beg_inventory: 74.0, beg_ppe: 250.0, beg_ap: 25.0, beg_accrued: 18.0,
        beg_debt: 120.0, common_stock: 100.0, beg_retn_earn: 180.0, new_debt_issuance: 0.0,
      },
      dcf: {
        company_name: "ManufCo 制造公司", projection_years: 5, revenue_y0: 300.0,
        rev_growth_y1: 0.06, rev_growth_y2: 0.05, rev_growth_y3: 0.04, rev_growth_y4: 0.04, rev_growth_y5: 0.03,
        ebitda_margin_y1: 0.18, ebitda_margin_y2: 0.18, ebitda_margin_y3: 0.19, ebitda_margin_y4: 0.19, ebitda_margin_y5: 0.20,
        da_pct: 0.08, capex_pct: 0.12, nwc_pct: 0.10, tax_rate: 0.25,
        risk_free_rate: 0.04, equity_risk_premium: 0.05, beta: 1.10, pre_tax_cost_of_debt: 0.055, debt_weight: 0.35,
        terminal_growth: 0.02, exit_multiple: 7.0, tv_method: "gordon",
        net_debt: 120.0, cash: 30.0, minority_interest: 0.0, preferred_stock: 0.0,
        stock_based_comp: 3.0, pension_deficit: 5.0, shares_outstanding: 60.0, current_price: 15.0, valuation_timing: "end",
      },
      lbo: {
        company_name: "ManufCo 制造公司", ltm_revenue: 300.0, ltm_ebitda: 54.0, entry_ev_ebitda: 7.0,
        existing_net_debt: 120.0, existing_cash: 30.0, transaction_fees: 12.0, financing_fees: 18.0,
        sponsor_equity: 150.0, revolver_capacity: 70.0, term_loan_a: 100.0, term_loan_b: 120.0,
        senior_notes: 60.0, subordinated_debt: 20.0,
        revolver_rate: 0.045, tla_rate: 0.045, tlb_rate: 0.055, senior_notes_rate: 0.075, sub_rate: 0.11,
        tla_mandatory_amort: 0.05, tlb_mandatory_amort: 0.02, cash_sweep_pct: 0.50, min_cash_balance: 12.0,
        rev_growth_y1: 0.05, rev_growth_y2: 0.04, rev_growth_y3: 0.04, rev_growth_y4: 0.03, rev_growth_y5: 0.03,
        ebitda_margin: 0.18, da_pct: 0.08, capex_pct: 0.12, nwc_pct: 0.10, tax_rate: 0.25,
        exit_ev_ebitda: 8.0, exit_year: 5, cash_interest_rate: 0.02,
      },
    },
  },
  financial_services: {
    name: "金融/银行", name_en: "Financial Services", icon: "🏛️",
    description: "特殊指标(净息差、拨备率)。注意: 当前三模型非银行业最佳适配，参数仅作参考映射（净息差→interest_rate, 拨备→rd_pct）",
    params: {
      three_statement: {
        company_name: "FinCo 金融公司", projection_years: 5, revenue_y0: 150.0,
        rev_growth_y1: 0.08, rev_growth_y2: 0.07, rev_growth_y3: 0.06, rev_growth_y4: 0.05, rev_growth_y5: 0.05,
        cogs_pct: 0.20, sga_pct: 0.30, rd_pct: 0.08, da_pct: 0.02, interest_rate: 0.03, tax_rate: 0.25,
        dividend_pct: 0.35, capex_pct: 0.02, dso: 30.0, dio: 5.0, dpo: 20.0, accrued_days: 30.0,
        beg_cash: 100.0, beg_ar: 12.0, beg_inventory: 2.0, beg_ppe: 40.0, beg_ap: 10.0, beg_accrued: 25.0,
        beg_debt: 200.0, common_stock: 120.0, beg_retn_earn: 80.0, new_debt_issuance: 0.0,
      },
      dcf: {
        company_name: "FinCo 金融公司", projection_years: 5, revenue_y0: 150.0,
        rev_growth_y1: 0.08, rev_growth_y2: 0.07, rev_growth_y3: 0.06, rev_growth_y4: 0.05, rev_growth_y5: 0.05,
        ebitda_margin_y1: 0.30, ebitda_margin_y2: 0.30, ebitda_margin_y3: 0.31, ebitda_margin_y4: 0.31, ebitda_margin_y5: 0.32,
        da_pct: 0.02, capex_pct: 0.02, nwc_pct: 0.03, tax_rate: 0.25,
        risk_free_rate: 0.04, equity_risk_premium: 0.06, beta: 0.95, pre_tax_cost_of_debt: 0.04, debt_weight: 0.40,
        terminal_growth: 0.02, exit_multiple: 10.0, tv_method: "gordon",
        net_debt: 200.0, cash: 100.0, minority_interest: 0.0, preferred_stock: 15.0,
        stock_based_comp: 3.0, pension_deficit: 0.0, shares_outstanding: 40.0, current_price: 18.0, valuation_timing: "end",
      },
      lbo: {
        company_name: "FinCo 金融公司", ltm_revenue: 150.0, ltm_ebitda: 45.0, entry_ev_ebitda: 9.0,
        existing_net_debt: 200.0, existing_cash: 100.0, transaction_fees: 8.0, financing_fees: 15.0,
        sponsor_equity: 200.0, revolver_capacity: 40.0, term_loan_a: 80.0, term_loan_b: 100.0,
        senior_notes: 60.0, subordinated_debt: 20.0,
        revolver_rate: 0.04, tla_rate: 0.04, tlb_rate: 0.05, senior_notes_rate: 0.065, sub_rate: 0.09,
        tla_mandatory_amort: 0.05, tlb_mandatory_amort: 0.02, cash_sweep_pct: 0.40, min_cash_balance: 20.0,
        rev_growth_y1: 0.06, rev_growth_y2: 0.05, rev_growth_y3: 0.05, rev_growth_y4: 0.04, rev_growth_y5: 0.04,
        ebitda_margin: 0.30, da_pct: 0.02, capex_pct: 0.02, nwc_pct: 0.03, tax_rate: 0.25,
        exit_ev_ebitda: 9.0, exit_year: 5, cash_interest_rate: 0.02,
      },
    },
  },
};

function jsListPresets() {
  var result = [];
  for (var pid in jsPresets) {
    var p = jsPresets[pid];
    result.push({ id: pid, name: p.name, name_en: p.name_en, icon: p.icon, description: p.description });
  }
  return result;
}

function jsGetPreset(pid, modelType) {
  if (!(pid in jsPresets)) throw new Error("Preset not found: " + pid);
  var preset = jsPresets[pid];
  if (!(modelType in preset.params)) throw new Error("Model type not in preset");
  // 与默认值合并：行业预设值覆盖默认，新增字段（tax_basis / preferred_weight 等）自动补全
  return normalizeParams(modelType, preset.params[modelType]);
}


// ==================== 4. jsCustomItems (custom_items.py) ====================

var WHITELIST_FUNCS = { MAX: true, MIN: true, ABS: true, ROUND: true, SUM: true };
var DANGEROUS_CHARS = "[]{}=:;\"'\\`~!@#$%&|<>?";

function tokenize(expr) {
  var tokens = [];
  var i = 0, n = expr.length;
  while (i < n) {
    var c = expr[i];
    if (c === ' ' || c === '\t' || c === '\n' || c === '\r') { i++; continue; }
    // 数字
    if (c >= '0' && c <= '9' || (c === '.' && i + 1 < n && expr[i+1] >= '0' && expr[i+1] <= '9')) {
      var j = i, dotSeen = false;
      while (j < n && (expr[j] >= '0' && expr[j] <= '9' || expr[j] === '.')) {
        if (expr[j] === '.') { if (dotSeen) break; dotSeen = true; }
        j++;
      }
      var numStr = expr.substring(i, j);
      var val = parseFloat(numStr);
      if (!dotSeen && /^\d+$/.test(numStr)) val = parseInt(numStr);
      tokens.push({ type: 'NUMBER', value: val, pos: i });
      i = j; continue;
    }
    // 标识符
    if (c >= 'a' && c <= 'z' || c >= 'A' && c <= 'Z' || c === '_') {
      var j2 = i;
      while (j2 < n && (expr[j2] >= 'a' && expr[j2] <= 'z' || expr[j2] >= 'A' && expr[j2] <= 'Z' || expr[j2] >= '0' && expr[j2] <= '9' || expr[j2] === '_')) j2++;
      tokens.push({ type: 'IDENT', value: expr.substring(i, j2), pos: i });
      i = j2; continue;
    }
    // 运算符
    if ('+-*/^'.indexOf(c) >= 0) { tokens.push({ type: 'OP', value: c, pos: i }); i++; continue; }
    if (c === '(') { tokens.push({ type: 'LPAREN', value: c, pos: i }); i++; continue; }
    if (c === ')') { tokens.push({ type: 'RPAREN', value: c, pos: i }); i++; continue; }
    if (c === ',') { tokens.push({ type: 'COMMA', value: c, pos: i }); i++; continue; }
    // 危险字符
    if (c === '.') throw new Error("非法字符 '.' (属性访问不被支持)，位置 " + i);
    if (DANGEROUS_CHARS.indexOf(c) >= 0) throw new Error("非法字符 '" + c + "'，位置 " + i);
    throw new Error("非法字符 '" + c + "'，位置 " + i);
  }
  tokens.push({ type: 'EOF', value: null, pos: n });
  return tokens;
}

function Parser(tokens) {
  this.tokens = tokens;
  this.pos = 0;
}
Parser.prototype.peek = function() { return this.tokens[this.pos]; };
Parser.prototype.advance = function() { var t = this.tokens[this.pos]; this.pos++; return t; };
Parser.prototype.expect = function(type) {
  var t = this.peek();
  if (t.type !== type) throw new Error("期望 " + type + "，实际 " + t.type + " (位置 " + t.pos + ")");
  return this.advance();
};
Parser.prototype.parse = function() {
  var node = this.parseExpr();
  if (this.peek().type !== 'EOF') {
    var t = this.peek();
    throw new Error("意外的 token: " + t.type + " " + t.value + " (位置 " + t.pos + ")");
  }
  return node;
};
Parser.prototype.parseExpr = function() {
  var left = this.parseTerm();
  while (this.peek().type === 'OP' && (this.peek().value === '+' || this.peek().value === '-')) {
    var op = this.advance().value;
    var right = this.parseTerm();
    left = { type: 'binop', op: op, left: left, right: right };
  }
  return left;
};
Parser.prototype.parseTerm = function() {
  var left = this.parseFactor();
  while (this.peek().type === 'OP' && (this.peek().value === '*' || this.peek().value === '/')) {
    var op = this.advance().value;
    var right = this.parseFactor();
    left = { type: 'binop', op: op, left: left, right: right };
  }
  return left;
};
Parser.prototype.parseFactor = function() {
  var left = this.parseBase();
  if (this.peek().type === 'OP' && this.peek().value === '^') {
    this.advance();
    var right = this.parseFactor();
    left = { type: 'binop', op: '^', left: left, right: right };
  }
  return left;
};
Parser.prototype.parseBase = function() {
  var t = this.peek();
  // 一元正负号
  if (t.type === 'OP' && (t.value === '+' || t.value === '-')) {
    var op = this.advance().value;
    var operand = this.parseBase();
    return { type: 'unary', op: op, operand: operand };
  }
  // 数字
  if (t.type === 'NUMBER') { this.advance(); return { type: 'num', value: t.value }; }
  // 括号
  if (t.type === 'LPAREN') {
    this.advance();
    var node = this.parseExpr();
    this.expect('RPAREN');
    return node;
  }
  // 标识符
  if (t.type === 'IDENT') {
    this.advance();
    // 函数调用
    if (this.peek().type === 'LPAREN') {
      this.advance();
      var args = [];
      if (this.peek().type !== 'RPAREN') {
        args.push(this.parseExpr());
        while (this.peek().type === 'COMMA') { this.advance(); args.push(this.parseExpr()); }
      }
      this.expect('RPAREN');
      var upperName = t.value.toUpperCase();
      if (!WHITELIST_FUNCS[upperName]) {
        throw new Error("函数 '" + t.value + "' 不在白名单中 (位置 " + t.pos + ")。允许: ABS, MAX, MIN, ROUND, SUM");
      }
      return { type: 'call', name: upperName, args: args };
    }
    return { type: 'var', name: t.value };
  }
  throw new Error("意外的 token: " + t.type + " " + t.value + " (位置 " + t.pos + ")");
};

function parse(tokens) { return new Parser(tokens).parse(); }

function evaluate(ast, context) {
  var t = ast.type;
  if (t === 'num') return parseFloat(ast.value);
  if (t === 'var') {
    if (!(ast.name in context)) throw new Error("未知变量: '" + ast.name + "'");
    var val = context[ast.name];
    if (val === null || val === undefined) return 0.0;
    return parseFloat(val);
  }
  if (t === 'unary') {
    var v = evaluate(ast.operand, context);
    return ast.op === '-' ? -v : v;
  }
  if (t === 'binop') {
    var left = evaluate(ast.left, context);
    var right = evaluate(ast.right, context);
    var op = ast.op;
    if (op === '+') return left + right;
    if (op === '-') return left - right;
    if (op === '*') return left * right;
    if (op === '/') { if (right === 0) throw new Error("除零错误"); return left / right; }
    if (op === '^') { var r = Math.pow(left, right); if (!isFinite(r)) throw new Error("幂运算错误"); return r; }
    throw new Error("未知运算符: " + op);
  }
  if (t === 'call') {
    var name = ast.name;
    var args = ast.args.map(function(a) { return evaluate(a, context); });
    if (name === 'MAX') { if (!args.length) throw new Error("MAX() 至少需要 1 个参数"); return Math.max.apply(null, args); }
    if (name === 'MIN') { if (!args.length) throw new Error("MIN() 至少需要 1 个参数"); return Math.min.apply(null, args); }
    if (name === 'ABS') { if (args.length !== 1) throw new Error("ABS() 需要 1 个参数"); return Math.abs(args[0]); }
    if (name === 'ROUND') {
      if (args.length < 1 || args.length > 2) throw new Error("ROUND() 需要 1 或 2 个参数");
      var digits = args.length === 2 ? parseInt(args[1]) : 0;
      var f = Math.pow(10, digits);
      return Math.round(args[0] * f) / f;
    }
    if (name === 'SUM') { var s = 0; for (var k = 0; k < args.length; k++) s += args[k]; return s; }
    throw new Error("未知函数: " + name);
  }
  throw new Error("未知 AST 节点类型: " + t);
}

function toExcelFormula(ast, cellMap) {
  var t = ast.type;
  if (t === 'num') {
    var v = ast.value;
    if (typeof v === 'number' && v === Math.floor(v)) return String(v);
    return String(v);
  }
  if (t === 'var') {
    if (!(ast.name in cellMap)) throw new Error("变量 '" + ast.name + "' 未在 cell_map 中找到单元格引用");
    return cellMap[ast.name];
  }
  if (t === 'unary') {
    var inner = toExcelFormula(ast.operand, cellMap);
    return "(" + ast.op + inner + ")";
  }
  if (t === 'binop') {
    var left = toExcelFormula(ast.left, cellMap);
    var right = toExcelFormula(ast.right, cellMap);
    return "(" + left + ast.op + right + ")";
  }
  if (t === 'call') {
    var name = ast.name;
    var args = ast.args.map(function(a) { return toExcelFormula(a, cellMap); });
    return name + "(" + args.join(",") + ")";
  }
  throw new Error("未知 AST 节点类型: " + t);
}

function expressionToExcel(expr, cellMap) {
  var tokens = tokenize(expr);
  var ast = parse(tokens);
  return "=" + toExcelFormula(ast, cellMap);
}

function validateExpression(expr) {
  try { var tokens = tokenize(expr); var ast = parse(tokens); return { valid: true, error: "" }; }
  catch (e) { return { valid: false, error: String(e.message || e) }; }
}


// ==================== 5. jsCalculations (calculations.py) ====================

function _evalCustomItems(customItems, years, contexts) {
  if (!customItems || !customItems.length) return [];
  var result = [];
  for (var ci = 0; ci < customItems.length; ci++) {
    var item = customItems[ci];
    var entry = { name: item.name, format: item.format, formula: item.formula, values: [], error: "" };
    var ast;
    try {
      var tokens = tokenize(item.formula);
      ast = parse(tokens);
    } catch (e) {
      entry.error = "公式语法错误: " + (e.message || e);
      result.push(entry); continue;
    }
    var values = [];
    for (var i = 0; i < years.length; i++) {
      var ctx = i < contexts.length ? contexts[i] : {};
      try { values.push(_round2(evaluate(ast, ctx))); }
      catch (e) { values.push(null); }
    }
    entry.values = values;
    result.push(entry);
  }
  return result;
}

function calculateIrr(cashFlows, maxIter, tol) {
  maxIter = maxIter || 100; tol = tol || 1e-8;
  var n = cashFlows.length;
  if (n < 2) return 0.0;
  function npv(rate) {
    var total = 0;
    for (var i = 0; i < n; i++) total += cashFlows[i] / Math.pow(1 + rate, i);
    return total;
  }
  var lo = -0.9999, hi = 10.0;
  var npvLo = npv(lo), npvHi = npv(hi);
  if (npvLo * npvHi > 0) {
    var found = false;
    for (var x = -99; x <= 1000; x++) {
      var testRate = x * 0.01;
      if (npv(testRate) * npvLo < 0) { hi = testRate; npvHi = npv(testRate); found = true; break; }
    }
    if (!found) return 0.0;
  }
  for (var iter = 0; iter < maxIter; iter++) {
    var mid = (lo + hi) / 2;
    var npvMid = npv(mid);
    if (Math.abs(npvMid) < tol) return mid;
    if (npvMid * npvLo < 0) { hi = mid; npvHi = npvMid; }
    else { lo = mid; npvLo = npvMid; }
  }
  return (lo + hi) / 2;
}

function calculateThreeStatement(p) {
  var n = p.projection_years;
  var years = []; for (var i = 0; i <= n; i++) years.push(i);

  // Income Statement
  var revenue = new Array(n+1).fill(0);
  revenue[0] = p.revenue_y0;
  for (var i = 1; i <= n; i++) { var g = growthYear(p, i); revenue[i] = revenue[i-1] * (1+g); }
  var cogs = revenue.map(function(r, j) { return r * rateYear(p, "cogs_pct", j); });
  var grossProfit = revenue.map(function(r,j) { return r - cogs[j]; });
  var sga = revenue.map(function(r, j) { return r * rateYear(p, "sga_pct", j); });
  var rd = revenue.map(function(r, j) { return r * rateYear(p, "rd_pct", j); });
  var ebitda = grossProfit.map(function(gp,j) { return gp - sga[j] - rd[j]; });
  var daIs = revenue.map(function(r, j) { return r * rateYear(p, "da_pct", j); });
  var ebit = ebitda.map(function(e,j) { return e - daIs[j]; });

  // Supporting Schedules
  var ppeBeg = new Array(n+1).fill(0); ppeBeg[0] = p.beg_ppe;
  var capex = revenue.map(function(r, j) { return r * rateYear(p, "capex_pct", j); });
  var ppeEnd = new Array(n+1).fill(0); ppeEnd[0] = p.beg_ppe;
  for (var i = 1; i <= n; i++) { ppeBeg[i] = ppeEnd[i-1]; ppeEnd[i] = ppeBeg[i] + capex[i] - daIs[i]; }

  var debtBeg = new Array(n+1).fill(0); debtBeg[0] = p.beg_debt;
  var newDebt = new Array(n+1).fill(0);
  for (var i = 1; i <= n; i++) newDebt[i] = p.new_debt_issuance;
  var debtEnd = new Array(n+1).fill(0); debtEnd[0] = p.beg_debt;
  for (var i = 1; i <= n; i++) { debtBeg[i] = debtEnd[i-1]; debtEnd[i] = debtBeg[i] + newDebt[i]; }

  var interest = new Array(n+1).fill(0); interest[0] = p.interest_rate * p.beg_debt;
  for (var i = 1; i <= n; i++) interest[i] = p.interest_rate * debtEnd[i-1];

  var ebt = ebit.map(function(e,j) { return e - interest[j]; });
  var taxes = ebt.map(function(e) { return Math.max(0, e * p.tax_rate); });
  var netIncome = ebt.map(function(e,j) { return e - taxes[j]; });
  var dividends = netIncome.map(function(ni) { return ni * p.dividend_pct; });

  // Balance Sheet
  var cash = new Array(n+1).fill(0); cash[0] = p.beg_cash;
  var ar = new Array(n+1).fill(0); ar[0] = p.beg_ar;
  var inventory = new Array(n+1).fill(0); inventory[0] = p.beg_inventory;
  var ppeBs = new Array(n+1).fill(0); ppeBs[0] = p.beg_ppe;
  for (var i = 1; i <= n; i++) { ar[i] = revenue[i]/365*p.dso; inventory[i] = cogs[i]/365*p.dio; ppeBs[i] = ppeEnd[i]; }

  var ap = new Array(n+1).fill(0); ap[0] = p.beg_ap;
  var accrued = new Array(n+1).fill(0); accrued[0] = p.beg_accrued;
  var debtBs = new Array(n+1).fill(0); debtBs[0] = p.beg_debt;
  for (var i = 1; i <= n; i++) { ap[i] = cogs[i]/365*p.dpo; accrued[i] = revenue[i]/365*p.accrued_days; debtBs[i] = debtEnd[i]; }

  var commonStock = new Array(n+1).fill(p.common_stock);
  var re = new Array(n+1).fill(0); re[0] = p.beg_retn_earn;
  for (var i = 1; i <= n; i++) re[i] = re[i-1] + netIncome[i] - dividends[i];

  // Cash Flow Statement
  var deltaAr = new Array(n+1).fill(0), deltaInv = new Array(n+1).fill(0), deltaAp = new Array(n+1).fill(0), deltaAcc = new Array(n+1).fill(0);
  var cfo = new Array(n+1).fill(0), cfi = new Array(n+1).fill(0), cff = new Array(n+1).fill(0), netChange = new Array(n+1).fill(0);
  var capexCfs = new Array(n+1).fill(0), netDebtChange = new Array(n+1).fill(0), dividendsPaid = new Array(n+1).fill(0);
  for (var i = 1; i <= n; i++) {
    deltaAr[i] = -(ar[i]-ar[i-1]); deltaInv[i] = -(inventory[i]-inventory[i-1]);
    deltaAp[i] = ap[i]-ap[i-1]; deltaAcc[i] = accrued[i]-accrued[i-1];
    cfo[i] = netIncome[i] + daIs[i] + deltaAr[i] + deltaInv[i] + deltaAp[i] + deltaAcc[i];
    capexCfs[i] = -capex[i]; cfi[i] = capexCfs[i];
    netDebtChange[i] = debtBs[i]-debtBs[i-1]; dividendsPaid[i] = -dividends[i];
    cff[i] = netDebtChange[i] + dividendsPaid[i];
    netChange[i] = cfo[i] + cfi[i] + cff[i];
  }
  for (var i = 1; i <= n; i++) cash[i] = cash[i-1] + netChange[i];

  var totalAssets = [], totalLiab = [], totalEquity = [], totalLe = [], balanceCheck = [];
  for (var i = 0; i <= n; i++) {
    totalAssets.push(cash[i]+ar[i]+inventory[i]+ppeBs[i]);
    totalLiab.push(ap[i]+accrued[i]+debtBs[i]);
    totalEquity.push(commonStock[i]+re[i]);
    totalLe.push(totalLiab[i]+totalEquity[i]);
    balanceCheck.push(Math.abs(totalAssets[i]-totalLe[i])<0.01 ? "BALANCED" : "OUT OF BALANCE");
  }
  var ebitdaMargin = [0], netMargin = [0];
  for (var i = 1; i <= n; i++) { ebitdaMargin.push(_safeDiv(ebitda[i],revenue[i])); netMargin.push(_safeDiv(netIncome[i],revenue[i])); }

  function buildCtx(i) {
    return {
      revenue: revenue[i], cogs: cogs[i], gross_profit: grossProfit[i], sga: sga[i], rd: rd[i],
      ebitda: ebitda[i], da: daIs[i], ebit: ebit[i], interest: interest[i], ebt: ebt[i],
      taxes: taxes[i], net_income: netIncome[i], dividends: dividends[i], cash: cash[i],
      ar: ar[i], inventory: inventory[i], ppe: ppeBs[i], total_assets: totalAssets[i],
      ap: ap[i], accrued: accrued[i], debt: debtBs[i], total_liabilities: totalLiab[i],
      common_stock: commonStock[i], retained_earnings: re[i], total_equity: totalEquity[i],
      total_le: totalLe[i], capex: capex[i], new_debt: newDebt[i],
    };
  }
  var contexts = []; for (var i = 0; i <= n; i++) contexts.push(buildCtx(i));
  var customResults = _evalCustomItems(p.custom_items, years, contexts);

  return {
    model_type: "three_statement", company_name: p.company_name, years: years, tax_basis: p.tax_basis,
    income_statement: {
      revenue: revenue.map(_round2), cogs: cogs.map(_round2), gross_profit: grossProfit.map(_round2),
      sga: sga.map(_round2), rd: rd.map(_round2), ebitda: ebitda.map(_round2), da: daIs.map(_round2),
      ebit: ebit.map(_round2), interest: interest.map(_round2), ebt: ebt.map(_round2),
      taxes: taxes.map(_round2), net_income: netIncome.map(_round2), dividends: dividends.map(_round2),
    },
    balance_sheet: {
      cash: cash.map(_round2), ar: ar.map(_round2), inventory: inventory.map(_round2), ppe: ppeBs.map(_round2),
      total_assets: totalAssets.map(_round2), ap: ap.map(_round2), accrued: accrued.map(_round2),
      debt: debtBs.map(_round2), total_liabilities: totalLiab.map(_round2), common_stock: commonStock.map(_round2),
      retained_earnings: re.map(_round2), total_equity: totalEquity.map(_round2), total_le: totalLe.map(_round2),
      balance_check: balanceCheck,
    },
    cash_flow: {
      net_income: netIncome.map(_round2), da: daIs.map(_round2), delta_ar: deltaAr.map(_round2),
      delta_inventory: deltaInv.map(_round2), delta_ap: deltaAp.map(_round2), delta_accrued: deltaAcc.map(_round2),
      cfo: cfo.map(_round2), capex: capexCfs.map(_round2), cfi: cfi.map(_round2),
      net_debt_change: netDebtChange.map(_round2), dividends_paid: dividendsPaid.map(_round2),
      cff: cff.map(_round2), net_change_in_cash: netChange.map(_round2),
    },
    key_metrics: { ebitda_margin: ebitdaMargin.map(_round2), net_margin: netMargin.map(_round2) },
    custom_items: customResults,
  };
}

function calculateDcf(p) {
  var n = p.projection_years;
  var years = []; for (var i = 0; i <= n; i++) years.push(i);

  var revenue = new Array(n+1).fill(0); revenue[0] = p.revenue_y0;
  for (var i = 1; i <= n; i++) { var g = growthYear(p, i); revenue[i] = revenue[i-1] * (1+g); }

  var ebitda = new Array(n+1).fill(0); ebitda[0] = revenue[0] * marginYear(p, 1);
  for (var i = 1; i <= n; i++) { ebitda[i] = revenue[i] * marginYear(p, i); }

  var da = revenue.map(function(r, j) { return r * rateYear(p, "da_pct", j); });
  var ebit = ebitda.map(function(e,j) { return e - da[j]; });
  var nopat = ebit.map(function(e) { return e * (1 - p.tax_rate); });

  var deltaNwc = new Array(n+1).fill(0);
  for (var i = 1; i <= n; i++) deltaNwc[i] = (revenue[i] - revenue[i-1]) * rateYear(p, "nwc_pct", i);
  var capex = revenue.map(function(r, j) { return r * rateYear(p, "capex_pct", j); });
  var ufcf = new Array(n+1).fill(0);
  for (var i = 1; i <= n; i++) ufcf[i] = nopat[i] + da[i] - deltaNwc[i] - capex[i];

  var costOfEquity = p.risk_free_rate + p.equity_risk_premium * p.beta;
  var afterTaxKd = p.pre_tax_cost_of_debt * (1 - p.tax_rate);
  var w = dcfWeights(p);
  var costOfPreferred = _num(p.cost_of_preferred, 0);
  // WACC = We·Ke + Wp·Kp + Wd·Kd·(1-t)
  var wacc = costOfEquity * w.we + costOfPreferred * w.wp + afterTaxKd * w.wd;

  var lastUfcf = ufcf[n], lastEbitda = ebitda[n];
  var tvGordon = wacc !== p.terminal_growth ? lastUfcf * (1 + p.terminal_growth) / (wacc - p.terminal_growth) : 0.0;
  var tvExit = lastEbitda * p.exit_multiple;
  var selectedTv = p.tv_method === "gordon" ? tvGordon : tvExit;
  var pvTv = selectedTv / Math.pow(1 + wacc, n);

  var isMid = p.valuation_timing === "mid";
  var pvFcfs = new Array(n+1).fill(0);
  for (var i = 1; i <= n; i++) {
    var period = isMid ? i - 0.5 : i;
    pvFcfs[i] = ufcf[i] / Math.pow(1 + wacc, period);
  }
  var sumPvFcf = 0; for (var i = 1; i <= n; i++) sumPvFcf += pvFcfs[i];
  var ev = sumPvFcf + pvTv;
  var equityValue = ev - p.net_debt + p.cash - p.minority_interest - p.preferred_stock - p.stock_based_comp - p.pension_deficit;
  var impliedPrice = _safeDiv(equityValue, p.shares_outstanding);
  var upside = p.current_price ? _safeDiv(impliedPrice, p.current_price) - 1 : 0.0;

  var ebitdaMarginCalc = [0], ebitMargin = [0];
  for (var i = 1; i <= n; i++) { ebitdaMarginCalc.push(_safeDiv(ebitda[i], revenue[i])); ebitMargin.push(_safeDiv(ebit[i], revenue[i])); }

  function buildCtx(i) {
    return {
      revenue: revenue[i], ebitda: ebitda[i], da: da[i], ebit: ebit[i], nopat: nopat[i],
      delta_nwc: deltaNwc[i], capex: capex[i], ufcf: ufcf[i], wacc: wacc,
      cost_of_equity: costOfEquity, after_tax_kd: afterTaxKd, tv_gordon: tvGordon, tv_exit: tvExit,
      selected_tv: selectedTv, pv_tv: pvTv, sum_pv_fcf: sumPvFcf, enterprise_value: ev,
      equity_value: equityValue, implied_price: impliedPrice, current_price: p.current_price,
      shares_outstanding: p.shares_outstanding, net_debt: p.net_debt, cash: p.cash,
    };
  }
  var contexts = []; for (var i = 0; i <= n; i++) contexts.push(buildCtx(i));
  var customResults = _evalCustomItems(p.custom_items, years, contexts);

  return {
    model_type: "dcf", company_name: p.company_name, years: years, tax_basis: p.tax_basis,
    operating_model: {
      revenue: revenue.map(_round2), ebitda: ebitda.map(_round2), da: da.map(_round2),
      ebit: ebit.map(_round2), nopat: nopat.map(_round2), delta_nwc: deltaNwc.map(_round2),
      capex: capex.map(_round2), ufcf: ufcf.map(_round2),
    },
    wacc: {
      cost_of_equity: _round2(costOfEquity), after_tax_cost_of_debt: _round2(afterTaxKd),
      equity_weight: _round2(w.we), debt_weight: _round2(w.wd), wacc: _round2(wacc),
      preferred_weight: _round2(w.wp), cost_of_preferred: _round2(costOfPreferred),
      weights_sum: _round2(w.we + w.wp + w.wd), tax_basis: p.tax_basis,
    },
    terminal_value: {
      tv_method: p.tv_method, tv_gordon: _round2(tvGordon), tv_exit: _round2(tvExit),
      selected_tv: _round2(selectedTv), pv_tv: _round2(pvTv),
    },
    valuation: {
      sum_pv_fcf: _round2(sumPvFcf), pv_tv: _round2(pvTv), enterprise_value: _round2(ev),
      equity_value: _round2(equityValue), shares_outstanding: p.shares_outstanding,
      implied_price: _round2(impliedPrice), current_price: p.current_price, upside: _round2(upside),
    },
    key_metrics: { ebitda_margin: ebitdaMarginCalc.map(_round2), ebit_margin: ebitMargin.map(_round2) },
    custom_items: customResults,
  };
}

function calculateLbo(p) {
  var n = Math.max(1, Math.min(10, parseInt(p.exit_year, 10) || 5));
  var years = []; for (var i = 0; i <= n; i++) years.push(i);

  var entryEv = p.ltm_ebitda * p.entry_ev_ebitda;
  var purchaseEquity = entryEv - p.existing_net_debt;
  var refiDebt = p.existing_net_debt + p.existing_cash;
  var transFees = p.transaction_fees, finFees = p.financing_fees;
  var totalUses = purchaseEquity + refiDebt + transFees + finFees;
  var revolverDraw = totalUses - p.sponsor_equity - p.term_loan_a - p.term_loan_b - p.senior_notes - p.subordinated_debt - p.existing_cash;
  var totalSources = p.sponsor_equity + revolverDraw + p.term_loan_a + p.term_loan_b + p.senior_notes + p.subordinated_debt + p.existing_cash;
  var suBalance = Math.abs(totalSources - totalUses) < 0.01;

  var revenue = new Array(n+1).fill(0); revenue[0] = p.ltm_revenue;
  for (var i = 1; i <= n; i++) { var g = growthYear(p, i); revenue[i] = revenue[i-1] * (1+g); }
  var ebitda = new Array(n+1).fill(0); ebitda[0] = p.ltm_ebitda;
  for (var i = 1; i <= n; i++) ebitda[i] = revenue[i] * rateYear(p, "ebitda_margin", i);
  var da = revenue.map(function(r, j) { return r * rateYear(p, "da_pct", j); });
  var ebit = ebitda.map(function(e,j) { return e - da[j]; });
  var capex = revenue.map(function(r, j) { return r * rateYear(p, "capex_pct", j); });
  var deltaNwc = new Array(n+1).fill(0);
  for (var i = 1; i <= n; i++) deltaNwc[i] = -revenue[i] * rateYear(p, "nwc_pct", i);

  var begCash = new Array(n+1).fill(0); begCash[0] = p.existing_cash;
  var cfads = new Array(n+1).fill(0), cashAvail = new Array(n+1).fill(0);
  var revBeg = new Array(n+1).fill(0); revBeg[0] = revolverDraw;
  var revRepay = new Array(n+1).fill(0), revEnd = new Array(n+1).fill(0); revEnd[0] = revolverDraw;
  var tlaBeg = new Array(n+1).fill(0); tlaBeg[0] = p.term_loan_a;
  var tlaMandatory = new Array(n+1).fill(0), tlaEnd = new Array(n+1).fill(0); tlaEnd[0] = p.term_loan_a;
  var tlbBeg = new Array(n+1).fill(0); tlbBeg[0] = p.term_loan_b;
  var tlbMandatory = new Array(n+1).fill(0), tlbSweep = new Array(n+1).fill(0), tlbEnd = new Array(n+1).fill(0); tlbEnd[0] = p.term_loan_b;
  var snBeg = new Array(n+1).fill(0); snBeg[0] = p.senior_notes;
  var snEnd = new Array(n+1).fill(0); snEnd[0] = p.senior_notes;
  var subBeg = new Array(n+1).fill(0); subBeg[0] = p.subordinated_debt;
  var subEnd = new Array(n+1).fill(0); subEnd[0] = p.subordinated_debt;
  var totalDebt = new Array(n+1).fill(0), cashEnd = new Array(n+1).fill(0);
  totalDebt[0] = revEnd[0]+tlaEnd[0]+tlbEnd[0]+snEnd[0]+subEnd[0]; cashEnd[0] = begCash[0];
  var intRev = new Array(n+1).fill(0), intTla = new Array(n+1).fill(0), intTlb = new Array(n+1).fill(0);
  var intSn = new Array(n+1).fill(0), intSub = new Array(n+1).fill(0), intIncome = new Array(n+1).fill(0);
  var totalInterest = new Array(n+1).fill(0), ebt = new Array(n+1).fill(0), taxes = new Array(n+1).fill(0), netIncome = new Array(n+1).fill(0);

  for (var i = 1; i <= n; i++) {
    intRev[i] = p.revolver_rate * revBeg[i]; intTla[i] = p.tla_rate * tlaBeg[i];
    intTlb[i] = p.tlb_rate * tlbBeg[i]; intSn[i] = p.senior_notes_rate * snBeg[i]; intSub[i] = p.sub_rate * subBeg[i];
    totalInterest[i] = intRev[i]+intTla[i]+intTlb[i]+intSn[i]+intSub[i];
    intIncome[i] = p.cash_interest_rate * begCash[i];
    ebt[i] = ebit[i] - totalInterest[i] + intIncome[i];
    taxes[i] = Math.max(0, ebt[i] * p.tax_rate);
    netIncome[i] = ebt[i] - taxes[i];
    var cfo = netIncome[i] + da[i] + deltaNwc[i];
    cfads[i] = cfo - capex[i];
    cashAvail[i] = Math.max(0, begCash[i] + cfads[i] - p.min_cash_balance);
    revRepay[i] = Math.min(revBeg[i], cashAvail[i]);
    revEnd[i] = revBeg[i] - revRepay[i];
    tlaMandatory[i] = Math.min(tlaBeg[i], p.term_loan_a * p.tla_mandatory_amort);
    tlaEnd[i] = tlaBeg[i] - tlaMandatory[i];
    tlbMandatory[i] = Math.min(tlbBeg[i], p.term_loan_b * p.tlb_mandatory_amort);
    var remainingCash = Math.max(0, cashAvail[i] - revRepay[i] - tlaMandatory[i] - tlbMandatory[i]);
    tlbSweep[i] = Math.min(tlbBeg[i] - tlbMandatory[i], remainingCash * p.cash_sweep_pct);
    tlbEnd[i] = tlbBeg[i] - tlbMandatory[i] - tlbSweep[i];
    snEnd[i] = snBeg[i]; subEnd[i] = subBeg[i];
    totalDebt[i] = revEnd[i]+tlaEnd[i]+tlbEnd[i]+snEnd[i]+subEnd[i];
    cashEnd[i] = Math.max(p.min_cash_balance, begCash[i] + cfads[i] - revRepay[i] - tlaMandatory[i] - tlbMandatory[i] - tlbSweep[i]);
    if (i+1 <= n) { begCash[i+1]=cashEnd[i]; revBeg[i+1]=revEnd[i]; tlaBeg[i+1]=tlaEnd[i]; tlbBeg[i+1]=tlbEnd[i]; snBeg[i+1]=snEnd[i]; subBeg[i+1]=subEnd[i]; }
  }

  var exitYear = Math.min(p.exit_year, n);
  var exitEbitda = ebitda[exitYear];
  var exitEv = exitEbitda * p.exit_ev_ebitda;
  var exitEquity = exitEv - totalDebt[exitYear] + cashEnd[exitYear];
  var moic = _safeDiv(exitEquity, p.sponsor_equity);
  var irrCfs = new Array(n+1).fill(0); irrCfs[0] = -p.sponsor_equity;
  for (var i = 1; i <= n; i++) { if (i === exitYear) irrCfs[i] = exitEquity; else irrCfs[i] = 0; }
  var irr = calculateIrr(irrCfs.slice(0, exitYear+1));

  var ebitdaMargin = [0];
  for (var i = 1; i <= n; i++) ebitdaMargin.push(_safeDiv(ebitda[i], revenue[i]));

  function buildCtx(i) {
    return {
      revenue: revenue[i], ebitda: ebitda[i], da: da[i], ebit: ebit[i],
      interest_revolver: intRev[i], interest_tla: intTla[i], interest_tlb: intTlb[i],
      interest_sn: intSn[i], interest_sub: intSub[i], total_interest: totalInterest[i],
      interest_income: intIncome[i], ebt: ebt[i], taxes: taxes[i], net_income: netIncome[i],
      total_debt: totalDebt[i], cash: cashEnd[i], revolver: revEnd[i], tla: tlaEnd[i],
      tlb: tlbEnd[i], senior_notes: snEnd[i], subordinated_debt: subEnd[i], cfads: cfads[i],
    };
  }
  var contexts = []; for (var i = 0; i <= n; i++) contexts.push(buildCtx(i));
  var customResults = _evalCustomItems(p.custom_items, years, contexts);

  return {
    model_type: "lbo", company_name: p.company_name, years: years, tax_basis: p.tax_basis,
    sources_uses: {
      entry_ev: _round2(entryEv), purchase_equity: _round2(purchaseEquity), refinance_debt: _round2(refiDebt),
      transaction_fees: _round2(transFees), financing_fees: _round2(finFees), total_uses: _round2(totalUses),
      sponsor_equity: _round2(p.sponsor_equity), revolver_draw: _round2(revolverDraw), term_loan_a: _round2(p.term_loan_a),
      term_loan_b: _round2(p.term_loan_b), senior_notes: _round2(p.senior_notes), subordinated_debt: _round2(p.subordinated_debt),
      existing_cash: _round2(p.existing_cash), total_sources: _round2(totalSources),
      balance_check: suBalance ? "BALANCED" : "IMBALANCED",
    },
    income_statement: {
      revenue: revenue.map(_round2), ebitda: ebitda.map(_round2), da: da.map(_round2), ebit: ebit.map(_round2),
      interest_revolver: intRev.map(_round2), interest_tla: intTla.map(_round2), interest_tlb: intTlb.map(_round2),
      interest_sn: intSn.map(_round2), interest_sub: intSub.map(_round2), total_interest: totalInterest.map(_round2),
      interest_income: intIncome.map(_round2), ebt: ebt.map(_round2), taxes: taxes.map(_round2), net_income: netIncome.map(_round2),
    },
    debt_schedule: {
      beg_cash: begCash.map(_round2), cfads: cfads.map(_round2), cash_available: cashAvail.map(_round2),
      revolver_beg: revBeg.map(_round2), revolver_repay: revRepay.map(_round2), revolver_end: revEnd.map(_round2),
      tla_beg: tlaBeg.map(_round2), tla_mandatory: tlaMandatory.map(_round2), tla_end: tlaEnd.map(_round2),
      tlb_beg: tlbBeg.map(_round2), tlb_mandatory: tlbMandatory.map(_round2), tlb_sweep: tlbSweep.map(_round2), tlb_end: tlbEnd.map(_round2),
      sn_beg: snBeg.map(_round2), sn_end: snEnd.map(_round2), sub_beg: subBeg.map(_round2), sub_end: subEnd.map(_round2),
      total_debt: totalDebt.map(_round2), cash_end: cashEnd.map(_round2),
    },
    exit_returns: {
      exit_year: exitYear, exit_ebitda: _round2(exitEbitda), exit_ev_ebitda: p.exit_ev_ebitda,
      exit_ev: _round2(exitEv), total_debt_at_exit: _round2(totalDebt[exitYear]), cash_at_exit: _round2(cashEnd[exitYear]),
      equity_value: _round2(exitEquity), sponsor_equity_invested: _round2(p.sponsor_equity), moic: _round2(moic),
      irr: _round2(irr), irr_display: (irr * 100).toFixed(1) + "%",
    },
    key_metrics: { ebitda_margin: ebitdaMargin.map(_round2) },
    custom_items: customResults,
  };
}

function jsCalculate(modelType, params) {
  var p = normalizeParams(modelType, params);
  if (modelType === "three_statement") return calculateThreeStatement(p);
  if (modelType === "dcf") return calculateDcf(p);
  if (modelType === "lbo") return calculateLbo(p);
  throw new Error("Unknown model type: " + modelType);
}


// ==================== 6. jsBuildExcel (builders/*.py) ====================

// ==================== 5b. ExcelJS 专业样式导出 ====================
var XL_FMT = {
  num: '#,##0.00;(#,##0.00);-',
  num0: '#,##0;(#,##0);-',
  pct: '0.0%;(0.0%);-',
  pct2: '0.00%;(0.00%);-',
  mult: '0.00"x"',
  ratio2: '0.00',
  price: '#,##0.00;(#,##0.00);-',
  int: '#,##0;(#,##0);-'
};
var XL_C = {
  header: 'FF1F4E78', headerFont: 'FFFFFFFF', sub: 'FFD6E4F0', label: 'FFF2F2F2',
  input: 'FFFFF2CC', inputFont: 'FF0000FF', total: 'FFE2EFDA', border: 'FFBFBFBF',
  dark: 'FF1F4E78', grey: 'FF808080', ok: 'FFC6EFCE', bad: 'FFFFC7CE', plainFont: 'FF333333'
};

function xlThin(color) {
  var s = { style: 'thin', color: { argb: color || XL_C.border } };
  return { top: s, left: s, bottom: s, right: s };
}

function xlStyleCell(cell, kind, fmt) {
  var k = kind || 'plain';
  cell.border = xlThin();
  cell.alignment = { vertical: 'center' };
  if (fmt) cell.numFmt = fmt;
  if (k === 'title') {
    cell.border = {};
    cell.font = { bold: true, size: 14, color: { argb: XL_C.dark } };
  } else if (k === 'header') {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: XL_C.header } };
    cell.font = { bold: true, size: 10, color: { argb: XL_C.headerFont } };
    cell.alignment = { horizontal: 'center', vertical: 'center' };
  } else if (k === 'sub') {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: XL_C.sub } };
    cell.font = { bold: true, size: 10, color: { argb: XL_C.dark } };
  } else if (k === 'label') {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: XL_C.label } };
    cell.font = { size: 10, color: { argb: XL_C.plainFont } };
  } else if (k === 'input') {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: XL_C.input } };
    cell.font = { size: 10, color: { argb: XL_C.inputFont } };
    cell.alignment = { horizontal: 'right', vertical: 'center' };
  } else if (k === 'total') {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: XL_C.total } };
    cell.font = { bold: true, size: 10 };
  } else if (k === 'note') {
    cell.border = {};
    cell.font = { italic: true, size: 9, color: { argb: XL_C.grey } };
    cell.alignment = { wrapText: true, vertical: 'top', horizontal: 'left' };
  } else if (k === 'ok') {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: XL_C.ok } };
    cell.font = { bold: true, size: 10 };
    cell.alignment = { horizontal: 'center', vertical: 'center' };
  } else if (k === 'bad') {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: XL_C.bad } };
    cell.font = { bold: true, size: 10 };
    cell.alignment = { horizontal: 'center', vertical: 'center' };
  } else {
    cell.font = { size: 10, color: { argb: XL_C.plainFont } };
    cell.alignment = { horizontal: 'right', vertical: 'center' };
  }
}

// 声明式 sheet 构建器：所有行号显式（公式按固定行号引用）
function xlSheet(wb, name, ncols, opts) {
  opts = opts || {};
  var ws = wb.addWorksheet(name, {});
  if (opts.freeze) ws.views = [{ state: 'frozen', xSplit: opts.freeze[0] || 0, ySplit: opts.freeze[1] || 0 }];
  var widths = opts.widths || [];
  for (var ci = 1; ci <= ncols; ci++) ws.getColumn(ci).width = widths[ci - 1] || (ci === 1 ? 42 : 13);
  var r = 0;

  function styleRange(row, kind) {
    for (var c = 1; c <= ncols; c++) xlStyleCell(ws.getCell(row, c), kind);
  }
  function setVal(cell, raw) {
    if (typeof raw === 'string' && raw.charAt(0) === '=') cell.value = { formula: raw.substring(1) };
    else cell.value = raw;
  }
  var api = {
    ws: ws, ncols: ncols,
    row: function () { return r; },
    title: function (text) {
      r++;
      ws.getCell(r, 1).value = text;
      styleRange(r, 'title');
      if (ncols > 1) ws.mergeCells(r, 1, r, ncols);
      ws.getRow(r).height = 24;
    },
    blank: function () { r++; },
    section: function (text) {
      r++;
      ws.getCell(r, 1).value = text;
      styleRange(r, 'sub');
      if (ncols > 1) ws.mergeCells(r, 1, r, ncols);
    },
    note: function (text, height) {
      r++;
      ws.getCell(r, 1).value = text;
      styleRange(r, 'note');
      if (ncols > 1) ws.mergeCells(r, 1, r, ncols);
      ws.getRow(r).height = height || 42;
    },
    header: function (arr) {
      r++;
      for (var c = 1; c <= ncols; c++) {
        var cell = ws.getCell(r, c);
        if (arr[c - 1] !== null && arr[c - 1] !== undefined) cell.value = arr[c - 1];
        xlStyleCell(cell, 'header');
      }
      ws.getRow(r).height = 18;
    },
    // label + vals（从 B 列起，长度 ncols-1）；o: {fmt, numKind:'input'|'plain', total:bool}
    row: function (label, vals, o) {
      r++;
      o = o || {};
      var nk = o.numKind || 'plain';
      var lc = ws.getCell(r, 1);
      lc.value = label;
      xlStyleCell(lc, o.total ? 'total' : 'label');
      for (var c = 2; c <= ncols; c++) {
        var raw = vals ? vals[c - 2] : null;
        var kind = o.total ? 'total' : (typeof raw === 'string' && raw.charAt(0) === '=' ? 'plain' : nk);
        var fmt = o.fmt || null;
        if (raw !== null && raw !== undefined && raw !== '') {
          if (typeof raw === 'object') {
            if (raw.f) fmt = raw.f;
            if (raw.k) kind = raw.k;
            raw = raw.v;
          }
          setVal(ws.getCell(r, c), raw);
        }
        xlStyleCell(ws.getCell(r, c), kind, fmt);
      }
    }
  };
  return api;
}

// 比率是否按年细化（表单 toggle 开启时存在 _y1 键）
function xlRatioIsPer(p, baseKey) {
  var v = p[baseKey + '_y1'];
  return v !== undefined && v !== null && v !== '';
}
// Assumptions 比率行值：[B列统一值(Year0用), Year1..Year n]
function xlRatioVals(p, baseKey, n) {
  var per = xlRatioIsPer(p, baseKey);
  var base = parseFloat(p[baseKey]);
  if (isNaN(base)) base = null;
  var vals = [base];
  for (var i = 1; i <= n; i++) {
    if (per) {
      var yv = parseFloat(p[baseKey + '_y' + i]);
      vals.push(isNaN(yv) ? base : yv);
    } else vals.push(base);
  }
  return { per: per, vals: vals };
}
// 始终按年字段（增长率/DCF EBITDA margin）：[空(B), y1..yn]
function xlYearlyVals(p, baseKey, n) {
  var vals = [null];
  for (var i = 1; i <= n; i++) {
    var v = parseFloat(p[baseKey + '_y' + i]);
    vals.push(isNaN(v) ? 0 : v);
  }
  return vals;
}
// Assumptions 行引用：未细化或 Year0 → 锁 $B$row；细化年份 → 对应年列
function aRef(A, per, row, yi) {
  return A + '!' + (per && yi > 0 ? ycol(yi) + row : '$B$' + row);
}

function xlTaxLabel(p) {
  return p.tax_basis === 'mtr'
    ? 'Tax Rate 税率 (MTR 边际税率)'
    : 'Tax Rate 税率 (ETR 有效税率)';
}

async function jsBuildExcel(modelType, params) {
  var wb = new ExcelJS.Workbook();
  wb.creator = 'Financial Model Generator';
  wb.created = new Date();
  if (modelType === 'three_statement') buildThreeStatementExcel(wb, params);
  else if (modelType === 'dcf') buildDcfExcel(wb, params);
  else if (modelType === 'lbo') buildLboExcel(wb, params);
  else throw new Error('Unknown model type: ' + modelType);

  var buf = await wb.xlsx.writeBuffer();
  return new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

// --- 三表联动模型 Excel ---
function buildThreeStatementExcel(wb, p) {
  var n = p.projection_years;
  var NC = n + 2;
  var A = "Assumptions", IS = "'Income Statement'", SCH = "'Supporting Schedules'", BS = "'Balance Sheet'", CFS = "'Cash Flow Statement'";
  var F = XL_FMT;
  var wA = [42]; for (var wc = 1; wc <= n + 1; wc++) wA.push(12);
  function headArr() { var h = ["项目 Item", "Year 0"]; for (var i = 1; i <= n; i++) h.push("Year " + i); return h; }
  function yvals(fn) { var v = []; for (var yi = 0; yi <= n; yi++) v.push(fn(yi)); return v; }

  var cogs = xlRatioVals(p, "cogs_pct", n);
  var sgaR = xlRatioVals(p, "sga_pct", n);
  var rdR = xlRatioVals(p, "rd_pct", n);
  var daR = xlRatioVals(p, "da_pct", n);
  var cxR = xlRatioVals(p, "capex_pct", n);

  // ---------------- Assumptions ----------------
  var sh = xlSheet(wb, "Assumptions", NC, { widths: wA, freeze: [1, 3] });
  sh.title("Assumptions 假设参数");
  sh.blank();
  sh.header(["参数 Parameter", "Year 0"].concat((function () { var h = []; for (var i = 1; i <= n; i++) h.push("Year " + i); return h; })()));
  sh.section("运营假设 Operating Assumptions");
  sh.row("Revenue (base) 基期收入", [p.revenue_y0], { fmt: F.num, numKind: "input" });
  sh.row("Revenue Growth % 收入增长率（每年可不同）", xlYearlyVals(p, "rev_growth", n), { fmt: F.pct, numKind: "input" });
  sh.row("COGS % of Revenue 成本率", cogs.vals, { fmt: F.pct, numKind: "input" });
  sh.row("SG&A % of Revenue 销管费用率", sgaR.vals, { fmt: F.pct, numKind: "input" });
  sh.row("R&D % of Revenue 研发费用率", rdR.vals, { fmt: F.pct, numKind: "input" });
  sh.row("D&A % of Revenue 折旧摊销率", daR.vals, { fmt: F.pct, numKind: "input" });
  sh.row("Interest Rate on Debt 债务利率", [p.interest_rate], { fmt: F.pct, numKind: "input" });
  sh.row(xlTaxLabel(p), [p.tax_rate], { fmt: F.pct, numKind: "input" });
  sh.row("Dividend Payout % 股利分配率", [p.dividend_pct], { fmt: F.pct, numKind: "input" });
  sh.row("CapEx % of Revenue 资本支出率", cxR.vals, { fmt: F.pct, numKind: "input" });
  sh.blank();
  sh.section("营运资本天数 Working Capital Days");
  sh.row("DSO 应收天数", [p.dso], { fmt: F.int, numKind: "input" });
  sh.row("DIO 库存天数", [p.dio], { fmt: F.int, numKind: "input" });
  sh.row("DPO 应付天数", [p.dpo], { fmt: F.int, numKind: "input" });
  sh.row("Accrued Days 应计天数", [p.accrued_days], { fmt: F.int, numKind: "input" });
  sh.blank();
  sh.section("期初资产负债表 Beginning Balance Sheet");
  sh.row("Beginning Cash 期初现金", [p.beg_cash], { fmt: F.num, numKind: "input" });
  sh.row("Beginning AR 期初应收", [p.beg_ar], { fmt: F.num, numKind: "input" });
  sh.row("Beginning Inventory 期初库存", [p.beg_inventory], { fmt: F.num, numKind: "input" });
  sh.row("Beginning PP&E 期初固定资产", [p.beg_ppe], { fmt: F.num, numKind: "input" });
  sh.row("Beginning AP 期初应付", [p.beg_ap], { fmt: F.num, numKind: "input" });
  sh.row("Beginning Accrued 期初应计", [p.beg_accrued], { fmt: F.num, numKind: "input" });
  sh.row("Beginning Debt 期初债务", [p.beg_debt], { fmt: F.num, numKind: "input" });
  sh.row("Common Stock 普通股", [p.common_stock], { fmt: F.num, numKind: "input" });
  sh.row("Beginning Retained Earnings 期初留存收益", [p.beg_retn_earn], { fmt: F.num, numKind: "input" });
  sh.blank();
  sh.section("其他 Other");
  var ndVals = [null]; for (var nd = 1; nd <= n; nd++) ndVals.push(p.new_debt_issuance);
  sh.row("New Debt Issuance (annual) 新增债务", ndVals, { fmt: F.num, numKind: "input" });

  // ---------------- Income Statement ----------------
  var si = xlSheet(wb, "Income Statement", NC, { widths: wA, freeze: [1, 3] });
  si.title("Income Statement 利润表");
  si.blank();
  si.header(headArr());
  si.section("利润表主体 Profit & Loss");
  si.row("Revenue 营业收入", yvals(function (yi) {
    if (yi === 0) return "=" + A + "!B5";
    return "=" + ycol(yi - 1) + "5*(1+" + A + "!" + ycol(yi) + "6)";
  }), { fmt: F.num });
  si.row("Revenue Growth % 增长率", yvals(function (yi) {
    if (yi === 0) return null;
    return "=" + ycol(yi) + "5/" + ycol(yi - 1) + "5-1";
  }), { fmt: F.pct });
  si.row("COGS 营业成本", yvals(function (yi) { return "=" + ycol(yi) + "5*" + aRef(A, cogs.per, 7, yi); }), { fmt: F.num });
  si.row("Gross Profit 毛利", yvals(function (yi) { return "=" + ycol(yi) + "5-" + ycol(yi) + "7"; }), { fmt: F.num, total: true });
  si.row("SG&A 销管费用", yvals(function (yi) { return "=" + ycol(yi) + "5*" + aRef(A, sgaR.per, 8, yi); }), { fmt: F.num });
  si.row("R&D 研发费用", yvals(function (yi) { return "=" + ycol(yi) + "5*" + aRef(A, rdR.per, 9, yi); }), { fmt: F.num });
  si.row("EBITDA", yvals(function (yi) { return "=" + ycol(yi) + "8-" + ycol(yi) + "9-" + ycol(yi) + "10"; }), { fmt: F.num, total: true });
  si.row("D&A 折旧摊销", yvals(function (yi) {
    if (yi === 0) return "=" + ycol(yi) + "5*" + aRef(A, daR.per, 10, 0);
    return "=" + SCH + "!" + ycol(yi) + "6";
  }), { fmt: F.num });
  si.row("EBIT 营业利润", yvals(function (yi) { return "=" + ycol(yi) + "11-" + ycol(yi) + "12"; }), { fmt: F.num, total: true });
  si.row("Interest Expense 利息费用", yvals(function (yi) {
    if (yi === 0) return "=" + A + "!$B$11*" + A + "!B29";
    return "=" + SCH + "!" + ycol(yi) + "17";
  }), { fmt: F.num });
  si.row("EBT 税前利润", yvals(function (yi) { return "=" + ycol(yi) + "13-" + ycol(yi) + "14"; }), { fmt: F.num, total: true });
  si.row("Taxes 所得税", yvals(function (yi) { return "=MAX(0," + ycol(yi) + "15*" + A + "!$B$12)"; }), { fmt: F.num });
  si.row("Net Income 净利润", yvals(function (yi) { return "=" + ycol(yi) + "15-" + ycol(yi) + "16"; }), { fmt: F.num, total: true });
  si.row("Dividends 股利", yvals(function (yi) { return "=" + ycol(yi) + "17*" + A + "!$B$13"; }), { fmt: F.num });
  si.blank();
  if (p.custom_items && p.custom_items.length) {
    si.section("自定义行项 Custom Line Items");
    function buildIsCellMap(col) {
      return {
        revenue: col + "5", cogs: col + "7", gross_profit: col + "8", sga: col + "9", rd: col + "10",
        ebitda: col + "11", da: col + "12", ebit: col + "13", interest: col + "14", ebt: col + "15",
        taxes: col + "16", net_income: col + "17", dividends: col + "18",
        cash: BS + "!" + col + "5", ar: BS + "!" + col + "6", inventory: BS + "!" + col + "7", ppe: BS + "!" + col + "8",
        total_assets: BS + "!" + col + "9", ap: BS + "!" + col + "13", accrued: BS + "!" + col + "14", debt: BS + "!" + col + "15",
        total_liabilities: BS + "!" + col + "16", common_stock: BS + "!" + col + "20", retained_earnings: BS + "!" + col + "21",
        total_equity: BS + "!" + col + "22", total_le: BS + "!" + col + "25", capex: SCH + "!" + col + "8", new_debt: SCH + "!" + col + "19"
      };
    }
    for (var idx = 0; idx < p.custom_items.length; idx++) {
      var ci = p.custom_items[idx];
      var civals = yvals(function (yi) {
        var col = ycol(yi);
        try { return expressionToExcel(ci.formula, buildIsCellMap(col)); }
        catch (e) { return '="#ERR: ' + String(e.message || e).substring(0, 30) + '"'; }
      });
      // yvals 闭包内 ci 用 var 已固定（顺序执行），无需额外绑定
      si.row(ci.name, civals, { fmt: F.num });
    }
  }

  // ---------------- Supporting Schedules ----------------
  var ss = xlSheet(wb, "Supporting Schedules", NC, { widths: wA, freeze: [1, 3] });
  ss.title("Supporting Schedules 辅助表");
  ss.blank();
  ss.header(headArr());
  ss.section("PP&E 滚动 PP&E Roll-forward");
  ss.blank();
  ss.row("D&A 折旧摊销", yvals(function (yi) { return "=" + IS + "!" + ycol(yi) + "5*" + aRef(A, daR.per, 10, yi); }), { fmt: F.num });
  ss.row("Beginning PP&E 期初固定资产", yvals(function (yi) {
    if (yi === 0) return "=" + A + "!B26";
    return "=" + ycol(yi - 1) + "10";
  }), { fmt: F.num });
  ss.row("CapEx 资本支出", yvals(function (yi) { return "=" + IS + "!" + ycol(yi) + "5*" + aRef(A, cxR.per, 14, yi); }), { fmt: F.num });
  ss.row("Less: D&A 减:折旧", yvals(function (yi) { return "=-" + ycol(yi) + "6"; }), { fmt: F.num });
  ss.row("Ending PP&E 期末固定资产", yvals(function (yi) {
    if (yi === 0) return "=" + ycol(yi) + "7";
    return "=" + ycol(yi) + "7+" + ycol(yi) + "8+" + ycol(yi) + "9";
  }), { fmt: F.num, total: true });
  ss.blank(); ss.blank(); ss.blank();
  ss.section("债务滚动 Debt Roll-forward");
  ss.blank(); ss.blank();
  ss.row("Interest Expense 利息费用", yvals(function (yi) {
    if (yi === 0) return "=" + A + "!$B$11*" + A + "!B29";
    return "=" + A + "!$B$11*" + BS + "!" + ycol(yi - 1) + "15";
  }), { fmt: F.num });
  ss.row("Beginning Debt 期初债务", yvals(function (yi) {
    if (yi === 0) return "=" + A + "!B29";
    return "=" + ycol(yi - 1) + "20";
  }), { fmt: F.num });
  ss.row("New Debt Issuance 新增债务", yvals(function (yi) {
    if (yi === 0) return "=0";
    return "=" + A + "!" + ycol(yi) + "34";
  }), { fmt: F.num });
  ss.row("Ending Debt 期末债务", yvals(function (yi) {
    if (yi === 0) return "=" + ycol(yi) + "18";
    return "=" + ycol(yi) + "18+" + ycol(yi) + "19";
  }), { fmt: F.num, total: true });

  // ---------------- Balance Sheet ----------------
  var sb = xlSheet(wb, "Balance Sheet", NC, { widths: wA, freeze: [1, 3] });
  sb.title("Balance Sheet 资产负债表");
  sb.blank();
  sb.header(headArr());
  sb.section("资产 Assets");
  sb.row("Cash 现金", yvals(function (yi) {
    if (yi === 0) return "=" + A + "!B23";
    return "=" + ycol(yi - 1) + "5+" + CFS + "!" + ycol(yi) + "23";
  }), { fmt: F.num });
  sb.row("Accounts Receivable 应收账款", yvals(function (yi) {
    if (yi === 0) return "=" + A + "!B24";
    return "=" + IS + "!" + ycol(yi) + "5/365*" + A + "!$B$17";
  }), { fmt: F.num });
  sb.row("Inventory 存货", yvals(function (yi) {
    if (yi === 0) return "=" + A + "!B25";
    return "=" + IS + "!" + ycol(yi) + "7/365*" + A + "!$B$18";
  }), { fmt: F.num });
  sb.row("PP&E, net 固定资产净额", yvals(function (yi) {
    if (yi === 0) return "=" + A + "!B26";
    return "=" + SCH + "!" + ycol(yi) + "10";
  }), { fmt: F.num });
  sb.row("Total Assets 总资产", yvals(function (yi) { return "=SUM(" + ycol(yi) + "5:" + ycol(yi) + "8)"; }), { fmt: F.num, total: true });
  sb.blank(); sb.blank();
  sb.section("负债 Liabilities");
  sb.row("Accounts Payable 应付账款", yvals(function (yi) {
    if (yi === 0) return "=" + A + "!B27";
    return "=" + IS + "!" + ycol(yi) + "7/365*" + A + "!$B$19";
  }), { fmt: F.num });
  sb.row("Accrued Expenses 应计费用", yvals(function (yi) {
    if (yi === 0) return "=" + A + "!B28";
    return "=" + IS + "!" + ycol(yi) + "5/365*" + A + "!$B$20";
  }), { fmt: F.num });
  sb.row("Debt 债务", yvals(function (yi) {
    if (yi === 0) return "=" + A + "!B29";
    return "=" + SCH + "!" + ycol(yi) + "20";
  }), { fmt: F.num });
  sb.row("Total Liabilities 总负债", yvals(function (yi) { return "=SUM(" + ycol(yi) + "13:" + ycol(yi) + "15)"; }), { fmt: F.num, total: true });
  sb.blank(); sb.blank();
  sb.section("权益 Equity");
  sb.row("Common Stock 普通股", yvals(function () { return "=" + A + "!$B$30"; }), { fmt: F.num });
  sb.row("Retained Earnings 留存收益", yvals(function (yi) {
    if (yi === 0) return "=" + A + "!B31";
    return "=" + ycol(yi - 1) + "21+" + IS + "!" + ycol(yi) + "17-" + IS + "!" + ycol(yi) + "18";
  }), { fmt: F.num });
  sb.row("Total Equity 总权益", yvals(function (yi) { return "=" + ycol(yi) + "20+" + ycol(yi) + "21"; }), { fmt: F.num, total: true });
  sb.blank(); sb.blank();
  sb.row("Total Liabilities + Equity 总负债权益", yvals(function (yi) { return "=" + ycol(yi) + "16+" + ycol(yi) + "22"; }), { fmt: F.num, total: true });
  sb.blank();
  sb.row("平衡校验 Balance Check", yvals(function (yi) {
    var col = ycol(yi);
    return '=IF(ROUND(' + col + '9-' + col + '25,2)=0,"BALANCED","OUT OF BALANCE")';
  }), {});

  // ---------------- Cash Flow Statement ----------------
  var sc = xlSheet(wb, "Cash Flow Statement", NC, { widths: wA, freeze: [1, 3] });
  sc.title("Cash Flow Statement 现金流量表");
  sc.blank();
  sc.header(headArr());
  sc.section("经营活动 Operating Activities");
  sc.blank();
  sc.row("Net Income 净利润", yvals(function (yi) { if (yi === 0) return null; return "=" + IS + "!" + ycol(yi) + "17"; }), { fmt: F.num });
  sc.row("D&A 折旧摊销", yvals(function (yi) { if (yi === 0) return null; return "=" + SCH + "!" + ycol(yi) + "6"; }), { fmt: F.num });
  sc.row("Change in AR 应收变动", yvals(function (yi) {
    if (yi === 0) return null;
    return "=-(" + BS + "!" + ycol(yi) + "6-" + BS + "!" + ycol(yi - 1) + "6)";
  }), { fmt: F.num });
  sc.row("Change in Inventory 存货变动", yvals(function (yi) {
    if (yi === 0) return null;
    return "=-(" + BS + "!" + ycol(yi) + "7-" + BS + "!" + ycol(yi - 1) + "7)";
  }), { fmt: F.num });
  sc.row("Change in AP 应付变动", yvals(function (yi) {
    if (yi === 0) return null;
    return "=" + BS + "!" + ycol(yi) + "13-" + BS + "!" + ycol(yi - 1) + "13";
  }), { fmt: F.num });
  sc.row("Change in Accrued 应计变动", yvals(function (yi) {
    if (yi === 0) return null;
    return "=" + BS + "!" + ycol(yi) + "14-" + BS + "!" + ycol(yi - 1) + "14";
  }), { fmt: F.num });
  sc.row("Cash from Operations 经营现金流", yvals(function (yi) {
    if (yi === 0) return null;
    return "=SUM(" + ycol(yi) + "6:" + ycol(yi) + "11)";
  }), { fmt: F.num, total: true });
  sc.blank();
  sc.section("投资活动 Investing Activities");
  sc.row("CapEx 资本支出", yvals(function (yi) {
    if (yi === 0) return null;
    return "=-" + IS + "!" + ycol(yi) + "5*" + aRef(A, cxR.per, 14, yi);
  }), { fmt: F.num });
  sc.row("Cash from Investing 投资现金流", yvals(function (yi) { if (yi === 0) return null; return "=" + ycol(yi) + "15"; }), { fmt: F.num, total: true });
  sc.blank();
  sc.section("筹资活动 Financing Activities");
  sc.row("Net Debt Change 债务净变动", yvals(function (yi) {
    if (yi === 0) return null;
    return "=" + BS + "!" + ycol(yi) + "15-" + BS + "!" + ycol(yi - 1) + "15";
  }), { fmt: F.num });
  sc.row("Dividends Paid 支付股利", yvals(function (yi) {
    if (yi === 0) return null;
    return "=-" + IS + "!" + ycol(yi) + "18";
  }), { fmt: F.num });
  sc.row("Cash from Financing 筹资现金流", yvals(function (yi) {
    if (yi === 0) return null;
    return "=SUM(" + ycol(yi) + "19:" + ycol(yi) + "20)";
  }), { fmt: F.num, total: true });
  sc.blank();
  sc.row("Net Change in Cash 现金净变动", yvals(function (yi) {
    if (yi === 0) return null;
    return "=" + ycol(yi) + "12+" + ycol(yi) + "16+" + ycol(yi) + "21";
  }), { fmt: F.num, total: true });

  // ---------------- Dashboard ----------------
  var sd = xlSheet(wb, "Dashboard", NC, { widths: wA, freeze: [1, 3] });
  sd.title("Dashboard 关键指标摘要");
  sd.blank();
  sd.header(headArr());
  sd.blank();
  sd.row("Revenue 营业收入", yvals(function (yi) { return "=" + IS + "!" + ycol(yi) + "5"; }), { fmt: F.num });
  sd.row("EBITDA", yvals(function (yi) { return "=" + IS + "!" + ycol(yi) + "11"; }), { fmt: F.num });
  sd.row("EBIT 营业利润", yvals(function (yi) { return "=" + IS + "!" + ycol(yi) + "13"; }), { fmt: F.num });
  sd.row("Net Income 净利润", yvals(function (yi) { return "=" + IS + "!" + ycol(yi) + "17"; }), { fmt: F.num });
  sd.row("Total Assets 总资产", yvals(function (yi) { return "=" + BS + "!" + ycol(yi) + "9"; }), { fmt: F.num });
  sd.row("Total Debt 总债务", yvals(function (yi) { return "=" + BS + "!" + ycol(yi) + "15"; }), { fmt: F.num });
  sd.row("Cash 现金", yvals(function (yi) { return "=" + BS + "!" + ycol(yi) + "5"; }), { fmt: F.num });
  sd.row("Retained Earnings 留存收益", yvals(function (yi) { return "=" + BS + "!" + ycol(yi) + "21"; }), { fmt: F.num });
  sd.blank();
  sd.row("EBITDA Margin %", yvals(function (yi) {
    if (yi === 0) return null;
    return "=" + IS + "!" + ycol(yi) + "11/" + IS + "!" + ycol(yi) + "5";
  }), { fmt: F.pct });
  sd.row("Net Margin %", yvals(function (yi) {
    if (yi === 0) return null;
    return "=" + IS + "!" + ycol(yi) + "17/" + IS + "!" + ycol(yi) + "5";
  }), { fmt: F.pct });
}

// --- DCF 模型 Excel ---
function buildDcfExcel(wb, p) {
  var n = p.projection_years;
  var NC = n + 2;
  var A = "Assumptions", OP = "'Operating Model'", W = "WACC", TV = "'Terminal Value'", DV = "'DCF Valuation'";
  var F = XL_FMT;
  var wA = [42]; for (var wc = 1; wc <= n + 1; wc++) wA.push(12);
  function yhead() { var h = ["项目 Item", "Year 0"]; for (var i = 1; i <= n; i++) h.push("Year " + i); return h; }
  function yvals(fn) { var v = []; for (var yi = 0; yi <= n; yi++) v.push(fn(yi)); return v; }

  var daR = xlRatioVals(p, "da_pct", n);
  var cxR = xlRatioVals(p, "capex_pct", n);
  var nwR = xlRatioVals(p, "nwc_pct", n);

  // ---------------- Assumptions ----------------
  var sh = xlSheet(wb, "Assumptions", NC, { widths: wA, freeze: [1, 3] });
  sh.title("Assumptions 假设参数");
  sh.blank();
  sh.header(["参数 Parameter", "Year 0"].concat((function () { var h = []; for (var i = 1; i <= n; i++) h.push("Year " + i); return h; })()));
  sh.section("运营假设 Operating Assumptions");
  sh.row("Revenue (base) 基期收入", [p.revenue_y0], { fmt: F.num, numKind: "input" });
  sh.row("Revenue Growth % 收入增长率（每年可不同）", xlYearlyVals(p, "rev_growth", n), { fmt: F.pct, numKind: "input" });
  sh.row("EBITDA Margin % EBITDA利润率（每年可不同）", [null].concat((function () {
    var v = []; for (var i = 1; i <= n; i++) { var x = parseFloat(p["ebitda_margin_y" + i]); v.push(isNaN(x) ? 0 : x); } return v;
  })()), { fmt: F.pct, numKind: "input" });
  sh.row("D&A % of Revenue 折旧摊销率", daR.vals, { fmt: F.pct, numKind: "input" });
  sh.row("CapEx % of Revenue 资本支出率", cxR.vals, { fmt: F.pct, numKind: "input" });
  sh.row("NWC % of Revenue 净营运资本率", nwR.vals, { fmt: F.pct, numKind: "input" });
  sh.row(xlTaxLabel(p), [p.tax_rate], { fmt: F.pct, numKind: "input" });
  sh.blank();
  sh.section("WACC 资本成本");
  sh.row("Risk-free Rate 无风险利率 Rf", [p.risk_free_rate], { fmt: F.pct, numKind: "input" });
  sh.row("Equity Risk Premium 股权风险溢价 ERP", [p.equity_risk_premium], { fmt: F.pct, numKind: "input" });
  sh.row("Beta β", [p.beta], { fmt: F.ratio2, numKind: "input" });
  sh.row("Pre-tax Cost of Debt 税前债务成本 Kd", [p.pre_tax_cost_of_debt], { fmt: F.pct, numKind: "input" });
  sh.row("Target Debt Weight 目标债务权重 Wd", [p.debt_weight], { fmt: F.pct, numKind: "input" });
  sh.row("Preferred Weight 优先股权重 Wp", [p.preferred_weight], { fmt: F.pct, numKind: "input" });
  sh.row("Cost of Preferred 优先股成本 Kp", [p.cost_of_preferred], { fmt: F.pct, numKind: "input" });
  sh.blank();
  sh.section("终值 Terminal Value");
  sh.row("Terminal Growth Rate g 终值增长率", [p.terminal_growth], { fmt: F.pct, numKind: "input" });
  sh.row("Exit EV/EBITDA 退出倍数", [p.exit_multiple], { fmt: F.mult, numKind: "input" });
  sh.row("TV Method 终值方法 (gordon/exit)", [p.tv_method], { numKind: "input" });
  sh.blank();
  sh.section("EV → Equity 桥 Bridge");
  sh.row("Total Debt 总债务（扣减）", [p.net_debt], { fmt: F.num, numKind: "input" });
  sh.row("Cash 现金（加回）", [p.cash], { fmt: F.num, numKind: "input" });
  sh.row("Minority Interest 少数股东权益（扣减）", [p.minority_interest], { fmt: F.num, numKind: "input" });
  sh.row("Preferred Stock 优先股账面值（扣减）", [p.preferred_stock], { fmt: F.num, numKind: "input" });
  sh.row("Stock-Based Compensation 股权激励（扣减）", [p.stock_based_comp], { fmt: F.num, numKind: "input" });
  sh.row("Pension Deficit 养老金缺口（扣减）", [p.pension_deficit], { fmt: F.num, numKind: "input" });
  sh.row("Shares Outstanding (M) 流通股本（百万）", [p.shares_outstanding], { fmt: F.num0, numKind: "input" });
  sh.row("Current Share Price 当前股价", [p.current_price], { fmt: F.price, numKind: "input" });
  sh.row("Valuation Timing 估值时点 (mid/end)", [p.valuation_timing], { numKind: "input" });

  // ---------------- Operating Model ----------------
  var so = xlSheet(wb, "Operating Model", NC, { widths: wA, freeze: [1, 3] });
  so.title("Operating Model 运营模型");
  so.blank();
  so.header(yhead());
  so.section("无杠杆自由现金流 Unlevered Free Cash Flow");
  so.row("Revenue 营业收入", yvals(function (yi) {
    if (yi === 0) return "=" + A + "!B5";
    return "=" + ycol(yi - 1) + "5*(1+" + A + "!" + ycol(yi) + "6)";
  }), { fmt: F.num });
  so.row("Revenue Growth % 增长率", yvals(function (yi) {
    if (yi === 0) return null;
    return "=" + ycol(yi) + "5/" + ycol(yi - 1) + "5-1";
  }), { fmt: F.pct });
  so.row("EBITDA", yvals(function (yi) {
    if (yi > 0) return "=" + ycol(yi) + "5*" + A + "!" + ycol(yi) + "7";
    return "=" + ycol(yi) + "5*" + A + "!C7";
  }), { fmt: F.num, total: true });
  so.row("EBITDA Margin %", yvals(function (yi) {
    if (yi === 0) return null;
    return "=" + ycol(yi) + "7/" + ycol(yi) + "5";
  }), { fmt: F.pct });
  so.row("D&A 折旧摊销", yvals(function (yi) { return "=" + ycol(yi) + "5*" + aRef(A, daR.per, 8, yi); }), { fmt: F.num });
  so.row("EBIT 营业利润", yvals(function (yi) { return "=" + ycol(yi) + "7-" + ycol(yi) + "9"; }), { fmt: F.num, total: true });
  so.row("EBIT Margin %", yvals(function (yi) {
    if (yi === 0) return null;
    return "=" + ycol(yi) + "10/" + ycol(yi) + "5";
  }), { fmt: F.pct });
  so.row("NOPAT 税后营业利润", yvals(function (yi) { return "=" + ycol(yi) + "10*(1-" + A + "!$B$11)"; }), { fmt: F.num });
  so.row("Δ NWC 净营运资本变动", yvals(function (yi) {
    if (yi === 0) return null;
    return "=(" + ycol(yi) + "5-" + ycol(yi - 1) + "5)*" + aRef(A, nwR.per, 10, yi);
  }), { fmt: F.num });
  so.row("CapEx 资本支出", yvals(function (yi) { return "=" + ycol(yi) + "5*" + aRef(A, cxR.per, 9, yi); }), { fmt: F.num });
  so.row("Unlevered FCF 无杠杆自由现金流", yvals(function (yi) {
    if (yi === 0) return null;
    return "=" + ycol(yi) + "12+" + ycol(yi) + "9-" + ycol(yi) + "13-" + ycol(yi) + "14";
  }), { fmt: F.num, total: true });

  if (p.custom_items && p.custom_items.length) {
    so.blank();
    so.section("自定义行项 Custom Line Items");
    function buildOpCellMap(col) {
      return {
        revenue: col + "5", ebitda: col + "7", da: col + "9", ebit: col + "10", nopat: col + "12",
        delta_nwc: col + "13", capex: col + "14", ufcf: col + "15", wacc: W + "!$B$22",
        cost_of_equity: W + "!$B$8", cost_of_preferred: W + "!$B$14",
        preferred_weight: W + "!$B$17", equity_weight: W + "!$B$18",
        after_tax_kd: W + "!$B$12", tv_gordon: TV + "!$B$5", tv_exit: TV + "!$B$6",
        selected_tv: TV + "!$B$7", pv_tv: TV + "!$B$9", sum_pv_fcf: DV + "!$B$9",
        enterprise_value: DV + "!$B$11", equity_value: DV + "!$B$21", implied_price: DV + "!$B$23",
        current_price: A + "!$B$35", shares_outstanding: A + "!$B$34",
        net_debt: A + "!$B$28", cash: A + "!$B$29"
      };
    }
    for (var idx = 0; idx < p.custom_items.length; idx++) {
      var ci = p.custom_items[idx];
      var civals = yvals(function (yi) {
        var col = ycol(yi);
        try { return expressionToExcel(ci.formula, buildOpCellMap(col)); }
        catch (e) { return '="#ERR: ' + String(e.message || e).substring(0, 30) + '"'; }
      });
      so.row(ci.name, civals, { fmt: F.num });
    }
  }

  // ---------------- WACC ----------------
  var sw = xlSheet(wb, "WACC", 2, { widths: [44, 16], freeze: [0, 3] });
  sw.title("WACC 加权平均资本成本");
  sw.blank();
  sw.header(["项目 Item", "Value"]);
  sw.section("普通股成本 Cost of Equity (CAPM)");
  sw.row("Risk-free Rate (Rf) 无风险利率", ["=" + A + "!B14"], { fmt: F.pct });
  sw.row("Equity Risk Premium (ERP) 股权风险溢价", ["=" + A + "!B15"], { fmt: F.pct });
  sw.row("Beta β", ["=" + A + "!B16"], { fmt: F.ratio2 });
  sw.row("Cost of Equity (Ke = Rf + β×ERP) 普通股成本", ["=B5+B6*B7"], { fmt: F.pct, total: true });
  sw.section("债务与优先股成本 Debt & Preferred");
  sw.row("Pre-tax Cost of Debt 税前债务成本 Kd", ["=" + A + "!B17"], { fmt: F.pct });
  sw.row("Tax Rate 税率（" + (p.tax_basis === "mtr" ? "MTR 边际税率" : "ETR 有效税率") + "）", ["=" + A + "!B11"], { fmt: F.pct });
  sw.row("After-tax Cost of Debt 税后债务成本", ["=B10*(1-B11)"], { fmt: F.pct });
  sw.row("Preferred Weight 优先股权重 Wp", ["=" + A + "!B19"], { fmt: F.pct });
  sw.row("Cost of Preferred (Kp) 优先股成本", ["=" + A + "!B20"], { fmt: F.pct });
  sw.section("目标资本结构 Target Capital Structure");
  sw.row("Debt Weight 债务权重 Wd", ["=" + A + "!B17"], { fmt: F.pct });
  sw.row("Preferred Weight 优先股权重 Wp", ["=" + A + "!B19"], { fmt: F.pct });
  sw.row("Common Equity Weight 普通股权重 We (=1−Wd−Wp)", ["=1-B16-B17"], { fmt: F.pct });
  sw.row("权重合计 Sum of Weights（应=100%）", ["=B16+B17+B18"], { fmt: F.pct });
  sw.row("权重校验 Weight Check", ['=IF(ROUND(B19,4)=1,"OK","请检查权重")'], {});
  sw.section("加权平均资本成本 Weighted Average");
  sw.row("WACC = Ke×We + Kp×Wp + Kd×(1−t)×Wd", ["=B8*B18+B14*B17+B12*B16"], { fmt: F.pct, total: true });
  sw.blank();
  sw.note("说明：WACC 为三层资本结构加权——普通股 Ke（CAPM）×自动权重 We、优先股 Kp×Wp、税后债务成本 Kd×(1−税率)×Wd。仅债务利息可抵税，故债务项使用税后成本；优先股与普通股股利均不可抵税。税率口径（ETR 有效税率 / MTR 边际税率）取自 Assumptions。", 56);

  // ---------------- Terminal Value ----------------
  var lastCol = ycol(n);
  var st = xlSheet(wb, "Terminal Value", 2, { widths: [46, 18], freeze: [0, 3] });
  st.title("Terminal Value 终值");
  st.blank();
  st.header(["项目 Item", "Value"]);
  st.blank();
  st.row("TV (Gordon Growth) 戈登增长终值", ["=" + OP + "!" + lastCol + "15*(1+" + A + "!$B$23)/(" + W + "!$B$22-" + A + "!$B$23)"], { fmt: F.num });
  st.row("TV (Exit Multiple) 退出倍数终值", ["=" + OP + "!" + lastCol + "7*" + A + "!$B$24"], { fmt: F.num });
  st.row("Selected TV 选定终值", ['=IF(' + A + '!$B$25="gordon",B5,B6)'], { fmt: F.num, total: true });
  st.row("Discount Period (Years) 折现期数", ["=" + n], { fmt: F.int });
  st.row("PV of Terminal Value 终值现值", ["=B7/(1+" + W + "!$B$22)^B8"], { fmt: F.num, total: true });
  st.blank();
  st.note("两种终值方法说明：\n① 戈登增长法（永续增长）：假设预测期后企业以终值增长率 g 永续经营，TV = 末年FCF×(1+g)/(WACC−g)。\n② 退出倍数法：假设在退出年按“退出年 EBITDA × 退出 EV/EBITDA 倍数”出售企业。\n“Selected TV”根据 Assumptions 中的 TV Method（gordon/exit）自动二选一；该终值折现后（PV of TV）计入 DCF Valuation 的企业价值 EV。", 84);

  // ---------------- DCF Valuation ----------------
  var NCV = n + 1; // A + Year1..n
  var wv = [42]; for (var wc2 = 1; wc2 <= n; wc2++) wv.push(12);
  var sv = xlSheet(wb, "DCF Valuation", NCV, { widths: wv, freeze: [1, 3] });
  sv.title("DCF Valuation 估值汇总");
  sv.blank();
  var dhead = ["项目 Item"]; for (var i = 1; i <= n; i++) dhead.push("Year " + i);
  sv.header(dhead);
  sv.section("预测期现金流折现 Discounting");
  var isMid = p.valuation_timing === "mid";
  function dvVals(fn) { var v = []; for (var yi = 1; yi <= n; yi++) v.push(fn(yi)); return v; }
  sv.row("Unlevered FCF 自由现金流", dvVals(function (yi) { return "=" + OP + "!" + ycol(yi) + "15"; }), { fmt: F.num });
  sv.row("Discount Period 折现期数", dvVals(function (yi) { return "=" + (isMid ? yi - 0.5 : yi); }), { fmt: F.ratio2 });
  sv.row("Discount Factor 折现因子", dvVals(function (yi) { return "=1/(1+" + W + "!$B$22)^" + ycol(yi) + "6"; }), { fmt: F.ratio2 });
  sv.row("PV of FCF FCF现值", dvVals(function (yi) { return "=" + ycol(yi) + "5*" + ycol(yi) + "7"; }), { fmt: F.num });
  sv.row("Sum of PV of FCF FCF现值合计", ["=SUM(" + ycol(1) + "8:" + ycol(n) + "8)"], { fmt: F.num, total: true });
  sv.row("PV of Terminal Value 终值现值", ["=" + TV + "!B9"], { fmt: F.num, total: true });
  sv.row("Enterprise Value 企业价值 (EV=ΣPV(FCF)+PV(TV))", ["=B9+B10"], { fmt: F.num, total: true });
  sv.blank();
  sv.section("EV → Equity 桥 Bridge to Equity Value");
  sv.row("Enterprise Value 企业价值", ["=B11"], { fmt: F.num });
  sv.row("Less: Total Debt 减:总债务", ["=-" + A + "!B28"], { fmt: F.num });
  sv.row("Plus: Cash 加:现金", ["=" + A + "!B29"], { fmt: F.num });
  sv.row("Less: Minority Interest 减:少数股东权益", ["=-" + A + "!B30"], { fmt: F.num });
  sv.row("Less: Preferred Stock 减:优先股账面值", ["=-" + A + "!B31"], { fmt: F.num });
  sv.row("Less: Stock-Based Comp 减:股权激励", ["=-" + A + "!B32"], { fmt: F.num });
  sv.row("Less: Pension Deficit 减:养老金缺口", ["=-" + A + "!B33"], { fmt: F.num });
  sv.row("Equity Value 股权价值", ["=SUM(B14:B20)"], { fmt: F.num, total: true });
  sv.row("Shares Outstanding (M) 流通股本（百万）", ["=" + A + "!B34"], { fmt: F.num0 });
  sv.row("Implied Share Price 隐含股价", ["=B21/B22"], { fmt: F.price, total: true });
  sv.row("Current Share Price 当前股价", ["=" + A + "!B35"], { fmt: F.price });
  sv.row("Upside/(Downside) % 涨跌幅", ["=B23/B24-1"], { fmt: F.pct, total: true });

  // ---------------- Sensitivity ----------------
  var s2 = xlSheet(wb, "Sensitivity", 6, { widths: [18, 13, 13, 13, 13, 13] });
  s2.title("Sensitivity 敏感性分析");
  s2.note("行=终值增长率 g，列=WACC；单元格=隐含股价（按戈登增长法重算终值）。", 28);
  s2.blank();
  var gVals = [p.terminal_growth - 0.01, p.terminal_growth - 0.005, p.terminal_growth, p.terminal_growth + 0.005, p.terminal_growth + 0.01];
  var wVals = [0.08, 0.09, 0.10, 0.11, 0.12];
  var head5 = ["g \\ WACC"].concat(wVals);
  s2.header(head5);
  for (var hc = 2; hc <= 6; hc++) s2.ws.getCell(4, hc).numFmt = F.pct;
  for (var gi = 0; gi < gVals.length; gi++) {
    var rowVals = [];
    for (var gj = 0; gj < wVals.length; gj++) {
      var wc = ycol(gj) + "$4", gcc = "$A" + (5 + gi);
      var pvTerms = [];
      for (var t = 1; t <= n; t++) pvTerms.push(OP + "!" + ycol(t) + "15/(1+" + wc + ")^" + t);
      var tvExpr = OP + "!" + lastCol + "15*(1+" + gcc + ")/(" + wc + "-" + gcc + ")";
      var pvTv = tvExpr + "/(1+" + wc + ")^" + n;
      var ev = "(" + pvTerms.join("+") + ")+" + pvTv;
      var equity = "(" + ev + ")-" + A + "!$B$28+" + A + "!$B$29-" + A + "!$B$30-" + A + "!$B$31-" + A + "!$B$32-" + A + "!$B$33";
      rowVals.push({ v: "=((" + equity + ")/" + A + "!$B$34)", f: F.price });
    }
    s2.row(gVals[gi], rowVals, { fmt: F.pct, numKind: "input" });
  }
}

// --- LBO 模型 Excel ---
function buildLboExcel(wb, p) {
  var n = Math.max(1, Math.min(10, parseInt(p.exit_year, 10) || 5));
  var NC = n + 2;
  var A = "Assumptions", SU = "'Sources & Uses'", DS = "'Debt Schedule'", IS = "'Income Statement'", CF = "'Cash Flow (CFADS)'";
  var F = XL_FMT;
  var wA = [42, 13]; for (var wc = 1; wc <= n; wc++) wA.push(12);
  function yHeadL(first) { var h = ["项目 Item", first]; for (var i = 1; i <= n; i++) h.push("Year " + i); return h; }
  function yvals(fn) { var v = []; for (var yi = 0; yi <= n; yi++) v.push(fn(yi)); return v; }

  var mgR = xlRatioVals(p, "ebitda_margin", n);
  var daR = xlRatioVals(p, "da_pct", n);
  var cxR = xlRatioVals(p, "capex_pct", n);
  var nwR = xlRatioVals(p, "nwc_pct", n);

  // ---------------- Assumptions（行号与公式引用严格对齐） ----------------
  var sh = xlSheet(wb, "Assumptions", NC, { widths: wA, freeze: [1, 3] });
  sh.title("Assumptions 假设参数");
  sh.blank();
  sh.header(["参数 Parameter", "Value"].concat((function () { var h = []; for (var i = 1; i <= n; i++) h.push("Year " + i); return h; })()));
  sh.section("交易假设 Transaction Assumptions");
  sh.row("LTM Revenue 最近12个月收入", [p.ltm_revenue], { fmt: F.num, numKind: "input" });
  sh.row("LTM EBITDA 最近12个月EBITDA", [p.ltm_ebitda], { fmt: F.num, numKind: "input" });
  sh.row("Entry EV/EBITDA 入场倍数", [p.entry_ev_ebitda], { fmt: F.mult, numKind: "input" });
  sh.row("Entry EV 入场企业价值", ["=B6*B7"], { fmt: F.num });
  sh.row("Existing Net Debt 现有净债务", [p.existing_net_debt], { fmt: F.num, numKind: "input" });
  sh.row("Existing Cash 现有现金", [p.existing_cash], { fmt: F.num, numKind: "input" });
  sh.row("Transaction Fees 交易费用", [p.transaction_fees], { fmt: F.num, numKind: "input" });
  sh.row("Financing Fees 融资费用", [p.financing_fees], { fmt: F.num, numKind: "input" });
  sh.section("资本结构 Capital Structure");
  sh.row("Sponsor Equity 发起人股本", [p.sponsor_equity], { fmt: F.num, numKind: "input" });
  sh.row("Revolver Capacity 循环额度上限", [p.revolver_capacity], { fmt: F.num, numKind: "input" });
  sh.row("Term Loan A", [p.term_loan_a], { fmt: F.num, numKind: "input" });
  sh.row("Term Loan B", [p.term_loan_b], { fmt: F.num, numKind: "input" });
  sh.row("Senior Notes 优先票据", [p.senior_notes], { fmt: F.num, numKind: "input" });
  sh.row("Subordinated Debt 次级债务", [p.subordinated_debt], { fmt: F.num, numKind: "input" });
  sh.section("利率 Interest Rates");
  sh.row("Revolver Rate 循环额度利率", [p.revolver_rate], { fmt: F.pct, numKind: "input" });
  sh.row("Term Loan A Rate", [p.tla_rate], { fmt: F.pct, numKind: "input" });
  sh.row("Term Loan B Rate", [p.tlb_rate], { fmt: F.pct, numKind: "input" });
  sh.row("Senior Notes Rate", [p.senior_notes_rate], { fmt: F.pct, numKind: "input" });
  sh.row("Subordinated Rate 次级债务利率", [p.sub_rate], { fmt: F.pct, numKind: "input" });
  sh.section("偿债 Debt Repayment");
  sh.row("TLA Mandatory Amort % TLA强制摊销率", [p.tla_mandatory_amort], { fmt: F.pct, numKind: "input" });
  sh.row("TLB Mandatory Amort % TLB强制摊销率", [p.tlb_mandatory_amort], { fmt: F.pct, numKind: "input" });
  sh.row("Cash Sweep % 现金清偿比例", [p.cash_sweep_pct], { fmt: F.pct, numKind: "input" });
  sh.row("Min Cash Balance 最低现金余额", [p.min_cash_balance], { fmt: F.num, numKind: "input" });
  sh.section("运营假设 Operating Assumptions");
  sh.row("Revenue Growth % 收入增长率（每年可不同）", xlYearlyVals(p, "rev_growth", n), { fmt: F.pct, numKind: "input" });
  sh.row("EBITDA Margin % EBITDA利润率", mgR.vals, { fmt: F.pct, numKind: "input" });
  sh.row("D&A % of Revenue 折旧摊销率", daR.vals, { fmt: F.pct, numKind: "input" });
  sh.row("CapEx % of Revenue 资本支出率", cxR.vals, { fmt: F.pct, numKind: "input" });
  sh.row("ΔNWC % of Revenue 净营运资本率", nwR.vals, { fmt: F.pct, numKind: "input" });
  sh.row(xlTaxLabel(p), [p.tax_rate], { fmt: F.pct, numKind: "input" });
  sh.section("退出 Exit");
  sh.row("Exit EV/EBITDA 退出倍数", [p.exit_ev_ebitda], { fmt: F.mult, numKind: "input" });
  sh.row("Exit Year 退出年（=预测年数）", [p.exit_year], { fmt: F.int, numKind: "input" });
  sh.row("Cash Interest Rate 现金存款利率", [p.cash_interest_rate], { fmt: F.pct, numKind: "input" });

  // ---------------- Sources & Uses ----------------
  var su = xlSheet(wb, "Sources & Uses", 2, { widths: [46, 18], freeze: [0, 3] });
  su.title("Sources & Uses 资金来源与用途");
  su.blank();
  su.header(["项目 Item", "Amount"]);
  su.section("Uses 用途");
  su.blank();
  su.row("Purchase of Equity 购买股权", ["=" + A + "!B8-" + A + "!B9"], { fmt: F.num });
  su.row("Refinance Existing Debt 再融资债务", ["=" + A + "!B9+" + A + "!B10"], { fmt: F.num });
  su.row("Transaction Fees 交易费用", ["=" + A + "!B11"], { fmt: F.num });
  su.row("Financing Fees 融资费用", ["=" + A + "!B12"], { fmt: F.num });
  su.row("Total Uses 用途合计", ["=SUM(B6:B9)"], { fmt: F.num, total: true });
  su.blank();
  su.section("Sources 来源");
  su.row("Sponsor Equity 发起人股本", ["=" + A + "!B14"], { fmt: F.num });
  su.row("Revolver Draw (plug) 循环额度提取", ["=B10-B13-B15-B16-B17-B18-B19"], { fmt: F.num });
  su.row("Term Loan A", ["=" + A + "!B16"], { fmt: F.num });
  su.row("Term Loan B", ["=" + A + "!B17"], { fmt: F.num });
  su.row("Senior Notes 优先票据", ["=" + A + "!B18"], { fmt: F.num });
  su.row("Subordinated Debt 次级债务", ["=" + A + "!B19"], { fmt: F.num });
  su.row("Existing Cash 现有现金", ["=" + A + "!B10"], { fmt: F.num });
  su.row("Total Sources 来源合计", ["=SUM(B13:B19)"], { fmt: F.num, total: true });
  su.blank();
  su.row("平衡校验 Balance Check", ['=IF(ROUND(B20-B10,2)=0,"BALANCED","IMBALANCED")'], {});

  // ---------------- Debt Schedule ----------------
  var sd = xlSheet(wb, "Debt Schedule", NC, { widths: wA, freeze: [1, 3] });
  sd.title("Debt Schedule 债务滚动表");
  sd.blank();
  sd.header(yHeadL("Close"));
  sd.section("现金与偿债可用 Cash & CFADS");
  sd.row("Beginning Cash 期初现金", yvals(function (yi) {
    if (yi === 0) return "=" + A + "!B10";
    return "=" + ycol(yi - 1) + "35";
  }), { fmt: F.num });
  sd.row("CFADS 偿债可用现金流", yvals(function (yi) {
    if (yi === 0) return "=0";
    return "=" + CF + "!" + ycol(yi) + "13";
  }), { fmt: F.num });
  sd.row("Min Cash 最低现金", yvals(function () { return "=" + A + "!$B$30"; }), { fmt: F.num });
  sd.row("Cash Available 可用现金", yvals(function (yi) {
    return "=MAX(0," + ycol(yi) + "5+" + ycol(yi) + "6-" + ycol(yi) + "7)";
  }), { fmt: F.num, total: true });
  sd.section("Revolver 循环额度");
  sd.blank();
  sd.row("Beginning 期初", yvals(function (yi) {
    if (yi === 0) return "=" + SU + "!B14";
    return "=" + ycol(yi - 1) + "13";
  }), { fmt: F.num });
  sd.row("Repayment 偿还", yvals(function (yi) {
    return "=MIN(" + ycol(yi) + "11," + ycol(yi) + "8)";
  }), { fmt: F.num });
  sd.row("Ending 期末", yvals(function (yi) {
    return "=" + ycol(yi) + "11-" + ycol(yi) + "12";
  }), { fmt: F.num, total: true });
  sd.section("Term Loan A");
  sd.blank();
  sd.row("Beginning 期初", yvals(function (yi) {
    if (yi === 0) return "=" + A + "!B16";
    return "=" + ycol(yi - 1) + "18";
  }), { fmt: F.num });
  sd.row("Mandatory Amort 强制摊销", yvals(function (yi) {
    return "=MIN(" + ycol(yi) + "16," + A + "!$B$16*" + A + "!$B$27)";
  }), { fmt: F.num });
  sd.row("Ending 期末", yvals(function (yi) {
    return "=" + ycol(yi) + "16-" + ycol(yi) + "17";
  }), { fmt: F.num, total: true });
  sd.section("Term Loan B");
  sd.blank();
  sd.row("Beginning 期初", yvals(function (yi) {
    if (yi === 0) return "=" + A + "!B17";
    return "=" + ycol(yi - 1) + "24";
  }), { fmt: F.num });
  sd.row("Mandatory Amort 强制摊销", yvals(function (yi) {
    return "=MIN(" + ycol(yi) + "21," + A + "!$B$17*" + A + "!$B$28)";
  }), { fmt: F.num });
  sd.row("Cash Sweep 现金清偿", yvals(function (yi) {
    return "=MIN(" + ycol(yi) + "21-" + ycol(yi) + "22,MAX(0," + ycol(yi) + "8-" + ycol(yi) + "12-" + ycol(yi) + "17-" + ycol(yi) + "22)*" + A + "!$B$29)";
  }), { fmt: F.num });
  sd.row("Ending 期末", yvals(function (yi) {
    return "=" + ycol(yi) + "21-" + ycol(yi) + "22-" + ycol(yi) + "23";
  }), { fmt: F.num, total: true });
  sd.section("Senior Notes 优先票据");
  sd.blank();
  sd.row("Beginning 期初", yvals(function (yi) {
    if (yi === 0) return "=" + A + "!B18";
    return "=" + ycol(yi - 1) + "28";
  }), { fmt: F.num });
  sd.row("Ending 期末", yvals(function (yi) { return "=" + ycol(yi) + "27"; }), { fmt: F.num, total: true });
  sd.section("Subordinated Debt 次级债务");
  sd.blank();
  sd.row("Beginning 期初", yvals(function (yi) {
    if (yi === 0) return "=" + A + "!B19";
    return "=" + ycol(yi - 1) + "32";
  }), { fmt: F.num });
  sd.row("Ending 期末", yvals(function (yi) { return "=" + ycol(yi) + "31"; }), { fmt: F.num, total: true });
  sd.blank();
  sd.row("Total Debt 总债务", yvals(function (yi) {
    return "=" + ycol(yi) + "13+" + ycol(yi) + "18+" + ycol(yi) + "24+" + ycol(yi) + "28+" + ycol(yi) + "32";
  }), { fmt: F.num, total: true });
  sd.row("Cash Ending 期末现金", yvals(function (yi) {
    return "=MAX(" + ycol(yi) + "7," + ycol(yi) + "5+" + ycol(yi) + "6-" + ycol(yi) + "12-" + ycol(yi) + "17-" + ycol(yi) + "22-" + ycol(yi) + "23)";
  }), { fmt: F.num, total: true });

  // ---------------- Income Statement ----------------
  var si = xlSheet(wb, "Income Statement", NC, { widths: wA, freeze: [1, 3] });
  si.title("Income Statement 利润表");
  si.blank();
  si.header(yHeadL("LTM"));
  si.section("利润表主体 Profit & Loss");
  si.row("Revenue 营业收入", yvals(function (yi) {
    if (yi === 0) return "=" + A + "!B5";
    return "=" + ycol(yi - 1) + "5*(1+" + A + "!" + ycol(yi) + "32)";
  }), { fmt: F.num });
  si.row("Revenue Growth % 增长率", yvals(function (yi) {
    if (yi === 0) return null;
    return "=" + ycol(yi) + "5/" + ycol(yi - 1) + "5-1";
  }), { fmt: F.pct });
  si.row("EBITDA", yvals(function (yi) {
    if (yi === 0) return "=" + A + "!B6";
    return "=" + ycol(yi) + "5*" + aRef(A, mgR.per, 33, yi);
  }), { fmt: F.num, total: true });
  si.row("EBITDA Margin %", yvals(function (yi) {
    if (yi === 0) return null;
    return "=" + ycol(yi) + "7/" + ycol(yi) + "5";
  }), { fmt: F.pct });
  si.row("D&A 折旧摊销", yvals(function (yi) { return "=" + ycol(yi) + "5*" + aRef(A, daR.per, 34, yi); }), { fmt: F.num });
  si.row("EBIT 营业利润", yvals(function (yi) { return "=" + ycol(yi) + "7-" + ycol(yi) + "9"; }), { fmt: F.num, total: true });
  si.blank();
  si.row("Interest - Revolver 利息(循环)", yvals(function (yi) {
    return "=" + A + "!$B$21*" + DS + "!" + ycol(yi) + "11";
  }), { fmt: F.num });
  si.row("Interest - TL A 利息(TLA)", yvals(function (yi) {
    return "=" + A + "!$B$22*" + DS + "!" + ycol(yi) + "16";
  }), { fmt: F.num });
  si.row("Interest - TL B 利息(TLB)", yvals(function (yi) {
    return "=" + A + "!$B$23*" + DS + "!" + ycol(yi) + "21";
  }), { fmt: F.num });
  si.row("Interest - Sr Notes 利息(优先票据)", yvals(function (yi) {
    return "=" + A + "!$B$24*" + DS + "!" + ycol(yi) + "27";
  }), { fmt: F.num });
  si.row("Interest - Sub 利息(次级)", yvals(function (yi) {
    return "=" + A + "!$B$25*" + DS + "!" + ycol(yi) + "31";
  }), { fmt: F.num });
  si.row("Total Interest 利息合计", yvals(function (yi) {
    return "=SUM(" + ycol(yi) + "12:" + ycol(yi) + "16)";
  }), { fmt: F.num, total: true });
  si.row("Interest Income 利息收入", yvals(function (yi) {
    return "=" + A + "!$B$41*" + DS + "!" + ycol(yi) + "5";
  }), { fmt: F.num });
  si.row("EBT 税前利润", yvals(function (yi) {
    return "=" + ycol(yi) + "10-" + ycol(yi) + "17+" + ycol(yi) + "18";
  }), { fmt: F.num, total: true });
  si.row("Taxes 所得税", yvals(function (yi) {
    return "=MAX(0," + ycol(yi) + "19*" + A + "!$B$37)";
  }), { fmt: F.num });
  si.row("Net Income 净利润", yvals(function (yi) {
    return "=" + ycol(yi) + "19-" + ycol(yi) + "20";
  }), { fmt: F.num, total: true });

  if (p.custom_items && p.custom_items.length) {
    si.section("自定义行项 Custom Line Items");
    si.blank();
    function buildIsCellMap(col) {
      return {
        revenue: col + "5", ebitda: col + "7", da: col + "9", ebit: col + "10",
        interest_revolver: col + "12", interest_tla: col + "13", interest_tlb: col + "14",
        interest_sn: col + "15", interest_sub: col + "16", total_interest: col + "17",
        interest_income: col + "18", ebt: col + "19", taxes: col + "20", net_income: col + "21",
        total_debt: DS + "!" + col + "34", cash: DS + "!" + col + "35", revolver: DS + "!" + col + "13",
        tla: DS + "!" + col + "18", tlb: DS + "!" + col + "24", senior_notes: DS + "!" + col + "28",
        subordinated_debt: DS + "!" + col + "32", cfads: CF + "!" + col + "13"
      };
    }
    for (var idx = 0; idx < p.custom_items.length; idx++) {
      var ci = p.custom_items[idx];
      var civals = yvals(function (yi) {
        var col = ycol(yi);
        try { return expressionToExcel(ci.formula, buildIsCellMap(col)); }
        catch (e) { return '="#ERR: ' + String(e.message || e).substring(0, 30) + '"'; }
      });
      si.row(ci.name, civals, { fmt: F.num });
    }
  }

  // ---------------- Cash Flow (CFADS) ----------------
  var sc = xlSheet(wb, "Cash Flow (CFADS)", NC, { widths: wA, freeze: [1, 3] });
  sc.title("Cash Flow (CFADS) 偿债可用现金流");
  sc.blank();
  sc.header(yHeadL("LTM"));
  sc.blank();
  sc.row("Net Income 净利润", yvals(function (yi) {
    if (yi === 0) return null;
    return "=" + IS + "!" + ycol(yi) + "21";
  }), { fmt: F.num });
  sc.blank();
  sc.row("D&A 折旧摊销", yvals(function (yi) {
    if (yi === 0) return null;
    return "=" + IS + "!" + ycol(yi) + "9";
  }), { fmt: F.num });
  sc.row("Δ NWC 净营运资本变动", yvals(function (yi) {
    if (yi === 0) return null;
    return "=-" + IS + "!" + ycol(yi) + "5*" + aRef(A, nwR.per, 36, yi);
  }), { fmt: F.num });
  sc.row("Cash from Operations 经营现金流", yvals(function (yi) {
    if (yi === 0) return null;
    return "=" + ycol(yi) + "5+" + ycol(yi) + "7+" + ycol(yi) + "8";
  }), { fmt: F.num, total: true });
  sc.blank();
  sc.row("CapEx 资本支出", yvals(function (yi) {
    if (yi === 0) return null;
    return "=-" + IS + "!" + ycol(yi) + "5*" + aRef(A, cxR.per, 35, yi);
  }), { fmt: F.num });
  sc.blank();
  sc.row("CFADS 偿债可用现金流", yvals(function (yi) {
    if (yi === 0) return null;
    return "=" + ycol(yi) + "9+" + ycol(yi) + "11";
  }), { fmt: F.num, total: true });

  // ---------------- Exit & Returns ----------------
  var exitCol = ycol(n);
  var er = xlSheet(wb, "Exit & Returns", NC, { widths: wA, freeze: [1, 3] });
  er.title("Exit & Returns 退出与回报");
  er.blank();
  er.header(["项目 Item", "Value"]);
  er.blank();
  er.row("Exit Year (" + n + ") EBITDA 退出年EBITDA", ["=" + IS + "!" + exitCol + "7"], { fmt: F.num });
  er.row("Exit EV/EBITDA 退出倍数", ["=" + A + "!B39"], { fmt: F.mult });
  er.row("Exit Enterprise Value 退出企业价值", ["=B5*B6"], { fmt: F.num, total: true });
  er.row("Less: Total Debt 减:总债务", ["=-" + DS + "!" + exitCol + "34"], { fmt: F.num });
  er.row("Plus: Cash 加:现金", ["=" + DS + "!" + exitCol + "35"], { fmt: F.num });
  er.row("Equity Value at Exit 退出股权价值", ["=B7+B8+B9"], { fmt: F.num, total: true });
  er.blank();
  er.row("Sponsor Equity Invested 投入股本", ["=" + A + "!B14"], { fmt: F.num });
  er.row("MoIC (x) 投资倍数", ["=B10/B12"], { fmt: F.mult, total: true });
  er.blank();
  er.section("现金流时间轴 Cash Flow Timeline");
  var tl = [null, "Year 0"]; for (var ti = 1; ti <= n; ti++) tl.push("Year " + ti);
  er.header(tl);
  var cfRowVals = [];
  for (var yi = 0; yi <= n; yi++) {
    if (yi === 0) cfRowVals.push("=-" + A + "!B14");
    else if (yi === n) cfRowVals.push("=B10");
    else cfRowVals.push("=0");
  }
  er.row("Sponsor Cash Flows 发起人现金流", cfRowVals, { fmt: F.num });
  er.row("IRR 内部收益率", ["=IRR(B17:" + ycol(n) + "17)"], { fmt: F.pct, total: true });
  er.note("退出年 = Assumptions 中的 Exit Year（预测年数）。时间轴：Year 0 投入 Sponsor Equity（负值），退出年收回股权价值（Equity Value at Exit），其间无分红则为 0。IRR 由 Excel IRR 函数对该行现金流计算；MoIC = 退出股权价值 / 投入股本。", 56);

  // ---------------- Sensitivity ----------------
  var entryMults = [p.entry_ev_ebitda - 1, p.entry_ev_ebitda - 0.5, p.entry_ev_ebitda, p.entry_ev_ebitda + 0.5, p.entry_ev_ebitda + 1];
  var exitMults = [p.exit_ev_ebitda - 1, p.exit_ev_ebitda - 0.5, p.exit_ev_ebitda, p.exit_ev_ebitda + 0.5, p.exit_ev_ebitda + 1];
  var s2 = xlSheet(wb, "Sensitivity", 6, { widths: [18, 13, 13, 13, 13, 13] });
  s2.title("Sensitivity 敏感性分析");
  s2.note("行=入场倍数，列=退出倍数；单元格=MoIC（退出股权价值/入场股权）。", 28);
  s2.blank();
  s2.header(["Entry \\ Exit"].concat(exitMults));
  for (var hc2 = 2; hc2 <= 6; hc2++) s2.ws.getCell(4, hc2).numFmt = F.mult;
  for (var ei = 0; ei < entryMults.length; ei++) {
    var rv = [];
    for (var ej = 0; ej < exitMults.length; ej++) {
      var entryEquity = "(" + A + "!$B$6*" + entryMults[ei] + "-" + A + "!$B$9+" + A + "!$B$10)";
      var exitEv = IS + "!" + exitCol + "7*" + exitMults[ej];
      var exitEquity = "(" + exitEv + "-" + DS + "!" + exitCol + "34+" + DS + "!" + exitCol + "35)";
      rv.push({ v: "=(" + exitEquity + ")/(" + entryEquity + ")", f: F.mult });
    }
    s2.row(entryMults[ei], rv, { fmt: F.mult, numKind: "input" });
  }
}


// ==================== 7. jsParseReport (upload_parser.py) ====================

var FIELD_KEYWORDS = {
  company_name: ["company", "公司名称", "公司", "entity name", "企业名称"],
  projection_years: ["projection years", "预测年数", "预测年限", "years"],
  revenue_y0: ["revenue y0", "base revenue", "基期收入", "year 0 revenue", "revenue (base)", "revenue-0", "initial revenue"],
  ltm_revenue: ["ltm revenue", "ltm 收入", "last twelve months revenue", "滚动12月收入"],
  rev_growth_y1: ["revenue growth y1", "收入增长率 y1", "revenue growth year 1", "rev growth y1", "增长 y1", "y1 growth"],
  rev_growth_y2: ["revenue growth y2", "收入增长率 y2", "revenue growth year 2", "rev growth y2", "增长 y2", "y2 growth"],
  rev_growth_y3: ["revenue growth y3", "收入增长率 y3", "revenue growth year 3", "rev growth y3", "增长 y3", "y3 growth"],
  rev_growth_y4: ["revenue growth y4", "收入增长率 y4", "revenue growth year 4", "rev growth y4", "增长 y4", "y4 growth"],
  rev_growth_y5: ["revenue growth y5", "收入增长率 y5", "revenue growth year 5", "rev growth y5", "增长 y5", "y5 growth"],
  cogs_pct: ["cogs", "营业成本", "cost of revenue", "cost of sales", "成本率", "主营业务成本"],
  sga_pct: ["sg&a", "销管费用", "销售管理费用", "selling general", "销售费用"],
  rd_pct: ["r&d", "研发费用", "research and development", "研发投入", "研发"],
  da_pct: ["d&a", "折旧摊销", "depreciation", "amortization", "折旧", "摊销"],
  capex_pct: ["capex", "资本支出", "capital expenditure", "资本性支出"],
  nwc_pct: ["nwc", "净营运资本", "networking capital", "营运资金"],
  ebitda_margin: ["ebitda margin", "ebitda 利润率", "ebitda率", "ebitda margin %"],
  ebitda_margin_y1: ["ebitda margin y1", "ebitda 利润率 y1", "ebitda率 y1"],
  ebitda_margin_y2: ["ebitda margin y2", "ebitda 利润率 y2", "ebitda率 y2"],
  ebitda_margin_y3: ["ebitda margin y3", "ebitda 利润率 y3", "ebitda率 y3"],
  ebitda_margin_y4: ["ebitda margin y4", "ebitda 利润率 y4", "ebitda率 y4"],
  ebitda_margin_y5: ["ebitda margin y5", "ebitda 利润率 y5", "ebitda率 y5"],
  ebitda: ["ebitda", "息税折旧摊销前利润"],
  ltm_ebitda: ["ltm ebitda", "ltm 息税折旧摊销前利润", "last twelve months ebitda"],
  interest_rate: ["interest rate", "债务利率", "借款利率", "贷款利率"],
  tax_rate: ["tax rate", "税率", "所得税率", "所得税", "effective tax"],
  dividend_pct: ["dividend", "股利", "分红", "payout", "分配率"],
  cash_interest_rate: ["cash interest", "现金存款利率", "存款利率"],
  dso: ["dso", "应收天数", "days sales outstanding", "days sales", "应收账款天数"],
  dio: ["dio", "库存天数", "days inventory", "days inventory outstanding", "存货天数"],
  dpo: ["dpo", "应付天数", "days payable", "days payable outstanding", "应付账款天数"],
  accrued_days: ["accrued days", "应计费用天数", "accrued"],
  beg_cash: ["beginning cash", "期初现金", "cash (beg)", "期初货币资金"],
  beg_ar: ["beginning ar", "期初应收", "accounts receivable (beg)", "期初应收账款"],
  beg_inventory: ["beginning inventory", "期初库存", "期初存货", "inventory (beg)"],
  beg_ppe: ["beginning ppe", "期初固定资产", "ppe (beg)", "期初pp&e"],
  beg_ap: ["beginning ap", "期初应付", "accounts payable (beg)", "期初应付账款"],
  beg_accrued: ["beginning accrued", "期初应计", "accrued (beg)", "期初应计费用"],
  beg_debt: ["beginning debt", "期初债务", "debt (beg)", "期初负债"],
  common_stock: ["common stock", "普通股", "股本"],
  beg_retn_earn: ["retained earnings", "留存收益", "re (beg)", "期初留存收益", "期初未分配利润"],
  new_debt_issuance: ["new debt", "新增债务", "新增借款", "debt issuance"],
  risk_free_rate: ["risk-free rate", "无风险利率", "risk free rate", "rf", "无风险"],
  equity_risk_premium: ["equity risk premium", "股权风险溢价", "erp", "风险溢价"],
  beta: ["beta", "贝塔", "β"],
  pre_tax_cost_of_debt: ["pre-tax cost of debt", "税前债务成本", "cost of debt", "债务成本"],
  debt_weight: ["debt weight", "债务权重", "target debt", "目标债务"],
  terminal_growth: ["terminal growth", "终值增长率", "gordon growth", "永续增长率"],
  exit_multiple: ["exit multiple", "退出倍数", "exit ev/ebitda", "退出 ev/ebitda"],
  tv_method: ["tv method", "终值方法", "terminal value method"],
  net_debt: ["net debt", "净债务", "总债务"],
  cash: ["cash", "现金", "货币资金"],
  minority_interest: ["minority interest", "少数股东权益", "少数股东"],
  preferred_stock: ["preferred stock", "优先股"],
  stock_based_comp: ["stock-based comp", "股权激励", "sbc", "股份支付", "股权支付"],
  pension_deficit: ["pension deficit", "养老金缺口", "养老金"],
  shares_outstanding: ["shares outstanding", "流通股本", "总股本", "股本数"],
  current_price: ["current price", "当前股价", "现价", "股价"],
  valuation_timing: ["valuation timing", "估值时点", "timing"],
  entry_ev_ebitda: ["entry ev/ebitda", "入场倍数", "入场 ev/ebitda", "entry multiple"],
  existing_net_debt: ["existing net debt", "现有净债务", "现有债务"],
  existing_cash: ["existing cash", "现有现金"],
  transaction_fees: ["transaction fees", "交易费用", "交易费"],
  financing_fees: ["financing fees", "融资费用", "融资费"],
  sponsor_equity: ["sponsor equity", "sponsor 股权", "基金股权", "股本投入"],
  revolver_capacity: ["revolver capacity", "revolver 额度", "循环额度"],
  term_loan_a: ["term loan a", "tla", "term loan a"],
  term_loan_b: ["term loan b", "tlb", "term loan b"],
  senior_notes: ["senior notes", "高级票据", "优先票据"],
  subordinated_debt: ["subordinated debt", "次级债务", "次级债"],
  revolver_rate: ["revolver rate", "revolver 利率"],
  tla_rate: ["tla rate", "term loan a 利率", "tla 利率"],
  tlb_rate: ["tlb rate", "term loan b 利率", "tlb 利率"],
  senior_notes_rate: ["senior notes rate", "senior notes 利率", "优先票据利率"],
  sub_rate: ["sub rate", "subordinated 利率", "次级债务利率"],
  tla_mandatory_amort: ["tla mandatory amort", "tla 强制摊销"],
  tlb_mandatory_amort: ["tlb mandatory amort", "tlb 强制摊销"],
  cash_sweep_pct: ["cash sweep", "现金清偿", "清偿率"],
  min_cash_balance: ["min cash", "最低现金", "最低现金余额"],
  exit_ev_ebitda: ["exit ev/ebitda", "退出倍数", "exit multiple"],
  exit_year: ["exit year", "退出年"],
};

var PERCENT_FIELDS = {
  cogs_pct: true, sga_pct: true, rd_pct: true, da_pct: true, capex_pct: true, nwc_pct: true,
  interest_rate: true, tax_rate: true, dividend_pct: true, cash_interest_rate: true,
  rev_growth_y1: true, rev_growth_y2: true, rev_growth_y3: true, rev_growth_y4: true, rev_growth_y5: true,
  ebitda_margin: true, ebitda_margin_y1: true, ebitda_margin_y2: true, ebitda_margin_y3: true,
  ebitda_margin_y4: true, ebitda_margin_y5: true,
  risk_free_rate: true, equity_risk_premium: true, pre_tax_cost_of_debt: true, debt_weight: true,
  terminal_growth: true, tla_mandatory_amort: true, tlb_mandatory_amort: true, cash_sweep_pct: true,
  revolver_rate: true, tla_rate: true, tlb_rate: true, senior_notes_rate: true, sub_rate: true,
};

function _parseNumber(value, isPercent) {
  var s = String(value).trim().replace(/,/g, "").replace(/，/g, "");
  s = s.replace(/[¥$€£]/g, "").trim();
  if (isPercent) {
    if (s.indexOf("%") >= 0) {
      s = s.replace(/%/g, "").trim();
      var v = parseFloat(s);
      if (isNaN(v)) return { value: 0.0, note: "解析失败" };
      return { value: v / 100, note: "百分比(含%)" };
    }
    var v = parseFloat(s);
    if (isNaN(v)) return { value: 0.0, note: "解析失败" };
    if (Math.abs(v) > 1) return { value: v / 100, note: "百分比(>1转换)" };
    return { value: v, note: "小数" };
  } else {
    var v = parseFloat(s);
    if (isNaN(v)) return { value: 0.0, note: "解析失败" };
    return { value: v, note: "数值" };
  }
}

function _matchField(fieldName, pairs) {
  var keywords = FIELD_KEYWORDS[fieldName] || [];
  if (!keywords.length) return { value: null, label: "", warnings: [] };
  var candidates = [];
  for (var i = 0; i < pairs.length; i++) {
    var label = pairs[i][0], value = pairs[i][1];
    var labelLower = label.toLowerCase();
    for (var k = 0; k < keywords.length; k++) {
      var kwLower = keywords[k].toLowerCase();
      if (labelLower.indexOf(kwLower) >= 0) {
        var score = labelLower.trim() === kwLower ? 100 : 50 + kwLower.length;
        candidates.push({ label: label, value: value, keyword: keywords[k], score: score });
        break;
      }
    }
  }
  if (!candidates.length) return { value: null, label: "", warnings: [] };
  candidates.sort(function(a, b) { return b.score - a.score; });
  var best = candidates[0];
  var warnings = [];
  if (candidates.length > 1) {
    var otherLabels = candidates.slice(1, 3).map(function(c) { return c.label; });
    warnings.push("字段 '" + fieldName + "' 找到 " + candidates.length + " 个候选，已选 '" + best.label + "'，其他: " + otherLabels.join(", "));
  }
  return { value: best.value, label: best.label, warnings: warnings };
}

function jsParseReport(file, modelType) {
  return new Promise(function(resolve, reject) {
    var ext = file.name.split(".").pop().toLowerCase();
    var reader = new FileReader();
    reader.onload = function(e) {
      try {
        var data = e.target.result;
        var pairs = [];
        if (ext === "csv") {
          var text;
          try { text = new TextDecoder("utf-8-sig").decode(data); }
          catch(err) { text = new TextDecoder("gbk").decode(data); }
          var lines = text.split(/\r?\n/);
          for (var li = 0; li < lines.length; li++) {
            var row = lines[li].split(",");
            if (!row.length || row.length < 2) continue;
            var label = row[0].trim();
            var value = null;
            for (var ci = 1; ci < row.length; ci++) {
              if (row[ci] && row[ci].trim()) { value = row[ci].trim(); break; }
            }
            if (label && value !== null) pairs.push([label, value]);
          }
        } else if (ext === "xlsx" || ext === "xls") {
          var wb = XLSX.read(data, { type: "array" });
          for (var si = 0; si < wb.SheetNames.length; si++) {
            var ws = wb.Sheets[wb.SheetNames[si]];
            var rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: null });
            for (var ri = 0; ri < rows.length; ri++) {
              var row = rows[ri];
              if (!row || row.length < 2) continue;
              var label = row[0];
              if (label === null) continue;
              label = String(label).trim();
              var value = null;
              for (var ci = 1; ci < row.length; ci++) {
                if (row[ci] !== null && String(row[ci]).trim()) { value = String(row[ci]).trim(); break; }
              }
              if (label && value !== null) pairs.push([label, value]);
            }
          }
        } else {
          reject(new Error("Unsupported file extension: " + ext));
          return;
        }
        if (!pairs.length) {
          reject(new Error("文件为空或无法解析，请检查格式（需要标签列 + 值列）"));
          return;
        }
        var schemaFields = jsDefaults[modelType];
        var matched = {}, unmatched = [], allWarnings = [];
        for (var fieldName in schemaFields) {
          var defaultVal = schemaFields[fieldName];
          var isText = typeof defaultVal === "string";
          var isInt = typeof defaultVal === "number" && defaultVal === Math.floor(defaultVal) && fieldName !== "beta" && fieldName !== "exit_multiple" && fieldName !== "exit_year" && fieldName !== "projection_years" ? false : (fieldName === "projection_years" || fieldName === "exit_year");
          var matchResult = _matchField(fieldName, pairs);
          allWarnings = allWarnings.concat(matchResult.warnings);
          if (matchResult.value === null) { unmatched.push(fieldName); continue; }
          if (isText) {
            matched[fieldName] = matchResult.value;
          } else if (isInt) {
            var iv = parseInt(String(matchResult.value).replace(/,/g, ""));
            if (isNaN(iv)) { unmatched.push(fieldName); allWarnings.push("字段 '" + fieldName + "' 值 '" + matchResult.value + "' 无法转整数"); }
            else matched[fieldName] = iv;
          } else {
            var isPercent = !!PERCENT_FIELDS[fieldName];
            var result = _parseNumber(matchResult.value, isPercent);
            if (result.note === "解析失败") { unmatched.push(fieldName); allWarnings.push("字段 '" + fieldName + "' 值 '" + matchResult.value + "' 解析数值失败"); }
            else matched[fieldName] = Math.round(result.value * 10000) / 10000;
          }
        }
        resolve({
          matched: matched, unmatched: unmatched, warnings: allWarnings,
          model_type: modelType, filename: file.name,
          total_fields: Object.keys(schemaFields).length, matched_count: Object.keys(matched).length,
        });
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = function() { reject(new Error("文件读取失败")); };
    if (ext === "csv") reader.readAsArrayBuffer(file);
    else reader.readAsArrayBuffer(file);
  });
}