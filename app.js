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
  return preset.params[modelType];
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
  var growths = [p.rev_growth_y1, p.rev_growth_y2, p.rev_growth_y3, p.rev_growth_y4, p.rev_growth_y5];
  for (var i = 1; i <= n; i++) { var g = i <= growths.length ? growths[i-1] : 0; revenue[i] = revenue[i-1] * (1+g); }
  var cogs = revenue.map(function(r) { return r * p.cogs_pct; });
  var grossProfit = revenue.map(function(r,j) { return r - cogs[j]; });
  var sga = revenue.map(function(r) { return r * p.sga_pct; });
  var rd = revenue.map(function(r) { return r * p.rd_pct; });
  var ebitda = grossProfit.map(function(gp,j) { return gp - sga[j] - rd[j]; });
  var daIs = revenue.map(function(r) { return r * p.da_pct; });
  var ebit = ebitda.map(function(e,j) { return e - daIs[j]; });

  // Supporting Schedules
  var ppeBeg = new Array(n+1).fill(0); ppeBeg[0] = p.beg_ppe;
  var capex = revenue.map(function(r) { return r * p.capex_pct; });
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
    model_type: "three_statement", company_name: p.company_name, years: years,
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
  var growths = [p.rev_growth_y1, p.rev_growth_y2, p.rev_growth_y3, p.rev_growth_y4, p.rev_growth_y5];
  for (var i = 1; i <= n; i++) { var g = i <= growths.length ? growths[i-1] : 0; revenue[i] = revenue[i-1] * (1+g); }

  var margins = [p.ebitda_margin_y1, p.ebitda_margin_y2, p.ebitda_margin_y3, p.ebitda_margin_y4, p.ebitda_margin_y5];
  var ebitda = new Array(n+1).fill(0); ebitda[0] = revenue[0] * margins[0];
  for (var i = 1; i <= n; i++) { var m = i <= margins.length ? margins[i-1] : margins[margins.length-1]; ebitda[i] = revenue[i] * m; }

  var da = revenue.map(function(r) { return r * p.da_pct; });
  var ebit = ebitda.map(function(e,j) { return e - da[j]; });
  var nopat = ebit.map(function(e) { return e * (1 - p.tax_rate); });

  var deltaNwc = new Array(n+1).fill(0);
  for (var i = 1; i <= n; i++) deltaNwc[i] = (revenue[i] - revenue[i-1]) * p.nwc_pct;
  var capex = revenue.map(function(r) { return r * p.capex_pct; });
  var ufcf = new Array(n+1).fill(0);
  for (var i = 1; i <= n; i++) ufcf[i] = nopat[i] + da[i] - deltaNwc[i] - capex[i];

  var costOfEquity = p.risk_free_rate + p.equity_risk_premium * p.beta;
  var afterTaxKd = p.pre_tax_cost_of_debt * (1 - p.tax_rate);
  var equityWeight = 1 - p.debt_weight;
  var wacc = costOfEquity * equityWeight + afterTaxKd * p.debt_weight;

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
    model_type: "dcf", company_name: p.company_name, years: years,
    operating_model: {
      revenue: revenue.map(_round2), ebitda: ebitda.map(_round2), da: da.map(_round2),
      ebit: ebit.map(_round2), nopat: nopat.map(_round2), delta_nwc: deltaNwc.map(_round2),
      capex: capex.map(_round2), ufcf: ufcf.map(_round2),
    },
    wacc: {
      cost_of_equity: _round2(costOfEquity), after_tax_cost_of_debt: _round2(afterTaxKd),
      equity_weight: _round2(equityWeight), debt_weight: _round2(p.debt_weight), wacc: _round2(wacc),
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
  var n = 5;
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
  var growths = [p.rev_growth_y1, p.rev_growth_y2, p.rev_growth_y3, p.rev_growth_y4, p.rev_growth_y5];
  for (var i = 1; i <= n; i++) { var g = i <= growths.length ? growths[i-1] : 0; revenue[i] = revenue[i-1] * (1+g); }
  var ebitda = new Array(n+1).fill(0); ebitda[0] = p.ltm_ebitda;
  for (var i = 1; i <= n; i++) ebitda[i] = revenue[i] * p.ebitda_margin;
  var da = revenue.map(function(r) { return r * p.da_pct; });
  var ebit = ebitda.map(function(e,j) { return e - da[j]; });
  var capex = revenue.map(function(r) { return r * p.capex_pct; });
  var deltaNwc = new Array(n+1).fill(0);
  for (var i = 1; i <= n; i++) deltaNwc[i] = -revenue[i] * p.nwc_pct;

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
    model_type: "lbo", company_name: p.company_name, years: years,
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
  if (modelType === "three_statement") return calculateThreeStatement(params);
  if (modelType === "dcf") return calculateDcf(params);
  if (modelType === "lbo") return calculateLbo(params);
  throw new Error("Unknown model type: " + modelType);
}


// ==================== 6. jsBuildExcel (builders/*.py) ====================

function jsBuildExcel(modelType, params) {
  var wb = XLSX.utils.book_new();
  if (modelType === "three_statement") buildThreeStatementExcel(wb, params);
  else if (modelType === "dcf") buildDcfExcel(wb, params);
  else if (modelType === "lbo") buildLboExcel(wb, params);
  else throw new Error("Unknown model type: " + modelType);

  var wbout = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  return new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

// --- 三表联动模型 Excel ---
function buildThreeStatementExcel(wb, p) {
  var n = p.projection_years;
  var A = "Assumptions", IS = "'Income Statement'", SCH = "'Supporting Schedules'", BS = "'Balance Sheet'", CFS = "'Cash Flow Statement'";

  // Assumptions sheet
  var aAoa = [];
  aAoa[0] = ["Assumptions 假设参数"];
  aAoa[2] = ["参数 Parameter", "Year 0"]; for (var i = 1; i <= n; i++) aAoa[2].push("Year " + i);
  aAoa[3] = ["运营假设 Operating Assumptions"];
  aAoa[4] = ["Revenue (base) 基期收入", p.revenue_y0];
  aAoa[5] = ["Revenue Growth % 收入增长率", ""]; for (var i = 0; i < n; i++) aAoa[5][2+i] = [p.rev_growth_y1, p.rev_growth_y2, p.rev_growth_y3, p.rev_growth_y4, p.rev_growth_y5][i];
  aAoa[6] = ["COGS % of Revenue 成本率", p.cogs_pct];
  aAoa[7] = ["SG&A % of Revenue 销管费用率", p.sga_pct];
  aAoa[8] = ["R&D % of Revenue 研发费用率", p.rd_pct];
  aAoa[9] = ["D&A % of Revenue 折旧摊销率", p.da_pct];
  aAoa[10] = ["Interest Rate on Debt 债务利率", p.interest_rate];
  aAoa[11] = ["Tax Rate 税率", p.tax_rate];
  aAoa[12] = ["Dividend Payout % 股利分配率", p.dividend_pct];
  aAoa[13] = ["CapEx % of Revenue 资本支出率", p.capex_pct];
  aAoa[15] = ["营运资本天数 Working Capital Days"];
  aAoa[16] = ["DSO 应收天数", p.dso];
  aAoa[17] = ["DIO 库存天数", p.dio];
  aAoa[18] = ["DPO 应付天数", p.dpo];
  aAoa[19] = ["Accrued Days 应计天数", p.accrued_days];
  aAoa[21] = ["期初资产负债表 Beginning Balance Sheet"];
  aAoa[22] = ["Beginning Cash 期初现金", p.beg_cash];
  aAoa[23] = ["Beginning AR 期初应收", p.beg_ar];
  aAoa[24] = ["Beginning Inventory 期初库存", p.beg_inventory];
  aAoa[25] = ["Beginning PP&E 期初固定资产", p.beg_ppe];
  aAoa[26] = ["Beginning AP 期初应付", p.beg_ap];
  aAoa[27] = ["Beginning Accrued 期初应计", p.beg_accrued];
  aAoa[28] = ["Beginning Debt 期初债务", p.beg_debt];
  aAoa[29] = ["Common Stock 普通股", p.common_stock];
  aAoa[30] = ["Beginning Retained Earnings 期初留存收益", p.beg_retn_earn];
  aAoa[32] = ["其他 Other"];
  aAoa[33] = ["New Debt Issuance (annual) 新增债务"]; for (var i = 0; i < n; i++) { if (!aAoa[33]) aAoa[33] = []; aAoa[33][2+i] = p.new_debt_issuance; }
  XLSX.utils.book_append_sheet(wb, aoaToSheet(aAoa), "Assumptions");

  // Income Statement
  var isAoa = [];
  isAoa[0] = ["Income Statement 利润表"];
  isAoa[2] = ["项目 Item", "Year 0"]; for (var i = 1; i <= n; i++) isAoa[2].push("Year " + i);
  function isRow(r, fn) { isAoa[r-1] = []; isAoa[r-1][0] = fn(0, "label"); for (var yi = 0; yi <= n; yi++) isAoa[r-1][1+yi] = fn(yi, "val"); }
  isRow(5, function(yi, typ) { if (typ === "label") return "Revenue 营业收入"; if (yi === 0) return "=" + A + "!B5"; var prev = ycol(yi-1); return "=" + prev + "5*(1+" + A + "!" + ycol(yi) + "6)"; });
  isRow(6, function(yi, typ) { if (typ === "label") return "Revenue Growth % 增长率"; if (yi === 0) return null; return "=" + ycol(yi) + "5/" + ycol(yi-1) + "5-1"; });
  isRow(7, function(yi, typ) { if (typ === "label") return "COGS 营业成本"; return "=" + ycol(yi) + "5*" + A + "!$B$7"; });
  isRow(8, function(yi, typ) { if (typ === "label") return "Gross Profit 毛利"; return "=" + ycol(yi) + "5-" + ycol(yi) + "7"; });
  isRow(9, function(yi, typ) { if (typ === "label") return "SG&A 销管费用"; return "=" + ycol(yi) + "5*" + A + "!$B$8"; });
  isRow(10, function(yi, typ) { if (typ === "label") return "R&D 研发费用"; return "=" + ycol(yi) + "5*" + A + "!$B$9"; });
  isRow(11, function(yi, typ) { if (typ === "label") return "EBITDA"; return "=" + ycol(yi) + "8-" + ycol(yi) + "9-" + ycol(yi) + "10"; });
  isRow(12, function(yi, typ) { if (typ === "label") return "D&A 折旧摊销"; if (yi === 0) return "=" + ycol(yi) + "5*" + A + "!$B$10"; return "=" + SCH + "!" + ycol(yi) + "6"; });
  isRow(13, function(yi, typ) { if (typ === "label") return "EBIT 营业利润"; return "=" + ycol(yi) + "11-" + ycol(yi) + "12"; });
  isRow(14, function(yi, typ) { if (typ === "label") return "Interest Expense 利息费用"; if (yi === 0) return "=" + A + "!$B$11*" + A + "!B29"; return "=" + SCH + "!" + ycol(yi) + "17"; });
  isRow(15, function(yi, typ) { if (typ === "label") return "EBT 税前利润"; return "=" + ycol(yi) + "13-" + ycol(yi) + "14"; });
  isRow(16, function(yi, typ) { if (typ === "label") return "Taxes 所得税"; return "=MAX(0," + ycol(yi) + "15*" + A + "!$B$12)"; });
  isRow(17, function(yi, typ) { if (typ === "label") return "Net Income 净利润"; return "=" + ycol(yi) + "15-" + ycol(yi) + "16"; });
  isRow(18, function(yi, typ) { if (typ === "label") return "Dividends 股利"; return "=" + ycol(yi) + "17*" + A + "!$B$13"; });

  // Custom items in IS
  if (p.custom_items && p.custom_items.length) {
    isAoa[19] = ["自定义行项 Custom Line Items"];
    function buildIsCellMap(col) {
      return {
        revenue: col+"5", cogs: col+"7", gross_profit: col+"8", sga: col+"9", rd: col+"10",
        ebitda: col+"11", da: col+"12", ebit: col+"13", interest: col+"14", ebt: col+"15",
        taxes: col+"16", net_income: col+"17", dividends: col+"18",
        cash: BS+"!"+col+"5", ar: BS+"!"+col+"6", inventory: BS+"!"+col+"7", ppe: BS+"!"+col+"8",
        total_assets: BS+"!"+col+"9", ap: BS+"!"+col+"13", accrued: BS+"!"+col+"14", debt: BS+"!"+col+"15",
        total_liabilities: BS+"!"+col+"16", common_stock: BS+"!"+col+"20", retained_earnings: BS+"!"+col+"21",
        total_equity: BS+"!"+col+"22", total_le: BS+"!"+col+"25", capex: SCH+"!"+col+"8", new_debt: SCH+"!"+col+"19",
      };
    }
    for (var idx = 0; idx < p.custom_items.length; idx++) {
      var ci = p.custom_items[idx]; var r = 21 + idx;
      isAoa[r-1] = [ci.name];
      for (var yi = 0; yi <= n; yi++) {
        var col = ycol(yi); var cm = buildIsCellMap(col);
        try { isAoa[r-1][1+yi] = expressionToExcel(ci.formula, cm); } catch(e) { isAoa[r-1][1+yi] = '="#ERR: ' + String(e.message || e).substring(0,30) + '"'; }
      }
    }
  }
  XLSX.utils.book_append_sheet(wb, aoaToSheet(isAoa), "Income Statement");

  // Supporting Schedules
  var schAoa = [];
  schAoa[0] = ["Supporting Schedules 辅助表"];
  schAoa[2] = ["项目 Item", "Year 0"]; for (var i = 1; i <= n; i++) schAoa[2].push("Year " + i);
  schAoa[3] = ["PP&E 滚动 PP&E Roll-forward"];
  function schRow(r, fn) { schAoa[r-1] = []; schAoa[r-1][0] = fn(0, "label"); for (var yi = 0; yi <= n; yi++) schAoa[r-1][1+yi] = fn(yi, "val"); }
  schRow(6, function(yi, typ) { if (typ === "label") return "D&A 折旧摊销"; return "=" + IS + "!" + ycol(yi) + "5*" + A + "!$B$10"; });
  schRow(7, function(yi, typ) { if (typ === "label") return "Beginning PP&E 期初固定资产"; if (yi === 0) return "=" + A + "!B26"; return "=" + ycol(yi-1) + "10"; });
  schRow(8, function(yi, typ) { if (typ === "label") return "CapEx 资本支出"; return "=" + IS + "!" + ycol(yi) + "5*" + A + "!$B$14"; });
  schRow(9, function(yi, typ) { if (typ === "label") return "Less: D&A 减:折旧"; return "=-" + ycol(yi) + "6"; });
  schRow(10, function(yi, typ) { if (typ === "label") return "Ending PP&E 期末固定资产"; if (yi === 0) return "=" + ycol(yi) + "7"; return "=" + ycol(yi) + "7+" + ycol(yi) + "8+" + ycol(yi) + "9"; });
  schAoa[14] = ["债务滚动 Debt Roll-forward"];
  schRow(17, function(yi, typ) { if (typ === "label") return "Interest Expense 利息费用"; if (yi === 0) return "=" + A + "!$B$11*" + A + "!B29"; return "=" + A + "!$B$11*" + BS + "!" + ycol(yi-1) + "15"; });
  schRow(18, function(yi, typ) { if (typ === "label") return "Beginning Debt 期初债务"; if (yi === 0) return "=" + A + "!B29"; return "=" + ycol(yi-1) + "20"; });
  schRow(19, function(yi, typ) { if (typ === "label") return "New Debt Issuance 新增债务"; if (yi === 0) return "=0"; return "=" + A + "!" + ycol(yi) + "34"; });
  schRow(20, function(yi, typ) { if (typ === "label") return "Ending Debt 期末债务"; if (yi === 0) return "=" + ycol(yi) + "18"; return "=" + ycol(yi) + "18+" + ycol(yi) + "19"; });
  XLSX.utils.book_append_sheet(wb, aoaToSheet(schAoa), "Supporting Schedules");

  // Balance Sheet
  var bsAoa = [];
  bsAoa[0] = ["Balance Sheet 资产负债表"];
  bsAoa[2] = ["项目 Item", "Year 0"]; for (var i = 1; i <= n; i++) bsAoa[2].push("Year " + i);
  bsAoa[3] = ["资产 Assets"];
  function bsRow(r, fn) { bsAoa[r-1] = []; bsAoa[r-1][0] = fn(0, "label"); for (var yi = 0; yi <= n; yi++) bsAoa[r-1][1+yi] = fn(yi, "val"); }
  bsRow(5, function(yi, typ) { if (typ === "label") return "Cash 现金"; if (yi === 0) return "=" + A + "!B23"; return "=" + ycol(yi-1) + "5+" + CFS + "!" + ycol(yi) + "23"; });
  bsRow(6, function(yi, typ) { if (typ === "label") return "Accounts Receivable 应收账款"; if (yi === 0) return "=" + A + "!B24"; return "=" + IS + "!" + ycol(yi) + "5/365*" + A + "!$B$17"; });
  bsRow(7, function(yi, typ) { if (typ === "label") return "Inventory 存货"; if (yi === 0) return "=" + A + "!B25"; return "=" + IS + "!" + ycol(yi) + "7/365*" + A + "!$B$18"; });
  bsRow(8, function(yi, typ) { if (typ === "label") return "PP&E, net 固定资产净额"; if (yi === 0) return "=" + A + "!B26"; return "=" + SCH + "!" + ycol(yi) + "10"; });
  bsRow(9, function(yi, typ) { if (typ === "label") return "Total Assets 总资产"; return "=SUM(" + ycol(yi) + "5:" + ycol(yi) + "8)"; });
  bsAoa[11] = ["负债 Liabilities"];
  bsRow(13, function(yi, typ) { if (typ === "label") return "Accounts Payable 应付账款"; if (yi === 0) return "=" + A + "!B27"; return "=" + IS + "!" + ycol(yi) + "7/365*" + A + "!$B$19"; });
  bsRow(14, function(yi, typ) { if (typ === "label") return "Accrued Expenses 应计费用"; if (yi === 0) return "=" + A + "!B28"; return "=" + IS + "!" + ycol(yi) + "5/365*" + A + "!$B$20"; });
  bsRow(15, function(yi, typ) { if (typ === "label") return "Debt 债务"; if (yi === 0) return "=" + A + "!B29"; return "=" + SCH + "!" + ycol(yi) + "20"; });
  bsRow(16, function(yi, typ) { if (typ === "label") return "Total Liabilities 总负债"; return "=SUM(" + ycol(yi) + "13:" + ycol(yi) + "15)"; });
  bsAoa[18] = ["权益 Equity"];
  bsRow(20, function(yi, typ) { if (typ === "label") return "Common Stock 普通股"; return "=" + A + "!$B$30"; });
  bsRow(21, function(yi, typ) { if (typ === "label") return "Retained Earnings 留存收益"; if (yi === 0) return "=" + A + "!B31"; return "=" + ycol(yi-1) + "21+" + IS + "!" + ycol(yi) + "17-" + IS + "!" + ycol(yi) + "18"; });
  bsRow(22, function(yi, typ) { if (typ === "label") return "Total Equity 总权益"; return "=" + ycol(yi) + "20+" + ycol(yi) + "21"; });
  bsRow(25, function(yi, typ) { if (typ === "label") return "Total Liabilities + Equity 总负债权益"; return "=" + ycol(yi) + "16+" + ycol(yi) + "22"; });
  bsAoa[26] = ["平衡校验 Balance Check"];
  for (var yi = 0; yi <= n; yi++) { var col = ycol(yi); bsAoa[26][1+yi] = '=IF(ROUND(' + col + '9-' + col + '25,2)=0,"BALANCED","OUT OF BALANCE")'; }
  XLSX.utils.book_append_sheet(wb, aoaToSheet(bsAoa), "Balance Sheet");

  // Cash Flow Statement
  var cfsAoa = [];
  cfsAoa[0] = ["Cash Flow Statement 现金流量表"];
  cfsAoa[2] = ["项目 Item", "Year 0"]; for (var i = 1; i <= n; i++) cfsAoa[2].push("Year " + i);
  cfsAoa[3] = ["经营活动 Operating Activities"];
  function cfsRow(r, fn) { cfsAoa[r-1] = []; cfsAoa[r-1][0] = fn(0, "label"); for (var yi = 0; yi <= n; yi++) cfsAoa[r-1][1+yi] = fn(yi, "val"); }
  cfsRow(6, function(yi, typ) { if (typ === "label") return "Net Income 净利润"; if (yi === 0) return null; return "=" + IS + "!" + ycol(yi) + "17"; });
  cfsRow(7, function(yi, typ) { if (typ === "label") return "D&A 折旧摊销"; if (yi === 0) return null; return "=" + SCH + "!" + ycol(yi) + "6"; });
  cfsRow(8, function(yi, typ) { if (typ === "label") return "Change in AR 应收变动"; if (yi === 0) return null; return "=-(" + BS + "!" + ycol(yi) + "6-" + BS + "!" + ycol(yi-1) + "6)"; });
  cfsRow(9, function(yi, typ) { if (typ === "label") return "Change in Inventory 存货变动"; if (yi === 0) return null; return "=-(" + BS + "!" + ycol(yi) + "7-" + BS + "!" + ycol(yi-1) + "7)"; });
  cfsRow(10, function(yi, typ) { if (typ === "label") return "Change in AP 应付变动"; if (yi === 0) return null; return "=" + BS + "!" + ycol(yi) + "13-" + BS + "!" + ycol(yi-1) + "13"; });
  cfsRow(11, function(yi, typ) { if (typ === "label") return "Change in Accrued 应计变动"; if (yi === 0) return null; return "=" + BS + "!" + ycol(yi) + "14-" + BS + "!" + ycol(yi-1) + "14"; });
  cfsRow(12, function(yi, typ) { if (typ === "label") return "Cash from Operations 经营现金流"; if (yi === 0) return null; return "=SUM(" + ycol(yi) + "6:" + ycol(yi) + "11)"; });
  cfsAoa[13] = ["投资活动 Investing Activities"];
  cfsRow(15, function(yi, typ) { if (typ === "label") return "CapEx 资本支出"; if (yi === 0) return null; return "=-" + IS + "!" + ycol(yi) + "5*" + A + "!$B$14"; });
  cfsRow(16, function(yi, typ) { if (typ === "label") return "Cash from Investing 投资现金流"; if (yi === 0) return null; return "=" + ycol(yi) + "15"; });
  cfsAoa[17] = ["筹资活动 Financing Activities"];
  cfsRow(19, function(yi, typ) { if (typ === "label") return "Net Debt Change 债务净变动"; if (yi === 0) return null; return "=" + BS + "!" + ycol(yi) + "15-" + BS + "!" + ycol(yi-1) + "15"; });
  cfsRow(20, function(yi, typ) { if (typ === "label") return "Dividends Paid 支付股利"; if (yi === 0) return null; return "=-" + IS + "!" + ycol(yi) + "18"; });
  cfsRow(21, function(yi, typ) { if (typ === "label") return "Cash from Financing 筹资现金流"; if (yi === 0) return null; return "=SUM(" + ycol(yi) + "19:" + ycol(yi) + "20)"; });
  cfsRow(23, function(yi, typ) { if (typ === "label") return "Net Change in Cash 现金净变动"; if (yi === 0) return null; return "=" + ycol(yi) + "12+" + ycol(yi) + "16+" + ycol(yi) + "21"; });
  XLSX.utils.book_append_sheet(wb, aoaToSheet(cfsAoa), "Cash Flow Statement");

  // Dashboard
  var dashAoa = [];
  dashAoa[0] = ["Dashboard 关键指标摘要"];
  dashAoa[2] = ["指标 Metric", "Year 0"]; for (var i = 1; i <= n; i++) dashAoa[2].push("Year " + i);
  function dashRow(r, fn) { dashAoa[r-1] = []; dashAoa[r-1][0] = fn(0, "label"); for (var yi = 0; yi <= n; yi++) dashAoa[r-1][1+yi] = fn(yi, "val"); }
  dashRow(5, function(yi, typ) { if (typ === "label") return "Revenue 营业收入"; return "=" + IS + "!" + ycol(yi) + "5"; });
  dashRow(6, function(yi, typ) { if (typ === "label") return "EBITDA"; return "=" + IS + "!" + ycol(yi) + "11"; });
  dashRow(7, function(yi, typ) { if (typ === "label") return "EBIT 营业利润"; return "=" + IS + "!" + ycol(yi) + "13"; });
  dashRow(8, function(yi, typ) { if (typ === "label") return "Net Income 净利润"; return "=" + IS + "!" + ycol(yi) + "17"; });
  dashRow(9, function(yi, typ) { if (typ === "label") return "Total Assets 总资产"; return "=" + BS + "!" + ycol(yi) + "9"; });
  dashRow(10, function(yi, typ) { if (typ === "label") return "Total Debt 总债务"; return "=" + BS + "!" + ycol(yi) + "15"; });
  dashRow(11, function(yi, typ) { if (typ === "label") return "Cash 现金"; return "=" + BS + "!" + ycol(yi) + "5"; });
  dashRow(12, function(yi, typ) { if (typ === "label") return "Retained Earnings 留存收益"; return "=" + BS + "!" + ycol(yi) + "21"; });
  dashRow(14, function(yi, typ) { if (typ === "label") return "EBITDA Margin %"; if (yi === 0) return null; return "=" + IS + "!" + ycol(yi) + "11/" + IS + "!" + ycol(yi) + "5"; });
  dashRow(15, function(yi, typ) { if (typ === "label") return "Net Margin %"; if (yi === 0) return null; return "=" + IS + "!" + ycol(yi) + "17/" + IS + "!" + ycol(yi) + "5"; });
  XLSX.utils.book_append_sheet(wb, aoaToSheet(dashAoa), "Dashboard");
}

// --- DCF 模型 Excel ---
function buildDcfExcel(wb, p) {
  var n = p.projection_years;
  var A = "Assumptions", OP = "'Operating Model'", W = "WACC", TV = "'Terminal Value'", DV = "'DCF Valuation'";

  // Assumptions
  var aAoa = [];
  aAoa[0] = ["Assumptions 假设参数"];
  aAoa[2] = ["参数 Parameter", "Year 0"]; for (var i = 1; i <= n; i++) aAoa[2].push("Year " + i);
  aAoa[3] = ["运营假设 Operating Assumptions"];
  aAoa[4] = ["Revenue (base) 基期收入", p.revenue_y0];
  aAoa[5] = ["Revenue Growth % 收入增长率", ""]; for (var i = 0; i < n; i++) aAoa[5][2+i] = [p.rev_growth_y1, p.rev_growth_y2, p.rev_growth_y3, p.rev_growth_y4, p.rev_growth_y5][i];
  aAoa[6] = ["EBITDA Margin %"]; for (var i = 0; i < n; i++) aAoa[6][2+i] = [p.ebitda_margin_y1, p.ebitda_margin_y2, p.ebitda_margin_y3, p.ebitda_margin_y4, p.ebitda_margin_y5][i];
  aAoa[7] = ["D&A % of Revenue 折旧摊销率", p.da_pct];
  aAoa[8] = ["CapEx % of Revenue 资本支出率", p.capex_pct];
  aAoa[9] = ["NWC % of Revenue 净营运资本率", p.nwc_pct];
  aAoa[10] = ["Tax Rate 税率", p.tax_rate];
  aAoa[12] = ["WACC 资本成本"];
  aAoa[13] = ["Risk-free Rate 无风险利率", p.risk_free_rate];
  aAoa[14] = ["Equity Risk Premium 股权风险溢价", p.equity_risk_premium];
  aAoa[15] = ["Beta", p.beta];
  aAoa[16] = ["Pre-tax Cost of Debt 税前债务成本", p.pre_tax_cost_of_debt];
  aAoa[17] = ["Target Debt Weight 目标债务权重", p.debt_weight];
  aAoa[19] = ["终值 Terminal Value"];
  aAoa[20] = ["Terminal Growth Rate g 终值增长率", p.terminal_growth];
  aAoa[21] = ["Exit EV/EBITDA 退出倍数", p.exit_multiple];
  aAoa[22] = ["TV Method 终值方法", p.tv_method];
  aAoa[24] = ["EV→Equity 桥"];
  aAoa[25] = ["Total Debt 总债务", p.net_debt];
  aAoa[26] = ["Cash 现金", p.cash];
  aAoa[27] = ["Minority Interest 少数股东权益", p.minority_interest];
  aAoa[28] = ["Preferred Stock 优先股", p.preferred_stock];
  aAoa[29] = ["Stock-Based Compensation 股权激励", p.stock_based_comp];
  aAoa[30] = ["Pension Deficit 养老金缺口", p.pension_deficit];
  aAoa[31] = ["Shares Outstanding (M) 流通股本", p.shares_outstanding];
  aAoa[32] = ["Current Share Price 当前股价", p.current_price];
  aAoa[33] = ["Valuation Timing 估值时点 (mid,end)", p.valuation_timing];
  XLSX.utils.book_append_sheet(wb, aoaToSheet(aAoa), "Assumptions");

  // Operating Model
  var opAoa = [];
  opAoa[0] = ["Operating Model 运营模型"];
  opAoa[2] = ["项目 Item", "Year 0"]; for (var i = 1; i <= n; i++) opAoa[2].push("Year " + i);
  function opRow(r, fn) { opAoa[r-1] = []; opAoa[r-1][0] = fn(0, "label"); for (var yi = 0; yi <= n; yi++) opAoa[r-1][1+yi] = fn(yi, "val"); }
  opRow(5, function(yi, typ) { if (typ === "label") return "Revenue 营业收入"; if (yi === 0) return "=" + A + "!B5"; return "=" + ycol(yi-1) + "5*(1+" + A + "!" + ycol(yi) + "6)"; });
  opRow(6, function(yi, typ) { if (typ === "label") return "Revenue Growth % 增长率"; if (yi === 0) return null; return "=" + ycol(yi) + "5/" + ycol(yi-1) + "5-1"; });
  opRow(7, function(yi, typ) { if (typ === "label") return "EBITDA"; if (yi > 0) return "=" + ycol(yi) + "5*" + A + "!" + ycol(yi) + "7"; return "=" + ycol(yi) + "5*" + A + "!C7"; });
  opRow(8, function(yi, typ) { if (typ === "label") return "EBITDA Margin %"; if (yi === 0) return null; return "=" + ycol(yi) + "7/" + ycol(yi) + "5"; });
  opRow(9, function(yi, typ) { if (typ === "label") return "D&A 折旧摊销"; return "=" + ycol(yi) + "5*" + A + "!$B$8"; });
  opRow(10, function(yi, typ) { if (typ === "label") return "EBIT 营业利润"; return "=" + ycol(yi) + "7-" + ycol(yi) + "9"; });
  opRow(11, function(yi, typ) { if (typ === "label") return "EBIT Margin %"; if (yi === 0) return null; return "=" + ycol(yi) + "10/" + ycol(yi) + "5"; });
  opRow(12, function(yi, typ) { if (typ === "label") return "NOPAT 税后营业利润"; return "=" + ycol(yi) + "10*(1-" + A + "!$B$11)"; });
  opRow(13, function(yi, typ) { if (typ === "label") return "Δ NWC 净营运资本变动"; if (yi === 0) return null; return "=(" + ycol(yi) + "5-" + ycol(yi-1) + "5)*" + A + "!$B$10"; });
  opRow(14, function(yi, typ) { if (typ === "label") return "CapEx 资本支出"; return "=" + ycol(yi) + "5*" + A + "!$B$9"; });
  opRow(15, function(yi, typ) { if (typ === "label") return "Unlevered FCF 无杠杆自由现金流"; if (yi === 0) return null; return "=" + ycol(yi) + "12+" + ycol(yi) + "9-" + ycol(yi) + "13-" + ycol(yi) + "14"; });

  if (p.custom_items && p.custom_items.length) {
    opAoa[16] = ["自定义行项 Custom Line Items"];
    function buildOpCellMap(col) {
      return {
        revenue: col+"5", ebitda: col+"7", da: col+"9", ebit: col+"10", nopat: col+"12",
        delta_nwc: col+"13", capex: col+"14", ufcf: col+"15", wacc: W+"!$B$14", cost_of_equity: W+"!$B$8",
        after_tax_kd: W+"!$B$11", tv_gordon: TV+"!$B$5", tv_exit: TV+"!$B$6", selected_tv: TV+"!$B$7",
        pv_tv: TV+"!$B$9", sum_pv_fcf: DV+"!$B$10", enterprise_value: DV+"!$B$12", equity_value: DV+"!$B$22",
        implied_price: DV+"!$B$24", current_price: A+"!$B$33", shares_outstanding: A+"!$B$32",
        net_debt: A+"!$B$26", cash: A+"!$B$27",
      };
    }
    for (var idx = 0; idx < p.custom_items.length; idx++) {
      var ci = p.custom_items[idx]; var r = 18 + idx;
      opAoa[r-1] = [ci.name];
      for (var yi = 0; yi <= n; yi++) {
        var col = ycol(yi); var cm = buildOpCellMap(col);
        try { opAoa[r-1][1+yi] = expressionToExcel(ci.formula, cm); } catch(e) { opAoa[r-1][1+yi] = '="#ERR: ' + String(e.message || e).substring(0,30) + '"'; }
      }
    }
  }
  XLSX.utils.book_append_sheet(wb, aoaToSheet(opAoa), "Operating Model");

  // WACC
  var wAoa = [];
  wAoa[0] = ["WACC 加权平均资本成本"];
  wAoa[2] = ["项目 Item", "Value"];
  wAoa[4] = ["Risk-free Rate (Rf) 无风险利率", "=" + A + "!B14"];
  wAoa[5] = ["Equity Risk Premium (ERP) 股权风险溢价", "=" + A + "!B15"];
  wAoa[6] = ["Beta", "=" + A + "!B16"];
  wAoa[7] = ["Cost of Equity (CAPM) 股权成本", "=B5+B6*B7"];
  wAoa[8] = ["Pre-tax Cost of Debt 税前债务成本", "=" + A + "!B17"];
  wAoa[9] = ["Tax Rate 税率", "=" + A + "!B11"];
  wAoa[10] = ["After-tax Cost of Debt 税后债务成本", "=B9*(1-B10)"];
  wAoa[11] = ["Equity Weight 股权重", "=1-" + A + "!B18"];
  wAoa[12] = ["Debt Weight 债权重", "=" + A + "!B18"];
  wAoa[13] = ["WACC 加权平均资本成本", "=B8*B12+B11*B13"];
  XLSX.utils.book_append_sheet(wb, aoaToSheet(wAoa), "WACC");

  // Terminal Value
  var tvAoa = [];
  var lastCol = ycol(n);
  tvAoa[0] = ["Terminal Value 终值"];
  tvAoa[2] = ["项目 Item", "Value"];
  tvAoa[4] = ["Gordon Growth TV 戈登增长终值", "=" + OP + "!" + lastCol + "15*(1+" + A + "!$B$21)/(" + W + "!$B$14-" + A + "!$B$21)"];
  tvAoa[5] = ["Exit Multiple TV 退出倍数终值", "=" + OP + "!" + lastCol + "7*" + A + "!$B$22"];
  tvAoa[6] = ["Selected TV 选定终值", '=IF(' + A + '!$B$23="gordon",B5,B6)'];
  tvAoa[7] = ["Discount Period (Years) 折现期数", "=" + n];
  tvAoa[8] = ["PV of Terminal Value 终值现值", "=B7/(1+" + W + "!$B$14)^B8"];
  XLSX.utils.book_append_sheet(wb, aoaToSheet(tvAoa), "Terminal Value");

  // DCF Valuation
  var dvAoa = [];
  dvAoa[0] = ["DCF Valuation 估值汇总"];
  dvAoa[2] = ["项目 Item"]; for (var i = 1; i <= n; i++) dvAoa[2].push("Year " + i);
  var isMid = p.valuation_timing === "mid";
  function dvRow(r, fn) { dvAoa[r-1] = []; dvAoa[r-1][0] = fn(0, "label"); for (var yi = 0; yi < n; yi++) dvAoa[r-1][1+yi] = fn(yi+1, "val"); }
  dvRow(5, function(yi, typ) { if (typ === "label") return "Unlevered FCF 自由现金流"; return "=" + OP + "!" + ycol(yi) + "15"; });
  dvRow(6, function(yi, typ) { if (typ === "label") return "Discount Period 折现期数"; var period = isMid ? yi - 0.5 : yi; return "=" + period; });
  dvRow(7, function(yi, typ) { if (typ === "label") return "Discount Factor 折现因子"; return "=1/(1+" + W + "!$B$14)^" + ycol(yi) + "6"; });
  dvRow(8, function(yi, typ) { if (typ === "label") return "PV of FCF FCF现值"; return "=" + ycol(yi) + "5*" + ycol(yi) + "7"; });
  dvAoa[9] = ["Sum of PV of FCF FCF现值合计", "=SUM(" + ycol(1) + "8:" + ycol(n) + "8)"];
  dvAoa[10] = ["PV of Terminal Value 终值现值", "=" + TV + "!B9"];
  dvAoa[11] = ["Enterprise Value 企业价值", "=B10+B11"];
  dvAoa[13] = ["EV → Equity 桥"];
  dvAoa[14] = ["Enterprise Value 企业价值", "=B12"];
  dvAoa[15] = ["Less: Total Debt 减:总债务", "=-" + A + "!B26"];
  dvAoa[16] = ["Plus: Cash 加:现金", "=" + A + "!B27"];
  dvAoa[17] = ["Less: Minority Interest 减:少数股东权益", "=-" + A + "!B28"];
  dvAoa[18] = ["Less: Preferred Stock 减:优先股", "=-" + A + "!B29"];
  dvAoa[19] = ["Less: Stock-Based Comp 减:股权激励", "=-" + A + "!B30"];
  dvAoa[20] = ["Less: Pension Deficit 减:养老金缺口", "=-" + A + "!B31"];
  dvAoa[21] = ["Equity Value 股权价值", "=SUM(B15:B21)"];
  dvAoa[22] = ["Shares Outstanding (M) 流通股本", "=" + A + "!B32"];
  dvAoa[23] = ["Implied Share Price 隐含股价", "=B22/B23"];
  dvAoa[24] = ["Current Share Price 当前股价", "=" + A + "!B33"];
  dvAoa[25] = ["Upside/(Downside) % 涨跌幅", "=B24/B25-1"];
  XLSX.utils.book_append_sheet(wb, aoaToSheet(dvAoa), "DCF Valuation");

  // Sensitivity
  var sAoa = [];
  sAoa[0] = ["Sensitivity 敏感性分析"];
  sAoa[1] = ["行=终值增长率 g，列=WACC；单元格=隐含股价"];
  var gVals = [p.terminal_growth-0.01, p.terminal_growth-0.005, p.terminal_growth, p.terminal_growth+0.005, p.terminal_growth+0.01];
  var wVals = [0.08, 0.09, 0.10, 0.11, 0.12];
  sAoa[3] = ["g \\ WACC"];
  for (var j = 0; j < wVals.length; j++) sAoa[3][1+j] = wVals[j];
  for (var i = 0; i < gVals.length; i++) {
    sAoa[4+i] = [gVals[i]];
    for (var j = 0; j < wVals.length; j++) {
      var wc = ycol(j) + "$4", gc = "$A" + (5+i);
      var pvTerms = [];
      for (var t = 1; t <= n; t++) pvTerms.push(OP + "!" + ycol(t) + "15/(1+" + wc + ")^" + t);
      var tv = OP + "!" + lastCol + "15*(1+" + gc + ")/(" + wc + "-" + gc + ")";
      var pvTv = tv + "/(1+" + wc + ")^" + n;
      var ev = "(" + pvTerms.join("+") + ")+" + pvTv;
      var equity = "(" + ev + ")-" + A + "!$B$26+" + A + "!$B$27-" + A + "!$B$28-" + A + "!$B$29-" + A + "!$B$30-" + A + "!$B$31";
      sAoa[4+i][1+j] = "=((" + equity + ")/" + A + "!$B$32)";
    }
  }
  XLSX.utils.book_append_sheet(wb, aoaToSheet(sAoa), "Sensitivity");
}

// --- LBO 模型 Excel ---
function buildLboExcel(wb, p) {
  var n = 5;
  var A = "Assumptions", SU = "'Sources & Uses'", DS = "'Debt Schedule'", IS = "'Income Statement'", CF = "'Cash Flow (CFADS)'";

  // Assumptions
  var aAoa = [];
  aAoa[0] = ["Assumptions 假设参数"];
  aAoa[2] = ["参数 Parameter", "Value"]; for (var i = 1; i <= n; i++) aAoa[2].push("Year " + i);
  aAoa[3] = ["交易假设 Transaction Assumptions"];
  aAoa[4] = ["LTM Revenue", p.ltm_revenue];
  aAoa[5] = ["LTM EBITDA", p.ltm_ebitda];
  aAoa[6] = ["Entry EV/EBITDA", p.entry_ev_ebitda];
  aAoa[7] = ["Entry EV 入场企业价值", "=B6*B7"];
  aAoa[8] = ["Existing Net Debt 现有净债务", p.existing_net_debt];
  aAoa[9] = ["Existing Cash 现有现金", p.existing_cash];
  aAoa[10] = ["Transaction Fees 交易费用", p.transaction_fees];
  aAoa[11] = ["Financing Fees 融资费用", p.financing_fees];
  aAoa[13] = ["资本结构 Capital Structure"];
  aAoa[13] = ["Sponsor Equity", p.sponsor_equity];
  aAoa[14] = ["Revolver Capacity", p.revolver_capacity];
  aAoa[15] = ["Term Loan A", p.term_loan_a];
  aAoa[16] = ["Term Loan B", p.term_loan_b];
  aAoa[17] = ["Senior Notes", p.senior_notes];
  aAoa[18] = ["Subordinated Debt", p.subordinated_debt];
  aAoa[20] = ["利率 Interest Rates"];
  aAoa[20] = ["Revolver Rate", p.revolver_rate];
  aAoa[21] = ["Term Loan A Rate", p.tla_rate];
  aAoa[22] = ["Term Loan B Rate", p.tlb_rate];
  aAoa[23] = ["Senior Notes Rate", p.senior_notes_rate];
  aAoa[24] = ["Subordinated Debt Rate", p.sub_rate];
  aAoa[26] = ["偿债 Debt Repayment"];
  aAoa[26] = ["TLA Mandatory Amort %", p.tla_mandatory_amort];
  aAoa[27] = ["TLB Mandatory Amort %", p.tlb_mandatory_amort];
  aAoa[28] = ["Cash Sweep %", p.cash_sweep_pct];
  aAoa[29] = ["Min Cash Balance", p.min_cash_balance];
  aAoa[31] = ["运营假设 Operating Assumptions"];
  aAoa[31] = ["Revenue Growth %"]; for (var i = 0; i < n; i++) aAoa[31][2+i] = [p.rev_growth_y1, p.rev_growth_y2, p.rev_growth_y3, p.rev_growth_y4, p.rev_growth_y5][i];
  aAoa[32] = ["EBITDA Margin %"]; for (var i = 0; i < n; i++) aAoa[32][2+i] = p.ebitda_margin;
  aAoa[33] = ["D&A % of Revenue", p.da_pct];
  aAoa[34] = ["CapEx % of Revenue", p.capex_pct];
  aAoa[35] = ["ΔNWC % of Revenue", p.nwc_pct];
  aAoa[36] = ["Tax Rate", p.tax_rate];
  aAoa[38] = ["退出 Exit"];
  aAoa[38] = ["Exit EV/EBITDA", p.exit_ev_ebitda];
  aAoa[39] = ["Exit Year", p.exit_year];
  aAoa[40] = ["Cash Interest Rate", p.cash_interest_rate];
  XLSX.utils.book_append_sheet(wb, aoaToSheet(aAoa), "Assumptions");

  // Sources & Uses
  var suAoa = [];
  suAoa[0] = ["Sources & Uses 资金来源与用途"];
  suAoa[2] = ["项目 Item", "Amount"];
  suAoa[3] = ["Uses 用途"];
  suAoa[5] = ["Purchase of Equity 购买股权", "=" + A + "!B8-" + A + "!B9"];
  suAoa[6] = ["Refinance Existing Debt 再融资债务", "=" + A + "!B9+" + A + "!B10"];
  suAoa[7] = ["Transaction Fees 交易费用", "=" + A + "!B11"];
  suAoa[8] = ["Financing Fees 融资费用", "=" + A + "!B12"];
  suAoa[9] = ["Total Uses 用途合计", "=SUM(B6:B9)"];
  suAoa[11] = ["Sources 来源"];
  suAoa[12] = ["Sponsor Equity", "=" + A + "!B14"];
  suAoa[13] = ["Revolver Draw (plug) 循环额度提取", "=B10-B13-B15-B16-B17-B18-B19"];
  suAoa[14] = ["Term Loan A", "=" + A + "!B16"];
  suAoa[15] = ["Term Loan B", "=" + A + "!B17"];
  suAoa[16] = ["Senior Notes", "=" + A + "!B18"];
  suAoa[17] = ["Subordinated Debt", "=" + A + "!B19"];
  suAoa[18] = ["Existing Cash 现有现金", "=" + A + "!B10"];
  suAoa[19] = ["Total Sources 来源合计", "=SUM(B13:B19)"];
  suAoa[21] = ["平衡校验 Balance Check", '=IF(ROUND(B20-B10,2)=0,"BALANCED","IMBALANCED")'];
  XLSX.utils.book_append_sheet(wb, aoaToSheet(suAoa), "Sources & Uses");

  // Debt Schedule
  var dsAoa = [];
  dsAoa[0] = ["Debt Schedule 债务滚动表"];
  dsAoa[2] = ["项目 Item", "Close"]; for (var i = 1; i <= n; i++) dsAoa[2].push("Year " + i);
  dsAoa[3] = ["现金与偿债可用 Cash & CFADS"];
  function dsRow(r, fn) { dsAoa[r-1] = []; dsAoa[r-1][0] = fn(0, "label"); for (var yi = 0; yi <= n; yi++) dsAoa[r-1][1+yi] = fn(yi, "val"); }
  dsRow(5, function(yi, typ) { if (typ === "label") return "Beginning Cash 期初现金"; if (yi === 0) return "=" + A + "!B10"; return "=" + ycol(yi-1) + "35"; });
  dsRow(6, function(yi, typ) { if (typ === "label") return "CFADS 偿债可用现金流"; if (yi === 0) return "=0"; return "=" + CF + "!" + ycol(yi) + "13"; });
  dsRow(7, function(yi, typ) { if (typ === "label") return "Min Cash 最低现金"; return "=" + A + "!$B$30"; });
  dsRow(8, function(yi, typ) { if (typ === "label") return "Cash Available 可用现金"; return "=MAX(0," + ycol(yi) + "5+" + ycol(yi) + "6-" + ycol(yi) + "7)"; });
  dsAoa[9] = ["Revolver 循环额度"];
  dsRow(11, function(yi, typ) { if (typ === "label") return "Beginning 期初"; if (yi === 0) return "=" + SU + "!B14"; return "=" + ycol(yi-1) + "13"; });
  dsRow(12, function(yi, typ) { if (typ === "label") return "Repayment 偿还"; return "=MIN(" + ycol(yi) + "11," + ycol(yi) + "8)"; });
  dsRow(13, function(yi, typ) { if (typ === "label") return "Ending 期末"; return "=" + ycol(yi) + "11-" + ycol(yi) + "12"; });
  dsAoa[14] = ["Term Loan A"];
  dsRow(16, function(yi, typ) { if (typ === "label") return "Beginning 期初"; if (yi === 0) return "=" + A + "!B16"; return "=" + ycol(yi-1) + "18"; });
  dsRow(17, function(yi, typ) { if (typ === "label") return "Mandatory Amort 强制摊销"; return "=MIN(" + ycol(yi) + "16," + A + "!$B$16*" + A + "!$B$27)"; });
  dsRow(18, function(yi, typ) { if (typ === "label") return "Ending 期末"; return "=" + ycol(yi) + "16-" + ycol(yi) + "17"; });
  dsAoa[19] = ["Term Loan B"];
  dsRow(21, function(yi, typ) { if (typ === "label") return "Beginning 期初"; if (yi === 0) return "=" + A + "!B17"; return "=" + ycol(yi-1) + "24"; });
  dsRow(22, function(yi, typ) { if (typ === "label") return "Mandatory Amort 强制摊销"; return "=MIN(" + ycol(yi) + "21," + A + "!$B$17*" + A + "!$B$28)"; });
  dsRow(23, function(yi, typ) { if (typ === "label") return "Cash Sweep 现金清偿"; return "=MIN(" + ycol(yi) + "21-" + ycol(yi) + "22,MAX(0," + ycol(yi) + "8-" + ycol(yi) + "12-" + ycol(yi) + "17-" + ycol(yi) + "22)*" + A + "!$B$29)"; });
  dsRow(24, function(yi, typ) { if (typ === "label") return "Ending 期末"; return "=" + ycol(yi) + "21-" + ycol(yi) + "22-" + ycol(yi) + "23"; });
  dsAoa[25] = ["Senior Notes 次优先票据"];
  dsRow(27, function(yi, typ) { if (typ === "label") return "Beginning 期初"; if (yi === 0) return "=" + A + "!B18"; return "=" + ycol(yi-1) + "28"; });
  dsRow(28, function(yi, typ) { if (typ === "label") return "Ending 期末"; return "=" + ycol(yi) + "27"; });
  dsAoa[29] = ["Subordinated Debt 次级债务"];
  dsRow(31, function(yi, typ) { if (typ === "label") return "Beginning 期初"; if (yi === 0) return "=" + A + "!B19"; return "=" + ycol(yi-1) + "32"; });
  dsRow(32, function(yi, typ) { if (typ === "label") return "Ending 期末"; return "=" + ycol(yi) + "31"; });
  dsRow(34, function(yi, typ) { if (typ === "label") return "Total Debt 总债务"; return "=" + ycol(yi) + "13+" + ycol(yi) + "18+" + ycol(yi) + "24+" + ycol(yi) + "28+" + ycol(yi) + "32"; });
  dsRow(35, function(yi, typ) { if (typ === "label") return "Cash Ending 期末现金"; return "=MAX(" + ycol(yi) + "7," + ycol(yi) + "5+" + ycol(yi) + "6-" + ycol(yi) + "12-" + ycol(yi) + "17-" + ycol(yi) + "22-" + ycol(yi) + "23)"; });
  XLSX.utils.book_append_sheet(wb, aoaToSheet(dsAoa), "Debt Schedule");

  // Income Statement
  var isAoa = [];
  isAoa[0] = ["Income Statement 利润表"];
  isAoa[2] = ["项目 Item", "LTM"]; for (var i = 1; i <= n; i++) isAoa[2].push("Year " + i);
  function isRow(r, fn) { isAoa[r-1] = []; isAoa[r-1][0] = fn(0, "label"); for (var yi = 0; yi <= n; yi++) isAoa[r-1][1+yi] = fn(yi, "val"); }
  isRow(5, function(yi, typ) { if (typ === "label") return "Revenue 营业收入"; if (yi === 0) return "=" + A + "!B5"; return "=" + ycol(yi-1) + "5*(1+" + A + "!" + ycol(yi) + "32)"; });
  isRow(6, function(yi, typ) { if (typ === "label") return "Revenue Growth %"; if (yi === 0) return null; return "=" + ycol(yi) + "5/" + ycol(yi-1) + "5-1"; });
  isRow(7, function(yi, typ) { if (typ === "label") return "EBITDA"; if (yi === 0) return "=" + A + "!B6"; return "=" + ycol(yi) + "5*" + A + "!" + ycol(yi) + "33"; });
  isRow(8, function(yi, typ) { if (typ === "label") return "EBITDA Margin %"; if (yi === 0) return null; return "=" + ycol(yi) + "7/" + ycol(yi) + "5"; });
  isRow(9, function(yi, typ) { if (typ === "label") return "D&A 折旧摊销"; return "=" + ycol(yi) + "5*" + A + "!$B$34"; });
  isRow(10, function(yi, typ) { if (typ === "label") return "EBIT 营业利润"; return "=" + ycol(yi) + "7-" + ycol(yi) + "9"; });
  isRow(12, function(yi, typ) { if (typ === "label") return "Interest - Revolver 利息(循环)"; return "=" + A + "!$B$21*" + DS + "!" + ycol(yi) + "11"; });
  isRow(13, function(yi, typ) { if (typ === "label") return "Interest - TL A 利息(TLA)"; return "=" + A + "!$B$22*" + DS + "!" + ycol(yi) + "16"; });
  isRow(14, function(yi, typ) { if (typ === "label") return "Interest - TL B 利息(TLB)"; return "=" + A + "!$B$23*" + DS + "!" + ycol(yi) + "21"; });
  isRow(15, function(yi, typ) { if (typ === "label") return "Interest - Sr Notes 利息(优先票据)"; return "=" + A + "!$B$24*" + DS + "!" + ycol(yi) + "27"; });
  isRow(16, function(yi, typ) { if (typ === "label") return "Interest - Sub 利息(次级)"; return "=" + A + "!$B$25*" + DS + "!" + ycol(yi) + "31"; });
  isRow(17, function(yi, typ) { if (typ === "label") return "Total Interest 利息合计"; return "=SUM(" + ycol(yi) + "12:" + ycol(yi) + "16)"; });
  isRow(18, function(yi, typ) { if (typ === "label") return "Interest Income 利息收入"; return "=" + A + "!$B$41*" + DS + "!" + ycol(yi) + "5"; });
  isRow(19, function(yi, typ) { if (typ === "label") return "EBT 税前利润"; return "=" + ycol(yi) + "10-" + ycol(yi) + "17+" + ycol(yi) + "18"; });
  isRow(20, function(yi, typ) { if (typ === "label") return "Taxes 所得税"; return "=MAX(0," + ycol(yi) + "19*" + A + "!$B$37)"; });
  isRow(21, function(yi, typ) { if (typ === "label") return "Net Income 净利润"; return "=" + ycol(yi) + "19-" + ycol(yi) + "20"; });

  if (p.custom_items && p.custom_items.length) {
    isAoa[22] = ["自定义行项 Custom Line Items"];
    function buildIsCellMap(col) {
      return {
        revenue: col+"5", ebitda: col+"7", da: col+"9", ebit: col+"10",
        interest_revolver: col+"12", interest_tla: col+"13", interest_tlb: col+"14",
        interest_sn: col+"15", interest_sub: col+"16", total_interest: col+"17",
        interest_income: col+"18", ebt: col+"19", taxes: col+"20", net_income: col+"21",
        total_debt: DS+"!"+col+"34", cash: DS+"!"+col+"35", revolver: DS+"!"+col+"13",
        tla: DS+"!"+col+"18", tlb: DS+"!"+col+"24", senior_notes: DS+"!"+col+"28",
        subordinated_debt: DS+"!"+col+"32", cfads: CF+"!"+col+"13",
      };
    }
    for (var idx = 0; idx < p.custom_items.length; idx++) {
      var ci = p.custom_items[idx]; var r = 24 + idx;
      isAoa[r-1] = [ci.name];
      for (var yi = 0; yi <= n; yi++) {
        var col = ycol(yi); var cm = buildIsCellMap(col);
        try { isAoa[r-1][1+yi] = expressionToExcel(ci.formula, cm); } catch(e) { isAoa[r-1][1+yi] = '="#ERR: ' + String(e.message || e).substring(0,30) + '"'; }
      }
    }
  }
  XLSX.utils.book_append_sheet(wb, aoaToSheet(isAoa), "Income Statement");

  // Cash Flow (CFADS)
  var cfAoa = [];
  cfAoa[0] = ["Cash Flow (CFADS) 偿债可用现金流"];
  cfAoa[2] = ["项目 Item", "LTM"]; for (var i = 1; i <= n; i++) cfAoa[2].push("Year " + i);
  function cfRow(r, fn) { cfAoa[r-1] = []; cfAoa[r-1][0] = fn(0, "label"); for (var yi = 0; yi <= n; yi++) cfAoa[r-1][1+yi] = fn(yi, "val"); }
  cfRow(5, function(yi, typ) { if (typ === "label") return "Net Income 净利润"; if (yi === 0) return null; return "=" + IS + "!" + ycol(yi) + "21"; });
  cfRow(7, function(yi, typ) { if (typ === "label") return "D&A 折旧摊销"; if (yi === 0) return null; return "=" + IS + "!" + ycol(yi) + "9"; });
  cfRow(8, function(yi, typ) { if (typ === "label") return "Δ NWC 净营运资本变动"; if (yi === 0) return null; return "=-" + IS + "!" + ycol(yi) + "5*" + A + "!$B$36"; });
  cfRow(9, function(yi, typ) { if (typ === "label") return "Cash from Operations 经营现金流"; if (yi === 0) return null; return "=" + ycol(yi) + "5+" + ycol(yi) + "7+" + ycol(yi) + "8"; });
  cfRow(11, function(yi, typ) { if (typ === "label") return "CapEx 资本支出"; if (yi === 0) return null; return "=-" + IS + "!" + ycol(yi) + "5*" + A + "!$B$35"; });
  cfRow(13, function(yi, typ) { if (typ === "label") return "CFADS 偿债可用现金流"; if (yi === 0) return null; return "=" + ycol(yi) + "9+" + ycol(yi) + "11"; });
  XLSX.utils.book_append_sheet(wb, aoaToSheet(cfAoa), "Cash Flow (CFADS)");

  // Exit & Returns
  var exitCol = ycol(p.exit_year);
  var erAoa = [];
  erAoa[0] = ["Exit & Returns 退出与回报"];
  erAoa[2] = ["项目 Item", "Value"];
  erAoa[4] = ["Exit Year (" + p.exit_year + ") EBITDA 退出年EBITDA", "=" + IS + "!" + exitCol + "7"];
  erAoa[5] = ["Exit EV/EBITDA 退出倍数", "=" + A + "!B39"];
  erAoa[6] = ["Exit Enterprise Value 退出企业价值", "=B5*B6"];
  erAoa[7] = ["Less: Total Debt 减:总债务", "=-" + DS + "!" + exitCol + "34"];
  erAoa[8] = ["Plus: Cash 加:现金", "=" + DS + "!" + exitCol + "35"];
  erAoa[9] = ["Equity Value at Exit 退出股权价值", "=B7+B8+B9"];
  erAoa[11] = ["Sponsor Equity Invested 投入股本", "=" + A + "!B14"];
  erAoa[12] = ["MoIC (x) 投资倍数", "=B10/B12"];
  erAoa[14] = ["现金流时间轴 Cash Flow Timeline"];
  for (var yi = 0; yi < 6; yi++) erAoa[14][1+yi] = yi > 0 ? "Year " + yi : "Year 0";
  erAoa[15] = ["Sponsor Cash Flows"];
  for (var yi = 0; yi < 6; yi++) {
    if (yi === 0) erAoa[15][1+yi] = "=-" + A + "!B14";
    else if (yi === p.exit_year) erAoa[15][1+yi] = "=B10";
    else if (yi < p.exit_year) erAoa[15][1+yi] = "=0";
  }
  erAoa[17] = ["IRR 内部收益率", "=IRR(B16:" + exitCol + "16)"];
  XLSX.utils.book_append_sheet(wb, aoaToSheet(erAoa), "Exit & Returns");

  // Sensitivity
  var entryMults = [p.entry_ev_ebitda-1, p.entry_ev_ebitda-0.5, p.entry_ev_ebitda, p.entry_ev_ebitda+0.5, p.entry_ev_ebitda+1];
  var exitMults = [p.exit_ev_ebitda-1, p.exit_ev_ebitda-0.5, p.exit_ev_ebitda, p.exit_ev_ebitda+0.5, p.exit_ev_ebitda+1];
  var lboSAoa = [];
  lboSAoa[0] = ["Sensitivity 敏感性分析"];
  lboSAoa[1] = ["行=入场倍数，列=退出倍数；单元格=MoIC"];
  lboSAoa[3] = ["Entry \\ Exit"];
  for (var j = 0; j < exitMults.length; j++) lboSAoa[3][1+j] = exitMults[j];
  for (var i = 0; i < entryMults.length; i++) {
    lboSAoa[4+i] = [entryMults[i]];
    for (var j = 0; j < exitMults.length; j++) {
      var entryEquity = "(" + A + "!$B$6*" + entryMults[i] + "-" + A + "!$B$9+" + A + "!$B$10)";
      var exitEv = IS + "!" + exitCol + "7*" + exitMults[j];
      var exitEquity = "(" + exitEv + "-" + DS + "!" + exitCol + "34+" + DS + "!" + exitCol + "35)";
      lboSAoa[4+i][1+j] = "=(" + exitEquity + ")/(" + entryEquity + ")";
    }
  }
  XLSX.utils.book_append_sheet(wb, aoaToSheet(lboSAoa), "Sensitivity");
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