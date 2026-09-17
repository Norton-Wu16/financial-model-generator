"""三个财务模型的参数 schema（Pydantic）。

每个模型对应一组参数，默认值为示例数据，用户可在前端修改后提交。
参数顺序和分组与前端表单一致。
"""
from pydantic import BaseModel, Field
from typing import List, Literal

from backend.custom_items import CustomItem


# ==================== 三表联动模型 ====================
class ThreeStatementParams(BaseModel):
    # 公司
    company_name: str = Field("示例公司 Sample Co.", description="公司名称 Company Name")
    projection_years: int = Field(5, ge=1, le=10, description="预测年数 Projection Years")
    # 收入与利润率
    revenue_y0: float = Field(100.0, description="基期收入 Year 0 Revenue")
    rev_growth_y1: float = Field(0.10, description="收入增长率 Y1 Revenue Growth Y1")
    rev_growth_y2: float = Field(0.08, description="收入增长率 Y2 Revenue Growth Y2")
    rev_growth_y3: float = Field(0.06, description="收入增长率 Y3 Revenue Growth Y3")
    rev_growth_y4: float = Field(0.05, description="收入增长率 Y4 Revenue Growth Y4")
    rev_growth_y5: float = Field(0.04, description="收入增长率 Y5 Revenue Growth Y5")
    cogs_pct: float = Field(0.60, description="COGS 占收入比 COGS % of Revenue")
    sga_pct: float = Field(0.15, description="SG&A 占收入比 SG&A % of Revenue")
    rd_pct: float = Field(0.05, description="R&D 占收入比 R&D % of Revenue")
    da_pct: float = Field(0.05, description="D&A 占收入比 D&A % of Revenue")
    interest_rate: float = Field(0.06, description="债务利率 Interest Rate on Debt")
    tax_rate: float = Field(0.25, description="税率 Tax Rate")
    dividend_pct: float = Field(0.30, description="股利分配率 Dividend Payout Ratio")
    capex_pct: float = Field(0.08, description="CapEx 占收入比 CapEx % of Revenue")
    # 营运资本天数
    dso: float = Field(45.0, description="应收天数 Days Sales Outstanding (DSO)")
    dio: float = Field(60.0, description="库存天数 Days Inventory Outstanding (DIO)")
    dpo: float = Field(30.0, description="应付天数 Days Payable Outstanding (DPO)")
    accrued_days: float = Field(20.0, description="应计费用天数 Accrued Expenses Days")
    # 期初资产负债表
    beg_cash: float = Field(20.0, description="期初现金 Beginning Cash")
    beg_ar: float = Field(12.0, description="期初应收 Beginning AR")
    beg_inventory: float = Field(16.0, description="期初库存 Beginning Inventory")
    beg_ppe: float = Field(80.0, description="期初 PP&E Beginning PP&E")
    beg_ap: float = Field(8.0, description="期初应付 Beginning AP")
    beg_accrued: float = Field(5.0, description="期初应计 Beginning Accrued Expenses")
    beg_debt: float = Field(50.0, description="期初债务 Beginning Debt")
    common_stock: float = Field(30.0, description="普通股 Common Stock")
    beg_retn_earn: float = Field(35.0, description="期初留存收益 Beginning Retained Earnings")
    new_debt_issuance: float = Field(0.0, description="每年新增债务发行 Annual New Debt Issuance")
    # 自定义行项
    custom_items: List[CustomItem] = Field(default_factory=list, description="自定义行项 Custom Line Items")


# ==================== DCF 模型 ====================
class DCFParams(BaseModel):
    company_name: str = Field("示例公司 Sample Co.", description="公司名称 Company Name")
    projection_years: int = Field(5, ge=1, le=10, description="预测年数 Projection Years")
    # 运营
    revenue_y0: float = Field(100.0, description="基期收入 Year 0 Revenue")
    rev_growth_y1: float = Field(0.15, description="收入增长率 Y1")
    rev_growth_y2: float = Field(0.12, description="收入增长率 Y2")
    rev_growth_y3: float = Field(0.10, description="收入增长率 Y3")
    rev_growth_y4: float = Field(0.08, description="收入增长率 Y4")
    rev_growth_y5: float = Field(0.06, description="收入增长率 Y5")
    ebitda_margin_y1: float = Field(0.25, description="EBITDA 利润率 Y1")
    ebitda_margin_y2: float = Field(0.25, description="EBITDA 利润率 Y2")
    ebitda_margin_y3: float = Field(0.25, description="EBITDA 利润率 Y3")
    ebitda_margin_y4: float = Field(0.25, description="EBITDA 利润率 Y4")
    ebitda_margin_y5: float = Field(0.25, description="EBITDA 利润率 Y5")
    da_pct: float = Field(0.05, description="D&A 占收入比 D&A % of Revenue")
    capex_pct: float = Field(0.06, description="CapEx 占收入比 CapEx % of Revenue")
    nwc_pct: float = Field(0.10, description="NWC 占收入比 NWC % of Revenue")
    tax_rate: float = Field(0.25, description="税率 Tax Rate")
    # WACC
    risk_free_rate: float = Field(0.04, description="无风险利率 Risk-free Rate (Rf)")
    equity_risk_premium: float = Field(0.05, description="股权风险溢价 Equity Risk Premium (ERP)")
    beta: float = Field(1.20, description="Beta")
    pre_tax_cost_of_debt: float = Field(0.06, description="税前债务成本 Pre-tax Cost of Debt")
    debt_weight: float = Field(0.30, description="目标债务权重 Target Debt Weight")
    # 终值
    terminal_growth: float = Field(0.03, description="终值增长率 Terminal Growth Rate (g)")
    exit_multiple: float = Field(10.0, description="退出 EV/EBITDA 倍数 Exit Multiple")
    tv_method: Literal["gordon", "exit"] = Field("gordon", description="终值方法 Terminal Value Method")
    # EV -> Equity 桥
    net_debt: float = Field(20.0, description="净债务 Net Debt")
    cash: float = Field(30.0, description="现金 Cash")
    minority_interest: float = Field(0.0, description="少数股东权益 Minority Interest")
    preferred_stock: float = Field(0.0, description="优先股 Preferred Stock")
    stock_based_comp: float = Field(5.0, description="股权激励 Stock-Based Compensation")
    pension_deficit: float = Field(0.0, description="养老金缺口 Pension Deficit")
    shares_outstanding: float = Field(50.0, description="流通股本（百万股）Shares Outstanding (M)")
    current_price: float = Field(20.0, description="当前股价 Current Share Price")
    valuation_timing: Literal["mid", "end"] = Field("mid", description="估值时点 Valuation Timing")
    # 自定义行项
    custom_items: List[CustomItem] = Field(default_factory=list, description="自定义行项 Custom Line Items")


# ==================== LBO 模型 ====================
class LBOParams(BaseModel):
    company_name: str = Field("示例公司 Sample Co.", description="公司名称 Company Name")
    # 交易假设
    ltm_revenue: float = Field(250.0, description="LTM 收入 LTM Revenue")
    ltm_ebitda: float = Field(50.0, description="LTM EBITDA")
    entry_ev_ebitda: float = Field(10.0, description="入场 EV/EBITDA 倍数 Entry Multiple")
    existing_net_debt: float = Field(100.0, description="现有净债务 Existing Net Debt (to refinance)")
    existing_cash: float = Field(5.0, description="现有现金 Existing Cash")
    transaction_fees: float = Field(10.0, description="交易费用 Transaction Fees")
    financing_fees: float = Field(15.0, description="融资费用 Financing Fees")
    # 资本结构
    sponsor_equity: float = Field(150.0, description=" Sponsor 股权 Sponsor Equity")
    revolver_capacity: float = Field(50.0, description="Revolver 额度 Revolver Capacity")
    term_loan_a: float = Field(100.0, description="Term Loan A 金额")
    term_loan_b: float = Field(150.0, description="Term Loan B 金额")
    senior_notes: float = Field(75.0, description="Senior Notes 金额")
    subordinated_debt: float = Field(25.0, description="Subordinated Debt 金额")
    # 利率
    revolver_rate: float = Field(0.04, description="Revolver 利率")
    tla_rate: float = Field(0.04, description="Term Loan A 利率")
    tlb_rate: float = Field(0.05, description="Term Loan B 利率")
    senior_notes_rate: float = Field(0.07, description="Senior Notes 利率")
    sub_rate: float = Field(0.10, description="Subordinated Debt 利率")
    # 偿债
    tla_mandatory_amort: float = Field(0.05, description="TLA 强制摊销 %")
    tlb_mandatory_amort: float = Field(0.02, description="TLB 强制摊销 %")
    cash_sweep_pct: float = Field(0.50, description="现金清偿 % Cash Sweep %")
    min_cash_balance: float = Field(10.0, description="最低现金余额 Minimum Cash Balance")
    # 运营
    rev_growth_y1: float = Field(0.08, description="收入增长率 Y1")
    rev_growth_y2: float = Field(0.07, description="收入增长率 Y2")
    rev_growth_y3: float = Field(0.06, description="收入增长率 Y3")
    rev_growth_y4: float = Field(0.05, description="收入增长率 Y4")
    rev_growth_y5: float = Field(0.04, description="收入增长率 Y5")
    ebitda_margin: float = Field(0.20, description="EBITDA 利润率")
    da_pct: float = Field(0.04, description="D&A 占收入比")
    capex_pct: float = Field(0.05, description="CapEx 占收入比")
    nwc_pct: float = Field(0.08, description="ΔNWC 占收入比")
    tax_rate: float = Field(0.25, description="税率 Tax Rate")
    # 退出
    exit_ev_ebitda: float = Field(10.0, description="退出 EV/EBITDA 倍数 Exit Multiple")
    exit_year: int = Field(5, ge=1, le=10, description="退出年 Exit Year")
    cash_interest_rate: float = Field(0.02, description="现金存款利率 Cash Interest Rate")
    # 自定义行项
    custom_items: List[CustomItem] = Field(default_factory=list, description="自定义行项 Custom Line Items")


class GenerateRequest(BaseModel):
    model_type: Literal["three_statement", "dcf", "lbo"]
    params: dict
