import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
font = FontProperties(fname='/System/Library/Fonts/STHeiti Light.ttc')
plt.rcParams['font.sans-serif'] = ['Heiti TC', 'Apple LiGothic', 'Arial Unicode MS']  
plt.rcParams['axes.unicode_minus'] = False
from datetime import datetime
import io

# --- 財務數據儲存類別 (No changes needed) ---
class FinancialData:
    def __init__(self):
        # 初始化所有財務數據，提供預設值
        self.data = {
            'operating_revenue': 0.0,
            'cost_of_goods_sold': 0.0,
            'operating_expenses': 0.0,
            'net_profit_after_tax': 0.0,
            'shareholders_equity': 0.0,
            'total_assets': 0.0,
            'current_assets': 0.0,
            'current_liabilities': 0.0,
            'inventory': 0.0,
            'accounts_receivable': 0.0,
            'interest_expense': 0.0,
            'net_profit_before_tax': 0.0,
            'operating_cash_flow': 0.0,
            'investing_cash_flow': 0.0,
            'financing_cash_flow': 0.0,
            'capital_expenditures': 0.0,
            'cash_dividends_paid': 0.0,
            'non_recurring_gain_loss': 0.0,
            'total_profit': 0.0, # 通常指稅前利潤或淨利潤，用於非經常性損益佔比
            'cash_and_equivalents': 0.0,
            'short_term_borrowing': 0.0,
            'accounts_payable_days': 0.0, # 假設已知或手動輸入
            'prev_year_net_profit_after_tax': 0.0,
            'prev_year_operating_revenue': 0.0,
            'prev_year_inventory': 0.0, # 去年存貨
            'prev_year_accounts_receivable': 0.0, # 去年應收帳款
            'prev_year_inventory_turnover_rate': 0.0,
            'prev_year_accounts_receivable_turnover_days': 0.0,
            'prev_year_gross_profit_margin': 0.0, # 去年毛利率
            'industry_avg_roe': 0.0,
            'industry_avg_revenue_growth_rate': 0.0,
            'cost_of_debt_interest_rate': 0.0, # 負債利率，例如 0.03 代表 3%
            'prev_total_liabilities': 0.0,
            'prev_total_assets': 0.0,
            'prev_net_debt': 0.0,
            'three_year_operating_cash_flows': [0.0, 0.0, 0.0] # 近三年營業現金流
        }

    def update_data(self, key, value):
        if key in self.data:
            self.data[key] = value
        else:
            st.warning(f"Warning: Key '{key}' not found in financial data.")

    def get_data(self, key):
        return self.data.get(key, 0.0)

# --- 財務計算與評估類別 (No changes needed) ---
class FinancialCalculator:
    def __init__(self, financial_data):
        self.financial_data = financial_data

    def get_value(self, key, default=0.0):
        """Helper to safely get data values, converting to float."""
        value = self.financial_data.get_data(key)
        if value is None or (isinstance(value, str) and value.strip() == ''):
            return default
        try:
            return float(value)
        except ValueError:
            return default
        except TypeError: # For cases where value might be a list but float is expected
            return default

    def calculate_ratios(self):
        data = self.financial_data.data
        ratios = {}

        # 獲取所有必要的原始數據
        operating_revenue = self.get_value('operating_revenue')
        cost_of_goods_sold = self.get_value('cost_of_goods_sold')
        operating_expenses = self.get_value('operating_expenses')
        net_profit_after_tax = self.get_value('net_profit_after_tax')
        shareholders_equity = self.get_value('shareholders_equity')
        total_assets = self.get_value('total_assets')
        net_profit_before_tax = self.get_value('net_profit_before_tax')
        interest_expense = self.get_value('interest_expense')
        operating_cash_flow = self.get_value('operating_cash_flow')
        capital_expenditures = self.get_value('capital_expenditures')
        financing_cash_flow = self.get_value('financing_cash_flow')
        current_assets = self.get_value('current_assets')
        current_liabilities = self.get_value('current_liabilities')
        inventory = self.get_value('inventory')
        accounts_receivable = self.get_value('accounts_receivable')
        cash_and_equivalents = self.get_value('cash_and_equivalents')
        short_term_borrowing = self.get_value('short_term_borrowing')
        prev_year_net_profit_after_tax = self.get_value('prev_year_net_profit_after_tax')
        prev_year_operating_revenue = self.get_value('prev_year_operating_revenue')
        prev_year_inventory = self.get_value('prev_year_inventory')
        prev_year_accounts_receivable = self.get_value('prev_year_accounts_receivable')
        total_liabilities = self.get_value('total_assets') - self.get_value('shareholders_equity') # 總負債 = 總資產 - 股東權益
        prev_total_liabilities = self.get_value('prev_total_liabilities')
        prev_total_assets = self.get_value('prev_total_assets')


        # 獲利能力比率
        ratios['gross_profit_margin'] = (operating_revenue - cost_of_goods_sold) / operating_revenue if operating_revenue != 0 else 0.0
        ratios['operating_profit_margin'] = (operating_revenue - cost_of_goods_sold - operating_expenses) / operating_revenue if operating_revenue != 0 else 0.0
        ratios['net_profit_margin'] = net_profit_after_tax / operating_revenue if operating_revenue != 0 else 0.0
        ratios['roe'] = net_profit_after_tax / shareholders_equity if shareholders_equity != 0 else 0.0
        ratios['roa'] = net_profit_after_tax / total_assets if total_assets != 0 else 0.0
        ratios['net_profit_growth_rate'] = (net_profit_after_tax - prev_year_net_profit_after_tax) / prev_year_net_profit_after_tax if prev_year_net_profit_after_tax != 0 else 0.0
        ratios['revenue_growth_rate'] = (operating_revenue - prev_year_operating_revenue) / prev_year_operating_revenue if prev_year_operating_revenue != 0 else 0.0
        ratios['profit_cash_content'] = operating_cash_flow / net_profit_after_tax if net_profit_after_tax != 0 else 0.0

        # 償債能力比率
        ratios['current_ratio'] = current_assets / current_liabilities if current_liabilities != 0 else 0.0
        ratios['quick_ratio'] = (current_assets - inventory) / current_liabilities if current_liabilities != 0 else 0.0
        ebit = net_profit_before_tax + interest_expense
        ratios['interest_coverage_ratio'] = ebit / interest_expense if interest_expense != 0 else float('inf')

        # 營運效率比率
        avg_inventory = (inventory + prev_year_inventory) / 2 if prev_year_inventory != 0 else inventory
        ratios['inventory_turnover_rate'] = cost_of_goods_sold / avg_inventory if avg_inventory != 0 else 0.0
        ratios['inventory_turnover_days'] = 365 / ratios['inventory_turnover_rate'] if ratios['inventory_turnover_rate'] != 0 else float('inf')

        avg_accounts_receivable = (accounts_receivable + prev_year_accounts_receivable) / 2 if prev_year_accounts_receivable != 0 else accounts_receivable
        ratios['accounts_receivable_turnover_rate'] = operating_revenue / avg_accounts_receivable if avg_accounts_receivable != 0 else 0.0
        ratios['accounts_receivable_turnover_days'] = 365 / ratios['accounts_receivable_turnover_rate'] if ratios['accounts_receivable_turnover_rate'] != 0 else float('inf')

        # 現金流量相關比率
        ratios['free_cash_flow'] = operating_cash_flow - capital_expenditures
        ratios['financing_to_operating_cash_flow_ratio'] = financing_cash_flow / operating_cash_flow if operating_cash_flow != 0 else 0.0

        # 負債比率
        ratios['debt_ratio'] = total_liabilities / total_assets if total_assets != 0 else 0.0

        # 財務費用佔營收比例
        ratios['financial_expense_to_revenue_ratio'] = interest_expense / operating_revenue if operating_revenue != 0 else 0.0

        # 淨負債
        ratios['net_debt'] = total_liabilities - cash_and_equivalents

        return ratios

    def assess_profit_quality(self, ratios):
        data = self.financial_data.data
        score = 0
        conclusion_list = [] # 使用列表來儲存多個結論
        details = {}

        # 獲取所需數據
        profit_cash_content = ratios.get('profit_cash_content', 0.0)
        ar_turnover_days = ratios.get('accounts_receivable_turnover_days', float('inf'))
        non_recurring_gain_loss = self.get_value('non_recurring_gain_loss')
        total_profit_for_non_recurring = self.get_value('total_profit') # 假設這個是計算非經常性損益佔比的基準利潤
        net_profit_growth_rate = ratios.get('net_profit_growth_rate', 0.0)

        # 計算非經常性損益佔比
        non_recurring_profit_ratio = abs(non_recurring_gain_loss) / total_profit_for_non_recurring if total_profit_for_non_recurring != 0 else 0.0

        # --- 是/否判斷 ---
        is_profit_cash_content_high = profit_cash_content >= 1.0 # 獲利含金量 >= 100%
        is_ar_days_short = ar_turnover_days <= 45 # 應收帳款周轉天數 <= 45天
        is_non_recurring_low = non_recurring_profit_ratio <= 0.10 # 非經常性損益佔比 <= 10%
        is_net_profit_growing = net_profit_growth_rate > 0 # 淨利成長率 > 0%

        details['獲利含金量 >= 100%?'] = "是" if is_profit_cash_content_high else "否"
        details['應收帳款周轉天數 <= 45天?'] = "是" if is_ar_days_short else "否"
        details['非經常性損益佔比 <= 10%?'] = "是" if is_non_recurring_low else "否"
        details['淨利成長率 > 0%?'] = "是" if is_net_profit_growing else "否"

        # --- 評分邏輯 (根據文件中的分數區間) ---
        # 1. 獲利含金量 (30%)
        if profit_cash_content >= 1.0: score += 30
        elif 0.8 <= profit_cash_content < 1.0: score += 24 # 80% of 30
        elif 0.7 <= profit_cash_content < 0.8: score += 21 # 70% of 30
        elif 0.5 <= profit_cash_content < 0.7: score += 16.5 # 55% of 30
        elif 0.35 <= profit_cash_content < 0.5: score += 9 # 30% of 30
        elif 0.0 <= profit_cash_content < 0.35: score += 6 # 20% of 30
        else: score += 3 # 10% of 30 (for negative)

        # 2. 應收帳款周轉天數 (30%)
        if ar_turnover_days <= 20: score += 30
        elif 21 <= ar_turnover_days <= 45: score += 24 # 80% of 30
        elif 46 <= ar_turnover_days <= 90: score += 18 # 60% of 30
        elif 91 <= ar_turnover_days <= 140: score += 12 # 40% of 30
        else: score += 3 # 10% of 30

        # 3. 非經常性損益佔比 <= 10% (15%)
        score += (15 if is_non_recurring_low else 7.5) # 100% or 50% of 15

        # 4. 淨利成長率 (25%)
        if net_profit_growth_rate >= 0.70: score += 25
        elif 0.30 <= net_profit_growth_rate < 0.70: score += 20 # 80% of 25
        elif 0.0 <= net_profit_growth_rate < 0.30: score += 15 # 60% of 25
        else: score += 10 # 40% of 25 (for negative)

        # --- 結論 (根據是/否判斷組合) ---
        if is_profit_cash_content_high and is_ar_days_short and is_non_recurring_low and is_net_profit_growing:
            conclusion_list.append("獲利品質極佳。公司盈利能力強勁且穩定，現金流健康，應收帳款管理高效。")
        elif is_profit_cash_content_high and is_ar_days_short and is_non_recurring_low and not is_net_profit_growing:
            conclusion_list.append("獲利品質良好但成長動能不足。盈利質量高，但淨利潤未能持續增長，需關注市場變化。")
        elif not is_profit_cash_content_high and is_ar_days_short and is_non_recurring_low and is_net_profit_growing:
            conclusion_list.append("獲利品質有待提升。淨利潤增長強勁且應收帳款管理良好，但獲利含金量不足，需警惕利潤虛增。")
        elif not is_ar_days_short and is_non_recurring_low:
            conclusion_list.append("獲利品質存在疑慮，應收帳款回款慢是主要問題，可能影響現金流。")
        elif not is_non_recurring_low:
            conclusion_list.append("獲利品質不穩定，非經常性損益佔比較高，可能掩蓋核心業務的真實表現。")
        else:
            conclusion_list.append("獲利品質綜合判斷，需根據具體數據進一步分析。")

        return {
            'score': min(score, 100),
            'conclusion': " ".join(conclusion_list),
            'details': details
        }

    def assess_cash_flow(self, ratios):
        data = self.financial_data.data
        score = 0
        conclusion_list = []
        details = {}

        # 獲取所需數據
        operating_cash_flow = self.get_value('operating_cash_flow')
        free_cash_flow = ratios.get('free_cash_flow', 0.0)
        net_profit_after_tax = self.get_value('net_profit_after_tax')
        investing_cash_flow = self.get_value('investing_cash_flow')
        financing_cash_flow = self.get_value('financing_cash_flow')
        capital_expenditures = self.get_value('capital_expenditures')
        three_year_operating_cash_flows = data.get('three_year_operating_cash_flows', [0.0, 0.0, 0.0]) # 確保是列表

        # 計算指標
        op_cf_vs_net_profit = operating_cash_flow / net_profit_after_tax if net_profit_after_tax != 0 else 0.0
        investing_cf_ratio_to_op_cf = investing_cash_flow / operating_cash_flow if operating_cash_flow != 0 else 0.0
        financing_cf_ratio_to_op_cf = financing_cash_flow / operating_cash_flow if operating_cash_flow != 0 else 0.0

        # --- 是/否判斷 ---
        is_op_cf_positive = operating_cash_flow > 0
        is_fcf_positive = free_cash_flow > 0
        is_op_cf_gt_net_profit = op_cf_vs_net_profit > 1.0 # 營業現金流 > 淨利
        is_investing_cf_negative = investing_cash_flow < 0 # 投資現金流為負 (通常表示投資擴張)
        is_financing_cf_negative = financing_cash_flow < 0 # 融資現金流為負 (通常表示還債或回購股票)

        details['營業現金流為正?'] = "是" if is_op_cf_positive else "否"
        details['自由現金流為正?'] = "是" if is_fcf_positive else "否"
        details['營業現金流 > 淨利?'] = "是" if is_op_cf_gt_net_profit else "否"
        details['投資現金流為負?'] = "是" if is_investing_cf_negative else "否"
        details['融資現金流為負?'] = "是" if is_financing_cf_negative else "否"

        # --- 評分邏輯 (根據文件中的分數區間) ---
        # 1. 營業活動現金流 (30%)
        if operating_cash_flow > 5000000: score += 30
        elif 2000000 <= operating_cash_flow <= 5000000: score += 21 # 70% of 30
        elif 500000 <= operating_cash_flow < 2000000: score += 12 # 40% of 30
        elif 0 <= operating_cash_flow < 500000: score += 7.5 # 25% of 30
        else: score += 0

        # 2. 自由現金流 (25%)
        if free_cash_flow > 3000000: score += 25
        elif 1000000 <= free_cash_flow <= 3000000: score += 17.5 # 70% of 25
        elif 500000 <= free_cash_flow < 1000000: score += 10 # 40% of 25
        elif -500000 <= free_cash_flow < 500000: score += 6.25 # 25% of 25
        else: score += 0

        # 3. 營業現金流 > 淨利 (20%)
        if op_cf_vs_net_profit >= 1.5: score += 20
        elif 1.0 <= op_cf_vs_net_profit < 1.5: score += 14 # 70% of 20
        elif 0.7 <= op_cf_vs_net_profit < 1.0: score += 8 # 40% of 20
        elif 0.3 <= op_cf_vs_net_profit < 0.7: score += 5 # 25% of 20
        else: score += 0

        # 4. 投資現金流為負 (15%) - 這裡文件描述較模糊，假設是投資現金流佔營業現金流的比例，且為負數
        # 如果投資現金流是負的，且佔營業現金流的比例在合理範圍內，表示公司有在投資
        if investing_cf_ratio_to_op_cf < 0 and abs(investing_cf_ratio_to_op_cf) <= 0.5: # 負數且佔比不大於50%
            score += 15
        elif investing_cf_ratio_to_op_cf < 0 and abs(investing_cf_ratio_to_op_cf) > 0.5: # 負數且佔比較大
            score += 7.5 # 50% of 15
        else: # 正數或零
            score += 0

        # 5. 融資現金流和營運現金流的比例 (10%)
        if -0.3 <= financing_cf_ratio_to_op_cf <= 0.3: score += 10
        elif 0.3 < financing_cf_ratio_to_op_cf <= 1.0: score += 8 # 80% of 10
        elif -1.0 <= financing_cf_ratio_to_op_cf < -0.3: score += 7 # 70% of 10
        elif financing_cf_ratio_to_op_cf > 1.0: score += 4 # 40% of 10
        else: score += 3 # 30% of 10 (for < -1.0)


        # --- 結論 (根據是/否判斷組合) ---
        # 邏輯架構（現金流量問題）
        if is_op_cf_positive and is_fcf_positive and is_op_cf_gt_net_profit and is_investing_cf_negative and is_financing_cf_negative:
            conclusion_list.append("公司有現金、投資、及有還債能力。現金流狀況極佳，各項指標表現優異，財務健康。")
        elif is_op_cf_positive and (not is_op_cf_gt_net_profit) and is_investing_cf_negative and (not is_financing_cf_negative) and (not is_fcf_positive):
            conclusion_list.append("大量進行投資中，期望未來會有回報。此為投資燒錢階段，需關注未來回報情況。")
        elif (not is_op_cf_positive) and (not is_op_cf_gt_net_profit) and is_investing_cf_negative and (not is_financing_cf_negative) and (not is_fcf_positive):
            conclusion_list.append("現金流入為負、投資燒錢，風險高。公司現金流狀況非常不佳，需警惕資金斷裂風險。")
        elif is_op_cf_positive and is_op_cf_gt_net_profit and (not is_investing_cf_negative) and is_financing_cf_negative and is_fcf_positive:
            conclusion_list.append("有穩定收入了，不再缺錢或借貸，可以開始賺錢了。公司現金流穩健，具備自我造血能力。")
        elif not is_op_cf_positive:
            conclusion_list.append("營業現金流為負，即使其他指標尚可，也可能隱藏獲利品質問題。")
        elif is_op_cf_positive and not is_fcf_positive:
            conclusion_list.append("營業現金流為正但自由現金流為負，現金在投資或營運上消耗較大，需警惕資金壓力。")
        elif is_op_cf_positive and is_fcf_positive and not is_op_cf_gt_net_profit:
            conclusion_list.append("營業現金流為正且自由現金流為正，但未顯著大於淨利，獲利含金量有待提高。")
        elif is_op_cf_positive and is_fcf_positive and is_op_cf_gt_net_profit and not is_investing_cf_negative:
            conclusion_list.append("多數指標良好，但投資現金流不是負值，可能表示投資活動不夠積極或沒有大量資本支出。")
        elif is_op_cf_positive and is_fcf_positive and is_op_cf_gt_net_profit and is_investing_cf_negative and not is_financing_cf_negative:
            conclusion_list.append("營運與投資現金流表現良好，但融資現金流不是負值，可能意味公司仍在依賴外部融資或有償還外部借款壓力。")
        else:
            conclusion_list.append("現金流狀況綜合判斷，需根據具體數據進一步分析。")

        return {
            'score': min(score, 100),
            'conclusion': " ".join(conclusion_list),
            'details': details
        }

    def assess_liquidity_risk(self, ratios):
        data = self.financial_data.data
        score = 0
        conclusion_list = []
        details = {}

        # 獲取所需數據
        current_ratio = ratios.get('current_ratio', 0.0)
        quick_ratio = ratios.get('quick_ratio', 0.0)
        operating_cash_flow = self.get_value('operating_cash_flow')
        cash_and_equivalents = self.get_value('cash_and_equivalents')
        short_term_borrowing = self.get_value('short_term_borrowing')
        interest_expense = self.get_value('interest_expense')
        prev_inventory_turnover_rate = self.get_value('prev_year_inventory_turnover_rate')
        inventory_turnover_rate = ratios.get('inventory_turnover_rate', 0.0)
        three_year_operating_cash_flows = data.get('three_year_operating_cash_flows', [0.0, 0.0, 0.0])

        # 計算指標
        cash_to_short_debt_ratio = cash_and_equivalents / short_term_borrowing if short_term_borrowing != 0 else float('inf')
        op_cf_to_interest_coverage = operating_cash_flow / interest_expense if interest_expense != 0 else float('inf')

        # 存貨周轉天數變化判斷
        current_inv_days = 365 / inventory_turnover_rate if inventory_turnover_rate != 0 else float('inf')
        prev_inv_days = 365 / prev_inventory_turnover_rate if prev_inventory_turnover_rate != 0 else float('inf')

        is_inventory_days_stable_or_down = False
        if prev_inv_days == float('inf') and current_inv_days == float('inf'):
            is_inventory_days_stable_or_down = True # 都無限大，視為穩定
        elif prev_inv_days == float('inf'): # 去年無限大，今年有值，視為改善
            is_inventory_days_stable_or_down = True
        elif current_inv_days <= prev_inv_days:
            is_inventory_days_stable_or_down = True


        # --- 是/否判斷 ---
        is_current_ratio_ok = current_ratio > 2.0
        is_quick_ratio_ok = quick_ratio > 1.0
        is_op_cf_positive = operating_cash_flow > 0
        is_cash_gt_short_debt = cash_to_short_debt_ratio >= 1.0
        is_op_cf_covers_interest = op_cf_to_interest_coverage > 1.0

        details['流動比率 > 2?'] = "是" if is_current_ratio_ok else "否"
        details['速動比率 > 1?'] = "是" if is_quick_ratio_ok else "否"
        details['營業現金流為正?'] = "是" if is_op_cf_positive else "否"
        details['存貨周轉天數穩定或下降?'] = "是" if is_inventory_days_stable_or_down else "否"
        details['現金及約當現金 > 短期借款?'] = "是" if is_cash_gt_short_debt else "否"
        details['營業現金流能覆蓋利息支出?'] = "是" if is_op_cf_covers_interest else "否"

        # --- 評分邏輯 (根據文件中的分數區間) ---
        # 1. 現金及約當現金 > 短期借款 (25%)
        if cash_to_short_debt_ratio >= 2.0: score += 25
        elif 1.0 <= cash_to_short_debt_ratio < 2.0: score += 16.7 # 2/3 of 25
        elif 0.5 <= cash_to_short_debt_ratio < 1.0: score += 8.3 # 1/3 of 25
        else: score += 0

        # 2. 營業現金流為正 (25%)
        positive_op_cf_count = sum(1 for cf in three_year_operating_cash_flows if cf > 0)
        is_op_cf_growing = all(three_year_operating_cash_flows[i] <= three_year_operating_cash_flows[i+1] for i in range(len(three_year_operating_cash_flows)-1)) if len(three_year_operating_cash_flows) > 1 else False

        if positive_op_cf_count == 3 and is_op_cf_growing: score += 25
        elif positive_op_cf_count == 3: score += 20 # 80% of 25
        elif is_op_cf_positive: score += 15 # 60% of 25 (最近一年為正)
        elif positive_op_cf_count == 2: score += 10 # 40% of 25
        elif positive_op_cf_count == 1: score += 5 # 20% of 25
        else: score += 0

        # 3. 流動比率 > 2 (15%)
        if current_ratio > 3.0: score += 15
        elif 2.5 <= current_ratio <= 3.0: score += 12 # 80% of 15
        elif 2.0 <= current_ratio < 2.5: score += 9 # 60% of 15
        elif 1.5 <= current_ratio < 2.0: score += 6 # 40% of 15
        elif 1.0 <= current_ratio < 1.5: score += 3 # 20% of 15
        else: score += 0

        # 4. 存貨周轉天數穩定或下降 (15%)
        if is_inventory_days_stable_or_down: score += 15
        else: score += 0

        # 5. 營業現金流能覆蓋利息支出 (10%)
        if op_cf_to_interest_coverage > 5.0: score += 10
        elif 3.0 <= op_cf_to_interest_coverage <= 5.0: score += 8 # 80% of 10
        elif 1.0 <= op_cf_to_interest_coverage < 3.0: score += 6 # 60% of 10
        elif 0.5 <= op_cf_to_interest_coverage < 1.0: score += 4 # 40% of 10
        elif 0.2 <= op_cf_to_interest_coverage < 0.5: score += 2 # 20% of 10
        else: score += 0

        # 6. 速動比率 > 1 (10%)
        if quick_ratio > 2.0: score += 10
        elif 1.5 <= quick_ratio <= 2.0: score += 8 # 80% of 10
        elif 1.2 <= quick_ratio < 1.5: score += 6 # 60% of 10
        elif 1.0 <= quick_ratio < 1.2: score += 4 # 40% of 10
        elif 0.7 <= quick_ratio < 1.0: score += 2 # 20% of 10
        else: score += 0

        # --- 結論 (根據是/否判斷組合) ---
        # 邏輯架構（流動性風險評估）
        if is_current_ratio_ok and is_quick_ratio_ok and is_op_cf_positive and is_inventory_days_stable_or_down and is_cash_gt_short_debt and is_op_cf_covers_interest:
            conclusion_list.append("公司短期流動性充足，無立即償債風險，現金管理穩健。財務狀況非常健康。")
        elif is_current_ratio_ok and not is_quick_ratio_ok:
            conclusion_list.append("流動性可能依賴存貨變現，需檢查存貨周轉率是否惡化。應警惕存貨積壓風險。")
        elif (not is_current_ratio_ok) and (not is_quick_ratio_ok) and is_op_cf_positive:
            conclusion_list.append("靠著本業現金流維持，但營運一出現問題便陷入資金困境。短期償債能力有疑慮，風險較高。")
        elif is_current_ratio_ok and is_quick_ratio_ok and (not is_op_cf_positive):
            conclusion_list.append("償債比率尚可，但現金流異常偏弱，需提防盈餘品質不佳導致短期資金風險。利潤可能未轉化為現金。")
        elif is_inventory_days_stable_or_down and (not is_current_ratio_ok or not is_quick_ratio_ok or not is_op_cf_positive or not is_cash_gt_short_debt or not is_op_cf_covers_interest):
            conclusion_list.append("有一定營運效率，但財務結構失衡、現金流不足，風險高。存貨管理良好，但整體流動性仍需改善。")
        elif not is_op_cf_positive:
            conclusion_list.append("即使流動比率達標，營運活動未產生現金，可能隱藏獲利品質問題。營業現金流為負是嚴重警訊。")
        else:
            conclusion_list.append("流動性風險綜合判斷，需根據具體數據進一步分析。")

        return {
            'score': min(score, 100),
            'conclusion': " ".join(conclusion_list),
            'details': details
        }

    def assess_debt_solvency(self, ratios):
        data = self.financial_data.data
        score = 0
        conclusion_list = []
        details = {}

        # 獲取所需數據
        interest_coverage_ratio = ratios.get('interest_coverage_ratio', 0.0)
        roa = ratios.get('roa', 0.0)
        cost_of_debt_interest_rate = self.get_value('cost_of_debt_interest_rate')
        free_cash_flow = ratios.get('free_cash_flow', 0.0)
        cash_dividends_paid = self.get_value('cash_dividends_paid')
        debt_ratio = ratios.get('debt_ratio', 0.0)
        financial_expense_to_revenue_ratio = ratios.get('financial_expense_to_revenue_ratio', 0.0)
        prev_total_liabilities = self.get_value('prev_total_liabilities')
        prev_total_assets = self.get_value('prev_total_assets')
        prev_debt_ratio = prev_total_liabilities / prev_total_assets if prev_total_assets != 0 else 0.0

        # --- 是/否判斷 ---
        is_interest_coverage_good = interest_coverage_ratio > 3.0
        is_roa_higher_than_debt_rate = roa > cost_of_debt_interest_rate if cost_of_debt_interest_rate != 0 else True # 如果沒有負債利率，預設為好
        is_fcf_sufficient_for_dividend = (free_cash_flow >= cash_dividends_paid and cash_dividends_paid > 0) or (cash_dividends_paid == 0 and free_cash_flow >= 0) # FCF足以支付股利，或無股利且FCF為正
        is_debt_ratio_high = debt_ratio > 0.50 # 負債比率 > 50%
        is_financial_expense_high = financial_expense_to_revenue_ratio > 0.05 # 財務費用佔營收比例 > 5%
        is_debt_ratio_increased = debt_ratio > prev_debt_ratio and prev_debt_ratio != 0 # 負債比率較去年上升

        details['利息保障倍數 > 3?'] = "是" if is_interest_coverage_good else "否"
        details['ROA > 負債利率?'] = "是" if is_roa_higher_than_debt_rate else "否"
        details['自由現金流足以支付股利?'] = "是" if is_fcf_sufficient_for_dividend else "否"
        details['負債比率 > 50%?'] = "是" if is_debt_ratio_high else "否"
        details['財務費用佔營收比例 > 5%?'] = "是" if is_financial_expense_high else "否"
        details['負債比率較去年上升?'] = "是" if is_debt_ratio_increased else "否"

        # --- 評分邏輯 (根據文件中的分數區間) ---
        # 1. 利息保障倍數 (25%)
        if interest_coverage_ratio > 5.0: score += 25
        elif 3.0 <= interest_coverage_ratio <= 5.0: score += 20 # 80% of 25
        elif 1.0 <= interest_coverage_ratio < 3.0: score += 10 # 40% of 25
        else: score += 0

        # 2. ROA 高於負債利率 (20%)
        if is_roa_higher_than_debt_rate: score += 20
        else: score += 0

        # 3. 自由現金流足以支付股利 (20%)
        if is_fcf_sufficient_for_dividend: score += 20
        else: score += 0

        # 4. 負債比率 (20%)
        if debt_ratio < 0.30: score += 20
        elif 0.30 <= debt_ratio < 0.50: score += 15 # 75% of 20
        elif 0.50 <= debt_ratio < 0.70: score += 10 # 50% of 20
        elif 0.70 <= debt_ratio < 0.90: score += 5 # 25% of 20
        else: score += 0

        # 5. 財務費用佔營收比例 (15%)
        if financial_expense_to_revenue_ratio < 0.01: score += 15
        elif 0.01 <= financial_expense_to_revenue_ratio < 0.03: score += 10 # 66% of 15
        elif 0.03 <= financial_expense_to_revenue_ratio < 0.05: score += 5 # 33% of 15
        else: score += 0

        # --- 結論 (根據是/否判斷組合) ---
        # 邏輯架構（負債與償債能力）
        if is_interest_coverage_good and (not is_debt_ratio_high): # 利息保障倍數「是」且負債比率不高
            conclusion_list.append("負債少、賺錢亦夠還債，用錢有效率。公司財務結構穩健，償債能力強勁。")
        elif is_interest_coverage_good and is_debt_ratio_high:
            conclusion_list.append("雖負債高，但當前獲利足以支撐利息，需關注未來利率變動風險。公司槓桿運用較高，但償債能力暫無問題。")
        elif is_debt_ratio_high and (not is_fcf_sufficient_for_dividend):
            conclusion_list.append("高負債下現金生成不足，可能需借新還舊，財務風險升高。公司資金壓力較大，償債能力堪憂。")
        elif (not is_interest_coverage_good) and (not is_roa_higher_than_debt_rate) and (not is_fcf_sufficient_for_dividend) and (not is_debt_ratio_high) and (not is_financial_expense_high):
            conclusion_list.append("低負債且償債能力強，但可能過度保守，錯失槓桿獲利機會。公司財務狀況穩健，但成長潛力可能受限。")
        elif not is_interest_coverage_good:
            conclusion_list.append("利息保障倍數不足，償債壓力大。負債與償還能力不佳。")
        elif is_interest_coverage_good and not is_roa_higher_than_debt_rate:
            conclusion_list.append("雖然利息有保障，但資產報酬率低於負債利率，借款成本高於資產效益。負債與償還能力不佳。")
        elif is_interest_coverage_good and is_roa_higher_than_debt_rate and not is_fcf_sufficient_for_dividend:
            conclusion_list.append("有能力賺錢且資產效益高於負債成本，但自由現金流不足以支付股利，資金周轉可能緊張。負債與償還能力不佳。")
        elif is_interest_coverage_good and is_roa_higher_than_debt_rate and is_fcf_sufficient_for_dividend and is_debt_ratio_high:
            conclusion_list.append("多數指標良好，但負債比率仍高於50%，存在較高槓桿風險。負債與償還能力不佳。")
        elif is_interest_coverage_good and is_roa_higher_than_debt_rate and is_fcf_sufficient_for_dividend and (not is_debt_ratio_high) and is_financial_expense_high:
            conclusion_list.append("償債能力強且負債比率不高，但財務費用佔營收比例過高，顯示財務槓桿使用效率不佳。負債與償還能力不佳。")
        else:
            conclusion_list.append("負債與償債能力綜合判斷，需根據具體數據進一步分析。")

        return {
            'score': min(score, 100),
            'conclusion': " ".join(conclusion_list),
            'details': details
        }

    def assess_operational_efficiency(self, ratios):
        data = self.financial_data.data
        score = 0
        conclusion_list = []
        details = {}

        # 獲取所需數據
        inventory_turnover_rate = ratios.get('inventory_turnover_rate', 0.0)
        prev_inventory_turnover_rate = self.get_value('prev_year_inventory_turnover_rate')
        accounts_receivable_turnover_days = ratios.get('accounts_receivable_turnover_days', float('inf'))
        prev_accounts_receivable_turnover_days = self.get_value('prev_year_accounts_receivable_turnover_days')
        gross_profit_margin = ratios.get('gross_profit_margin', 0.0)
        prev_gross_profit_margin = self.get_value('prev_year_gross_profit_margin')
        operating_revenue = self.get_value('operating_revenue')
        prev_year_operating_revenue = self.get_value('prev_year_operating_revenue')
        industry_avg_revenue_growth_rate = self.get_value('industry_avg_revenue_growth_rate')
        accounts_payable_days = self.get_value('accounts_payable_days')
        # 假設有去年的應付帳款天數，如果沒有則用當前值
        prev_accounts_payable_days = data.get('prev_year_accounts_payable_days', accounts_payable_days)
        inventory = self.get_value('inventory')
        prev_year_inventory = self.get_value('prev_year_inventory')
        accounts_receivable = self.get_value('accounts_receivable')
        prev_year_accounts_receivable = self.get_value('prev_year_accounts_receivable')


        # 計算指標變化
        inv_turnover_rate_change_pct = (inventory_turnover_rate - prev_inventory_turnover_rate) / prev_inventory_turnover_rate if prev_inventory_turnover_rate != 0 else 0.0
        ar_days_change_pct = (accounts_receivable_turnover_days - prev_accounts_receivable_turnover_days) / prev_accounts_receivable_turnover_days if prev_accounts_receivable_turnover_days != 0 else 0.0
        gross_margin_change_abs = abs(gross_profit_margin - prev_gross_profit_margin)
        revenue_growth_rate = (operating_revenue - prev_year_operating_revenue) / prev_year_operating_revenue if prev_year_operating_revenue != 0 else 0.0
        inventory_growth_rate = (inventory - prev_year_inventory) / prev_year_inventory if prev_year_inventory != 0 else 0.0
        ap_days_change_pct = (accounts_payable_days - prev_accounts_payable_days) / prev_accounts_payable_days if prev_accounts_payable_days != 0 else 0.0
        ar_growth_rate = (accounts_receivable - prev_year_accounts_receivable) / prev_year_accounts_receivable if prev_year_accounts_receivable != 0 else 0.0


        # --- 是/否判斷 ---
        is_inv_turnover_rate_down = inv_turnover_rate_change_pct < 0 # 存貨周轉率下降 (數值變小)
        is_gross_margin_stable = gross_margin_change_abs <= 0.03 # 毛利率波動 <= 3% (絕對差值)
        is_revenue_growth_gt_industry_avg = revenue_growth_rate > industry_avg_revenue_growth_rate
        is_ar_days_stable_or_down = ar_days_change_pct <= 0 # 應收帳款周轉天數穩定或下降 (數值變小或不變)
        is_ap_days_normal = abs(ap_days_change_pct) <= 0.05 # 應付帳款天數波動 <= 5%
        is_revenue_inv_growth_sync = inventory_growth_rate <= revenue_growth_rate # 存貨成長率 <= 營收成長率
        is_revenue_ar_growth_sync = ar_growth_rate <= revenue_growth_rate # 應收帳款成長率 <= 營收成長率

        details['存貨周轉率下降?'] = "是" if is_inv_turnover_rate_down else "否"
        details['毛利率維持穩定?'] = "是" if is_gross_margin_stable else "否"
        details['營收成長率 > 同業平均?'] = "是" if is_revenue_growth_gt_industry_avg else "否"
        details['應收帳款周轉天數穩定或下降?'] = "是" if is_ar_days_stable_or_down else "否"
        details['應付帳款天數無明顯異常?'] = "是" if is_ap_days_normal else "否"
        details['營收與存貨成長同步?'] = "是" if is_revenue_inv_growth_sync else "否"
        details['營收與應收帳款成長同步?'] = "是" if is_revenue_ar_growth_sync else "否"


        # --- 評分邏輯 (根據文件中的分數區間) ---
        # 1. 存貨周轉率變化 (25%)
        if inv_turnover_rate_change_pct > 0.10: score += 25
        elif 0.0 <= inv_turnover_rate_change_pct <= 0.10: score += 20 # 80% of 25
        elif -0.10 <= inv_turnover_rate_change_pct < 0.0: score += 10 # 40% of 25
        else: score += 0 # < -0.10

        # 2. 應收帳款周轉天數變化 (20%)
        if ar_days_change_pct < -0.10: score += 20 # 天數大幅減少
        elif -0.10 <= ar_days_change_pct <= 0.05: score += 15 # 天數小幅減少或穩定
        elif 0.05 < ar_days_change_pct <= 0.20: score += 5 # 天數增加 5%~20%
        else: score += 0 # > 0.20

        # 3. 營收與存貨成長同步性 (20%)
        if is_revenue_inv_growth_sync: score += 20
        elif inventory_growth_rate <= revenue_growth_rate + 0.05: score += 10 # 存貨成長率 > 營收成長率 5% 以內
        elif inventory_growth_rate <= revenue_growth_rate + 0.15: score += 5 # 存貨成長率 > 營收成長率 5%~15%
        else: score += 0

        # 4. 毛利率穩定性 (15%)
        if gross_margin_change_abs <= 0.03: score += 15
        elif gross_margin_change_abs <= 0.05: score += 10 # 波動 3%~5%
        elif gross_margin_change_abs <= 0.10: score += 5 # 波動 5%~10%
        else: score += 0

        # 5. 營收成長 vs. 同業 (10%)
        revenue_growth_gap = revenue_growth_rate - industry_avg_revenue_growth_rate
        if revenue_growth_gap > 0.02: score += 10
        elif abs(revenue_growth_gap) <= 0.02: score += 7 # 與同業差距 ±2% 以內
        elif revenue_growth_gap >= -0.05: score += 3 # 低於同業 2%~5%
        else: score += 0

        # 6. 應付帳款天數異常 (5%)
        if is_ap_days_normal: score += 5
        else: score += 0

        # 7. 營收與應收帳款成長同步 (5%)
        if is_revenue_ar_growth_sync: score += 5
        elif ar_growth_rate <= revenue_growth_rate + 0.05: score += 3 # 應收帳款成長率 > 營收成長率 5% 以內
        else: score += 0

        # --- 結論 (根據是/否判斷組合) ---
        # 邏輯架構（營運效率與周轉問題）
        if (not is_inv_turnover_rate_down) and is_gross_margin_stable: # 存貨周轉率未下降(穩定/上升)且毛利率穩定
            conclusion_list.append("營運效率良好，毛利穩定且存貨周轉情況不錯。這表示公司在營運和盈利能力上表現健康。")
        elif is_inv_turnover_rate_down and (not is_gross_margin_stable):
            conclusion_list.append("營運效率惡化，存貨周轉問題與毛利率波動並存，可能面臨滯銷或價格戰壓力。需警惕存貨跌價損失和盈利能力下降。")
        elif (not is_inv_turnover_rate_down) and is_ar_days_stable_or_down:
            conclusion_list.append("存貨和應收帳款周轉都表現良好，營運效率較高。這顯示公司資金回籠快，資產利用效率高。")
        elif is_inv_turnover_rate_down and is_ap_days_normal and (not is_revenue_inv_growth_sync):
            conclusion_list.append("存貨周轉率下降、營收與存貨成長不同步，儘管應付帳款天數正常，但整體營運效率不佳，可能存在存貨積壓問題。需警惕資金占用和經營風險。")
        elif (not is_ar_days_stable_or_down) and (not is_revenue_growth_gt_industry_avg) and is_gross_margin_stable:
            conclusion_list.append("毛利穩定且存貨周轉尚可，但應收帳款周轉惡化且營收成長不及同業。可能需提防客戶付款風險或市場份額流失。")
        elif is_revenue_growth_gt_industry_avg and is_inv_turnover_rate_down:
            conclusion_list.append("營收成長雖快，但存貨周轉率下降可能預示著盲目擴張或存貨管理問題。需警惕成長的質量。")
        elif not (is_inv_turnover_rate_down or is_gross_margin_stable or is_revenue_growth_gt_industry_avg or is_ar_days_stable_or_down or is_ap_days_normal or is_revenue_inv_growth_sync or is_revenue_ar_growth_sync): # 若所有關鍵判斷皆為"否"
            conclusion_list.append("各項營運效率指標均表現不佳，可能面臨嚴重的經營困境和資金壓力。建議立即審視公司策略。")
        else:
            conclusion_list.append("營運效率綜合判斷，需根據具體數據進一步分析。")

        return {
            'score': min(score, 100),
            'conclusion': " ".join(conclusion_list),
            'details': details
        }

    def assess_investment_expansion(self, ratios):
        data = self.financial_data.data
        score = 0
        conclusion_list = []
        details = {}

        # 獲取所需數據
        capital_expenditures = self.get_value('capital_expenditures')
        operating_cash_flow = self.get_value('operating_cash_flow')
        roe = ratios.get('roe', 0.0)
        industry_avg_roe = self.get_value('industry_avg_roe')
        current_debt_ratio = ratios.get('debt_ratio', 0.0) # 從 ratios 獲取
        prev_total_liabilities = self.get_value('prev_total_liabilities')
        prev_total_assets = self.get_value('prev_total_assets')
        prev_debt_ratio = prev_total_liabilities / prev_total_assets if prev_total_assets != 0 else 0.0
        free_cash_flow = ratios.get('free_cash_flow', 0.0)
        net_debt = ratios.get('net_debt', 0.0)
        prev_net_debt = self.get_value('prev_net_debt')
        revenue = self.get_value('operating_revenue') # 假設營收用於FCF狀態判斷

        # 計算指標
        capex_to_ocf_ratio = capital_expenditures / operating_cash_flow if operating_cash_flow != 0 else 0.0
        roe_diff_from_industry = roe - industry_avg_roe
        net_debt_change_pct = (net_debt - prev_net_debt) / abs(prev_net_debt) if prev_net_debt != 0 else 0.0
        debt_ratio_change_pct = (current_debt_ratio - prev_debt_ratio) / prev_debt_ratio if prev_debt_ratio != 0 else 0.0


        # --- 是/否判斷 ---
        is_capex_high = capex_to_ocf_ratio > 0.5 and capital_expenditures > 0 # 資本支出佔營業現金流比例 > 50%
        is_roe_gt_industry_avg = roe_diff_from_industry > 0 # ROE > 同業平均
        is_debt_ratio_increased = debt_ratio_change_pct > 0 # 負債比率較去年上升
        is_fcf_positive = free_cash_flow > 0 # 自由現金流為正
        is_net_debt_increased = net_debt_change_pct > 0 # 淨負債增加

        details['資本支出佔營業現金流比例 > 50%?'] = "是" if is_capex_high else "否"
        details['ROE > 同業平均?'] = "是" if is_roe_gt_industry_avg else "否"
        details['負債比率較去年上升?'] = "是" if is_debt_ratio_increased else "否"
        details['自由現金流為正?'] = "是" if is_fcf_positive else "否"
        details['淨負債增加?'] = "是" if is_net_debt_increased else "否"

        # --- 評分邏輯 (根據文件中的分數區間) ---
        # 1. 自由現金流狀態 (10分)
        if revenue > 0: # Avoid division by zero
            if free_cash_flow > 0.3 * revenue: score += 10 # 充裕現金流 (>+30%營收)
            elif free_cash_flow > 0.1 * revenue: score += 8 # 穩定現金流 (+10%~+30%營收)
            elif free_cash_flow > 0: score += 5 # 小幅正現金流 (0~+10%營收)
            else: score += 2 # 負數 (現金流困難)
        elif free_cash_flow > 0: score += 5
        else: score += 2

        # 2. 資本支出/營業現金流比例 (10分)
        if capex_to_ocf_ratio > 0.8: score += 10
        elif capex_to_ocf_ratio > 0.5: score += 8
        elif capex_to_ocf_ratio > 0.2: score += 5
        else: score += 2

        # 3. ROE 超過同業幅度 (10分)
        if roe_diff_from_industry > 0.05: score += 10 # 高 >5%
        elif roe_diff_from_industry > 0.02: score += 8 # 高 2%~5%
        elif abs(roe_diff_from_industry) <= 0.02: score += 5 # ±2%內
        else: score += 2 # 低於同業

        # 4. 淨負債變動 (10分)
        if net_debt_change_pct < -0.05: score += 10 # 減少 (>5%)
        elif -0.05 <= net_debt_change_pct <= 0.05: score += 8 # 小幅增加 (0~5%)
        elif 0.05 < net_debt_change_pct <= 0.15: score += 5 # 中幅增加 (5%~15%)
        else: score += 2 # 大幅增加 (>15%)

        # 5. 負債比率變化 (10分)
        if debt_ratio_change_pct < -0.05: score += 10 # 減少 (>5%)
        elif -0.05 <= debt_ratio_change_pct <= 0.05: score += 8 # 小幅上升 (0~5%)
        elif 0.05 < debt_ratio_change_pct <= 0.15: score += 5 # 中幅上升 (5%~15%)
        else: score += 2 # 大幅上升 (>15%)

        # --- 結論 (根據是/否判斷組合) ---
        # 邏輯架構（投資與擴張合理性）
        if not is_fcf_positive: # 邏輯開端是檢查自由現金流
            conclusion_list.append("自由現金流為負數，不論其他條件，公司短期內都面臨資金壓力，擴張與投資的可能性偏低。")
            if is_capex_high and is_roe_gt_industry_avg and (not is_debt_ratio_increased) and (not is_net_debt_increased):
                conclusion_list.append("然而，高資本支出帶來高回報且負債未顯著增加，顯示財務槓桿與資本支出同步提升，營運效率與投資回報皆表現亮眼，屬於具備良好資金運用能力的成長企業。")
            elif is_capex_high and (not is_roe_gt_industry_avg):
                conclusion_list.append("儘管高資本支出，但ROE未優於同業，投資效益未顯現，可能過度擴張或專案報酬率低。")
            elif (not is_capex_high) and (not is_roe_gt_industry_avg) and (not is_debt_ratio_increased) and (not is_net_debt_increased):
                conclusion_list.append("沒在花錢也沒賺錢，公司太保守或已無成長動能。")
        else: # 自由現金流為正數
            if is_capex_high and is_roe_gt_industry_avg:
                conclusion_list.append("高投資帶來高回報，擴張策略有效。公司處於積極且有效的成長階段。")
            elif is_capex_high and (not is_roe_gt_industry_avg):
                conclusion_list.append("投資效益未顯現，可能過度擴張或專案報酬率低。需審慎評估投資專案回報。")
            elif is_debt_ratio_increased:
                conclusion_list.append("負債比率較去年上升，需評估是否過度依賴融資支撐擴張。公司擴張策略可能伴隨較高財務風險。")
            elif is_roe_gt_industry_avg and (not is_capex_high):
                conclusion_list.append("投資花得少，回報普通但沒亂擴張。公司擴張保守，但投資回報尚可接受。")
            else:
                conclusion_list.append("投資與擴張策略穩健，公司在成長的同時保持了健康的財務狀況。")

        # 根據綜合分數給出總結性結論
        overall_conclusion = ""
        if score >= 40:
            overall_conclusion = "積極擴張且資金與回報俱佳，屬於「優質擴張企業」。"
        elif score >= 30:
            overall_conclusion = "穩健擴張中，部分指標如槓桿或現金流需持續觀察，屬於「成長型企業」。"
        elif score >= 20:
            overall_conclusion = "擴張或回報力道普通，或存在財務壓力，需審慎觀察。"
        else: # <20分
            overall_conclusion = "投資與擴張動能低，或槓桿風險過高，應審慎投資。"

        return {
            'score': min(score, 100),
            'conclusion': " ".join(conclusion_list),
            'overall_conclusion': overall_conclusion, # 加入綜合性結論
            'details': details
        }

# --- 財務術語小百科字典 (No changes needed) ---
TERMS_GLOSSARY = {
    'operating_revenue': "營業收入 (Operating Revenue):\n指企業在日常經營活動中銷售商品或提供服務所獲得的收入，是公司本業的核心收入。",
    'cost_of_goods_sold': "銷貨成本 (Cost of Goods Sold):\n指銷售商品或提供服務直接相關的成本，如原材料、直接人工和製造費用。",
    'operating_expenses': "營業費用 (Operating Expenses):\n指企業在經營活動中發生的，與生產銷售不直接相關但維持企業運營所需的費用，如銷售費用、管理費用和研發費用。",
    'net_profit_after_tax': "稅後淨利 (Net Profit After Tax):\n指企業在扣除所有成本、費用和稅款後，最終歸屬於股東的利潤，是衡量公司盈利能力的最終指標。",
    'shareholders_equity': "股東權益 (Shareholder's Equity):\n指公司資產扣除負債後的淨值，代表股東在公司中的所有權。包括股本、資本公積、保留盈餘等。",
    'total_assets': "總資產 (Total Assets):\n指公司所擁有的一切資源，包括流動資產、非流動資產等，是衡量公司規模的重要指標。",
    'current_assets': "流動資產 (Current Assets):\n指預期在一年內可以變現或消耗掉的資產，如現金、應收帳款、存貨等。",
    'current_liabilities': "流動負債 (Current Liabilities):\n指預期在一年內必須償還的債務，如短期借款、應付帳款等。",
    'inventory': "存貨 (Inventory):\n指公司持有以供銷售、生產過程中使用或將在生產過程中耗用的商品或材料。",
    'accounts_receivable': "應收帳款 (Accounts Receivable):\n指公司因銷售商品或提供服務而應向客戶收取的款項。",
    'interest_expense': "利息費用 (Interest Expense):\n指公司為使用借入資金而支付的成本。",
    'net_profit_before_tax': "稅前淨利 (Net Profit Before Tax):\n指公司在扣除所有成本和費用（不包括所得稅）後的利潤。",
    'operating_cash_flow': "營業活動現金流 (Operating Cash Flow):\n指公司透過日常營運活動（銷售商品、提供服務）所產生或消耗的現金，正數表示本業能賺現金。",
    'investing_cash_flow': "投資活動現金流 (Investing Cash Flow):\n指公司透過投資活動（如購買或出售固定資產、投資其他公司）所產生或消耗的現金，負數通常表示公司在擴張或投資新項目。",
    'financing_cash_flow': "籌資活動現金流 (Financing Cash Flow):\n指公司透過籌資活動（如發行股票、借款、償還債務、發放股利）所產生或消耗的現金。",
    'capital_expenditures': "資本支出 (Capital Expenditures):\n指公司用於購買、升級或維護固定資產（如廠房、設備）所花的資金，是投資活動現金流的重要組成部分。",
    'cash_dividends_paid': "現金股利 (Cash Dividends Paid):\n公司向股東支付的現金分配，通常是從公司利潤中撥出。",
    'non_recurring_gain_loss': "非經常性損益 (Non-recurring Gain/Loss):\n指公司在正常經營活動之外發生的，不具持續性的損益，如出售資產的利得或損失、訴訟賠償等。過高可能影響獲利品質的穩定性。",
    'total_profit': "利潤總額 (Total Profit):\n通常指稅前利潤或淨利潤，在此工具中主要用於計算非經常性損益佔比。",
    'cash_and_equivalents': "現金及約當現金 (Cash and Equivalents):\n指公司持有的現金和可以迅速轉換為現金的資產，如短期、流動性高且容易變現的投資，這些投資到期日通常在三個月以內。",
    'short_term_borrowing': "短期借款 (Short-term Borrowing):\n指公司在一年內必須償還的債務，通常是為了滿足短期營運資金需求。",
    'accounts_payable_days': "應付帳款天數 (Accounts Payable Days):\n衡量公司支付供應商款項的平均天數。天數越長表示公司利用供應商信用的能力越強，但過長可能影響供應商關係。",
    'prev_year_net_profit_after_tax': "去年稅後淨利 (Previous Year Net Profit After Tax):\n用於計算稅後淨利成長率，與本期稅後淨利比較，判斷盈利成長動能。",
    'prev_year_operating_revenue': "去年營業收入 (Previous Year Operating Revenue):\n用於計算營收成長率，與本期營業收入比較，判斷營收增長情況。",
    'prev_year_inventory_turnover_rate': "去年存貨周轉率 (Previous Year Inventory Turnover Rate):\n用於與本期存貨周轉率比較，判斷存貨管理效率的變化趨勢。",
    'prev_year_accounts_receivable_turnover_days': "去年應收帳款周轉天數 (Previous Year AR Turnover Days):\n用於與本期應收帳款周轉天數比較，判斷客戶付款速度的變化趨勢。",
    'prev_year_gross_profit_margin': "去年毛利率 (Previous Year Gross Profit Margin):\n用於與本期毛利率比較，判斷公司核心業務盈利能力的穩定性。",
    'industry_avg_roe': "行業平均ROE (Industry Average ROE):\n股東權益報酬率的行業平均值，用於與公司自身的ROE比較，判斷公司在行業內的盈利能力水平。",
    'industry_avg_revenue_growth_rate': "行業平均營收成長率 (Industry Average Revenue Growth Rate):\n行業的平均營收增長速度，用於與公司自身的營收成長率比較，判斷公司是否跑贏行業。",
    'cost_of_debt_interest_rate': "負債利率 (Cost of Debt Interest Rate):\n公司所承擔債務的平均利率成本，用於判斷資產報酬率是否足以覆蓋負債成本。",
    'prev_total_liabilities': "去年總負債 (Previous Year Total Liabilities):\n用於計算負債比率的年度變化，評估公司負債水平的趨勢。",
    'prev_total_assets': "去年總資產 (Previous Year Total Assets):\n用於計算負債比率的年度變化，評估公司負債水平的趨勢。",
    'prev_net_debt': "去年淨負債 (Previous Year Net Debt):\n用於計算淨負債的年度變化，判斷公司債務負擔的變動趨勢。",
    'three_year_operating_cash_flows': "近三年營業現金流 (Operating Cash Flows for last 3 years):\n公司過去連續三年的營業活動現金流量，用於評估營運現金流的穩定性和持續性。",

    # 比率術語解釋 (這些會動態生成，但我們可以為其添加通用解釋)
    'gross_profit_margin': "毛利率 (Gross Profit Margin):\n(營業收入 - 銷貨成本) / 營業收入。衡量公司核心業務的盈利能力，反映銷售價格與成本控制的效率。",
    'operating_profit_margin': "營業利益率 (Operating Profit Margin):\n(營業收入 - 銷貨成本 - 營業費用) / 營業收入。衡量公司從本業經營中獲取利潤的能力，反映營運效率。",
    'net_profit_margin': "淨利率 (Net Profit Margin):\n稅後淨利 / 營業收入。衡量公司最終盈利效率，反映所有成本和費用扣除後的最終利潤。",
    'roe': "股東權益報酬率 (Return on Equity - ROE):\n稅後淨利 / 股東權益。衡量公司利用股東資本創造利潤的效率。",
    'roa': "總資產報酬率 (Return on Assets - ROA):\n稅後淨利 / 總資產。衡量公司利用總資產創造利潤的效率。",
    'net_profit_growth_rate': "淨利成長率 (Net Profit Growth Rate):\n(當期稅後淨利 - 去年同期稅後淨利) / 去年同期稅後淨利。衡量公司淨利潤的增長速度。",
    'profit_cash_content': "獲利含金量 (Profit Cash Content):\n營業活動現金流 / 稅後淨利。衡量公司淨利潤中有多少比例是實際收到的現金，高於100%通常表示獲利品質較高。",
    'current_ratio': "流動比率 (Current Ratio):\n流動資產 / 流動負債。衡量公司短期償債能力，通常大於1表示流動性較佳。",
    'quick_ratio': "速動比率 (Quick Ratio):\n(流動資產 - 存貨) / 流動負債。比流動比率更保守的短期償債能力指標，排除了流動性較差的存貨。",
    'interest_coverage_ratio': "利息保障倍數 (Interest Coverage Ratio):\n稅前息前利潤 (EBIT) / 利息費用。衡量公司經營利潤支付利息的能力，數值越高表示償債能力越強。",
    'inventory_turnover_rate': "存貨周轉率 (Inventory Turnover Rate):\n銷貨成本 / 平均存貨。衡量公司銷售和補充存貨的效率，周轉率越高通常表示存貨管理越有效率。",
    'accounts_receivable_turnover_rate': "應收帳款周轉率 (Accounts Receivable Turnover Rate):\n營業收入 / 應收帳款。衡量公司收回應收帳款的速度，周轉率越高表示收款越快。",
    'accounts_receivable_turnover_days': "應收帳款周轉天數 (Accounts Receivable Turnover Days):\n365 / 應收帳款周轉率。衡量公司收回應收帳款所需的平均天數，天數越少表示收款越快。",
    'free_cash_flow': "自由現金流 (Free Cash Flow - FCF):\n營業活動現金流 - 資本支出。指公司在滿足自身營運和資本投資需求後，可以自由支配的現金，是衡量公司財務健康和價值的關鍵指標。",
    'financing_to_operating_cash_flow_ratio': "融資現金流與營運現金流比例 (Financing to Operating Cash Flow Ratio):\n籌資活動現金流 / 營業活動現金流。衡量公司營運活動產生的現金流有多少用於融資活動，或融資活動如何彌補營運現金流的不足。",
    'debt_ratio': "負債比率 (Debt Ratio):\n總負債 / 總資產。衡量公司資產中由負債提供資金的比例，反映公司的財務槓桿水平。",
    'financial_expense_to_revenue_ratio': "財務費用佔營收比例 (Financial Expense to Revenue Ratio):\n財務費用 / 營業收入。衡量公司財務成本在營收中的佔比，反映公司負債對盈利能力的影響。",
    'net_debt': "淨負債 (Net Debt):\n總負債 - 現金及約當現金。衡量公司扣除可用現金後的實際負債水平，更準確反映公司的債務負擔。",
    'revenue_growth_rate': "營收成長率 (Revenue Growth Rate):\n(當期營業收入 - 去年同期營業收入) / 去年同期營業收入。衡量公司營業收入的增長速度，反映市場拓展能力。",
}

# --- Streamlit Helper Functions ---
def plot_bar_chart(labels, values, title):
    """
    繪製一個簡單的條形圖並返回 Matplotlib Figure。
    """
    plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'Arial Unicode MS', 'SimHei']
    plt.rcParams['axes.unicode_minus'] = False

    fig, ax = plt.subplots(figsize=(8, 5)) # Increased size for better readability
    bars = ax.bar(labels, values, color='skyblue')
    ax.set_title(title + " - 關鍵指標", fontsize=14, fontproperties=font)
    ax.tick_params(axis='x', rotation=45, labelsize=10, ha="right") # Rotate and align labels
    ax.yaxis.get_major_formatter().set_scientific(False)

    # 根據數值範圍調整y軸標籤格式
    if values: # Check if values list is not empty
        max_val = max(values) if values else 0
        min_val = min(values) if values else 0
        if max_val > 1000:
            ax.ticklabel_format(style='plain', axis='y', useOffset=False)
            ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x:,.0f}'))
        elif -1.0 <= min_val and max_val <= 1.0 and not all(v == 0 for v in values):
            ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: '{:.0%}'.format(y)))

    plt.tight_layout() # Adjust layout to prevent labels overlapping
    return fig

def generate_overall_report_text(calculator, ratios, *analysis_results):
    """
    生成綜合報告的文本內容。
    """
    fd = calculator.financial_data
    report_lines = []

    report_lines.append(f"===== 綜合財務分析報告 ({datetime.now().strftime('%Y-%m-%d %H:%M')}) =====\n\n")
    report_lines.append("--- 關鍵財務比率一覽 ---\n")

    key_ratios_to_report = {
        'gross_profit_margin': '毛利率', 'operating_profit_margin': '營業利益率',
        'net_profit_margin': '淨利率', 'roe': '股東權益報酬率 (ROE)',
        'roa': '總資產報酬率 (ROA)', 'net_profit_growth_rate': '淨利成長率',
        'revenue_growth_rate': '營收成長率', 'profit_cash_content': '獲利含金量',
        'current_ratio': '流動比率', 'quick_ratio': '速動比率',
        'interest_coverage_ratio': '利息保障倍數', 'inventory_turnover_rate': '存貨周轉率',
        'accounts_receivable_turnover_days': '應收帳款周轉天數', 'free_cash_flow': '自由現金流',
        'debt_ratio': '負債比率', 'financial_expense_to_revenue_ratio': '財務費用佔營收比例',
        'net_debt': '淨負債', 'accounts_payable_days': '應付帳款天數',
    }

    table_data = []
    for key, display_name in key_ratios_to_report.items():
        value = ratios.get(key, fd.get_data(key))
        display_value = ""
        if isinstance(value, float):
            if any(s in key for s in ['_margin', '_rate', '_ratio', 'roe', 'roa', 'growth']):
                display_value = f"{value * 100:.2f}%"
            elif 'days' in key:
                display_value = f"{value:.0f}天"
            elif any(s in key for s in ['cash_flow', 'profit', 'assets', 'debt', 'revenue', 'expenses', 'inventory']):
                display_value = f"{value:,.2f} 元"
            else:
                display_value = f"{value:.2f}"
        else:
            display_value = str(value)
        table_data.append([display_name, display_value])

    # Simple text table formatting
    if table_data:
        col_widths = [max(len(str(item)) for item in col) for col in zip(*table_data)]
        header = " | ".join(str(item).ljust(width) for item, width in zip(["比率名稱", "數值"], col_widths))
        separator = "-+-".join('-' * width for width in col_widths)
        report_lines.append(header)
        report_lines.append(separator)
        for row in table_data:
            report_lines.append(" | ".join(str(item).ljust(width) for item, width in zip(row, col_widths)))
    report_lines.append("\n")

    analysis_titles = [
        "獲利品質分析", "現金流量分析", "流動性風險評估",
        "負債與償債能力", "營運效率與周轉", "投資與擴張合理性"
    ]

    for i, result in enumerate(analysis_results):
        report_lines.append(f"--- {analysis_titles[i]} 總結 ---")
        report_lines.append(f"評分: {result.get('score', 0):.2f} / 100")
        report_lines.append(f"結論: {result.get('conclusion', '無結論')}\n")

    return "\n".join(report_lines)

def display_analysis_tab(result, title, labels, values):
    """
    通用函數，用於顯示單個分析分頁的內容。
    """
    st.subheader(f"{title}")
    st.metric(label="總體評分", value=f"{result.get('score', 0):.2f} / 100")
    st.markdown(f"**結論:** {result.get('conclusion', '無結論')}")

    with st.expander("詳細評估細節 (是/否判斷)"):
        st.table(pd.DataFrame(list(result.get('details', {}).items()), columns=["評估項目", "結果"]))

    # Chart
    if labels and values:
        fig = plot_bar_chart(labels, values, title)
        st.pyplot(fig)
    else:
        st.info("無足夠數據繪製圖表。")

# --- Streamlit App ---

st.set_page_config(page_title="財務報表分析工具", layout="wide")
st.title("📊 財務報表分析工具")
st.write("輸入或載入您的財務數據，以獲得全面的財務健康評估。")

# --- Initialize Session State ---
if 'financial_data' not in st.session_state:
    st.session_state.financial_data = FinancialData()
if 'calculator' not in st.session_state:
    st.session_state.calculator = FinancialCalculator(st.session_state.financial_data)
if 'ratios' not in st.session_state:
    st.session_state.ratios = {}
if 'results' not in st.session_state:
    st.session_state.results = {}
if 'data_loaded' not in st.session_state:
    st.session_state.data_loaded = False


# --- Sidebar for Data Input ---
with st.sidebar:
    st.header("數據輸入")
    st.write("您可以從檔案載入或手動輸入數據。")

    # File Uploader
    uploaded_file = st.file_uploader("從檔案載入數據 (CSV/Excel)", type=['csv', 'xlsx', 'xls'])
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)

            st.success(f"已成功載入檔案: {uploaded_file.name}")
            st.dataframe(df.head())

            # Update FinancialData from DataFrame (assuming first row and matching columns)
            for key in st.session_state.financial_data.data.keys():
                if key in df.columns:
                    try:
                        value = df.loc[0, key]
                        if isinstance(st.session_state.financial_data.get_data(key), list):
                            if isinstance(value, str):
                                st.session_state.financial_data.update_data(key, [float(x.strip()) for x in value.split(',')])
                            else:
                                st.session_state.financial_data.update_data(key, [float(value)])
                        else:
                            st.session_state.financial_data.update_data(key, float(value))
                    except (ValueError, TypeError, KeyError):
                        st.warning(f"檔案中 '{key}' 的數據格式或名稱不正確，將使用預設值或0。")
            st.session_state.data_loaded = True

        except Exception as e:
            st.error(f"載入檔案時發生錯誤: {e}")

    # Manual Input
    st.subheader("手動輸入")
    with st.expander("展開以手動輸入數據", expanded=not st.session_state.data_loaded):
        fd = st.session_state.financial_data
        input_fields_map = {
            "營業收入": 'operating_revenue', "銷貨成本": 'cost_of_goods_sold',
            "營業費用": 'operating_expenses', "稅後淨利": 'net_profit_after_tax',
            "股東權益": 'shareholders_equity', "總資產": 'total_assets',
            "流動資產": 'current_assets', "流動負債": 'current_liabilities',
            "存貨": 'inventory', "應收帳款": 'accounts_receivable',
            "利息費用": 'interest_expense', "稅前淨利": 'net_profit_before_tax',
            "營業活動現金流": 'operating_cash_flow', "投資活動現金流": 'investing_cash_flow',
            "籌資活動現金流": 'financing_cash_flow', "資本支出": 'capital_expenditures',
            "現金股利": 'cash_dividends_paid', "非經常性損益": 'non_recurring_gain_loss',
            "利潤總額": 'total_profit', "現金及約當現金": 'cash_and_equivalents',
            "短期借款": 'short_term_borrowing', "應付帳款天數": 'accounts_payable_days',
            "去年稅後淨利": 'prev_year_net_profit_after_tax', "去年營業收入": 'prev_year_operating_revenue',
            "去年存貨": 'prev_year_inventory', "去年應收帳款": 'prev_year_accounts_receivable',
            "去年存貨周轉率": 'prev_year_inventory_turnover_rate', "去年應收帳款周轉天數": 'prev_year_accounts_receivable_turnover_days',
            "去年毛利率": 'prev_year_gross_profit_margin', "行業平均ROE": 'industry_avg_roe',
            "行業平均營收成長率": 'industry_avg_revenue_growth_rate', "負債利率": 'cost_of_debt_interest_rate',
            "去年總負債": 'prev_total_liabilities', "去年總資產": 'prev_total_assets',
            "去年淨負債": 'prev_net_debt',
        }

        # Create two columns for better layout
        col1, col2 = st.columns(2)
        fields = list(input_fields_map.items())
        mid_point = len(fields) // 2

        with col1:
            for label, key in fields[:mid_point]:
                fd.data[key] = st.number_input(
                    label=label,
                    value=float(fd.get_data(key)),
                    format="%.2f",
                    key=f"input_{key}", # Unique key for each input
                    help=TERMS_GLOSSARY.get(key, "暫無解釋。")
                )
        with col2:
            for label, key in fields[mid_point:]:
                 fd.data[key] = st.number_input(
                    label=label,
                    value=float(fd.get_data(key)),
                    format="%.2f",
                    key=f"input_{key}",
                    help=TERMS_GLOSSARY.get(key, "暫無解釋。")
                )

        # Special handling for list input
        three_year_cf_str = st.text_input(
            "近三年營業現金流 (逗號分隔)",
            value=", ".join(map(str, fd.get_data('three_year_operating_cash_flows'))),
            help=TERMS_GLOSSARY.get('three_year_operating_cash_flows')
        )
        try:
            fd.data['three_year_operating_cash_flows'] = [float(x.strip()) for x in three_year_cf_str.split(',') if x.strip()]
        except ValueError:
            st.error("近三年營業現金流格式不正確，請使用逗號分隔的數字。")
            fd.data['three_year_operating_cash_flows'] = [0.0, 0.0, 0.0]

# --- Analysis Button ---
if st.sidebar.button("🚀 執行所有分析", type="primary"):
    try:
        # Update FinancialData from manual inputs before analysis
        # (This is implicitly done by st.number_input updating the fd.data dict)

        st.session_state.ratios = st.session_state.calculator.calculate_ratios()
        st.session_state.results['profit_quality'] = st.session_state.calculator.assess_profit_quality(st.session_state.ratios)
        st.session_state.results['cash_flow'] = st.session_state.calculator.assess_cash_flow(st.session_state.ratios)
        st.session_state.results['liquidity'] = st.session_state.calculator.assess_liquidity_risk(st.session_state.ratios)
        st.session_state.results['debt_solvency'] = st.session_state.calculator.assess_debt_solvency(st.session_state.ratios)
        st.session_state.results['op_efficiency'] = st.session_state.calculator.assess_operational_efficiency(st.session_state.ratios)
        st.session_state.results['inv_expansion'] = st.session_state.calculator.assess_investment_expansion(st.session_state.ratios)
        st.toast("分析完成！請查看各分頁結果。", icon="🎉")
    except Exception as e:
        st.error(f"執行分析時發生錯誤: {e}\n請確認數據輸入是否完整且正確。")


# --- Main Area with Tabs ---
if st.session_state.ratios: # Only show tabs if analysis has run
    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
        "📈 綜合報告", "🏆 獲利品質", "💧 現金流量",
        "💰 流動性風險", "⚖️ 負債與償債", "⚙️ 營運效率", "🏗️ 投資與擴張"
    ])

    fd = st.session_state.financial_data
    ratios = st.session_state.ratios
    results = st.session_state.results

    with tab1: # 綜合報告
        st.header("綜合財務分析報告")
        overall_labels = ["ROE", "ROA", "淨利率", "流動比率", "負債比率", "自由現金流"]
        overall_values = [
            ratios.get('roe', 0.0), ratios.get('roa', 0.0),
            ratios.get('net_profit_margin', 0.0), ratios.get('current_ratio', 0.0),
            ratios.get('debt_ratio', 0.0), ratios.get('free_cash_flow', 0.0)
        ]
        fig = plot_bar_chart(overall_labels, overall_values, "綜合關鍵財務指標")
        st.pyplot(fig)

        report_text = generate_overall_report_text(
            st.session_state.calculator,
            ratios,
            results['profit_quality'], results['cash_flow'], results['liquidity'],
            results['debt_solvency'], results['op_efficiency'], results['inv_expansion']
        )
        st.text_area("報告內容", report_text, height=400)
        st.download_button(
            label="💾 儲存報告",
            data=report_text,
            file_name=f"financial_report_{datetime.now().strftime('%Y%m%d')}.txt",
            mime="text/plain"
        )


    with tab2: # 獲利品質
        display_analysis_tab(
            results['profit_quality'],
            "獲利品質分析",
            ["獲利含金量", "應收帳款天數", "非經常性損益佔比", "淨利成長率"],
            [
                ratios.get('profit_cash_content', 0.0),
                ratios.get('accounts_receivable_turnover_days', 0.0),
                (fd.get_data('non_recurring_gain_loss') / fd.get_data('total_profit')) if fd.get_data('total_profit')!=0 else 0,
                ratios.get('net_profit_growth_rate', 0.0)
            ]
        )

    with tab3: # 現金流量
        display_analysis_tab(
            results['cash_flow'],
            "現金流量分析",
            ["營業現金流", "自由現金流", "營業現金流/淨利", "投資現金流", "融資/營運現金流"],
            [
                fd.get_data('operating_cash_flow'),
                ratios.get('free_cash_flow', 0.0),
                (fd.get_data('operating_cash_flow') / fd.get_data('net_profit_after_tax')) if fd.get_data('net_profit_after_tax')!=0 else 0,
                fd.get_data('investing_cash_flow'),
                ratios.get('financing_to_operating_cash_flow_ratio', 0.0)
            ]
        )

    with tab4: # 流動性風險
        display_analysis_tab(
            results['liquidity'],
            "流動性風險評估",
            ["流動比率", "速動比率", "現金/短期借款", "利息保障倍數"],
            [
                ratios.get('current_ratio', 0.0),
                ratios.get('quick_ratio', 0.0),
                fd.get_data('cash_and_equivalents') / fd.get_data('short_term_borrowing') if fd.get_data('short_term_borrowing')!=0 else 0,
                ratios.get('interest_coverage_ratio', 0.0)
            ]
        )

    with tab5: # 負債與償債
        display_analysis_tab(
            results['debt_solvency'],
            "負債與償債能力",
            ["利息保障倍數", "ROA", "自由現金流/現金股利", "負債比率", "財務費用/營收"],
            [
                ratios.get('interest_coverage_ratio', 0.0),
                ratios.get('roa', 0.0),
                ratios.get('free_cash_flow', 0.0) / fd.get_data('cash_dividends_paid') if fd.get_data('cash_dividends_paid')!=0 else float('inf'), # Handle division by zero
                ratios.get('debt_ratio', 0.0),
                ratios.get('financial_expense_to_revenue_ratio', 0.0)
            ]
        )

    with tab6: # 營運效率
        display_analysis_tab(
            results['op_efficiency'],
            "營運效率與周轉",
            ["存貨周轉率", "應收帳款周轉天數", "毛利率", "應付帳款天數"],
            [
                ratios.get('inventory_turnover_rate', 0.0),
                ratios.get('accounts_receivable_turnover_days', 0.0),
                ratios.get('gross_profit_margin', 0.0),
                fd.get_data('accounts_payable_days')
            ]
        )

    with tab7: # 投資與擴張
        display_analysis_tab(
            results['inv_expansion'],
            "投資與擴張合理性",
            ["自由現金流", "資本支出/營運現金流", "ROE", "淨負債變動"],
            [
                ratios.get('free_cash_flow', 0.0),
                (fd.get_data('capital_expenditures') / fd.get_data('operating_cash_flow')) if fd.get_data('operating_cash_flow')!=0 else 0,
                ratios.get('roe', 0.0),
                ratios.get('net_debt', 0.0) - fd.get_data('prev_net_debt')
            ]
        )

else:
    st.info("請在左側輸入或載入數據，然後點擊 '執行所有分析' 按鈕以查看結果。")

st.sidebar.markdown("---")
st.sidebar.caption("© 2024 Financial Analyzer (Streamlit Version)")