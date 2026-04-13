"""
cal_house_loan.py

支持两部分贷款（商业贷款 + 公积金贷款）的月供与利息计算。
默认本金：商业 30 万元、公积金 100 万元（单位：元）。
计算并输出：等额本息（EMI）与等额本金两种还款方式下的月供、总利息、总还款等信息，
分别为两部分贷款计算后将结果相加得到组合贷款结果。

所有输入通过文件内变量预设，运行时不再交互输入。
"""

from typing import Tuple  # typing是Python的类型提示模块，Tuple表示元组类型


def calc_emi(principal: float, annual_rate: float, years: int) -> Tuple[float, float, float, float, float]:
    """等额本息（EMI）计算

    返回：(monthly_payment, total_payment, total_interest, first_month_principal, first_month_interest)
    """
    if principal <= 0 or years <= 0:
        return 0.0, 0.0, 0.0, 0.0, 0.0

    monthly_rate = annual_rate / 100.0 / 12.0
    months = years * 12
    if monthly_rate == 0:
        monthly_payment = principal / months
    else:
        monthly_payment = principal * monthly_rate * (1 + monthly_rate) ** months / ((1 + monthly_rate) ** months - 1)

    total_payment = monthly_payment * months
    total_interest = total_payment - principal

    # 第一个月利息 = principal * monthly_rate
    first_month_interest = principal * monthly_rate
    first_month_principal = monthly_payment - first_month_interest

    return monthly_payment, total_payment, total_interest, first_month_principal, first_month_interest


def calc_equal_principal(principal: float, annual_rate: float, years: int) -> Tuple[float, float, float, float]:
    """等额本金计算

    返回：(first_month_payment, last_month_payment, total_payment, total_interest)
    """
    if principal <= 0 or years <= 0:
        return 0.0, 0.0, 0.0, 0.0

    monthly_rate = annual_rate / 100.0 / 12.0
    months = years * 12
    monthly_principal = principal / months

    total_interest = 0.0
    first_month_payment = monthly_principal + principal * monthly_rate
    last_month_payment = monthly_principal + (principal - monthly_principal * (months - 1)) * monthly_rate

    # 累计利息 = sum_{k=0..months-1} (principal - monthly_principal*k) * monthly_rate
    # 等差求和形式
    # total_interest = monthly_rate * (months * principal - monthly_principal * (months*(months-1)/2))
    total_interest = monthly_rate * (months * principal - monthly_principal * (months * (months - 1) / 2.0))
    total_payment = principal + total_interest

    return first_month_payment, last_month_payment, total_payment, total_interest


def format_currency(x: float) -> str:
    return f"{x:,.2f} 元"


def run_calculation():
    # 默认参数：本金（元），年利率（%），期限（年）
    commercial_principal = 280000.0   # 商业贷款 28 万
    provident_principal = 1000000.0   # 公积金贷款 100 万

    # 默认利率，可根据当地政策调整
    commercial_annual_rate = 3.05
    provident_annual_rate = 2.6

    loan_years = 30  # 默认 30 年，可按需修改

    print("房贷计算器（组合贷款：商业 + 公积金）\n")

    print("参数设置：")
    print(f"  商业贷款本金：{format_currency(commercial_principal)}，年利率：{commercial_annual_rate}%")
    print(f"  公积金贷款本金：{format_currency(provident_principal)}，年利率：{provident_annual_rate}%")
    print(f"  贷款期限：{loan_years} 年\n")

    # 商业贷款 - 等额本息
    b_emi = calc_emi(commercial_principal, commercial_annual_rate, loan_years)
    # 公积金贷款 - 等额本息
    p_emi = calc_emi(provident_principal, provident_annual_rate, loan_years)

    total_monthly_emi = b_emi[0] + p_emi[0]
    total_payment_emi = b_emi[1] + p_emi[1]
    total_interest_emi = b_emi[2] + p_emi[2]

    print("=== 等额本息（EMI）结果 ===")
    print(f"每月月供（商业）：{format_currency(b_emi[0])}（含首月本金{format_currency(b_emi[3])}，首月利息{format_currency(b_emi[4])}）")
    print(f"每月月供（公积金）：{format_currency(p_emi[0])}（含首月本金{format_currency(p_emi[3])}，首月利息{format_currency(p_emi[4])}）")
    print(f"合计每月月供：{format_currency(total_monthly_emi)}")
    print(f"总还款：{format_currency(total_payment_emi)}，总利息：{format_currency(total_interest_emi)}\n")

    # 等额本金
    b_ep = calc_equal_principal(commercial_principal, commercial_annual_rate, loan_years)
    p_ep = calc_equal_principal(provident_principal, provident_annual_rate, loan_years)

    total_payment_ep = b_ep[2] + p_ep[2]
    total_interest_ep = b_ep[3] + p_ep[3]
    first_month_payment_ep = b_ep[0] + p_ep[0]
    last_month_payment_ep = b_ep[1] + p_ep[1]

    print("=== 等额本金 结果 ===")
    print(f"首月月供（商业）：{format_currency(b_ep[0])}，末月月供（商业）：{format_currency(b_ep[1])}")
    print(f"首月月供（公积金）：{format_currency(p_ep[0])}，末月月供（公积金）：{format_currency(p_ep[1])}")
    print(f"合计首月月供：{format_currency(first_month_payment_ep)}，合计末月月供：{format_currency(last_month_payment_ep)}")
    print(f"总还款：{format_currency(total_payment_ep)}，总利息：{format_currency(total_interest_ep)}\n")

    # 输出对比小结
    print("=== 对比小结 ===")
    print(f"等额本息每月均付：{format_currency(total_monthly_emi)}，总利息：{format_currency(total_interest_emi)}")
    print(f"等额本金首月付：{format_currency(first_month_payment_ep)}，末月付：{format_currency(last_month_payment_ep)}，总利息：{format_currency(total_interest_ep)}")


if __name__ == "__main__":
    run_calculation()