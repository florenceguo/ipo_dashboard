# -*- coding: utf-8 -*-
import json

# 检查发行统计数据
with open('excel_data_summary.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

print("=== 发行统计数据检查 ===")
issuance_data = data.get('发行统计', {}).get('data', [])
print(f"数据条数: {len(issuance_data)}")

if issuance_data:
    print("\n前3条数据:")
    for i, item in enumerate(issuance_data[:3]):
        print(f"第{i+1}条: {item}")
    
    print(f"\n数据字段: {list(issuance_data[0].keys())}")
    
    # 检查募资金额数据
    total_funds = []
    for item in issuance_data:
        fund = item.get('total_raised_fund', 0)
        if fund and fund != 0:
            total_funds.append(fund)
    
    print(f"\n非零募资金额数据: {len(total_funds)} 条")
    if total_funds:
        print(f"募资金额范围: {min(total_funds):.2f} - {max(total_funds):.2f}")
        print(f"平均募资金额: {sum(total_funds)/len(total_funds):.2f}")
    
    # 计算总募资金额
    total_raised = sum(item.get('total_raised_fund', 0) for item in issuance_data)
    print(f"\n总募资金额: {total_raised:.2f}")
    print(f"总募资金额(亿): {total_raised/100000000:.2f}")