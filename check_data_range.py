# -*- coding: utf-8 -*-
import json

# 检查数据时间范围
with open('excel_data_summary.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

print("=== 数据时间范围检查 ===")

# 检查板块涨跌幅数据
sector_data = data.get('板块涨跌幅', {}).get('data', [])
print(f"板块数据总数: {len(sector_data)}")

if sector_data:
    labels = [item['week_label'] for item in sector_data if item.get('week_label')]
    if labels:
        print(f"周标签范围: {min(labels)} 到 {max(labels)}")
        
        # 统计各年的数据
        year_counts = {}
        for label in labels:
            if len(label) >= 2:
                year = '20' + label[:2]
                year_counts[year] = year_counts.get(year, 0) + 1
        
        print("各年数据统计:")
        for year in sorted(year_counts.keys()):
            print(f"  {year}: {year_counts[year]} 条")
        
        # 显示2025年的数据样例
        data_2025 = [item for item in sector_data if item.get('week_label', '').startswith('25')]
        print(f"\n2025年数据条数: {len(data_2025)}")
        if data_2025:
            print("2025年数据样例:")
            for item in data_2025[:5]:
                print(f"  {item['week_label']}: 上证主板={item.get('上证主板')}, 科创板={item.get('科创板')}, 北交所={item.get('北交所')}")

# 检查其他数据表的时间范围
for sheet_name in ['周度收益', '发行统计', '中签率统计']:
    sheet_data = data.get(sheet_name, {}).get('data', [])
    if sheet_data:
        labels = [item.get('week_label') for item in sheet_data if item.get('week_label')]
        if labels:
            labels = [l for l in labels if l]  # 过滤空值
            print(f"\n{sheet_name} 时间范围: {min(labels)} 到 {max(labels)} (共{len(labels)}条)")

