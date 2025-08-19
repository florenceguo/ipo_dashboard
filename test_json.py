# -*- coding: utf-8 -*-
import json

try:
    with open('excel_data_summary.json', 'r', encoding='utf-8') as f:
        data = json.load(f)
    print("✅ JSON文件有效！")
    print("包含工作表:", list(data.keys()))
    
    # 检查板块涨跌幅数据
    sector_data = data.get('板块涨跌幅', {}).get('data', [])
    print(f"板块涨跌幅数据行数: {len(sector_data)}")
    if sector_data:
        print("第一行数据:", sector_data[0])
        
except json.JSONDecodeError as e:
    print(f"❌ JSON解析错误: {e}")
except Exception as e:
    print(f"❌ 其他错误: {e}")

