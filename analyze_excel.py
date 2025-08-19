# -*- coding: utf-8 -*-
"""
Excel数据分析和转换脚本
专门用于处理ipoinfo.py生成的Excel文件，转换为网站可用的JSON格式
"""
import pandas as pd
import openpyxl
import json
import os
import glob
from datetime import datetime

def find_latest_excel_file():
    """查找最新的Excel文件"""
    # 查找所有可能的Excel文件
    patterns = [
        '新股周度统计*.xlsx',
        '*.xlsx'
    ]
    
    excel_files = []
    for pattern in patterns:
        excel_files.extend(glob.glob(pattern))
    
    if not excel_files:
        raise FileNotFoundError("未找到Excel文件！请先运行 ipoinfo.py 生成Excel文件。")
    
    # 按修改时间排序，返回最新的文件
    latest_file = max(excel_files, key=os.path.getmtime)
    print(f"找到最新Excel文件: {latest_file}")
    
    # 显示文件修改时间
    mod_time = datetime.fromtimestamp(os.path.getmtime(latest_file))
    print(f"文件修改时间: {mod_time.strftime('%Y-%m-%d %H:%M:%S')}")
    
    return latest_file

def clean_data(obj):
    """清理数据中的NaN值，转换为JSON兼容格式"""
    if isinstance(obj, dict):
        return {k: clean_data(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [clean_data(item) for item in obj]
    elif obj != obj:  # 检查NaN (NaN != NaN 为 True)
        return None
    elif obj == float('inf') or obj == float('-inf'):
        return None
    else:
        return obj

def convert_excel_to_json(excel_file):
    """将Excel文件转换为JSON格式"""
    try:
        # 使用openpyxl读取工作表名称
        wb = openpyxl.load_workbook(excel_file)
        sheet_names = wb.sheetnames
        print(f"发现工作表: {sheet_names}")
        
        # 读取每个工作表的数据
        data_summary = {}
        
        for sheet in sheet_names:
            try:
                print(f"\n处理工作表: {sheet}")
                df = pd.read_excel(excel_file, sheet_name=sheet)
                print(f"数据形状: {df.shape}")
                
                if len(df) > 0:
                    print(f"列名: {list(df.columns)}")
                    print("前3行数据:")
                    print(df.head(3))
                
                # 保存数据摘要
                data_summary[sheet] = {
                    'shape': df.shape,
                    'columns': list(df.columns),
                    'data': df.to_dict('records') if len(df) > 0 else []
                }
                
            except Exception as e:
                print(f"读取工作表 {sheet} 时出错: {e}")
                continue
        
        # 清理数据中的NaN值
        clean_data_summary = clean_data(data_summary)
        
        # 生成JSON文件名
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        json_filename = 'excel_data_summary.json'
        backup_filename = f'excel_data_backup_{timestamp}.json'
        
        # 如果已存在JSON文件，先备份
        if os.path.exists(json_filename):
            os.rename(json_filename, backup_filename)
            print(f"已备份旧数据文件为: {backup_filename}")
        
        # 保存新的JSON文件
        with open(json_filename, 'w', encoding='utf-8') as f:
            json.dump(clean_data_summary, f, ensure_ascii=False, indent=2, default=str)
        
        print(f"\n✅ 数据转换完成!")
        print(f"输出文件: {json_filename}")
        print(f"文件大小: {os.path.getsize(json_filename):,} 字节")
        
        # 显示数据统计
        print(f"\n📊 数据统计:")
        for sheet_name, sheet_data in clean_data_summary.items():
            if 'data' in sheet_data:
                record_count = len(sheet_data['data'])
                print(f"  {sheet_name}: {record_count} 条记录")
        
        return json_filename
        
    except Exception as e:
        print(f"❌ 转换过程中出错: {e}")
        raise

def main():
    """主函数"""
    print("=" * 50)
    print("📊 Excel数据转换工具")
    print("=" * 50)
    
    try:
        # 1. 查找最新的Excel文件
        excel_file = find_latest_excel_file()
        
        # 2. 转换为JSON
        json_file = convert_excel_to_json(excel_file)
        
        print(f"\n🎉 转换成功完成!")
        print(f"Excel文件: {excel_file}")
        print(f"JSON文件: {json_file}")
        print(f"\n💡 现在可以刷新网站查看最新数据!")
        
    except Exception as e:
        print(f"\n❌ 转换失败: {e}")
        print(f"\n📝 解决步骤:")
        print(f"1. 确保已运行 ipoinfo.py 生成Excel文件")
        print(f"2. 检查Excel文件是否损坏")
        print(f"3. 确保Python环境中有pandas和openpyxl库")

if __name__ == "__main__":
    main()