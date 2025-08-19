# -*- coding: utf-8 -*-
import json
import gzip

# 创建优化版本的数据文件
def optimize_data():
    # 读取原始数据
    with open('excel_data_summary.json', 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    # 创建压缩版本的数据
    optimized_data = {}
    
    # 只保留网站需要的核心数据，移除冗余字段
    for key, value in data.items():
        if 'data' in value:
            # 限制数据行数，提高加载速度
            if key == '原始数据':
                # 原始数据只保留最近50条记录
                value['data'] = value['data'][:50]
            
            # 移除不必要的字段
            if key == '北交所':
                # 保留所有北交所数据
                pass
            elif key == '周度收益':
                # 保留所有周度收益数据
                pass
            elif key == '中签率统计':
                # 保留所有中签率数据
                pass
            elif key == '发行统计':
                # 保留所有发行统计数据
                pass
            elif key == '板块涨跌幅':
                # 保留所有板块数据
                pass
            elif key == '市盈率统计':
                # 保留所有市盈率数据
                pass
        
        optimized_data[key] = value
    
    # 保存优化后的数据
    with open('excel_data_optimized.json', 'w', encoding='utf-8') as f:
        json.dump(optimized_data, f, ensure_ascii=False, separators=(',', ':'))
    
    # 创建压缩版本
    with open('excel_data_optimized.json', 'rb') as f_in:
        with gzip.open('excel_data_optimized.json.gz', 'wb') as f_out:
            f_out.writelines(f_in)
    
    # 输出文件大小对比
    import os
    original_size = os.path.getsize('excel_data_summary.json')
    optimized_size = os.path.getsize('excel_data_optimized.json')
    compressed_size = os.path.getsize('excel_data_optimized.json.gz')
    
    print(f"原始文件大小: {original_size:,} 字节 ({original_size/1024:.1f} KB)")
    print(f"优化文件大小: {optimized_size:,} 字节 ({optimized_size/1024:.1f} KB)")
    print(f"压缩文件大小: {compressed_size:,} 字节 ({compressed_size/1024:.1f} KB)")
    print(f"压缩比: {compressed_size/original_size:.1%}")

if __name__ == "__main__":
    optimize_data()

