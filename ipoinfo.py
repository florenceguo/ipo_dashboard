# -*- coding: utf-8 -*-
"""
Created on Wed Aug 13 17:24:30 2025

@author: ogi1_
"""



import pandas as pd
import numpy as np
from WindPy import w
import matplotlib.pyplot as plt
import seaborn as sns
import matplotlib as mpl
import statsmodels.api as sm
from statsmodels.tsa.stattools import coint
from statsmodels.tsa.stattools import adfuller
from statsmodels.graphics.tsaplots import plot_acf, plot_pacf
import seaborn as sns
import matplotlib.pyplot as plt
import matplotlib as mpl
from datetime import datetime
from datetime import datetime
import openpyxl
import os

mpl.rcParams['font.sans-serif'] = ['SimHei']
mpl.rcParams['font.serif'] = ['SimHei']

# riskparity = riskparity()
sns.set_style("white",{"font.sans-serif":["simhei", "Arial"]})


# 创建标准化的周度索引
def get_week_number(date):
    year = date.year
    # 获取该年第一天
    first_day = pd.Timestamp(f'{year}-01-01')
    # 计算是第几周（从1开始）
    week_num = ((date - first_day).days // 7) + 1
    # 确保周数在1-52之间
    week_num = min(max(1, week_num), 52)
    return f'{str(year)[-2:]}W{week_num:02d}'
 
def update_ipodata(startdate = '2024-01-01',enddate = '2025-06-09',path = None,reconnect = True, ifplot = True):
   
    if reconnect:
        w.close()
        if not w.isconnected():
            w.start()
            
    ipostocks = w.wset("newstockissueinformation","startdate=%s;enddate=%s;datetype=online;board=all;field=wind_code,sec_name,purchase_code,prospectus_date,online_issue_date,listing_date,ipo_board,issue_price,ipo_pe,industry_pe,industrype_static,industrype_ttm,offline_amt,offline_issue_amt,limitofsubscriptionshare,limitofonlinepurchaseshare,online_dto_ratio,expected_raised_fund,actual_raised_fund,offline_btc_ratio,offline_os_ratio,offline_blocked_funds,windindustry,stocktype,exchange"%(startdate,enddate),usedf = True)[1]
    
   
    ipostocks['offline_maxbuyamt'] = ipostocks['issue_price']*ipostocks['limitofsubscriptionshare']*10000
    ipostocks['online_maxbuyamt'] = ipostocks['issue_price']*ipostocks['limitofonlinepurchaseshare']
    
    ipoinfo = ipostocks[['wind_code','sec_name','listing_date','ipo_board','windindustry','ipo_pe','industry_pe','actual_raised_fund','offline_maxbuyamt','online_maxbuyamt']]
    ##还需要获取A类中签率，B类中签率，上市首日涨跌幅
    stocks = ipoinfo.wind_code.unique().tolist()
    codestr = str(stocks)[1:-1].replace("'",'')
    lotteryrate_a = w.wss(codestr, "ipo_lotteryrate_abc","instituteType=1",usedf = True)[1]
    lotteryrate_b = w.wss(codestr, "ipo_lotteryrate_abc,ipo_pctchange","instituteType=2",usedf = True)[1]
    lotteryrate_online = w.wss(codestr,"ipo_cashratio",usedf = True)[1]
    
    ipo_sub = pd.concat([lotteryrate_a,lotteryrate_b,lotteryrate_online],axis = 1)
    ipo_sub.columns = ['lottery_a','lottery_b','pctchg','lottery_online']
    ipo_sub = ipo_sub/100
    
    ipoinfo = ipoinfo.set_index('wind_code').join(ipo_sub,how = 'left')
    
    ipoinfo['rtamt'] = ipoinfo.apply(lambda x: min(200000000,x.offline_maxbuyamt)*x.pctchg*x.lottery_b,axis = 1)
    ipoinfo['rtamt_bj'] = ipoinfo.apply(lambda x:min(10000000,x.online_maxbuyamt)*x.pctchg*x.lottery_online,axis = 1)
    ipoinfo['lottery_a2b'] = ipoinfo.lottery_a/ipoinfo.lottery_b
    
    # 将listing_date转换为datetime类型
    ipoinfo['listing_date'] = pd.to_datetime(ipoinfo['listing_date'])
    
    ipoinfo.dropna(subset = ['listing_date'],inplace = True)
    
    # 添加周度标签
    ipoinfo['week_label'] = ipoinfo['listing_date'].apply(get_week_number)
    ###北交所月度年化收益
    monthly_returns_bj = ipoinfo.groupby(ipoinfo.listing_date.dt.to_period('M')).rtamt_bj.sum()
    monthly_annualized_returns_bj = monthly_returns_bj*12/10000000
    # 按周汇总收益
    weekly_returns = ipoinfo.groupby('week_label')['rtamt'].sum()
    
    # 计算年化收益贡献
    weekly_annualized_returns = weekly_returns * 52 / 200000000
    
    # 过滤掉没有记录的周（值为0或NaN的周）
    weekly_annualized_returns = weekly_annualized_returns[weekly_annualized_returns != 0].dropna()
    
    # 获取需要显示的周标签（每月第一周）
    def is_first_week_of_month(week_label):
        year = int('20' + week_label[:2])
        week_num = int(week_label[3:])
        # 计算该周的第一天
        first_day = pd.Timestamp(f'{year}-01-01')
        week_start = first_day + pd.Timedelta(days=(week_num-1)*7)
        # 判断是否是每月的第一周
        return week_start.day <= 7
    
    # 获取需要显示的周标签
    display_weeks = [week for week in weekly_annualized_returns.index if is_first_week_of_month(week)]
    
    # 设置全局字体大小
    plt.rcParams['font.size'] = 10
    
    # 计算每周的中签率统计
    weekly_lottery_stats = ipoinfo.groupby('week_label').agg({
        'lottery_a': 'mean',
        'lottery_b': 'mean',
        'lottery_a2b': 'mean'
    })
    
    # 只保留有实际数据的周（剔除所有指标都为0的周）
    weekly_lottery_stats = weekly_lottery_stats[
        (weekly_lottery_stats['lottery_a'] != 0) | 
        (weekly_lottery_stats['lottery_b'] != 0) | 
        (weekly_lottery_stats['lottery_a2b'] != 0)
    ].dropna()
    
    # 获取需要显示的周标签（每月第一周）
    display_weeks_lottery = [week for week in weekly_lottery_stats.index if is_first_week_of_month(week)]
    
    # 计算每周的股票个数和募资总额
    weekly_stats = ipoinfo.groupby('week_label').agg({
        'sec_name': 'count',  # 使用sec_name来计数股票个数
        'actual_raised_fund': 'sum'  # 募资总额
    }).rename(columns={'sec_name': 'stock_count', 'actual_raised_fund': 'total_raised_fund'})
    
    # 只保留有实际数据的周
    weekly_stats = weekly_stats[weekly_stats['stock_count'] > 0]
    
    # 获取需要显示的周标签（每月第一周）
    display_weeks_stats = [week for week in weekly_stats.index if is_first_week_of_month(week)]
    
    # 添加月度标签
    ipoinfo['month_label'] = ipoinfo['listing_date'].dt.to_period('M')
    
    # 计算每个板块的月度平均涨跌幅
    monthly_board_returns = ipoinfo.groupby(['month_label', 'ipo_board'])['pctchg'].mean().unstack()
    
    # 只保留有实际数据的月份（剔除所有板块都为NaN的月份）
    monthly_board_returns = monthly_board_returns.dropna(how='all')
    
    # 对于每个板块，如果某个月没有数据，保持为NaN（不填充为0）
    # 这样在绘图时会自动跳过这些点
    
    # 计算每周的IPO市盈率和行业市盈率均值
    weekly_pe_stats = ipoinfo.groupby('week_label').agg({
        'ipo_pe': 'mean',
        'industry_pe': 'mean'
    })
    
    # 只保留有实际数据的周（剔除所有指标都为NaN的周）
    weekly_pe_stats = weekly_pe_stats.dropna(how='all')
    
    # 获取需要显示的周标签（每月第一周）
    display_weeks_pe = [week for week in weekly_pe_stats.index if is_first_week_of_month(week)]
    
    # 保存所有数据到Excel
    if path is None:
        fp = '新股周度统计.xlsx'
    else:
        fp = path
    with pd.ExcelWriter(fp, engine='openpyxl') as writer:
        monthly_annualized_returns_bj.to_frame('北交所年化收益贡献').to_excel(writer,sheet_name = '北交所')
        # 保存周度收益数据
        weekly_annualized_returns.to_frame('年化收益贡献').to_excel(writer, sheet_name='周度收益')
        
        # 保存中签率统计数据
        weekly_lottery_stats.to_excel(writer, sheet_name='中签率统计')
        
        # 保存股票个数和募资总额统计数据
        weekly_stats.to_excel(writer, sheet_name='发行统计')
        
        # 保存板块月度平均涨跌幅数据
        monthly_board_returns.to_excel(writer, sheet_name='板块涨跌幅')
        
        # 保存IPO市盈率和行业市盈率统计数据
        weekly_pe_stats.to_excel(writer, sheet_name='市盈率统计')
        
        # 保存原始数据
        ipoinfo.to_excel(writer, sheet_name='原始数据')
        
        # 保存收益贡献图
        if ifplot:
            plt.figure(figsize=(7, 5.5))  # 放大图表尺寸
            ax = plt.bar(range(len(weekly_annualized_returns)), weekly_annualized_returns.values)
            plt.title('新股周度年化收益贡献', fontsize=10)
            plt.ylabel('年化收益贡献', fontsize=10)
            plt.gca().yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x:.2%}'))
            for i, v in enumerate(weekly_annualized_returns.values):
                plt.text(i, v, f'{v:.1%}', ha='center', va='bottom', fontsize=8)
            plt.axhline(y=0.015, color='red', linestyle='--', label='悲观 (1.5%)')
            plt.axhline(y=0.035, color='orange', linestyle='--', label='中性 (3.5%)')
            plt.axhline(y=0.07, color='green', linestyle='--', label='乐观 (7%)')
            plt.legend(fontsize=8)
            plt.xticks(range(len(weekly_annualized_returns)), weekly_annualized_returns.index, rotation=90, fontsize=8)
            for i, label in enumerate(plt.gca().get_xticklabels()):
                if label.get_text() not in display_weeks:
                    label.set_visible(False)
            plt.tight_layout()
            plt.gcf().canvas.draw()  # 强制刷新
            plt.savefig('temp_returns.png', dpi=300, bbox_inches='tight')
            img1 = openpyxl.drawing.image.Image('temp_returns.png')
            img1.width = 600  # 放大图片宽度
            img1.height = 450  # 放大图片高度
            worksheet = writer.sheets['周度收益']
            worksheet.add_image(img1, 'A10')
            plt.close()
        
        # 收益率贡献图北交所
        if ifplot:
            plt.figure(figsize=(7, 5.5))  # 放大图表尺寸
            ax11 = plt.bar(range(len(monthly_annualized_returns_bj)), monthly_annualized_returns_bj.values)
            plt.title('北交所月度年化收益贡献', fontsize=10)
            plt.ylabel('年化收益贡献', fontsize=10)
            plt.gca().yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x:.2%}'))
            for i, v in enumerate(monthly_annualized_returns_bj.values):
                plt.text(i, v, f'{v:.1%}', ha='center', va='bottom', fontsize=8)
            plt.axhline(y=0.01, color='red', linestyle='--', label='悲观 (1.0%)')
            plt.axhline(y=0.025, color='orange', linestyle='--', label='中性 (2.5%)')
            plt.axhline(y=0.035, color='green', linestyle='--', label='乐观 (3.5%)')
            plt.legend(fontsize=8)
            plt.xticks(range(len(monthly_annualized_returns_bj)), monthly_annualized_returns_bj.index, rotation=90, fontsize=8)
           
            plt.tight_layout()
            plt.gcf().canvas.draw()  # 强制刷新
            plt.savefig('temp_returns_bj.png', dpi=300, bbox_inches='tight')
            img11 = openpyxl.drawing.image.Image('temp_returns_bj.png')
            img11.width = 600  # 放大图片宽度
            img11.height = 450  # 放大图片高度
            worksheet = writer.sheets['北交所']
            worksheet.add_image(img11, 'A10')
            plt.close()
        
        
        
        # 保存中签率统计图
        if ifplot:
            plt.figure(figsize=(7, 5.5))  # 放大图表尺寸
            ax1 = plt.gca()
            ax1.plot(range(len(weekly_lottery_stats)), weekly_lottery_stats['lottery_a'], 'b-', label='A类中签率', linewidth=1.5)
            ax1.plot(range(len(weekly_lottery_stats)), weekly_lottery_stats['lottery_b'], 'r-', label='B类中签率', linewidth=1.5)
            ax1.set_ylabel('中签率', fontsize=10)
            ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x*10000:.0f}%%'))
            ax1.tick_params(axis='y', labelsize=8)
            ax1.legend(loc='upper left', fontsize=8)
            ax2 = ax1.twinx()
            ax2.bar(range(len(weekly_lottery_stats)), weekly_lottery_stats['lottery_a2b'], alpha=0.3, color='g', label='A/B类中签率比')
            ax2.set_ylabel('A/B类中签率比', fontsize=10)
            ax2.tick_params(axis='y', labelsize=8)
            ax2.legend(loc='upper right', fontsize=8)
            plt.title('新股周度中签率统计', fontsize=10)
            plt.xticks(range(len(weekly_lottery_stats)), weekly_lottery_stats.index, fontsize=8)
            for i, label in enumerate(ax1.get_xticklabels()):
                if label.get_text() not in display_weeks_lottery:
                    label.set_visible(False)
            plt.tight_layout()
            # 手动旋转所有x轴标签
            for label in ax1.get_xticklabels():
                label.set_rotation(90)
            for label in ax2.get_xticklabels():
                label.set_rotation(90)
            plt.gcf().canvas.draw()
            plt.savefig('temp_lottery.png', dpi=300, bbox_inches='tight')
            img2 = openpyxl.drawing.image.Image('temp_lottery.png')
            img2.width = 600
            img2.height = 450
            worksheet = writer.sheets['中签率统计']
            worksheet.add_image(img2, 'A10')
            plt.show()
            plt.close()
    
        # 保存股票个数和募资总额统计图
        if ifplot:
            plt.figure(figsize=(7, 5.5))  # 放大图表尺寸
            ax1 = plt.gca()
            ax1.bar(range(len(weekly_stats)), weekly_stats['stock_count'], alpha=0.6, color='blue', label='股票个数')
            ax1.set_ylabel('股票个数', fontsize=10)
            ax1.tick_params(axis='y', labelsize=8)
            ax2 = ax1.twinx()
            ax2.plot(range(len(weekly_stats)), weekly_stats['total_raised_fund']/100000000, 'r-', linewidth=1.5, label='募资总额')
            ax2.set_ylabel('募资总额（亿元）', fontsize=10)
            ax2.tick_params(axis='y', labelsize=8)
            lines1, labels1 = ax1.get_legend_handles_labels()
            lines2, labels2 = ax2.get_legend_handles_labels()
            ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper left', fontsize=8)
            plt.title('新股周度发行统计', fontsize=10)
            plt.xticks(range(len(weekly_stats)), weekly_stats.index, fontsize=8)
            for i, label in enumerate(ax1.get_xticklabels()):
                if label.get_text() not in display_weeks_stats:
                    label.set_visible(False)
            plt.tight_layout()
            # 手动旋转所有x轴标签
            for label in ax1.get_xticklabels():
                label.set_rotation(90)
            for label in ax2.get_xticklabels():
                label.set_rotation(90)
            plt.gcf().canvas.draw()
            plt.savefig('temp_stats.png', dpi=300, bbox_inches='tight')
            img3 = openpyxl.drawing.image.Image('temp_stats.png')
            img3.width = 600
            img3.height = 450
            worksheet = writer.sheets['发行统计']
            worksheet.add_image(img3, 'A10')
            plt.show()
            plt.close()
    
        # 保存板块月度平均涨跌幅图
        if ifplot:
            plt.figure(figsize=(7, 5.5))  # 放大图表尺寸
            
            # 为每个板块绘制线图，自动跳过NaN值
            for board in monthly_board_returns.columns:
                # 获取该板块的数据，过滤掉NaN值
                board_data = monthly_board_returns[board].dropna()
                if len(board_data) > 0:  # 只有当板块有数据时才绘制
                    x_positions = [monthly_board_returns.index.get_loc(idx) for idx in board_data.index]
                    plt.plot(x_positions, board_data.values, 'o-', label=board, linewidth=2, markersize=4)
            
            plt.title('各板块月度平均涨跌幅趋势', fontsize=10)
            plt.ylabel('平均涨跌幅', fontsize=10)
            plt.gca().yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x:.2%}'))
            plt.legend(loc='upper left', fontsize=8)
            plt.xticks(range(len(monthly_board_returns.index)), 
                      [str(period) for period in monthly_board_returns.index], 
                      rotation=45, fontsize=8)
            plt.grid(True, linestyle='--', alpha=0.3, axis='y')
            plt.tight_layout()
            plt.gcf().canvas.draw()
            plt.savefig('temp_board_returns.png', dpi=300, bbox_inches='tight')
            img4 = openpyxl.drawing.image.Image('temp_board_returns.png')
            img4.width = 600
            img4.height = 450
            worksheet = writer.sheets['板块涨跌幅']
            worksheet.add_image(img4, 'A10')
            plt.show()
            plt.close()
        
        # 保存IPO市盈率和行业市盈率统计图
        if ifplot:
            plt.figure(figsize=(7, 5.5))  # 放大图表尺寸
            ax1 = plt.gca()
            ax1.plot(range(len(weekly_pe_stats)), weekly_pe_stats['ipo_pe'], 'b-', label='IPO市盈率', linewidth=1.5)
            ax1.plot(range(len(weekly_pe_stats)), weekly_pe_stats['industry_pe'], 'r-', label='行业市盈率', linewidth=1.5)
            ax1.set_ylabel('市盈率', fontsize=10)
            ax1.tick_params(axis='y', labelsize=8)
            ax1.legend(loc='upper left', fontsize=8)
            plt.title('新股周度市盈率统计', fontsize=10)
            plt.xticks(range(len(weekly_pe_stats)), weekly_pe_stats.index, rotation=90, fontsize=8)
            for i, label in enumerate(ax1.get_xticklabels()):
                if label.get_text() not in display_weeks_pe:
                    label.set_visible(False)
            plt.grid(True, linestyle='--', alpha=0.3, axis='y')
            plt.tight_layout()
            plt.gcf().canvas.draw()
            plt.savefig('temp_pe_stats.png', dpi=300, bbox_inches='tight')
            img5 = openpyxl.drawing.image.Image('temp_pe_stats.png')
            img5.width = 600
            img5.height = 450
            worksheet = writer.sheets['市盈率统计']
            worksheet.add_image(img5, 'A10')
            plt.close()
    
    # 删除临时图片文件
    if ifplot:
        try:
            os.remove('temp_returns.png')
            os.remove('temp_lottery.png')
            os.remove('temp_stats.png')
            os.remove('temp_board_returns.png')
            os.remove('temp_pe_stats.png')
            os.remove('temp_returns_bj.png')
        except FileNotFoundError:
            pass  # 如果文件不存在，忽略错误
    return ipoinfo

##不同规模，不同板块的打新收益测算
def rt_estimation(ipoinfo,startdate,enddate,aum = 120000000,board = [],rf = 0.014):
    '''
    board = ['上证主板', '科创板']
    '''
    

    start_date = datetime.strptime(startdate, '%Y-%m-%d')
    end_date = datetime.strptime(enddate, '%Y-%m-%d')
    delta = (end_date - start_date).days
    ipoinfo_sub = ipoinfo[(ipoinfo.listing_date>=startdate) & (ipoinfo.listing_date<=enddate)]
    ipoinfo_sub['rtamt'] = ipoinfo_sub.apply(lambda x: min(aum,x.offline_maxbuyamt)*x.pctchg*x.lottery_b,axis = 1)
    if len(board)>0:
        ipoinfo_sub = ipoinfo_sub[ipoinfo_sub.ipo_board.isin(board)]
    iport = (365/delta)*(ipoinfo_sub.rtamt.sum()/aum)
    freecash_rt = rf*(aum-1.2*80000000)/aum
    tot_rt = iport+freecash_rt
    return iport,freecash_rt,tot_rt

if __name__ == "__main__":
    startdate = '2024-03-31'
    enddate = '2025-08-18'
    fpath = r'C:\quant\ai\IPO数据相关\新股周度统计%s.xlsx'%enddate
    
    # 设置是否生成图表 - 您可以修改这个参数
    ifplot = False  # True: 生成图表和Excel, False: 只生成Excel数据
    
    ipoinfo = update_ipodata(startdate = startdate,enddate = enddate,path = fpath,reconnect = True, ifplot = ifplot)
    ipoinfo['yearmon'] = ipoinfo.listing_date.dt.year*100+ipoinfo.listing_date.dt.month
    rt_stats = ipoinfo.groupby(['yearmon','ipo_board']).pctchg.mean()
    a = rt_stats.to_frame("首日涨跌幅").unstack('ipo_board')
    
    #iport,freecash_rt,tot_rt = rt_estimation(ipoinfo,'2025-01-01','2025-07-21',aum = 220000000,board = ['上证主板', '科创板'])