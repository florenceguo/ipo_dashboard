// 全局变量存储数据
let excelData = {};
let charts = {};

// 页面加载完成后初始化
document.addEventListener('DOMContentLoaded', function() {
    loadExcelData();
    initializeDateInputs();
});

// 初始化日期输入框
function initializeDateInputs() {
    const today = new Date();
    const oneYearAgo = new Date(today.getFullYear() - 1, today.getMonth(), today.getDate());
    
    document.getElementById('startDate').value = oneYearAgo.toISOString().split('T')[0];
    document.getElementById('endDate').value = today.toISOString().split('T')[0];
}

// 加载Excel数据
async function loadExcelData() {
    try {
        // 添加时间戳避免缓存问题
        const timestamp = new Date().getTime();
        const response = await fetch(`excel_data_summary.json?t=${timestamp}`);
        if (!response.ok) {
            throw new Error('无法加载数据文件');
        }
        
        excelData = await response.json();
        console.log('数据加载成功:', excelData);
        
        // 隐藏加载界面，显示内容
        document.getElementById('loading').style.display = 'none';
        document.getElementById('content').style.display = 'block';
        
        // 初始化所有图表和统计
        initializeCharts();
        updateStatistics();
        generateMarketInsights();
        
    } catch (error) {
        console.error('加载数据失败:', error);
        document.getElementById('loading').innerHTML = `
            <div class="error">
                ❌ 数据加载失败: ${error.message}<br>
                请确保 excel_data_summary.json 文件存在且格式正确
            </div>
        `;
    }
}

// 初始化所有图表
function initializeCharts() {
    createWeeklyReturnsChart();
    createBeijingExchangeChart();
    createLotteryRateChart();
    createIssuanceChart();
    createSectorPerformanceChart();
    createPeRatioChart();
}

// 创建新股周度年化收益图表
function createWeeklyReturnsChart() {
    const ctx = document.getElementById('weeklyReturnsChart').getContext('2d');
    const data = excelData['周度收益']?.data || [];
    
    const labels = data.map(item => item.week_label);
    const returns = data.map(item => (item['年化收益贡献'] * 100).toFixed(2));
    
    charts.weeklyReturns = new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [{
                label: '年化收益贡献 (%)',
                data: returns,
                borderColor: '#3498db',
                backgroundColor: 'rgba(52, 152, 219, 0.1)',
                fill: true,
                tension: 0.4,
                pointBackgroundColor: '#3498db',
                pointBorderColor: '#ffffff',
                pointBorderWidth: 2,
                pointRadius: 4
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: false
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        callback: function(value) {
                            return value + '%';
                        }
                    }
                },
                x: {
                    display: true,
                    ticks: {
                        maxTicksLimit: 10
                    }
                }
            },
            elements: {
                point: {
                    hoverRadius: 6
                }
            }
        }
    });
}

// 创建北交所月度年化收益图表
function createBeijingExchangeChart() {
    const ctx = document.getElementById('beijingExchangeChart').getContext('2d');
    const data = excelData['北交所']?.data || [];
    
    const labels = data.map(item => {
        const date = new Date(item.listing_date);
        return `${date.getFullYear()}年${date.getMonth() + 1}月`;
    });
    const returns = data.map(item => (item['北交所年化收益贡献'] * 100).toFixed(2));
    
    charts.beijingExchange = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: '年化收益贡献 (%)',
                data: returns,
                backgroundColor: 'rgba(230, 126, 34, 0.8)',
                borderColor: '#e67e22',
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: false
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        callback: function(value) {
                            return value + '%';
                        }
                    }
                }
            }
        }
    });
}

// 创建中签率统计图表
function createLotteryRateChart() {
    const ctx = document.getElementById('lotteryRateChart').getContext('2d');
    const data = excelData['中签率统计']?.data || [];
    
    const labels = data.map(item => item.week_label);
    const lotteryA = data.map(item => (item.lottery_a * 10000).toFixed(3));
    const lotteryB = data.map(item => (item.lottery_b * 10000).toFixed(3));
    
    charts.lotteryRate = new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [
                {
                    label: 'A类中签率 (‰)',
                    data: lotteryA,
                    borderColor: '#e74c3c',
                    backgroundColor: 'rgba(231, 76, 60, 0.1)',
                    fill: false,
                    tension: 0.4
                },
                {
                    label: 'B类中签率 (‰)',
                    data: lotteryB,
                    borderColor: '#27ae60',
                    backgroundColor: 'rgba(39, 174, 96, 0.1)',
                    fill: false,
                    tension: 0.4
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'top'
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        callback: function(value) {
                            return value + '‰';
                        }
                    }
                },
                x: {
                    ticks: {
                        maxTicksLimit: 8
                    }
                }
            }
        }
    });
}

// 创建发行统计图表
function createIssuanceChart() {
    const ctx = document.getElementById('issuanceChart').getContext('2d');
    const data = excelData['发行统计']?.data || [];
    
    console.log('发行统计图表数据:', data.slice(0, 3)); // 调试信息
    
    const labels = data.map(item => item.week_label);
    const stockCount = data.map(item => item.stock_count || 0);
    const raisedFund = data.map(item => {
        const fund = item.total_raised_fund || 0;
        return parseFloat(fund.toFixed(2)); // 数据已经是亿为单位，直接使用
    });
    
    console.log('处理后的数据:', {
        labels: labels.slice(0, 5),
        stockCount: stockCount.slice(0, 5),
        raisedFund: raisedFund.slice(0, 5)
    });
    
    // 计算募资金额的合适范围，过滤掉0值
    const nonZeroRaised = raisedFund.filter(val => val > 0);
    let raisedMin = 0, raisedMax = 10;
    
    if (nonZeroRaised.length > 0) {
        const maxRaised = Math.max(...nonZeroRaised);
        const minRaised = Math.min(...nonZeroRaised);
        const raisedRange = maxRaised - minRaised;
        raisedMin = Math.max(0, minRaised - raisedRange * 0.1);
        raisedMax = maxRaised + raisedRange * 0.1;
    }
    
    charts.issuance = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [
                {
                    label: '发行数量',
                    data: stockCount,
                    backgroundColor: 'rgba(155, 89, 182, 0.8)',
                    borderColor: '#9b59b6',
                    borderWidth: 1,
                    yAxisID: 'y'
                },
                {
                    label: '募资金额 (亿)',
                    data: raisedFund,
                    type: 'line',
                    borderColor: '#f39c12',
                    backgroundColor: 'rgba(243, 156, 18, 0.3)',
                    fill: false,
                    tension: 0.4,
                    borderWidth: 3,
                    pointRadius: 5,
                    pointBackgroundColor: '#f39c12',
                    pointBorderColor: '#ffffff',
                    pointBorderWidth: 2,
                    yAxisID: 'y1'
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: {
                mode: 'index',
                intersect: false,
            },
            plugins: {
                legend: {
                    position: 'top',
                    labels: {
                        usePointStyle: true,
                        padding: 20
                    }
                },
                tooltip: {
                    callbacks: {
                        afterLabel: function(context) {
                            if (context.datasetIndex === 1) {
                                return '募资总额: ' + context.parsed.y + '亿元';
                            } else {
                                return '发行数量: ' + context.parsed.y + '只';
                            }
                        }
                    }
                }
            },
            scales: {
                x: {
                    ticks: {
                        maxTicksLimit: 10
                    }
                },
                y: {
                    type: 'linear',
                    display: true,
                    position: 'left',
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: '发行数量 (只)',
                        color: '#9b59b6',
                        font: {
                            weight: 'bold'
                        }
                    },
                    ticks: {
                        color: '#9b59b6'
                    }
                },
                y1: {
                    type: 'linear',
                    display: true,
                    position: 'right',
                    min: raisedMin,
                    max: raisedMax,
                    title: {
                        display: true,
                        text: '募资金额 (亿元)',
                        color: '#f39c12',
                        font: {
                            weight: 'bold'
                        }
                    },
                    ticks: {
                        color: '#f39c12',
                        callback: function(value) {
                            return value.toFixed(1) + '亿';
                        }
                    },
                    grid: {
                        drawOnChartArea: false,
                    },
                }
            }
        }
    });
}

// 创建板块表现图表
function createSectorPerformanceChart() {
    const ctx = document.getElementById('sectorPerformanceChart').getContext('2d');
    const rawData = excelData['板块涨跌幅']?.data || [];
    
    console.log('板块涨跌幅原始数据:', rawData.slice(0, 5));
    
    // 检查数据时间范围
    const allWeekLabels = rawData.map(item => item.week_label).filter(label => label);
    const uniqueYears = [...new Set(allWeekLabels.map(label => {
        const yearSuffix = label.substring(0, 2);
        return 2000 + parseInt(yearSuffix);
    }))].sort();
    console.log('数据年份范围:', uniqueYears);
    console.log('周标签范围:', Math.min(...allWeekLabels.map(l => l)), '到', Math.max(...allWeekLabels.map(l => l)));
    
    // 将周度数据转换为月度数据
    const monthlyData = {};
    const sectors = ['上证主板', '深证主板', '科创板', '创业板', '北交所'];
    
    rawData.forEach(item => {
        if (!item.week_label) return;
        
        // 从周标签提取年月 (例如: 24W13 -> 2024-03, 25W13 -> 2025-03)
        const yearSuffix = item.week_label.substring(0, 2);
        const week = parseInt(item.week_label.substring(3));
        
        // 处理年份：24 -> 2024, 25 -> 2025, etc.
        const year = 2000 + parseInt(yearSuffix);
        
        // 估算月份 (粗略计算: 周数/4.3)
        const month = Math.min(12, Math.max(1, Math.ceil(week / 4.3)));
        const monthKey = `${year}-${month.toString().padStart(2, '0')}`;
        
        if (!monthlyData[monthKey]) {
            monthlyData[monthKey] = {};
            sectors.forEach(sector => {
                monthlyData[monthKey][sector] = [];
            });
        }
        
        sectors.forEach(sector => {
            const value = item[sector];
            if (value !== null && value !== undefined && !isNaN(value)) {
                monthlyData[monthKey][sector].push(value);
            }
        });
    });
    
    // 过滤出有数据的月份，并计算每月平均值
    const validMonths = [];
    const validMonthlyData = {};
    
    Object.keys(monthlyData).sort().forEach(month => {
        // 检查这个月是否至少有一个板块有数据
        const hasData = sectors.some(sector => monthlyData[month][sector].length > 0);
        
        if (hasData) {
            validMonths.push(month);
            validMonthlyData[month] = monthlyData[month];
        }
    });
    
    const datasets = [];
    const colors = ['#FF6384', '#36A2EB', '#FFCE56', '#4BC0C0', '#9966FF'];
    
    sectors.forEach((sector, index) => {
        const monthlyAvgs = validMonths.map(month => {
            const values = validMonthlyData[month][sector];
            // 如果该板块在这个月没有数据，返回 null（不显示点）
            if (values.length === 0) return null;
            
            const avg = values.reduce((a, b) => a + b, 0) / values.length * 100;
            // 只有当平均值不为0时才返回数值，否则返回null
            return avg !== 0 ? parseFloat(avg.toFixed(2)) : null;
        });
        
        // 只有当板块有至少一个非空非零数据点时才添加到图表中
        const hasValidData = monthlyAvgs.some(val => val !== null && val !== 0);
        
        if (hasValidData) {
            datasets.push({
                label: sector,
                data: monthlyAvgs,
                borderColor: colors[index],
                backgroundColor: colors[index] + '20',
                fill: false,
                tension: 0.4,
                borderWidth: 3,
                pointRadius: function(context) {
                    // 只有当数据点不为null时才显示点
                    return context.parsed.y !== null ? 5 : 0;
                },
                pointBackgroundColor: colors[index],
                pointBorderColor: '#ffffff',
                pointBorderWidth: 2,
                spanGaps: true, // 允许跳过空值，连接前后有效点
                segment: {
                    // 自定义线段样式，跳过null值时不显示线段
                    borderDash: function(ctx) {
                        return ctx.p0.skip || ctx.p1.skip ? [5, 5] : undefined;
                    }
                }
            });
        }
    });
    
    // 格式化月份标签显示
    const formattedMonths = validMonths.map(month => {
        const [year, monthNum] = month.split('-');
        return `${year}年${parseInt(monthNum)}月`;
    });
    
    console.log('处理后的月度数据:', {
        totalMonths: formattedMonths.length,
        months: formattedMonths,
        datasets: datasets.map(d => ({
            label: d.label, 
            dataPoints: d.data.filter(v => v !== null).length,
            sampleData: d.data.slice(0, 5)
        }))
    });
    
    // 额外调试：显示月度数据的详细信息
    console.log('月度数据详情:', Object.keys(validMonthlyData).sort());
    
    charts.sectorPerformance = new Chart(ctx, {
        type: 'line',
        data: {
            labels: formattedMonths,
            datasets: datasets
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: {
                mode: 'index',
                intersect: false,
            },
            plugins: {
                title: {
                    display: true,
                    text: '各板块月度平均涨跌幅趋势',
                    font: {
                        size: 14,
                        weight: 'bold'
                    }
                },
                legend: {
                    position: 'top',
                    labels: {
                        usePointStyle: true,
                        padding: 15
                    }
                },
                tooltip: {
                    filter: function(tooltipItem) {
                        // 只显示有数据的点的tooltip
                        return tooltipItem.parsed.y !== null;
                    },
                    callbacks: {
                        title: function(tooltipItems) {
                            return tooltipItems[0].label;
                        },
                        label: function(context) {
                            if (context.parsed.y !== null) {
                                return context.dataset.label + ': ' + context.parsed.y + '%';
                            }
                            return null; // 不显示null值的tooltip
                        }
                    }
                }
            },
            scales: {
                x: {
                    display: true,
                    title: {
                        display: true,
                        text: '时间'
                    },
                    ticks: {
                        maxTicksLimit: 8
                    }
                },
                y: {
                    display: true,
                    title: {
                        display: true,
                        text: '平均涨跌幅 (%)'
                    },
                    ticks: {
                        callback: function(value) {
                            return value + '%';
                        }
                    },
                    grid: {
                        color: 'rgba(0,0,0,0.1)'
                    }
                }
            },
            elements: {
                point: {
                    hoverRadius: 8
                },
                line: {
                    borderJoinStyle: 'round'
                }
            }
        }
    });
}

// 创建市盈率统计图表
function createPeRatioChart() {
    const ctx = document.getElementById('peRatioChart').getContext('2d');
    const data = excelData['市盈率统计']?.data || [];
    
    const labels = data.map(item => item.week_label);
    const ipoPe = data.map(item => item.ipo_pe ? item.ipo_pe.toFixed(1) : null);
    const industryPe = data.map(item => item.industry_pe ? item.industry_pe.toFixed(1) : null);
    
    charts.peRatio = new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [
                {
                    label: 'IPO市盈率',
                    data: ipoPe,
                    borderColor: '#8e44ad',
                    backgroundColor: 'rgba(142, 68, 173, 0.1)',
                    fill: false,
                    tension: 0.4,
                    spanGaps: true
                },
                {
                    label: '行业市盈率',
                    data: industryPe,
                    borderColor: '#34495e',
                    backgroundColor: 'rgba(52, 73, 94, 0.1)',
                    fill: false,
                    tension: 0.4,
                    spanGaps: true
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'top'
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: '市盈率'
                    }
                },
                x: {
                    ticks: {
                        maxTicksLimit: 8
                    }
                }
            }
        }
    });
}

// 更新统计数据
function updateStatistics() {
    // 周度收益统计
    const weeklyData = excelData['周度收益']?.data || [];
    if (weeklyData.length > 0) {
        const returns = weeklyData.map(item => item['年化收益贡献']);
        const avgReturn = (returns.reduce((a, b) => a + b, 0) / returns.length * 100).toFixed(2);
        const maxReturn = (Math.max(...returns) * 100).toFixed(2);
        
        document.getElementById('avgWeeklyReturn').textContent = avgReturn + '%';
        document.getElementById('maxWeeklyReturn').textContent = maxReturn + '%';
    }
    
    // 北交所统计
    const bjData = excelData['北交所']?.data || [];
    if (bjData.length > 0) {
        const returns = bjData.map(item => item['北交所年化收益贡献']);
        const avgReturn = (returns.reduce((a, b) => a + b, 0) / returns.length * 100).toFixed(2);
        const totalReturn = (returns.reduce((a, b) => a + b, 0) * 100).toFixed(2);
        
        document.getElementById('avgBjReturn').textContent = avgReturn + '%';
        document.getElementById('totalBjReturn').textContent = totalReturn + '%';
    }
    
    // 中签率统计
    const lotteryData = excelData['中签率统计']?.data || [];
    if (lotteryData.length > 0) {
        const lotteryA = lotteryData.map(item => item.lottery_a);
        const lotteryB = lotteryData.map(item => item.lottery_b);
        const avgLotteryA = (lotteryA.reduce((a, b) => a + b, 0) / lotteryA.length * 10000).toFixed(3);
        const avgLotteryB = (lotteryB.reduce((a, b) => a + b, 0) / lotteryB.length * 10000).toFixed(3);
        
        document.getElementById('avgLotteryA').textContent = avgLotteryA + '‰';
        document.getElementById('avgLotteryB').textContent = avgLotteryB + '‰';
    }
    
    // 发行统计
    const issuanceData = excelData['发行统计']?.data || [];
    if (issuanceData.length > 0) {
        const totalStocks = issuanceData.reduce((a, b) => a + (b.stock_count || 0), 0);
        const totalRaised = issuanceData.reduce((a, b) => a + (b.total_raised_fund || 0), 0); // 数据已经是亿为单位
        
        document.getElementById('totalStocks').textContent = totalStocks;
        document.getElementById('totalRaised').textContent = totalRaised.toFixed(1) + '亿';
        
        console.log('发行统计数据:', {
            totalStocks,
            totalRaised: totalRaised.toFixed(1),
            dataLength: issuanceData.length,
            sampleData: issuanceData.slice(0, 3)
        });
    }
    
    // 板块表现统计
    const sectorData = excelData['板块涨跌幅']?.data || [];
    if (sectorData.length > 0) {
        const sectors = ['上证主板', '深证主板', '科创板', '创业板', '北交所'];
        let bestSector = '';
        let bestReturn = 0;
        
        sectors.forEach(sector => {
            const validReturns = sectorData
                .map(item => item[sector])
                .filter(val => val !== null && val !== undefined && !isNaN(val));
            
            if (validReturns.length > 0) {
                const avgReturn = validReturns.reduce((a, b) => a + b, 0) / validReturns.length;
                if (avgReturn > bestReturn) {
                    bestReturn = avgReturn;
                    bestSector = sector;
                }
            }
        });
        
        document.getElementById('bestSector').textContent = bestSector;
        document.getElementById('bestSectorReturn').textContent = (bestReturn * 100).toFixed(2) + '%';
    }
    
    // 市盈率统计
    const peData = excelData['市盈率统计']?.data || [];
    if (peData.length > 0) {
        const ipoPeValues = peData.map(item => item.ipo_pe).filter(val => val !== null && val !== undefined);
        const industryPeValues = peData.map(item => item.industry_pe).filter(val => val !== null && val !== undefined);
        
        if (ipoPeValues.length > 0) {
            const avgIpoPe = (ipoPeValues.reduce((a, b) => a + b, 0) / ipoPeValues.length).toFixed(1);
            document.getElementById('avgIpoPe').textContent = avgIpoPe;
        }
        
        if (industryPeValues.length > 0) {
            const avgIndustryPe = (industryPeValues.reduce((a, b) => a + b, 0) / industryPeValues.length).toFixed(1);
            document.getElementById('avgIndustryPe').textContent = avgIndustryPe;
        }
    }
}

// 多选下拉框相关函数
function toggleDropdown() {
    const options = document.getElementById('multiSelectOptions');
    const header = document.querySelector('.multi-select-header');
    const arrow = document.querySelector('.arrow');
    
    if (options.classList.contains('show')) {
        options.classList.remove('show');
        header.classList.remove('active');
        arrow.classList.remove('rotated');
    } else {
        options.classList.add('show');
        header.classList.add('active');
        arrow.classList.add('rotated');
    }
}

function toggleAll() {
    const allCheckbox = document.getElementById('board_all');
    const boardCheckboxes = document.querySelectorAll('[id^="board_"]:not(#board_all)');
    
    boardCheckboxes.forEach(checkbox => {
        checkbox.checked = allCheckbox.checked;
    });
    
    updateSelection();
}

function updateSelection() {
    const boardCheckboxes = document.querySelectorAll('[id^="board_"]:not(#board_all)');
    const allCheckbox = document.getElementById('board_all');
    const selectedText = document.getElementById('selectedText');
    
    const checkedBoxes = Array.from(boardCheckboxes).filter(cb => cb.checked);
    
    // 更新全选复选框状态
    if (checkedBoxes.length === boardCheckboxes.length) {
        allCheckbox.checked = true;
        allCheckbox.indeterminate = false;
    } else if (checkedBoxes.length === 0) {
        allCheckbox.checked = false;
        allCheckbox.indeterminate = false;
    } else {
        allCheckbox.checked = false;
        allCheckbox.indeterminate = true;
    }
    
    // 更新显示文本
    if (checkedBoxes.length === 0) {
        selectedText.textContent = '请选择板块';
    } else if (checkedBoxes.length === boardCheckboxes.length) {
        selectedText.textContent = '全部板块';
    } else if (checkedBoxes.length === 1) {
        selectedText.textContent = checkedBoxes[0].value;
    } else {
        selectedText.textContent = `已选择 ${checkedBoxes.length} 个板块`;
    }
}

// 获取选中的板块
function getSelectedBoards() {
    const boardCheckboxes = document.querySelectorAll('[id^="board_"]:not(#board_all)');
    return Array.from(boardCheckboxes)
        .filter(cb => cb.checked)
        .map(cb => cb.value);
}

// 点击页面其他地方关闭下拉框
document.addEventListener('click', function(event) {
    const container = document.querySelector('.multi-select-container');
    if (container && !container.contains(event.target)) {
        const options = document.getElementById('multiSelectOptions');
        const header = document.querySelector('.multi-select-header');
        const arrow = document.querySelector('.arrow');
        
        options.classList.remove('show');
        header.classList.remove('active');
        arrow.classList.remove('rotated');
    }
});

// 计算打新收益
function calculateReturns() {
    const netAssets = parseFloat(document.getElementById('netAssets').value) * 10000; // 转换为元
    const riskFreeRate = parseFloat(document.getElementById('riskFreeRate').value) / 100;
    const startDate = document.getElementById('startDate').value;
    const endDate = document.getElementById('endDate').value;
    const selectedBoards = getSelectedBoards();
    
    if (!netAssets || !riskFreeRate || !startDate || !endDate) {
        alert('请填写完整的计算参数！');
        return;
    }
    
    try {
        // 从原始数据中筛选符合条件的记录
        const rawData = excelData['原始数据']?.data || [];
        const filteredData = rawData.filter(item => {
            const listingDate = new Date(item.listing_date);
            const start = new Date(startDate);
            const end = new Date(endDate);
            
            return listingDate >= start && 
                   listingDate <= end && 
                   (selectedBoards.length === 0 || selectedBoards.includes(item.ipo_board));
        });
        
        if (filteredData.length === 0) {
            alert('在指定条件下没有找到相关新股数据！');
            return;
        }
        
        // 计算打新收益
        let totalReturn = 0;
        let validCount = 0;
        
        filteredData.forEach(item => {
            if (item.lottery_b && item.pctchg && item.offline_maxbuyamt) {
                const investAmount = Math.min(netAssets, item.offline_maxbuyamt);
                const stockReturn = investAmount * item.pctchg * item.lottery_b;
                totalReturn += stockReturn;
                validCount++;
            }
        });
        
        // 计算时间跨度
        const startDateTime = new Date(startDate).getTime();
        const endDateTime = new Date(endDate).getTime();
        const daysDiff = (endDateTime - startDateTime) / (1000 * 60 * 60 * 24);
        
        // 计算年化收益率
        const annualizedReturn = (totalReturn / netAssets) * (365 / daysDiff);
        const freeAmount = netAssets - (1.2 * 8000000); // 假设需要预留120万用于打新
        const freeCashReturn = freeAmount > 0 ? (freeAmount / netAssets) * riskFreeRate : 0;
        const totalAnnualizedReturn = annualizedReturn + freeCashReturn;
        
        // 显示结果
        const resultContent = `
            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px;">
                <div class="stat-item">
                    <div class="stat-value">${validCount}</div>
                    <div class="stat-label">参与新股数量</div>
                </div>
                <div class="stat-item">
                    <div class="stat-value">${(totalReturn / 10000).toFixed(2)}万</div>
                    <div class="stat-label">打新总收益</div>
                </div>
                <div class="stat-item">
                    <div class="stat-value">${(annualizedReturn * 100).toFixed(2)}%</div>
                    <div class="stat-label">打新年化收益率</div>
                </div>
                <div class="stat-item">
                    <div class="stat-value">${(freeCashReturn * 100).toFixed(2)}%</div>
                    <div class="stat-label">闲置资金收益率</div>
                </div>
                <div class="stat-item">
                    <div class="stat-value">${(totalAnnualizedReturn * 100).toFixed(2)}%</div>
                    <div class="stat-label">总年化收益率</div>
                </div>
                <div class="stat-item">
                    <div class="stat-value">${selectedBoards.join(', ') || '全部板块'}</div>
                    <div class="stat-label">选择板块</div>
                </div>
            </div>
            <div style="margin-top: 15px; padding: 15px; background: #f8f9fa; border-radius: 8px;">
                <strong>计算说明：</strong><br>
                • 计算期间：${startDate} 至 ${endDate} (${Math.round(daysDiff)}天)<br>
                • 净资产：${(netAssets / 10000).toFixed(0)}万元<br>
                • 无风险收益率：${(riskFreeRate * 100).toFixed(1)}%<br>
                • 预留打新资金：120万元<br>
                • 闲置资金：${Math.max(0, (freeAmount / 10000)).toFixed(0)}万元
            </div>
        `;
        
        document.getElementById('resultContent').innerHTML = resultContent;
        document.getElementById('calculationResult').style.display = 'block';
        
    } catch (error) {
        console.error('计算收益时出错:', error);
        alert('计算过程中出现错误，请检查输入参数！');
    }
}

// 工具函数：格式化数字
function formatNumber(num, decimals = 2) {
    if (num === null || num === undefined || isNaN(num)) return '-';
    return Number(num).toFixed(decimals);
}

// 工具函数：格式化百分比
function formatPercentage(num, decimals = 2) {
    if (num === null || num === undefined || isNaN(num)) return '-';
    return (Number(num) * 100).toFixed(decimals) + '%';
}

// 生成市场洞察和趋势分析
function generateMarketInsights() {
    try {
        const insights = [];
        
        // 分析周度收益趋势
        const weeklyData = excelData['周度收益']?.data || [];
        if (weeklyData.length > 0) {
            const recentWeeks = weeklyData.slice(-8); // 最近8周
            const earlierWeeks = weeklyData.slice(-16, -8); // 之前8周
            
            const recentAvg = recentWeeks.reduce((a, b) => a + (b['年化收益贡献'] || 0), 0) / recentWeeks.length;
            const earlierAvg = earlierWeeks.length > 0 ? earlierWeeks.reduce((a, b) => a + (b['年化收益贡献'] || 0), 0) / earlierWeeks.length : recentAvg;
            
            const trend = ((recentAvg - earlierAvg) / earlierAvg * 100).toFixed(1);
            const trendClass = trend > 0 ? 'trend-positive' : trend < 0 ? 'trend-negative' : 'trend-neutral';
            const trendText = trend > 0 ? '上升' : trend < 0 ? '下降' : '稳定';
            
            insights.push(`近期新股收益表现${trendText}，最近8周平均年化收益为 <span class="insight-highlight">${(recentAvg * 100).toFixed(2)}%</span>，相比前期 <span class="${trendClass}">${trendText}${Math.abs(trend)}%</span>。`);
        }
        
        // 分析发行规模趋势
        const issuanceData = excelData['发行统计']?.data || [];
        if (issuanceData.length > 0) {
            const recentIssues = issuanceData.slice(-4); // 最近4周
            const totalRecentStocks = recentIssues.reduce((a, b) => a + (b.stock_count || 0), 0);
            const totalRecentFunds = recentIssues.reduce((a, b) => a + (b.total_raised_fund || 0), 0);
            const avgWeeklyStocks = (totalRecentStocks / recentIssues.length).toFixed(1);
            const avgWeeklyFunds = (totalRecentFunds / recentIssues.length).toFixed(1);
            
            insights.push(`最近4周新股发行保持活跃，平均每周发行 <span class="insight-highlight">${avgWeeklyStocks}只</span> 新股，募资 <span class="insight-highlight">${avgWeeklyFunds}亿元</span>。`);
        }
        
        // 分析中签率趋势
        const lotteryData = excelData['中签率统计']?.data || [];
        if (lotteryData.length > 0) {
            const recentLottery = lotteryData.slice(-4);
            const avgLotteryB = recentLottery.reduce((a, b) => a + (b.lottery_b || 0), 0) / recentLottery.length;
            const lotteryLevel = avgLotteryB > 0.0002 ? '较高' : avgLotteryB > 0.0001 ? '中等' : '较低';
            
            insights.push(`当前B类投资者中签率处于 <span class="insight-highlight">${lotteryLevel}</span> 水平，平均为 <span class="insight-highlight">${(avgLotteryB * 10000).toFixed(3)}‰</span>。`);
        }
        
        // 分析板块表现趋势
        const sectorData = excelData['板块涨跌幅']?.data || [];
        if (sectorData.length > 0) {
            const sectors = ['上证主板', '深证主板', '科创板', '创业板', '北交所'];
            const sectorPerformance = {};
            const sectorTrends = {};
            
            sectors.forEach(sector => {
                // 最近4周数据
                const recentReturns = sectorData
                    .slice(-4)
                    .map(item => item[sector])
                    .filter(val => val !== null && val !== undefined && !isNaN(val));
                
                // 前4周数据用于对比趋势
                const earlierReturns = sectorData
                    .slice(-8, -4)
                    .map(item => item[sector])
                    .filter(val => val !== null && val !== undefined && !isNaN(val));
                
                if (recentReturns.length > 0) {
                    const recentAvg = recentReturns.reduce((a, b) => a + b, 0) / recentReturns.length;
                    sectorPerformance[sector] = recentAvg;
                    
                    if (earlierReturns.length > 0) {
                        const earlierAvg = earlierReturns.reduce((a, b) => a + b, 0) / earlierReturns.length;
                        const trendChange = ((recentAvg - earlierAvg) / Math.abs(earlierAvg)) * 100;
                        sectorTrends[sector] = trendChange;
                    }
                }
            });
            
            // 找出表现最佳的板块
            const bestSector = Object.entries(sectorPerformance)
                .sort((a, b) => b[1] - a[1])[0];
            
            // 找出趋势最好的板块
            const bestTrendSector = Object.entries(sectorTrends)
                .sort((a, b) => b[1] - a[1])[0];
            
            if (bestSector) {
                let sectorInsight = `各板块中，<span class="insight-highlight">${bestSector[0]}</span> 表现最为突出，近期平均涨幅达 <span class="trend-positive">${(bestSector[1] * 100).toFixed(1)}%</span>`;
                
                if (bestTrendSector && bestTrendSector[1] > 10) {
                    const trendText = bestTrendSector[1] > 0 ? '上升' : '下降';
                    const trendClass = bestTrendSector[1] > 0 ? 'trend-positive' : 'trend-negative';
                    sectorInsight += `。<span class="insight-highlight">${bestTrendSector[0]}</span> 近期趋势向好，相比前期 <span class="${trendClass}">${trendText}${Math.abs(bestTrendSector[1]).toFixed(1)}%</span>`;
                }
                
                insights.push(sectorInsight + '。');
            }
        }
        
        // 分析市盈率水平
        const peData = excelData['市盈率统计']?.data || [];
        if (peData.length > 0) {
            const recentPe = peData.slice(-4);
            const avgIpoPe = recentPe.reduce((a, b) => a + (b.ipo_pe || 0), 0) / recentPe.length;
            const avgIndustryPe = recentPe.reduce((a, b) => a + (b.industry_pe || 0), 0) / recentPe.length;
            const peRatio = (avgIpoPe / avgIndustryPe).toFixed(2);
            const peLevel = peRatio > 1.1 ? '溢价' : peRatio < 0.9 ? '折价' : '合理';
            
            insights.push(`当前新股估值水平 <span class="insight-highlight">${peLevel}</span>，IPO市盈率平均为 <span class="insight-highlight">${avgIpoPe.toFixed(1)}倍</span>，相对行业市盈率${peRatio}倍。`);
        }
        
        // 北交所专项分析
        const bjData = excelData['北交所']?.data || [];
        if (bjData.length > 0) {
            const recentBj = bjData.slice(-3); // 最近3个月
            const avgBjReturn = recentBj.reduce((a, b) => a + (b['北交所年化收益贡献'] || 0), 0) / recentBj.length;
            const bjTrend = avgBjReturn > 0.02 ? '表现优异' : avgBjReturn > 0.01 ? '表现平稳' : '相对较弱';
            
            insights.push(`北交所新股近期${bjTrend}，平均月度年化收益贡献为 <span class="insight-highlight">${(avgBjReturn * 100).toFixed(2)}%</span>。`);
        }
        
        // 投资建议
        const overallSentiment = insights.length > 0 ? generateInvestmentAdvice(excelData) : '';
        if (overallSentiment) {
            insights.push(`<br/><strong>💡 投资建议：</strong>${overallSentiment}`);
        }
        
        // 更新页面显示
        const insightsElement = document.getElementById('marketInsights');
        if (insights.length > 0) {
            insightsElement.innerHTML = insights.join('<br/><br/>');
        } else {
            insightsElement.innerHTML = '数据分析中，请稍后查看市场洞察...';
        }
        
    } catch (error) {
        console.error('生成市场洞察时出错:', error);
        document.getElementById('marketInsights').innerHTML = '暂时无法生成市场分析，请稍后再试。';
    }
}

// 生成投资建议
function generateInvestmentAdvice(data) {
    try {
        const weeklyData = data['周度收益']?.data || [];
        const lotteryData = data['中签率统计']?.data || [];
        
        if (weeklyData.length === 0) return '';
        
        const recentReturns = weeklyData.slice(-4).map(item => item['年化收益贡献'] || 0);
        const avgReturn = recentReturns.reduce((a, b) => a + b, 0) / recentReturns.length;
        
        const recentLottery = lotteryData.slice(-4);
        const avgLotteryB = recentLottery.length > 0 ? 
            recentLottery.reduce((a, b) => a + (b.lottery_b || 0), 0) / recentLottery.length : 0;
        
        let advice = '';
        
        if (avgReturn > 0.05) {
            advice = '当前新股市场收益率较高，建议 <span class="trend-positive">积极参与</span> 打新，但需注意风险控制。';
        } else if (avgReturn > 0.02) {
            advice = '新股市场收益适中，建议 <span class="trend-neutral">适度参与</span>，重点关注优质标的。';
        } else {
            advice = '新股收益相对较低，建议 <span class="trend-negative">谨慎参与</span>，优选高质量公司。';
        }
        
        if (avgLotteryB < 0.0001) {
            advice += ' 当前中签率较低，建议增加申购资金或关注北交所机会。';
        }
        
        return advice;
        
    } catch (error) {
        console.error('生成投资建议时出错:', error);
        return '';
    }
}
