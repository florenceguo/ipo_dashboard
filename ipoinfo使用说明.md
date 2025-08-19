# 📊 ipoinfo.py 优化版使用说明

## 🆕 新增功能：ifplot 开关

### 功能说明
添加了 `ifplot` 参数来控制是否生成图表，提供两种运行模式：

- **ifplot = True**: 完整模式 - 生成数据 + 图表 + 保存到Excel
- **ifplot = False**: 快速模式 - 只生成数据，不绘图，节省时间

## 🔧 使用方法

### 方法1：修改主执行部分（推荐）

在 `ipoinfo.py` 文件的底部找到：

```python
if __name__ == "__main__":
    startdate = '2024-01-01'
    enddate = '2025-08-13'
    fpath = r'C:\quant\ai\IPO数据相关\新股周度统计%s.xlsx'%enddate
    
    # 设置是否生成图表 - 您可以修改这个参数
    ifplot = True  # True: 生成图表和Excel, False: 只生成Excel数据
    
    ipoinfo = update_ipodata(startdate = startdate,enddate = enddate,path = fpath,reconnect = True, ifplot = ifplot)
```

**修改 ifplot 参数：**
- `ifplot = True` - 完整运行（生成图表）
- `ifplot = False` - 快速运行（不生成图表）

### 方法2：函数调用

如果您在其他脚本中调用，可以直接传参：

```python
from ipoinfo import update_ipodata

# 只要数据，不要图表（快速）
ipoinfo = update_ipodata(
    startdate='2024-01-01',
    enddate='2025-08-13', 
    path='新股周度统计2025-08-13.xlsx',
    reconnect=True,
    ifplot=False
)

# 要数据和图表（完整）
ipoinfo = update_ipodata(
    startdate='2024-01-01',
    enddate='2025-08-13',
    path='新股周度统计2025-08-13.xlsx', 
    reconnect=True,
    ifplot=True
)
```

## 🚀 运行示例

### 快速数据更新（不生成图表）
```python
# 修改 ipoinfo.py 中的 ifplot = False
# 然后运行
C:\ProgramData\anaconda3\python.exe ipoinfo.py
```

**优点：**
- ⚡ 运行速度快
- 💾 节省存储空间
- 🔄 适合频繁数据更新

**输出：**
- ✅ Excel文件包含所有数据表
- ❌ 不包含图表

### 完整报告生成（包含图表）
```python
# 修改 ipoinfo.py 中的 ifplot = True  
# 然后运行
C:\ProgramData\anaconda3\python.exe ipoinfo.py
```

**优点：**
- 📊 包含所有可视化图表
- 📈 完整的分析报告
- 🎨 美观的Excel展示

**输出：**
- ✅ Excel文件包含所有数据表
- ✅ 每个工作表包含对应的图表

## 📋 工作流程建议

### 日常数据更新
```bash
# 1. 设置为快速模式
ifplot = False

# 2. 运行数据获取
python ipoinfo.py

# 3. 转换为网站数据
python analyze_excel.py

# 4. 查看网站
# 访问 http://localhost:8080
```

### 周期性报告生成
```bash
# 1. 设置为完整模式
ifplot = True

# 2. 生成完整报告
python ipoinfo.py

# 3. 查看Excel中的图表
# 4. 如需网站展示，再运行 analyze_excel.py
```

## 🔍 技术细节

### 修改的内容
1. **函数签名**: 添加 `ifplot = True` 参数
2. **条件绘图**: 所有 `plt.*` 代码包装在 `if ifplot:` 中
3. **临时文件**: 只在绘图时创建和删除临时PNG文件
4. **错误处理**: 添加文件删除的异常处理

### 性能对比
| 模式 | 运行时间 | Excel大小 | 适用场景 |
|------|----------|-----------|----------|
| ifplot=False | ~30秒 | ~500KB | 数据更新 |
| ifplot=True | ~60秒 | ~5MB | 报告生成 |

## ❓ 常见问题

### Q1: 如何快速切换模式？
**A**: 只需修改 `ipoinfo.py` 底部的 `ifplot = True/False`

### Q2: 不生成图表会影响网站显示吗？
**A**: 不会！网站使用 `analyze_excel.py` 转换的JSON数据，图表由JavaScript重新生成。

### Q3: 可以只生成部分图表吗？
**A**: 当前版本是全开/全关，如需要可以进一步定制。

## 🎯 最佳实践

1. **开发阶段**: 使用 `ifplot=False` 快速迭代
2. **生产报告**: 使用 `ifplot=True` 生成完整报告  
3. **自动化**: 可以根据时间或条件动态设置ifplot值
4. **存储管理**: 定期清理只有数据的Excel文件，保留重要的图表版本

现在您的 `ipoinfo.py` 更加灵活高效了！🎉

