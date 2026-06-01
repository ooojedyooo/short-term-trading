---
name: 短线收益
description: 股票短线交易盈亏分析系统（v4.5含佣金印花税+跨账户配对）。触发词：分析短线收益、短线收益、股票交易分析、盈亏分析、处理交易记录、分析股票盈亏、交易记录分析、股票交易盈亏、投资收益分析、开始分析、来活了、分析吧、跑分析、算收益、看盈亏、新文件。当用户在短线收益工作目录下要求分析股票交易记录、计算盈亏、生成投资报告时使用此技能。
---

# 短线收益分析技能

## 概述

此技能用于自动化分析A股短线交易记录（v4.5），支持三种账户（两融/手机/平安），自动匹配买卖记录，计算含佣金（万一/双向/最低5元）和印花税（万五/卖出单边）的净盈亏，生成Excel汇总和HTML可视化报告，并自动提交到GitHub。

## 核心功能

### 1. 自动化数据处理

- 自动识别Excel交易记录文件（两融.xlsx / 平安.xls）
- OCR识别手机App截图（.jpg/.png）
- 跨账户合并匹配（v4.5）：同日所有账户数据合并后再按股票代码配对
- 单边交易标记：未平仓标记为 ⚠️仅买入未平仓 / ⚠️仅卖出未平仓 / ⚠️多买未平仓 / ⚠️多卖未平仓
- 自动归档已处理文件到 history/

### 2. 智能计算逻辑（v4.5 含交易成本）

```
匹配数量 = min(总买入数量, 总卖出数量)  # 跨账户合并后
买入均价 = 总买入金额 / 总买入数量
卖出均价 = 总卖出金额 / 总卖出数量
毛盈亏 = 卖出金额 - 买入金额
佣金 = max(买入金额×0.01%, 5元) + max(卖出金额×0.01%, 5元)  — 万一/双向/最低5元
印花税 = 卖出金额 × 0.05%                                        — 万五/卖出单边
交易成本 = 佣金 + 印花税
净盈亏 = 毛盈亏 - 交易成本
盈亏比例 = 净盈亏 / 买入金额 × 100%
```

### 3. 数据去重机制

- 同一日期整日替换（非按来源替换），因为跨账户配对会改变所有记录
- 确保Excel汇总文件不重复

### 4. 报告生成

**单日HTML报告**：
- 文件路径：`reports/YYYY-MM-DD-股票交易盈亏报告.html`
- 内容：单日交易明细、盈亏统计、连续亏损/回撤统计

**汇总可视化报告**：
- 文件路径：`reports/汇总可视化报告.html`
- 标签页导航：总览 / 个股 / 日历
- 总览页：统计卡片 + 月度/每日盈亏柱状图 + 累计收益曲线 + 股票盈亏排行 + 饼图
- 个股页：累计盈亏排行榜 + 盈亏对比柱状图（>15只时隐藏±50元以内小票）+ 胜率对比
- 日历页：ECharts日历热力图 + 连续亏损预警

## 执行命令

当用户请求分析时，执行：

```bash
cd d:\短线收益
"C:\Users\Ryan\AppData\Local\Programs\Python\Python313\python.exe" 股票交易分析系统.py
```

### ⚠️ Python环境（重要！）

**必须使用系统Python**：`C:\Users\Ryan\AppData\Local\Programs\Python\Python313\python.exe`

- 系统Python已安装全部依赖（pandas, openpyxl, easyocr, pillow, pytesseract等）
- **绝对不要用managed venv**（`C:\Users\Ryan\.workbuddy\binaries\python\envs\default\`），那是空隔离环境，会报缺依赖
- 之前出现过误用managed venv导致重复装包的问题，已记录

## 完整工作流

分析完成后，必须执行以下步骤：

1. 展示分析结果（毛盈亏、佣金、印花税、净盈亏）
2. 提交并推送到 Git：
```bash
cd d:\短线收益
git add -A
git commit -m "YYYY-MM-DD 交易分析：..."
git push
```
3. 预览单日报告：`reports/YYYY-MM-DD-股票交易盈亏报告.html`
4. 更新本地记忆文件 `d:\短线收益\.workbuddy\memory/YYYY-MM-DD.md`

## ⚠️ 重要注意事项

### 补充账户数据（必读！）

当用户说"少给了XX账户数据，补上"时，之前已处理的文件已归档到 `history/` 目录。由于系统采用"整日替换"去重机制，仅用新账户数据重跑会**覆盖掉已有记录**。

**正确做法**：
```bash
# 1. 将已归档的同日文件拷回工作目录
cp history/YYYY-MM-DD-两融-当日成交汇总.xlsx .
cp history/YYYYMMDD_*.jpg .
# 2. 确认新补充的文件也在工作目录
# 3. 然后重跑分析
"C:\Users\Ryan\AppData\Local\Programs\Python\Python313\python.exe" 股票交易分析系统.py
```

这样所有账户文件同时存在，系统会合并三账户数据后统一配对，整日替换才能得到完整结果。

### OCR依赖问题

- 首次OCR识别会触发easyocr模型下载，可能较慢
- 如果OCR失败提示缺包，确认使用的是系统Python（见上方"Python环境"）
- 系统Python路径：`C:\Users\Ryan\AppData\Local\Programs\Python\Python313\python.exe`

### Git推送问题

- Git配置了本地代理 http://127.0.0.1:7890
- 如果代理未启动，push会失败（Connection refused）
- 临时解决方案：`git -c http.proxy="" -c https.proxy="" push`（需网络能直连GitHub）
- 如果直连也超时，提醒用户开启代理后重试
- GitHub凭据可能过期（credential-manager无法自动续期），需用户手动 `git push` 重新登录

### 股票盈亏柱状图过滤

- 当股票数量 > 15只时，盈亏在 ±50元以内的自动隐藏
- 图表标题会显示"已隐藏X只盈亏±50元以内股票"
- ≤15只时全显示，不过滤
- 涉及3处图表：个股累计盈亏对比图、总览排行、月度钻取

## 文件命名规范

| 账户 | 文件格式 | 示例 |
|------|----------|------|
| 两融账户 | Excel | 2026-05-11-两融-当日成交汇总.xlsx |
| 手机账户 | 图片 | 20260511_150215_72_141.jpg |
| 平安账户截图 | 图片 | 20260511_平安.png |
| 平安账户导出 | Excel(TSV/GBK) | 20260511_平安.xls |

## 技术栈

- Python 3.x + pandas + openpyxl
- HTML + ECharts 可视化（v4.4起从Chart.js迁移到ECharts，支持日历热力图）
- Git 版本控制（GitHub: ooojedyooo/short-term-tracking）
