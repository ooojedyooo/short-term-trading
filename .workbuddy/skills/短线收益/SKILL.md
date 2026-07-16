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

### 5. 跨天配对分析（月度 + 年度流水线）

主分析系统只在**同一天**内配对买卖，跨日（T+0/短线）的零散买卖会被标成 ⚠️未平仓只扣交易成本。此脚本读取汇总表 `股票交易盈亏汇总.xlsx` 的未平仓行，按月/年整合重新配对，算出被漏算的真实盈亏。

**用途**：每月初复盘上月、每年初复盘前一年"按月配完仍剩余的"跨月持仓。

**执行命令**（均在 `d:\短线收益` 目录，用系统 Python）：

```bash
# 每月初：算上个月（--month 不带值自动取上月）
"C:\Users\Ryan\AppData\Local\Programs\Python\Python313\python.exe" 跨天配对分析.py --month
# 每年初：算前一年剩余跨月配对（--year 不带值自动取去年）
"C:\Users\Ryan\AppData\Local\Programs\Python\Python313\python.exe" 跨天配对分析.py --year
# 指定具体月份/年份
"C:\Users\Ryan\AppData\Local\Programs\Python\Python313\python.exe" 跨天配对分析.py --month 2026-06
"C:\Users\Ryan\AppData\Local\Programs\Python\Python313\python.exe" 跨天配对分析.py --year 2026
# 不带参数：全周期逐月总览（各月分别跨天配对）
"C:\Users\Ryan\AppData\Local\Programs\Python\Python313\python.exe" 跨天配对分析.py
```

**输出文件**：
- 月度：`reports/月度跨天配对分析_YYYY-MM.html` + `reports/剩余持仓_YYYY-MM.xlsx`
- 年度：`reports/年度跨天配对分析_YYYY.html` + `reports/剩余持仓_YYYY年度.xlsx`
- 总览：`reports/月度跨天配对分析.html`

**口径与语义**：
- 加权均价（WAC）配对，与系统一致；按"月"不跨月
- **月度模式**：聚合该月所有 ⚠️未平仓行 → 月内跨天配对 → 输出配对盈亏 + 剩余（剩余继续留在汇总表可追溯）
- **年度模式（忠实流水线）**：先把前一年各月分别月内配对取剩余（带金额）→ 各月剩余再跨月合并配对 → 输出"月度跨天释放" + "年度跨月释放"两部分。**切勿改成全年未平仓行一次性全局重配**（会抹平月内价差、低估真实盈亏）
- 剩余未匹配持仓始终保留在汇总表，脚本只读取计算、从不删除；导出的 `剩余持仓_*.xlsx` 含金额列，可直接喂给下一年度步骤

**已知局限**：①按月不跨月（5月纯买/6月纯卖分属两月不配对，需年度步骤打通）；②WAC非FIFO；③依赖汇总表数据准确性（已知 `validate_trades` 对手机OCR失效，万一OCR读偏会被带入跨天结果，关键票建议人工核对）

**修复记录**：曾修过3处bug——①名称变体（同码不同名如"中国稀土"vs"中国稀士"被拆两只漏配，改按证券代码分组）；②年度口径错误（全年重配低估盈亏，改忠实流水线）；③HTML报告表格数据未注入（`.format()`占位符数量不匹配，导致只有表头）

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
2. **刷新当月 + 当年跨天配对（每日自动同步，用户要求）**：当日分析写入汇总表后，紧接着用**当前**月份和年份刷新跨天配对，让"当月真实盈亏""当年真实盈亏"实时同步。跨天脚本只读汇总表、幂等安全，可天天重复跑：
```bash
cd d:\短线收益
# 注意：这里传的是【当前】月/年（实时刷新），不是默认的上月/去年
"C:\Users\Ryan\AppData\Local\Programs\Python\Python313\python.exe" 跨天配对分析.py --month YYYY-MM
"C:\Users\Ryan\AppData\Local\Programs\Python\Python313\python.exe" 跨天配对分析.py --year YYYY
```
   - 输出：`reports/月度跨天配对分析_YYYY-MM.html`、`reports/年度跨天配对分析_YYYY.html`、`reports/剩余持仓_*.xlsx`
   - ⚠️ 区分场景：**每日实时刷新**传当前月/年（如 `--month 2026-07 --year 2026`）；**月初/年初复盘**才用不带值的 `--month`/`--year`（自动取上月/去年）
3. 提交并推送到 Git（当日报告 + 刷新后的跨天报告一起提交）：
```bash
cd d:\短线收益
git add -A
git commit -m "YYYY-MM-DD 交易分析：...（含当月/当年跨天配对刷新）"
git push
```
4. 预览单日报告：`reports/YYYY-MM-DD-股票交易盈亏报告.html`
5. 更新本地记忆文件 `d:\短线收益\.workbuddy\memory/YYYY-MM-DD.md`

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
