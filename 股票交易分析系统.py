"""
股票交易盈亏分析系统
功能：自动分析股票交易记录，计算盈亏并生成报告
支持两种输入：
  1. Excel文件（券商导出）— 文件名格式：YYYY-MM-DD-两融-当日成交汇总.xlsx
  2. 图片文件（手机App截图）— 文件名格式：YYYY-MM-DD-手机交易.png/.jpg/.jpeg
作者：WorkBuddy
版本：v4.1
更新日期：2026-04-22

使用说明：
1. 将券商导出的Excel交易记录文件或手机App截图放到当前文件夹
2. 运行此脚本，自动处理所有未归档的文件
3. 处理后的原始文件会自动归档到history文件夹
4. HTML报告生成到reports文件夹（单日报告 + 汇总可视化报告）
5. Excel汇总文件支持去重更新（同日期+同来源的数据会覆盖）
6. 汇总可视化报告支持时间筛选和多级数据钻取
7. 同一天可以同时有Excel和图片两种输入，数据会自动合并

文件结构：
- 当前文件夹：待处理的Excel文件和图片文件
- reports/：HTML报告文件（单日报告 + 汇总可视化报告）
- history/：已处理的原始文件
- 股票交易盈亏汇总.xlsx：累计盈亏数据（去重更新）
"""

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime
import os
import shutil
import glob
import re

# ==================== 配置参数 ====================
EXCEL_OUTPUT = '股票交易盈亏汇总.xlsx'  # Excel汇总文件
REPORTS_DIR = 'reports'                  # HTML报告文件夹
HISTORY_DIR = 'history'                  # 原始文件归档文件夹

# 确保文件夹存在
os.makedirs(REPORTS_DIR, exist_ok=True)
os.makedirs(HISTORY_DIR, exist_ok=True)


# ==================== 核心处理函数 ====================

def extract_date_from_filename(filename):
    """从文件名提取日期，支持 YYYY-MM-DD 和 YYYYMMDD 格式"""
    date_match = re.search(r'(\d{4}-\d{2}-\d{2})', filename)
    if date_match:
        return date_match.group(1)
    date_match = re.search(r'(\d{4})(\d{2})(\d{2})', filename)
    if date_match:
        return f"{date_match.group(1)}-{date_match.group(2)}-{date_match.group(3)}"
    return datetime.now().strftime('%Y-%m-%d')


def get_source_from_filename(filename):
    """从文件名判断数据来源"""
    fname_lower = filename.lower()
    if fname_lower.endswith('.xlsx'):
        return '两融账户'
    elif any(fname_lower.endswith(ext) for ext in ['.png', '.jpg', '.jpeg']):
        return '手机账户'
    return '未知'


# ==================== 图片OCR解析 ====================

def parse_image_trades(image_path):
    """用手机App截图识别交易记录"""
    try:
        import easyocr
    except ImportError:
        print(f"[错误] 缺少OCR依赖，请运行：pip install easyocr")
        return pd.DataFrame()

    print(f"  正在识别图片：{os.path.basename(image_path)}")
    reader = easyocr.Reader(['ch_sim', 'en'], gpu=False, verbose=False)
    result = reader.readtext(image_path)

    items = []
    for bbox, text, conf in result:
        y_center = (bbox[0][1] + bbox[2][1]) / 2
        x_center = (bbox[0][0] + bbox[2][0]) / 2
        items.append({'text': text.strip(), 'y': y_center, 'x': x_center, 'conf': conf})

    items.sort(key=lambda r: r['y'])
    rows = []
    current_row = []
    last_y = None
    y_threshold = 25

    for item in items:
        if last_y is None or abs(item['y'] - last_y) <= y_threshold:
            current_row.append(item)
        else:
            if current_row:
                rows.append(current_row)
            current_row = [item]
        last_y = item['y']

    if current_row:
        rows.append(current_row)

    for row in rows:
        row.sort(key=lambda r: r['x'])

    records = []
    for row in rows:
        texts = [r['text'] for r in row]
        text_combined = ' '.join(texts)

        code_match = re.search(r'\b(\d{6})\b', text_combined)
        if not code_match:
            continue

        stock_code = code_match.group(1)

        direction = None
        if '买入' in text_combined or '买' in text_combined:
            direction = '证券买入'
        elif '卖出' in text_combined or '卖' in text_combined:
            direction = '证券卖出'
        else:
            continue

        stock_name = ''
        code_idx = text_combined.find(stock_code)
        if code_idx > 0:
            prefix = text_combined[:code_idx].strip()
            prefix = re.sub(r'\d{1,2}:\d{2}:\d{2}', '', prefix)
            prefix = re.sub(r'买入|卖出|买|卖|成交', '', prefix)
            prefix = prefix.strip()
            cn_match = re.findall(r'[\u4e00-\u9fff]+', prefix)
            if cn_match:
                stock_name = cn_match[-1]

        numbers_text = text_combined.replace(stock_code, '')
        numbers = re.findall(r'\d+\.?\d*', numbers_text)
        numbers = [float(n) for n in numbers if float(n) > 0]

        if len(numbers) < 3:
            continue

        volume = None
        for n in sorted(numbers):
            if n >= 100 and n == int(n):
                volume = int(n)
                break
        if volume is None:
            volume = int(min(numbers))

        amount = max(numbers)
        price = round(amount / volume, 3) if volume > 0 else 0

        for n in numbers:
            if n != amount and n != volume and abs(n - price) < 1:
                price = n
                break

        records.append({
            '证券代码': stock_code,
            '证券名称': stock_name if stock_name else '未知',
            '买卖类别': direction,
            '成交类型': '成交',
            '成交数量': volume,
            '成交价格': price,
            '成交金额': amount
        })

    df = pd.DataFrame(records)
    print(f"  识别到 {len(df)} 条交易记录")
    return df


def process_image_file(image_path):
    """处理单个图片文件"""
    trading_date = extract_date_from_filename(image_path)
    source = get_source_from_filename(os.path.basename(image_path))

    print(f"\n{'='*80}")
    print(f"处理图片：{os.path.basename(image_path)}")
    print(f"交易日期：{trading_date}")
    print(f"数据来源：{source}")
    print('='*80)

    df = parse_image_trades(image_path)

    if len(df) == 0:
        print("未识别到有效交易记录")
        return pd.DataFrame(), trading_date, 0, source

    buy_records = df[df['买卖类别'].str.contains('证券买入', na=False)].copy()
    sell_records = df[df['买卖类别'].str.contains('证券卖出', na=False)].copy()

    print(f"买入记录数：{len(buy_records)}")
    print(f"卖出记录数：{len(sell_records)}")

    profit_results = calculate_profits(df, buy_records, sell_records, trading_date, source)

    total_profit = sum(r['盈亏金额'] for r in profit_results)
    print(f"\n总盈亏：{total_profit:.2f} 元")
    print('='*80)

    result_df = pd.DataFrame(profit_results)
    return result_df, trading_date, total_profit, source


# ==================== Excel处理 ====================

def process_excel_file(input_file):
    """处理单个Excel文件"""
    trading_date = extract_date_from_filename(input_file)
    source = get_source_from_filename(os.path.basename(input_file))

    print(f"\n{'='*80}")
    print(f"处理文件：{os.path.basename(input_file)}")
    print(f"交易日期：{trading_date}")
    print(f"数据来源：{source}")
    print('='*80)

    df = pd.read_excel(input_file, skiprows=4, header=0)
    df.columns = ['证券代码', '证券名称', '买卖类别', '成交类型', '成交数量', '成交价格', '成交金额']
    df['证券代码'] = df['证券代码'].astype(str).str.replace('\t', '')
    df = df.dropna(subset=['证券代码'])
    df = df[df['证券代码'] != '证券代码']

    df['成交数量'] = pd.to_numeric(df['成交数量'], errors='coerce')
    df['成交价格'] = pd.to_numeric(df['成交价格'], errors='coerce')
    df['成交金额'] = pd.to_numeric(df['成交金额'], errors='coerce')

    buy_records = df[df['买卖类别'].str.contains('证券买入', na=False)].copy()
    sell_records = df[df['买卖类别'].str.contains('证券卖出', na=False)].copy()

    print(f"买入记录数：{len(buy_records)}")
    print(f"卖出记录数：{len(sell_records)}")

    profit_results = calculate_profits(df, buy_records, sell_records, trading_date, source)

    total_profit = sum(r['盈亏金额'] for r in profit_results)
    print(f"\n总盈亏：{total_profit:.2f} 元")
    print('='*80)

    result_df = pd.DataFrame(profit_results)
    return result_df, trading_date, total_profit, source


def calculate_profits(df, buy_records, sell_records, trading_date, source):
    """计算盈亏（通用逻辑，适用于Excel和图片输入）"""
    profit_results = []
    all_stocks = set(df['证券代码'].unique())

    for stock_code in all_stocks:
        stock_name = df[df['证券代码'] == stock_code]['证券名称'].iloc[0]

        buys = buy_records[buy_records['证券代码'] == stock_code]
        sells = sell_records[sell_records['证券代码'] == stock_code]

        if len(buys) == 0 or len(sells) == 0:
            print(f"跳过 {stock_name} ({stock_code})：缺少买入或卖出记录")
            continue

        total_buy_qty = buys['成交数量'].sum()
        total_sell_qty = sells['成交数量'].sum()
        matched_qty = min(total_buy_qty, total_sell_qty)

        if matched_qty == 0:
            continue

        total_buy_amt = buys['成交金额'].sum()
        avg_buy_price = total_buy_amt / total_buy_qty
        matched_buy_amt = avg_buy_price * matched_qty

        total_sell_amt = sells['成交金额'].sum()
        avg_sell_price = total_sell_amt / total_sell_qty
        matched_sell_amt = avg_sell_price * matched_qty

        profit = matched_sell_amt - matched_buy_amt
        profit_pct = (profit / matched_buy_amt) * 100 if matched_buy_amt != 0 else 0

        profit_results.append({
            '日期': trading_date,
            '数据来源': source,
            '证券代码': stock_code,
            '证券名称': stock_name,
            '买入数量': total_buy_qty,
            '卖出数量': total_sell_qty,
            '匹配数量': matched_qty,
            '买入均价': round(avg_buy_price, 4),
            '卖出均价': round(avg_sell_price, 4),
            '买入金额': round(matched_buy_amt, 2),
            '卖出金额': round(matched_sell_amt, 2),
            '盈亏金额': round(profit, 2),
            '盈亏比例': f"{profit_pct:.2f}%"
        })

        print(f"股票：{stock_name} ({stock_code})")
        print(f"  买入：数量={total_buy_qty:.0f}, 均价={avg_buy_price:.4f}, 金额={matched_buy_amt:.2f}")
        print(f"  卖出：数量={total_sell_qty:.0f}, 均价={avg_sell_price:.4f}, 金额={matched_sell_amt:.2f}")
        print(f"  盈亏：{profit:.2f} 元 ({profit_pct:.2f}%)")

    return profit_results


# ==================== Excel汇总 ====================

def append_to_excel(result_df, trading_date, source):
    """追加数据到Excel汇总文件（支持按日期+来源去重）"""
    if len(result_df) == 0:
        return

    if os.path.exists(EXCEL_OUTPUT):
        existing_df = pd.read_excel(EXCEL_OUTPUT)
        # 删除该日期+该来源的旧数据，保留其他日期和其他来源的数据
        existing_df = existing_df[~((existing_df['日期'] == trading_date) & (existing_df['数据来源'] == source))]
        combined_df = pd.concat([existing_df, result_df], ignore_index=True)
        combined_df = combined_df.sort_values(['日期', '数据来源']).reset_index(drop=True)
    else:
        combined_df = result_df

    wb = Workbook()
    ws = wb.active
    ws.title = "股票盈亏汇总"

    headers = ['日期', '数据来源', '证券代码', '证券名称', '买入数量', '卖出数量', '匹配数量',
               '买入均价', '卖出均价', '买入金额', '卖出金额', '盈亏金额', '盈亏比例']

    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)

    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')

    for idx, row in enumerate(combined_df.itertuples(index=False), 2):
        ws.cell(row=idx, column=1, value=row.日期)
        ws.cell(row=idx, column=2, value=row.数据来源)
        ws.cell(row=idx, column=3, value=row.证券代码)
        ws.cell(row=idx, column=4, value=row.证券名称)
        ws.cell(row=idx, column=5, value=row.买入数量)
        ws.cell(row=idx, column=6, value=row.卖出数量)
        ws.cell(row=idx, column=7, value=row.匹配数量)
        ws.cell(row=idx, column=8, value=row.买入均价)
        ws.cell(row=idx, column=9, value=row.卖出均价)
        ws.cell(row=idx, column=10, value=row.买入金额)
        ws.cell(row=idx, column=11, value=row.卖出金额)

        profit_cell = ws.cell(row=idx, column=12, value=row.盈亏金额)
        if row.盈亏金额 > 0:
            profit_cell.font = Font(color="FF0000", bold=True)
        else:
            profit_cell.font = Font(color="00B050", bold=True)

        ws.cell(row=idx, column=13, value=row.盈亏比例)

        for col in range(1, 14):
            ws.cell(row=idx, column=col).alignment = Alignment(horizontal='center', vertical='center')

    column_widths = [12, 12, 12, 14, 10, 10, 10, 12, 12, 14, 14, 14, 12]
    for i, width in enumerate(column_widths, 1):
        ws.column_dimensions[chr(64+i)].width = width

    wb.save(EXCEL_OUTPUT)
    action = "更新" if os.path.exists(EXCEL_OUTPUT) else "创建"
    print(f"Excel汇总已{action}（按日期+来源去重）：{EXCEL_OUTPUT}")


# ==================== HTML报告生成 ====================

def generate_html_report_from_summary(trading_date):
    """从汇总文件中提取当日数据生成HTML报告"""
    html_filename = f"{trading_date}-股票交易盈亏报告.html"
    html_path = os.path.join(REPORTS_DIR, html_filename)

    if not os.path.exists(EXCEL_OUTPUT):
        print(f"汇总文件不存在，无法生成 {trading_date} 的报告")
        return

    summary_df = pd.read_excel(EXCEL_OUTPUT)
    day_df = summary_df[summary_df['日期'] == trading_date].copy()

    if len(day_df) == 0:
        print(f"汇总文件中未找到 {trading_date} 的数据")
        return

    total_profit = day_df['盈亏金额'].sum()

    # 按数据来源分组统计
    sources = day_df['数据来源'].unique()
    source_summary = ""
    for src in sources:
        src_df = day_df[day_df['数据来源'] == src]
        src_profit = src_df['盈亏金额'].sum()
        source_summary += f"<span class='source-tag source-{src[:2]}'>{src}: {'+' if src_profit > 0 else ''}¥{src_profit:,.2f}</span>"

    html_content = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>股票交易盈亏报告 - {trading_date}</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{
            font-family: 'Microsoft YaHei', 'SimHei', Arial, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 20px;
            min-height: 100vh;
        }}
        .container {{
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 15px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            overflow: hidden;
        }}
        .header {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }}
        .header h1 {{ font-size: 32px; margin-bottom: 10px; font-weight: 600; }}
        .header .date {{ font-size: 18px; opacity: 0.9; margin-bottom: 8px; }}
        .header .sources {{ font-size: 14px; opacity: 0.85; }}
        .summary-cards {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            padding: 30px;
            background: #f8f9fa;
        }}
        .card {{
            background: white;
            padding: 25px;
            border-radius: 12px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
            text-align: center;
            transition: transform 0.3s;
        }}
        .card:hover {{ transform: translateY(-5px); }}
        .card .title {{ color: #666; font-size: 14px; margin-bottom: 10px; text-transform: uppercase; letter-spacing: 1px; }}
        .card .value {{ font-size: 32px; font-weight: bold; margin-bottom: 5px; }}
        .card .value.profit {{ color: #e74c3c; }}
        .card .value.loss {{ color: #27ae60; }}
        .card .value.neutral {{ color: #95a5a6; }}
        .table-container {{ padding: 30px; }}
        .table-title {{
            font-size: 24px;
            color: #2c3e50;
            margin-bottom: 20px;
            padding-left: 15px;
            border-left: 4px solid #667eea;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            background: white;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            border-radius: 8px;
            overflow: hidden;
            table-layout: fixed;
        }}
        thead {{ background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); }}
        th {{ color: white; padding: 15px 8px; text-align: center; font-weight: 600; font-size: 14px; }}
        td {{ padding: 15px 8px; text-align: center; border-bottom: 1px solid #ecf0f1; font-size: 14px; }}
        tbody tr:hover {{ background: #f8f9fa; }}
        .profit-amount {{ color: #e74c3c; font-weight: bold; }}
        .loss-amount {{ color: #27ae60; font-weight: bold; }}
        .footer {{
            padding: 20px;
            text-align: center;
            color: #7f8c8d;
            font-size: 14px;
            background: #f8f9fa;
            border-top: 1px solid #ecf0f1;
        }}
        .no-data {{ text-align: center; padding: 40px; color: #7f8c8d; }}
        .source-tag {{
            display: inline-block;
            padding: 4px 12px;
            border-radius: 12px;
            font-size: 12px;
            margin: 0 4px;
        }}
        .source-两融 {{ background: #e3f2fd; color: #1976d2; }}
        .source-手机 {{ background: #fff3e0; color: #f57c00; }}
        /* 列宽设置 */
        th:nth-child(1), td:nth-child(1) {{ width: 12%; }} /* 数据来源 */
        th:nth-child(2), td:nth-child(2) {{ width: 18%; }} /* 证券名称 */
        th:nth-child(3), td:nth-child(3) {{ width: 10%; }} /* 买入数量 */
        th:nth-child(4), td:nth-child(4) {{ width: 10%; }} /* 卖出数量 */
        th:nth-child(5), td:nth-child(5) {{ width: 10%; }} /* 匹配数量 */
        th:nth-child(6), td:nth-child(6) {{ width: 12%; }} /* 买入均价 */
        th:nth-child(7), td:nth-child(7) {{ width: 12%; }} /* 卖出均价 */
        th:nth-child(8), td:nth-child(8) {{ width: 14%; }} /* 买入金额 */
        th:nth-child(9), td:nth-child(9) {{ width: 14%; }} /* 卖出金额 */
        th:nth-child(10), td:nth-child(10) {{ width: 16%; }} /* 盈亏金额 */
        th:nth-child(11), td:nth-child(11) {{ width: 12%; }} /* 盈亏比例 */
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📈 股票交易盈亏报告</h1>
            <div class="date">交易日期：{trading_date}</div>
            <div class="sources">{source_summary}</div>
        </div>
        <div class="summary-cards">
            <div class="card">
                <div class="title">交易股票数</div>
                <div class="value neutral">{len(day_df)}</div>
            </div>
            <div class="card">
                <div class="title">总买入金额</div>
                <div class="value neutral">¥{day_df['买入金额'].sum():,.2f}</div>
            </div>
            <div class="card">
                <div class="title">总卖出金额</div>
                <div class="value neutral">¥{day_df['卖出金额'].sum():,.2f}</div>
            </div>
            <div class="card">
                <div class="title">总盈亏金额</div>
                <div class="value {'profit' if total_profit > 0 else 'loss'}">
                    {'+' if total_profit > 0 else ''}¥{total_profit:,.2f}
                </div>
            </div>
        </div>
        <div class="table-container">
            <h2 class="table-title">交易明细</h2>
"""

    if len(day_df) > 0:
        html_content += """
            <table>
                <thead>
                    <tr>
                        <th>数据来源</th>
                        <th>证券名称</th>
                        <th>买入数量</th>
                        <th>卖出数量</th>
                        <th>匹配数量</th>
                        <th>买入均价</th>
                        <th>卖出均价</th>
                        <th>买入金额</th>
                        <th>卖出金额</th>
                        <th>盈亏金额</th>
                        <th>盈亏比例</th>
                    </tr>
                </thead>
                <tbody>
"""
        for _, row in day_df.iterrows():
            profit_class = 'profit-amount' if row['盈亏金额'] > 0 else 'loss-amount'
            profit_sign = '+' if row['盈亏金额'] > 0 else ''
            html_content += f"""
                    <tr>
                        <td><span class="source-tag source-{row['数据来源'][:2]}">{row['数据来源']}</span></td>
                        <td><strong>{row['证券名称']}</strong></td>
                        <td>{row['买入数量']:,.0f}</td>
                        <td>{row['卖出数量']:,.0f}</td>
                        <td>{row['匹配数量']:,.0f}</td>
                        <td>¥{row['买入均价']:,.4f}</td>
                        <td>¥{row['卖出均价']:,.4f}</td>
                        <td>¥{row['买入金额']:,.2f}</td>
                        <td>¥{row['卖出金额']:,.2f}</td>
                        <td class="{profit_class}">{profit_sign}¥{row['盈亏金额']:,.2f}</td>
                        <td class="{profit_class}">{row['盈亏比例']}</td>
                    </tr>
"""
        html_content += """
                </tbody>
            </table>
"""
    else:
        html_content += """
            <div class="no-data">
                <p>暂无匹配的买卖记录</p>
            </div>
"""

    html_content += f"""
        </div>
        <div class="footer">
            <p>报告生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
        </div>
    </div>
</body>
</html>
"""

    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    print(f"HTML报告已生成（从汇总提取）：{html_path}")


# ==================== 汇总可视化报告（v3.0 交互式） ====================

def generate_summary_html():
    """生成交互式汇总可视化HTML报告"""
    if not os.path.exists(EXCEL_OUTPUT):
        print("未找到汇总文件，跳过汇总报告生成")
        return

    df = pd.read_excel(EXCEL_OUTPUT)
    if len(df) == 0:
        print("汇总文件无数据，跳过汇总报告生成")
        return

    df['日期'] = pd.to_datetime(df['日期']).dt.strftime('%Y-%m-%d')

    import json
    records = []
    for _, row in df.iterrows():
        records.append({
            'date': str(row['日期']),
            'source': str(row['数据来源']),
            'code': str(row['证券代码']),
            'name': str(row['证券名称']),
            'buyQty': int(row['买入数量']),
            'sellQty': int(row['卖出数量']),
            'matchQty': int(row['匹配数量']),
            'buyPrice': round(float(row['买入均价']), 4),
            'sellPrice': round(float(row['卖出均价']), 4),
            'buyAmount': round(float(row['买入金额']), 2),
            'sellAmount': round(float(row['卖出金额']), 2),
            'profit': round(float(row['盈亏金额']), 2),
            'profitPct': str(row['盈亏比例'])
        })

    data_json = json.dumps(records, ensure_ascii=False)
    now_str = datetime.now().strftime('%Y-%m-%d')
    gen_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    html_path = os.path.join(REPORTS_DIR, '汇总可视化报告.html')

    html_content = generate_html_template(data_json, now_str, gen_time)

    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    print(f"汇总可视化报告已生成：{html_path}")


def generate_html_template(data_json, now_str, gen_time):
    """生成HTML模板（v3.0完整内容）"""
    return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>股票交易汇总分析</title>
    <script src="https://cdn.jsdelivr.net/npm/echarts@5.4.3/dist/echarts.min.js"></script>
    <style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{
    font-family: 'Microsoft YaHei','SimHei',Arial,sans-serif;
    background: linear-gradient(135deg,#667eea 0%,#764ba2 100%);
    padding: 20px; min-height: 100vh;
}}
.container {{
    max-width: 1400px; margin: 0 auto; background: #fff;
    border-radius: 15px; box-shadow: 0 20px 60px rgba(0,0,0,0.3); overflow: hidden;
}}
.header {{
    background: linear-gradient(135deg,#667eea 0%,#764ba2 100%);
    color: #fff; padding: 30px 40px; display: flex; align-items: center; justify-content: space-between;
}}
.header h1 {{ font-size: 28px; font-weight: 600; }}
.header .subtitle {{ font-size: 14px; opacity: 0.85; margin-top: 4px; }}
.filter-bar {{
    display: flex; align-items: center; gap: 10px; padding: 16px 30px;
    background: #f0f2f5; border-bottom: 1px solid #e0e0e0; flex-wrap: wrap;
}}
.filter-bar label {{ font-size: 14px; color: #555; font-weight: 600; margin-right: 4px; }}
.filter-btn {{
    padding: 6px 16px; border: 1px solid #ccc; border-radius: 20px;
    background: #fff; color: #555; cursor: pointer; font-size: 13px;
    transition: all 0.2s;
}}
.filter-btn:hover {{ border-color: #667eea; color: #667eea; }}
.filter-btn.active {{ background: #667eea; color: #fff; border-color: #667eea; }}
.custom-date {{
    display: none; align-items: center; gap: 6px;
}}
.custom-date.show {{ display: flex; }}
.custom-date input {{
    padding: 5px 10px; border: 1px solid #ccc; border-radius: 6px;
    font-size: 13px; outline: none;
}}
.custom-date input:focus {{ border-color: #667eea; }}
.custom-date button {{
    padding: 5px 14px; background: #667eea; color: #fff; border: none;
    border-radius: 6px; cursor: pointer; font-size: 13px;
}}
.breadcrumb {{
    padding: 12px 30px; background: #fafbfc; border-bottom: 1px solid #eee;
    display: flex; align-items: center; gap: 8px; font-size: 14px;
}}
.breadcrumb span {{ color: #999; cursor: pointer; }}
.breadcrumb span:hover {{ color: #667eea; }}
.breadcrumb span.current {{ color: #333; font-weight: 600; cursor: default; }}
.breadcrumb .sep {{ color: #ccc; }}
.stats-grid {{
    display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    gap: 16px; padding: 24px 30px; background: #f8f9fa;
}}
.stat-card {{
    background: #fff; padding: 20px; border-radius: 10px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08); text-align: center;
}}
.stat-card .label {{ color: #888; font-size: 13px; margin-bottom: 6px; text-transform: uppercase; letter-spacing: 1px; }}
.stat-card .value {{ font-size: 24px; font-weight: bold; }}
.stat-card .value.profit {{ color: #e74c3c; }}
.stat-card .value.loss {{ color: #27ae60; }}
.stat-card .value.neutral {{ color: #3498db; }}
.charts-area {{ padding: 24px 30px; }}
.chart-section {{ margin-bottom: 32px; }}
.chart-title {{
    font-size: 20px; color: #2c3e50; margin-bottom: 14px;
    padding-left: 12px; border-left: 4px solid #667eea;
}}
.chart-wrap {{ width: 100%; height: 420px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }}
.two-charts {{ display: grid; grid-template-columns: 1fr 1fr; gap: 24px; }}
.two-charts .chart-wrap {{ height: 380px; }}
.detail-section {{ padding: 0 30px 30px; }}
.detail-table {{
    width: 100%; border-collapse: collapse; font-size: 14px;
    border-radius: 8px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.06);
}}
.detail-table thead {{ background: linear-gradient(135deg,#667eea 0%,#764ba2 100%); }}
.detail-table th {{ color: #fff; padding: 12px 8px; text-align: center; font-size: 13px; }}
.detail-table td {{ padding: 10px 8px; text-align: center; border-bottom: 1px solid #f0f0f0; }}
.detail-table tbody tr:hover {{ background: #f8f9fa; }}
.detail-table .profit-cell {{ color: #e74c3c; font-weight: 600; }}
.detail-table .loss-cell {{ color: #27ae60; font-weight: 600; }}
.empty-state {{ text-align: center; padding: 60px 20px; color: #aaa; font-size: 16px; }}
.footer {{
    padding: 16px; text-align: center; color: #999; font-size: 13px;
    background: #f8f9fa; border-top: 1px solid #eee;
}}
@media (max-width: 768px) {{
    .two-charts {{ grid-template-columns: 1fr; }}
    .header {{ flex-direction: column; gap: 10px; text-align: center; }}
    .stats-grid {{ grid-template-columns: repeat(2, 1fr); }}
    .filter-bar {{ justify-content: center; }}
}}
    </style>
</head>
<body>
<div class="container">
    <div class="header">
        <div>
            <h1>📊 股票交易汇总分析</h1>
            <div class="subtitle">点击图表数据项可钻取下级明细</div>
        </div>
    </div>
    <div class="filter-bar">
        <label>时间范围：</label>
        <button class="filter-btn active" data-type="mtd" onclick="setFilter('mtd')">本月</button>
        <button class="filter-btn" data-type="ytd" onclick="setFilter('ytd')">本年</button>
        <button class="filter-btn" data-type="all" onclick="setFilter('all')">全部</button>
        <button class="filter-btn" data-type="custom" onclick="setFilter('custom')">自定义</button>
        <div class="custom-date" id="customDate">
            <input type="date" id="startDate" onchange="applyCustomFilter()">
            <span style="color:#aaa">至</span>
            <input type="date" id="endDate" onchange="applyCustomFilter()">
            <button onclick="applyCustomFilter()">确定</button>
        </div>
    </div>
    <div class="breadcrumb" id="breadcrumb"></div>
    <div class="stats-grid" id="statsGrid"></div>
    <div class="charts-area" id="chartsArea"></div>
    <div class="detail-section" id="detailSection"></div>
    <div class="footer">
        <p>报告生成时间：{gen_time} | 数据来源：股票交易盈亏汇总.xlsx</p>
    </div>
</div>
<script>
const ALL_DATA = {data_json};
const NOW = '{now_str}';
let state = {{ view: 'overview', filterType: 'mtd', customStart: '', customEnd: '', drillParam: null }};
let navStack = [];
const chartMap = {{}};
function fmtMoney(n, sign) {{
    const s = n >= 0 ? (sign ? '+' : '') : '';
    return s + '¥' + Math.abs(n).toLocaleString('zh-CN', {{minimumFractionDigits:2, maximumFractionDigits:2}});
}}
function fmtPct(n, sign) {{
    const s = n >= 0 ? (sign ? '+' : '') : '';
    return s + n.toFixed(2) + '%';
}}
function profitColor(v) {{ return v >= 0 ? '#e74c3c' : '#27ae60'; }}
function profitCls(v) {{ return v >= 0 ? 'profit' : 'loss'; }}
function groupBy(arr, keyFn) {{
    const getKey = typeof keyFn === 'function' ? keyFn : r => r[keyFn];
    const m = {{}};
    arr.forEach(r => {{ const k = getKey(r); if (!m[k]) m[k] = []; m[k].push(r); }});
    return m;
}}
function sumField(arr, f) {{ return arr.reduce((s,r) => s + (typeof f === 'function' ? f(r) : r[f]), 0); }}
function unique(arr, f) {{ const fn = typeof f === 'function' ? f : r => r[f]; return [...new Set(arr.map(fn))]; }}
function disposeAll() {{
    Object.values(chartMap).forEach(c => {{ try {{ c.dispose(); }} catch(e){{}} }});
    Object.keys(chartMap).forEach(k => delete chartMap[k]);
}}
function mkChart(id) {{
    const dom = document.getElementById(id);
    if (!dom) return null;
    const c = echarts.init(dom);
    chartMap[id] = c;
    return c;
}}
function getFilteredRecords() {{
    let recs = [...ALL_DATA];
    if (state.filterType === 'mtd') {{
        const m = NOW.substring(0,7);
        recs = recs.filter(r => r.date.startsWith(m));
    }} else if (state.filterType === 'ytd') {{
        const y = NOW.substring(0,4);
        recs = recs.filter(r => r.date.startsWith(y));
    }} else if (state.filterType === 'custom' && state.customStart && state.customEnd) {{
        recs = recs.filter(r => r.date >= state.customStart && r.date <= state.customEnd);
    }}
    return recs;
}}
function setFilter(type) {{
    state.filterType = type;
    document.querySelectorAll('.filter-btn').forEach(b => {{
        b.classList.toggle('active', b.dataset.type === type);
    }});
    document.getElementById('customDate').classList.toggle('show', type === 'custom');
    if (type !== 'custom') {{
        state.view = 'overview'; state.drillParam = null; navStack = [];
    }}
    render();
}}
function applyCustomFilter() {{
    state.customStart = document.getElementById('startDate').value;
    state.customEnd = document.getElementById('endDate').value;
    if (state.customStart && state.customEnd) {{
        state.view = 'overview'; state.drillParam = null; navStack = []; render();
    }}
}}
function drillTo(view, param) {{
    navStack.push({{view: state.view, param: state.drillParam}});
    state.view = view; state.drillParam = param; render();
    window.scrollTo({{top: 0, behavior: 'smooth'}});
}}
function goBack() {{
    if (navStack.length === 0) return;
    const prev = navStack.pop();
    state.view = prev.view; state.drillParam = prev.param; render();
    window.scrollTo({{top: 0, behavior: 'smooth'}});
}}
function renderBreadcrumb() {{
    const el = document.getElementById('breadcrumb');
    let html = '<span onclick="goBackOverview()">汇总概览</span>';
    if (state.view === 'month') {{
        html += '<span class="sep">›</span><span class="current">' + state.drillParam + ' 月度明细</span>';
    }} else if (state.view === 'day') {{
        html += '<span class="sep">›</span><span onclick="goBackToMonthFromDay()">月度</span><span class="sep">›</span><span class="current">' + state.drillParam + ' 日度明细</span>';
    }} else if (state.view === 'stock') {{
        html += '<span class="sep">›</span><span class="current">' + state.drillParam + ' 个股分析</span>';
    }}
    el.innerHTML = html;
}}
function goBackOverview() {{ state.view = 'overview'; state.drillParam = null; navStack = []; render(); }}
function goBackToMonthFromDay() {{
    const month = state.drillParam ? state.drillParam.substring(0,7) : NOW.substring(0,7);
    state.view = 'month'; state.drillParam = month; render();
}}
function renderStats(cards) {{
    const el = document.getElementById('statsGrid');
    el.innerHTML = cards.map(c => '<div class="stat-card"><div class="label">' + c.label + '</div><div class="value ' + c.cls + '">' + c.value + '</div></div>').join('');
}}
function computeStats(recs) {{
    const tp = sumField(recs, 'profit');
    const tb = sumField(recs, 'buyAmount');
    const ts = sumField(recs, 'sellAmount');
    const tr = tb > 0 ? tp / tb * 100 : 0;
    const dates = unique(recs, 'date');
    const stocks = unique(recs, 'name');
    const winCount = recs.filter(r => r.profit > 0).length;
    const loseCount = recs.filter(r => r.profit < 0).length;
    return [
        {{label:'交易天数', value:dates.length, cls:'neutral'}},
        {{label:'交易笔数', value:recs.length, cls:'neutral'}},
        {{label:'涉及股票', value:stocks.length, cls:'neutral'}},
        {{label:'盈利笔数', value:winCount, cls:'profit'}},
        {{label:'亏损笔数', value:loseCount, cls:'loss'}},
        {{label:'胜率', value: recs.length > 0 ? (winCount/recs.length*100).toFixed(1)+'%' : 'N/A', cls:'neutral'}},
        {{label:'总买入金额', value: fmtMoney(tb), cls:'neutral'}},
        {{label:'总卖出金额', value: fmtMoney(ts), cls:'neutral'}},
        {{label:'总盈亏', value: fmtMoney(tp,true), cls: profitCls(tp)}},
        {{label:'总收益率', value: fmtPct(tr,true), cls: profitCls(tr)}},
    ];
}}
function renderDetailTable(recs, title) {{
    const el = document.getElementById('detailSection');
    if (recs.length === 0) {{ el.innerHTML = ''; return; }}
    el.innerHTML = '<h2 class="chart-title" style="margin-bottom:14px">' + title + '</h2>' +
        '<table class="detail-table"><thead><tr><th>日期</th><th>来源</th><th>证券名称</th><th>匹配数量</th><th>买入均价</th><th>卖出均价</th><th>买入金额</th><th>卖出金额</th><th>盈亏金额</th><th>盈亏比例</th></tr></thead><tbody>' +
        recs.map(r => {{
            const pc = r.profit >= 0 ? 'profit-cell' : 'loss-cell';
            return '<tr><td>' + r.date + '</td><td>' + r.source + '</td><td><b>' + r.name + '</b></td><td>' + r.matchQty + '</td><td>¥' + r.buyPrice.toFixed(4) + '</td><td>¥' + r.sellPrice.toFixed(4) + '</td><td>' + fmtMoney(r.buyAmount) + '</td><td>' + fmtMoney(r.sellAmount) + '</td><td class="' + pc + '">' + fmtMoney(r.profit,true) + '</td><td class="' + pc + '">' + r.profitPct + '</td></tr>';
        }}).join('') + '</tbody></table>';
}}
function renderOverview() {{
    const recs = getFilteredRecords();
    if (recs.length === 0) {{
        disposeAll(); renderBreadcrumb();
        document.getElementById('statsGrid').innerHTML = '';
        document.getElementById('chartsArea').innerHTML = '<div class="empty-state">当前筛选条件下暂无数据</div>';
        document.getElementById('detailSection').innerHTML = '';
        return;
    }}
    renderStats(computeStats(recs));
    const monthGroups = groupBy(recs, r => r.date.substring(0,7));
    const months = Object.keys(monthGroups).sort();
    const monthProfits = months.map(m => Math.round(sumField(monthGroups[m],'profit')*100)/100);
    const dayGroups = groupBy(recs, 'date');
    const days = Object.keys(dayGroups).sort();
    const dayProfits = days.map(d => Math.round(sumField(dayGroups[d],'profit')*100)/100);
    let cumulative = [], run = 0;
    dayProfits.forEach(p => {{ run += p; cumulative.push(Math.round(run*100)/100); }});
    const stockGroups = groupBy(recs, 'name');
    const stockList = Object.keys(stockGroups).map(n => ({{
        name: n, profit: Math.round(sumField(stockGroups[n],'profit')*100)/100,
        count: stockGroups[n].length
    }})).sort((a,b) => b.profit - a.profit);
    const chartsArea = document.getElementById('chartsArea');
    chartsArea.innerHTML = `
        <div class="chart-section"><h2 class="chart-title">📈 月度盈亏趋势 <span style="font-size:13px;color:#999;font-weight:normal">（点击柱体钻取月度明细）</span></h2><div id="chartMonth" class="chart-wrap"></div></div>
        <div class="chart-section"><h2 class="chart-title">📅 每日盈亏趋势 <span style="font-size:13px;color:#999;font-weight:normal">（点击柱体钻取日度明细）</span></h2><div id="chartDaily" class="chart-wrap"></div></div>
        <div class="chart-section"><h2 class="chart-title">💰 累计收益曲线</h2><div id="chartCumulative" class="chart-wrap"></div></div>
        <div class="chart-section"><h2 class="chart-title">🎯 各股票盈亏分析 <span style="font-size:13px;color:#999;font-weight:normal">（点击柱体/饼块查看个股明细）</span></h2><div class="two-charts"><div id="chartStockBar" class="chart-wrap"></div><div id="chartStockPie" class="chart-wrap"></div></div></div>`;
    document.getElementById('detailSection').innerHTML = '';
    const mc = mkChart('chartMonth');
    mc.setOption({{
        tooltip: {{ trigger:'axis', formatter: p => p[0].name+'月<br/>盈亏：'+fmtMoney(p[0].value,true) }},
        xAxis: {{ type:'category', data: months.map(m => m+'月'), axisLabel:{{color:'#555'}} }},
        yAxis: {{ type:'value', name:'盈亏（元）', axisLabel:{{formatter:'¥{{value}}'}} }},
        series: [{{ type:'bar', data: monthProfits, itemStyle:{{ color: p => profitColor(p.value) }}, label:{{ show:true, position:'top', formatter: p => fmtMoney(p.value,true), color:'#333', fontSize:11 }} }}]
    }});
    mc.on('click', p => drillTo('month', p.name.replace('月','')));
    const dc = mkChart('chartDaily');
    dc.setOption({{
        tooltip: {{ trigger:'axis', formatter: p => p[0].name+'<br/>盈亏：'+fmtMoney(p[0].value,true) }},
        xAxis: {{ type:'category', data: days, axisLabel:{{color:'#555', rotate:45}} }},
        yAxis: {{ type:'value', name:'盈亏（元）', axisLabel:{{formatter:'¥{{value}}'}} }},
        series: [{{ type:'bar', data: dayProfits, itemStyle:{{ color: p => profitColor(p.value) }}, label:{{ show:true, position:'top', formatter: p => fmtMoney(p.value,true), color:'#333', fontSize:11 }} }}]
    }});
    dc.on('click', p => drillTo('day', p.name));
    const cc = mkChart('chartCumulative');
    const maxCum = Math.max(...cumulative), minCum = Math.min(...cumulative);
    let markPts = [];
    if (cumulative.length > 0) {{
        markPts.push({{type:'max', name:'最高 ¥'+maxCum.toFixed(0)}});
        markPts.push({{type:'min', name:'最低 ¥'+minCum.toFixed(0)}});
    }}
    cc.setOption({{
        tooltip: {{ trigger:'axis', formatter: p => p[0].name+'<br/>累计：'+fmtMoney(p[0].value,true) }},
        xAxis: {{ type:'category', data: days, axisLabel:{{color:'#555', rotate:45}} }},
        yAxis: {{ type:'value', name:'累计收益（元）', axisLabel:{{formatter:'¥{{value}}'}} }},
        series: [{{ type:'line', data: cumulative, smooth:true, itemStyle:{{ color:'#667eea' }}, areaStyle:{{ color:{{ type:'linear', x:0,y:0,x2:0,y2:1, colorStops:[{{offset:0,color:'rgba(102,126,234,0.35)'}},{{offset:1,color:'rgba(102,126,234,0.02)'}}]}}}}, markPoint:{{ data: markPts }}, markLine:{{ data:[{{type:'average',name:'平均'}}] }} }}]
    }});
    const sb = mkChart('chartStockBar');
    sb.setOption({{
        tooltip: {{ trigger:'axis', formatter: p => p[0].name+'<br/>盈亏：'+fmtMoney(p[0].value,true) }},
        grid: {{ left:'3%', right:'4%', bottom:'15%', containLabel:true }},
        xAxis: {{ type:'category', data: stockList.map(s=>s.name), axisLabel:{{color:'#555', rotate:30}} }},
        yAxis: {{ type:'value', name:'盈亏（元）', axisLabel:{{formatter:'¥{{value}}'}} }},
        series: [{{ type:'bar', data: stockList.map(s=>s.profit), itemStyle:{{ color: p => profitColor(p.value) }}, label:{{ show:true, position:'top', formatter: p => fmtMoney(p.value,true), color:'#333', fontSize:10 }} }}]
    }});
    sb.on('click', p => drillTo('stock', p.name));
    const sp = mkChart('chartStockPie');
    const pieData = stockList.filter(s => s.profit > 0).map(s => ({{name:s.name, value:s.profit}}));
    const loseData = stockList.filter(s => s.profit < 0);
    if (loseData.length > 0) {{
        pieData.push({{name:'亏损合计', value: Math.abs(loseData.reduce((s,x)=>s+x.profit,0)), itemStyle:{{color:'#27ae60'}}}});
    }}
    sp.setOption({{
        tooltip:{{ trigger:'item', formatter:'{{b}}<br/>¥{{c}} ({{d}}%)' }},
        series:[{{ type:'pie', radius:['35%','65%'], center:['50%','55%'], itemStyle:{{ borderRadius:8, borderColor:'#fff', borderWidth:2 }}, label:{{ show:true, formatter:'{{b}}\\n¥{{c}}' }}, emphasis:{{ label:{{ show:true, fontSize:14, fontWeight:'bold' }} }}, data: pieData }}]
    }});
    sp.on('click', p => {{ if (p.name === '亏损合计') return; drillTo('stock', p.name); }});
}}
function renderMonthView() {{
    const allFiltered = getFilteredRecords();
    const month = state.drillParam;
    const recs = allFiltered.filter(r => r.date.startsWith(month));
    if (recs.length === 0) {{ goBackOverview(); return; }}
    const monthProfit = sumField(recs, 'profit');
    const monthBuy = sumField(recs, 'buyAmount');
    const monthDays = unique(recs, 'date');
    const monthStocks = unique(recs, 'name');
    renderStats([
        {{label:'月份', value: month, cls:'neutral'}},
        {{label:'交易天数', value: monthDays.length, cls:'neutral'}},
        {{label:'交易笔数', value: recs.length, cls:'neutral'}},
        {{label:'涉及股票', value: monthStocks.length, cls:'neutral'}},
        {{label:'月度盈亏', value: fmtMoney(monthProfit,true), cls: profitCls(monthProfit)}},
        {{label:'月度收益率', value: fmtPct(monthBuy>0?monthProfit/monthBuy*100:0,true), cls: profitCls(monthProfit)}},
    ]);
    const dayGroups = groupBy(recs, 'date');
    const days = Object.keys(dayGroups).sort();
    const dayProfits = days.map(d => Math.round(sumField(dayGroups[d],'profit')*100)/100);
    let cum = [], run = 0;
    dayProfits.forEach(p => {{ run += p; cum.push(Math.round(run*100)/100); }});
    const chartsArea = document.getElementById('chartsArea');
    chartsArea.innerHTML = `
        <div class="chart-section"><h2 class="chart-title">📅 ${{month}} 每日盈亏 <span style="font-size:13px;color:#999;font-weight:normal">（点击柱体查看当日明细）</span></h2><div id="chartMonthDaily" class="chart-wrap"></div></div>
        <div class="chart-section"><h2 class="chart-title">💰 ${{month}} 月内累计收益</h2><div id="chartMonthCum" class="chart-wrap"></div></div>`;
    const dc = mkChart('chartMonthDaily');
    dc.setOption({{
        tooltip:{{ trigger:'axis', formatter: p => p[0].name+'<br/>盈亏：'+fmtMoney(p[0].value,true) }},
        xAxis:{{ type:'category', data: days.map(d=>d.substring(8)), axisLabel:{{color:'#555'}} }},
        yAxis:{{ type:'value', name:'盈亏（元）', axisLabel:{{formatter:'¥{{value}}'}} }},
        series:[{{ type:'bar', data: dayProfits, itemStyle:{{ color: p => profitColor(p.value) }}, label:{{ show:true, position:'top', formatter: p => fmtMoney(p.value,true), color:'#333', fontSize:11 }} }}]
    }});
    dc.on('click', p => drillTo('day', month + '-' + p.name));
    const ccc = mkChart('chartMonthCum');
    ccc.setOption({{
        tooltip:{{ trigger:'axis', formatter: p => p[0].name+'<br/>累计：'+fmtMoney(p[0].value,true) }},
        xAxis:{{ type:'category', data: days.map(d=>d.substring(8)), axisLabel:{{color:'#555'}} }},
        yAxis:{{ type:'value', name:'累计收益（元）', axisLabel:{{formatter:'¥{{value}}'}} }},
        series:[{{ type:'line', data: cum, smooth:true, itemStyle:{{ color:'#667eea' }}, areaStyle:{{ color:{{ type:'linear', x:0,y:0,x2:0,y2:1, colorStops:[{{offset:0,color:'rgba(102,126,234,0.3)'}},{{offset:1,color:'rgba(102,126,234,0.02)'}}]}}}}, markPoint:{{ data:[{{type:'max',name:'最高'}},{{type:'min',name:'最低'}}] }} }}]
    }});
    renderDetailTable(recs, '📋 ' + month + ' 全部交易明细');
}}
function renderDayView() {{
    const allFiltered = getFilteredRecords();
    const date = state.drillParam;
    const recs = allFiltered.filter(r => r.date === date);
    if (recs.length === 0) {{ goBack(); return; }}
    const dayProfit = sumField(recs, 'profit');
    const dayBuy = sumField(recs, 'buyAmount');
    const dayRate = dayBuy > 0 ? dayProfit / dayBuy * 100 : 0;
    const winCount = recs.filter(r => r.profit > 0).length;
    renderStats([
        {{label:'日期', value: date, cls:'neutral'}},
        {{label:'交易笔数', value: recs.length, cls:'neutral'}},
        {{label:'盈利笔数', value: winCount, cls:'profit'}},
        {{label:'亏损笔数', value: recs.length - winCount, cls:'loss'}},
        {{label:'当日盈亏', value: fmtMoney(dayProfit,true), cls: profitCls(dayProfit)}},
        {{label:'当日收益率', value: fmtPct(dayRate,true), cls: profitCls(dayRate)}},
    ]);
    const sorted = [...recs].sort((a,b) => b.profit - a.profit);
    const chartsArea = document.getElementById('chartsArea');
    chartsArea.innerHTML = `<div class="chart-section"><h2 class="chart-title">📊 ${{date}} 各股票盈亏</h2><div id="chartDayStock" class="chart-wrap" style="height:350px"></div></div>`;
    const sc = mkChart('chartDayStock');
    sc.setOption({{
        tooltip:{{ trigger:'axis', formatter: p => p[0].name+'<br/>盈亏：'+fmtMoney(p[0].value,true) }},
        grid:{{ left:'3%', right:'4%', bottom:'10%', containLabel:true }},
        xAxis:{{ type:'category', data: sorted.map(r=>r.name), axisLabel:{{color:'#555', rotate:20}} }},
        yAxis:{{ type:'value', name:'盈亏（元）', axisLabel:{{formatter:'¥{{value}}'}} }},
        series:[{{ type:'bar', data: sorted.map(r=>r.profit), itemStyle:{{ color: p => profitColor(p.value) }}, label:{{ show:true, position:'top', formatter: p => fmtMoney(p.value,true), color:'#333', fontSize:12 }} }}]
    }});
    renderDetailTable(recs, '📋 ' + date + ' 交易明细');
}}
function renderStockView() {{
    const allFiltered = getFilteredRecords();
    const stockName = state.drillParam;
    const recs = allFiltered.filter(r => r.name === stockName);
    if (recs.length === 0) {{ goBack(); return; }}
    const totalProfit = sumField(recs, 'profit');
    const totalBuy = sumField(recs, 'buyAmount');
    const totalRate = totalBuy > 0 ? totalProfit / totalBuy * 100 : 0;
    const tradeDays = unique(recs, 'date');
    const winCount = recs.filter(r => r.profit > 0).length;
    renderStats([
        {{label:'股票名称', value: stockName, cls:'neutral'}},
        {{label:'证券代码', value: recs[0].code, cls:'neutral'}},
        {{label:'交易次数', value: recs.length, cls:'neutral'}},
        {{label:'交易天数', value: tradeDays.length, cls:'neutral'}},
        {{label:'盈利次数', value: winCount, cls:'profit'}},
        {{label:'总盈亏', value: fmtMoney(totalProfit,true), cls: profitCls(totalProfit)}},
        {{label:'总收益率', value: fmtPct(totalRate,true), cls: profitCls(totalRate)}},
        {{label:'平均每次盈亏', value: fmtMoney(totalProfit/recs.length,true), cls: profitCls(totalProfit)}},
    ]);
    const dayGroups = groupBy(recs, 'date');
    const days = Object.keys(dayGroups).sort();
    const dayProfits = days.map(d => Math.round(sumField(dayGroups[d],'profit')*100)/100);
    let cum = [], run = 0;
    dayProfits.forEach(p => {{ run += p; cum.push(Math.round(run*100)/100); }});
    const chartsArea = document.getElementById('chartsArea');
    const singleDay = days.length === 1;
    if (singleDay) {{
        chartsArea.innerHTML = `<div class="chart-section"><h2 class="chart-title">📊 ${{stockName}} 盈亏分布</h2><div id="chartStockSingle" class="chart-wrap" style="height:300px"></div></div>`;
        const sc = mkChart('chartStockSingle');
        sc.setOption({{
            tooltip:{{ trigger:'item', formatter:'{{b}}<br/>¥{{c}}' }},
            series:[{{ type:'pie', radius:['40%','65%'], data: recs.map(r => ({{ name: r.date, value: r.profit, itemStyle:{{ color: profitColor(r.profit) }} }})), label:{{ formatter:'{{b}}\\n{{c}}' }} }}]
        }});
    }} else {{
        chartsArea.innerHTML = `
            <div class="chart-section"><h2 class="chart-title">📅 ${{stockName}} 每日盈亏 <span style="font-size:13px;color:#999;font-weight:normal">（点击柱体查看当日明细）</span></h2><div id="chartStockDaily" class="chart-wrap"></div></div>
            <div class="chart-section"><h2 class="chart-title">💰 ${{stockName}} 累计收益曲线</h2><div id="chartStockCum" class="chart-wrap"></div></div>`;
        const dc = mkChart('chartStockDaily');
        dc.setOption({{
            tooltip:{{ trigger:'axis', formatter: p => p[0].name+'<br/>盈亏：'+fmtMoney(p[0].value,true) }},
            xAxis:{{ type:'category', data: days, axisLabel:{{color:'#555', rotate:45}} }},
            yAxis:{{ type:'value', name:'盈亏（元）', axisLabel:{{formatter:'¥{{value}}'}} }},
            series:[{{ type:'bar', data: dayProfits, itemStyle:{{ color: p => profitColor(p.value) }}, label:{{ show:true, position:'top', formatter: p => fmtMoney(p.value,true), color:'#333', fontSize:11 }} }}]
        }});
        dc.on('click', p => drillTo('day', p.name));
        const cc = mkChart('chartStockCum');
        cc.setOption({{
            tooltip:{{ trigger:'axis', formatter: p => p[0].name+'<br/>累计：'+fmtMoney(p[0].value,true) }},
            xAxis:{{ type:'category', data: days, axisLabel:{{color:'#555', rotate:45}} }},
            yAxis:{{ type:'value', name:'累计收益（元）', axisLabel:{{formatter:'¥{{value}}'}} }},
            series:[{{ type:'line', data: cum, smooth:true, itemStyle:{{ color:'#667eea' }}, areaStyle:{{ color:{{ type:'linear', x:0,y:0,x2:0,y2:1, colorStops:[{{offset:0,color:'rgba(102,126,234,0.3)'}},{{offset:1,color:'rgba(102,126,234,0.02)'}}]}}}}, markPoint:{{ data:[{{type:'max',name:'最高'}},{{type:'min',name:'最低'}}] }} }}]
        }});
    }}
    renderDetailTable(recs, '📋 ' + stockName + ' 全部交易记录');
}}
function render() {{
    disposeAll(); renderBreadcrumb();
    switch (state.view) {{
        case 'overview': renderOverview(); break;
        case 'month':   renderMonthView(); break;
        case 'day':     renderDayView(); break;
        case 'stock':   renderStockView(); break;
    }}
}}
window.addEventListener('resize', () => {{
    Object.values(chartMap).forEach(c => {{ try {{ c.resize(); }} catch(e){{}} }});
}});
render();
</script>
</body>
</html>"""


# ==================== 辅助函数 ====================

def archive_file(input_file):
    """归档原始文件"""
    filename = os.path.basename(input_file)
    archive_path = os.path.join(HISTORY_DIR, filename)

    if os.path.exists(archive_path):
        base, ext = os.path.splitext(filename)
        timestamp = datetime.now().strftime('%H%M%S')
        filename = f"{base}_{timestamp}{ext}"
        archive_path = os.path.join(HISTORY_DIR, filename)

    shutil.move(input_file, archive_path)
    print(f"原始文件已归档：{archive_path}")


def find_input_files():
    """查找待处理的所有输入文件（Excel + 图片）"""
    excel_files = [f for f in glob.glob('*.xlsx')
                   if f not in ['股票交易盈亏汇总.xlsx']
                   and not f.startswith('~$')]

    image_extensions = ['*.png', '*.jpg', '*.jpeg']
    image_files = []
    for ext in image_extensions:
        image_files.extend(glob.glob(ext))

    return excel_files, image_files


# ==================== 主程序 ====================

def main():
    print("\n" + "="*80)
    print("股票交易盈亏分析系统 v4.1")
    print("支持输入：Excel文件（券商导出）+ 图片文件（手机App截图）")
    print("="*80)

    excel_files, image_files = find_input_files()
    all_files = excel_files + image_files

    if not all_files:
        print("\n未找到待处理的文件")
        print("请将以下类型文件放到当前文件夹：")
        print("  - Excel文件：券商导出的交易记录（如：2026-04-22-两融-当日成交汇总.xlsx）")
        print("  - 图片文件：手机App截图（如：2026-04-22-手机交易.png）")
        generate_summary_html()
        return

    print(f"\n找到 {len(excel_files)} 个Excel文件，{len(image_files)} 个图片文件")

    # 按日期分组处理（同一天可能有Excel+图片两种输入）
    files_by_date = {}

    for f in excel_files:
        d = extract_date_from_filename(f)
        if d not in files_by_date:
            files_by_date[d] = {'excel': [], 'image': []}
        files_by_date[d]['excel'].append(f)

    for f in image_files:
        d = extract_date_from_filename(f)
        if d not in files_by_date:
            files_by_date[d] = {'excel': [], 'image': []}
        files_by_date[d]['image'].append(f)

    processed_dates = []

    # 按日期顺序处理
    for trading_date in sorted(files_by_date.keys()):
        date_files = files_by_date[trading_date]

        print(f"\n{'='*80}")
        print(f"处理日期：{trading_date}")
        print(f"  Excel文件：{len(date_files['excel'])} 个")
        print(f"  图片文件：{len(date_files['image'])} 个")
        print('='*80)

        # 处理Excel文件
        for excel_file in date_files['excel']:
            try:
                result_df, _, _, source = process_excel_file(excel_file)
                if len(result_df) > 0:
                    append_to_excel(result_df, trading_date, source)
                archive_file(excel_file)
            except Exception as e:
                print(f"\n[错误] 处理Excel文件 {excel_file} 时出错：{str(e)}")
                continue

        # 处理图片文件
        for image_file in date_files['image']:
            try:
                result_df, _, _, source = process_image_file(image_file)
                if len(result_df) > 0:
                    append_to_excel(result_df, trading_date, source)
                archive_file(image_file)
            except Exception as e:
                print(f"\n[错误] 处理图片文件 {image_file} 时出错：{str(e)}")
                continue

        # 从汇总文件提取当日全部数据生成单日报告
        generate_html_report_from_summary(trading_date)
        processed_dates.append(trading_date)

        print(f"\n日期 {trading_date} 处理完成")

    # 生成汇总可视化报告
    generate_summary_html()

    print("\n" + "="*80)
    print("[完成] 所有文件处理完成！")
    print("="*80)


if __name__ == "__main__":
    main()
