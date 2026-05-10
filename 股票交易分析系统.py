"""
股票交易盈亏分析系统
功能：自动分析股票交易记录，计算盈亏并生成报告
支持三种输入：
  1. Excel文件（券商导出）— 文件名格式：YYYY-MM-DD-两融-当日成交汇总.xlsx
  2. 图片文件（手机App截图）— 文件名格式：YYYY-MM-DD-手机交易.png/.jpg/.jpeg
  3. 图片文件（平安证券截图）— 文件名格式：YYYYMMDD_平安.png
作者：WorkBuddy
版本：v4.5
更新日期：2026-05-10

使用说明：
1. 将券商导出的Excel交易记录文件、手机App截图或平安证券截图放到当前文件夹
2. 运行此脚本，自动处理所有未归档的文件
3. 处理后的原始文件会自动归档到history文件夹
4. HTML报告生成到reports文件夹（单日报告 + 汇总可视化报告）
5. Excel汇总文件支持去重更新（同日期+同来源的数据会覆盖）
6. 汇总可视化报告支持时间筛选和多级数据钻取
7. 同一天可以同时有Excel、手机截图和平安截图多种输入，数据会自动合并
8. v4.3新增：佣金（万一/双向/最低5元）和印花税（万五/卖出单边）计算
9. v4.4新增：个股累计盈亏总览标签页、盈亏日历热力图（含连续亏损预警）
10. v4.5新增：跨账户同股配对、单边交易标记（未平仓）、连续亏损/回撤统计

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

# 交易成本配置
COMMISSION_RATE = 0.0001      # 佣金费率：万一（0.01%），双向征收
MIN_COMMISSION = 5             # 最低佣金：5元/笔
STAMP_DUTY_RATE = 0.0005      # 印花税费率：万五（0.05%），仅卖出单边征收


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
    elif fname_lower.endswith('.xls'):
        if '平安' in fname_lower:
            return '平安账户'
        return '两融账户'
    elif any(fname_lower.endswith(ext) for ext in ['.png', '.jpg', '.jpeg']):
        if '平安' in fname_lower:
            return '平安账户'
        return '手机账户'
    return '未知'


# ==================== 图片OCR解析 ====================

def parse_image_trades(image_path):
    """用手机App截图识别交易记录 - v2.0 改进版
    策略：不按行分组，而是提取所有关键信息后按股票代码聚合
    """
    # 修复Windows下torch/getpass环境变量问题
    import os as _os
    if not _os.environ.get('USERNAME'):
        _os.environ['USERNAME'] = 'Ryan'
    if not _os.environ.get('USER'):
        _os.environ['USER'] = 'Ryan'

    try:
        import easyocr
    except ImportError:
        print(f"[错误] 缺少OCR依赖，请运行：pip install easyocr")
        return pd.DataFrame()

    print(f"  正在识别图片：{os.path.basename(image_path)}")
    reader = easyocr.Reader(['ch_sim', 'en'], gpu=False, verbose=False)
    result = reader.readtext(image_path)

    # 提取所有文本项，保留位置信息
    items = []
    for bbox, text, conf in result:
        y_center = (bbox[0][1] + bbox[2][1]) / 2
        x_center = (bbox[0][0] + bbox[2][0]) / 2
        items.append({'text': text.strip(), 'y': y_center, 'x': x_center, 'conf': conf})

    items.sort(key=lambda r: r['y'])

    # 第一步：按y坐标分组（行间距阈值40px，更宽松）
    rows = []
    current_row = []
    last_y = None
    y_threshold = 40

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

    # 第二步：从每行提取关键信息
    raw_records = []
    for row in rows:
        texts = [r['text'] for r in row]
        text_combined = ' '.join(texts)

        # 提取股票代码（6位数字）
        code_match = re.search(r'\b(\d{6})\b', text_combined)
        if not code_match:
            continue
        stock_code = code_match.group(1)

        # 提取方向
        direction = None
        if '买入' in text_combined or '买' in text_combined:
            direction = '证券买入'
        elif '卖出' in text_combined or '卖' in text_combined:
            direction = '证券卖出'
        else:
            continue

        # 提取股票名称：在代码之前或之后的中文
        stock_name = ''
        code_idx = text_combined.find(stock_code)
        
        # 尝试代码前面
        if code_idx > 0:
            prefix = text_combined[:code_idx].strip()
            prefix = re.sub(r'\d{1,2}:\d{2}:\d{2}', '', prefix)
            prefix = re.sub(r'买入|卖出|买|卖|成交', '', prefix)
            prefix = prefix.strip()
            cn_match = re.findall(r'[\u4e00-\u9fff]+', prefix)
            if cn_match:
                stock_name = cn_match[-1]
        
        # 如果前面没找到，尝试代码后面
        if not stock_name:
            suffix = text_combined[code_idx + 6:].strip()  # 代码后6位之后
            suffix = re.sub(r'\d+\.?\d*', '', suffix)  # 移除数字
            suffix = re.sub(r'买入|卖出|买|卖|成交', '', suffix)
            suffix = suffix.strip()
            cn_match = re.findall(r'[\u4e00-\u9fff]+', suffix)
            if cn_match:
                stock_name = cn_match[0]  # 取第一个

        # 提取数字（价格、金额、数量）
        numbers_text = text_combined.replace(stock_code, '')
        numbers = re.findall(r'\d+\.?\d*', numbers_text)
        numbers = [float(n) for n in numbers if float(n) > 0]

        if len(numbers) < 3:
            continue

        # 识别成交量（整数且>=100）
        volume = None
        for n in sorted(numbers):
            if n >= 100 and n == int(n):
                volume = int(n)
                break
        if volume is None:
            volume = int(min(numbers))

        # 识别成交金额（最大值）
        amount = max(numbers)

        # 识别成交价
        price = round(amount / volume, 3) if volume > 0 else 0
        for n in numbers:
            if n != amount and n != volume and abs(n - price) < 1:
                price = n
                break

        raw_records.append({
            '证券代码': stock_code,
            '证券名称': stock_name,
            '买卖类别': direction,
            '成交数量': volume,
            '成交价格': price,
            '成交金额': amount,
            'y': sum(r['y'] for r in row) / len(row)  # 记录行y坐标用于调试
        })

    # 第三步：按股票代码聚合，合并同一股票的名称
    # 用出现次数最多的名称作为该股票的正式名称
    from collections import Counter

    code_names = {}
    for rec in raw_records:
        code = rec['证券代码']
        name = rec['证券名称']
        if name and name != '未知':
            if code not in code_names:
                code_names[code] = []
            code_names[code].append(name)

    # 为每个代码选择最佳名称
    best_names = {}
    for code, names in code_names.items():
        counter = Counter(names)
        best_names[code] = counter.most_common(1)[0][0]

    # 填充缺失的名称
    for rec in raw_records:
        code = rec['证券代码']
        if not rec['证券名称'] or rec['证券名称'] == '未知':
            if code in best_names:
                rec['证券名称'] = best_names[code]
            else:
                rec['证券名称'] = '未知'

    # 去重：同一方向+同一代码+同一数量+同一金额视为重复
    seen = set()
    unique_records = []
    for rec in raw_records:
        key = (rec['证券代码'], rec['买卖类别'], rec['成交数量'], rec['成交金额'])
        if key not in seen:
            seen.add(key)
            unique_records.append({
                '证券代码': rec['证券代码'],
                '证券名称': rec['证券名称'],
                '买卖类别': rec['买卖类别'],
                '成交类型': '成交',
                '成交数量': rec['成交数量'],
                '成交价格': rec['成交价格'],
                '成交金额': rec['成交金额']
            })

    df = pd.DataFrame(unique_records)
    print(f"  识别到 {len(df)} 条交易记录")
    return df


def process_image_file(image_path):
    """处理单个图片文件（根据来源路由到不同解析器）"""
    trading_date = extract_date_from_filename(image_path)
    source = get_source_from_filename(os.path.basename(image_path))

    print(f"\n{'='*80}")
    print(f"处理图片：{os.path.basename(image_path)}")
    print(f"交易日期：{trading_date}")
    print(f"数据来源：{source}")
    print('='*80)

    if source == '平安账户':
        df = parse_pingan_image_trades(image_path)
    else:
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


# ==================== 平安证券截图OCR解析 ====================

def _infer_code_by_name(partial_code, stock_name):
    """通过股票名称和部分代码推断完整代码
    当OCR只识别出5位或不完整代码时，尝试修正
    """
    if not stock_name or stock_name == '未知':
        return None

    # 如果同一批次中已经有其他行识别出了完整代码+相同名称，优先匹配
    # （在parse_pingan_image_trades中通过known_codes参数传递）

    # 常见科创板代码修正：68开头缺位的
    digits = re.sub(r'[^0-9]', '', partial_code)
    if len(digits) == 5 and (digits.startswith('68') or digits.startswith('63')):
        # 尝试在688xxx范围内搜索已知股票
        pass  # 由外层调用者通过已知代码匹配

    return None


# 已知股票代码映射（用于OCR识别不完整时通过名称匹配）
KNOWN_STOCK_NAMES = {
    '厦钨新能': '688778', '厦钨新能源': '688778', '度钨新能源': '688778',
    '辰钨新能源': '688778', '钨新能源': '688778',
}


def _lookup_code_by_name(name):
    """通过模糊名称匹配查找股票代码"""
    if not name or name == '未知':
        return None
    # 精确匹配
    if name in KNOWN_STOCK_NAMES:
        return KNOWN_STOCK_NAMES[name]
    # 模糊匹配：名称包含关键词
    for key, code in KNOWN_STOCK_NAMES.items():
        if len(key) >= 3 and key[:3] in name:
            return code
        if len(name) >= 3 and name[:3] in key:
            return code
    return None


def parse_pingan_image_trades(image_path):
    """解析平安证券成交明细截图
    平安截图特点：严格的列式表格，有固定表头
    列：成交时间 | 证券代码 | 证券名称 | 操作 | 成交量 | 成交均价 | 成交金额
    """
    # 修复Windows下torch/getpass环境变量问题
    import os as _os
    if not _os.environ.get('USERNAME'):
        _os.environ['USERNAME'] = 'Ryan'
    if not _os.environ.get('USER'):
        _os.environ['USER'] = 'Ryan'

    try:
        import easyocr
    except ImportError:
        print(f"[错误] 缺少OCR依赖，请运行：pip install easyocr")
        return pd.DataFrame()

    import numpy as np
    from PIL import Image

    print(f"  正在识别平安截图：{os.path.basename(image_path)}")
    reader = easyocr.Reader(['ch_sim', 'en'], gpu=False, verbose=False)

    # 使用PIL读取并通过numpy传给easyocr（避免OpenCV路径问题）
    img = Image.open(image_path)
    # 如果图片太小，放大3倍提高OCR识别率
    w, h = img.size
    scale = 1
    if h < 300:
        scale = 3
        img = img.resize((w * scale, h * scale), Image.LANCZOS)
        print(f"  图片已放大 ({w}x{h} → {w*scale}x{h*scale})")
    img_array = np.array(img)
    result = reader.readtext(img_array)

    # 提取所有文本项，保留位置信息（还原坐标到原始比例）
    items = []
    for bbox, text, conf in result:
        y_center = (bbox[0][1] + bbox[2][1]) / 2 / scale
        x_center = (bbox[0][0] + bbox[2][0]) / 2 / scale
        items.append({'text': text.strip(), 'y': y_center, 'x': x_center, 'conf': conf})

    items.sort(key=lambda r: r['y'])

    # 第一步：识别表头行，确定各列的x范围
    # 表头关键词映射（OCR可能误读，允许多个匹配模式）
    header_patterns = {
        '成交时间': ['成交时间'],
        '证券代码': ['证券代码', '代码'],
        '证券名称': ['证券名称', '名称'],
        '操作': ['操作'],
        '成交量': ['成交量', '成交粼', '成交数量', '数量'],
        '成交均价': ['成交均价', '均价'],
        '成交金额': ['成交金额', '金额', '成交额'],
    }

    header_keywords = {k: None for k in header_patterns}
    header_y = None
    used_positions = set()

    for item in items:
        text = item['text']
        # 先精确匹配
        for col_name, patterns in header_patterns.items():
            if header_keywords[col_name] is not None:
                continue  # 已匹配
            for pat in patterns:
                if pat in text:
                    header_keywords[col_name] = item['x']
                    header_y = item['y']
                    used_positions.add(item['x'])
                    break

    # 如果某些列没识别到表头，用已识别的列推算
    found_headers = {k: v for k, v in header_keywords.items() if v is not None}
    if len(found_headers) < 3:
        print("  [警告] 未能识别足够表头列，尝试回退到手机截图解析模式")
        return parse_image_trades(image_path)  # 回退到通用解析

    # 确定列边界（每列取两个相邻表头x的中点作为分界）
    sorted_headers = sorted(found_headers.items(), key=lambda kv: kv[1])
    column_ranges = {}
    for i, (col_name, col_x) in enumerate(sorted_headers):
        if i == 0:
            left = 0
        else:
            left = (sorted_headers[i-1][1] + col_x) / 2
        if i == len(sorted_headers) - 1:
            right = 99999
        else:
            right = (col_x + sorted_headers[i+1][1]) / 2
        column_ranges[col_name] = (left, right)

    print(f"  识别到表头列：{list(column_ranges.keys())}")

    # 第二步：按y坐标分组成数据行（跳过表头行）
    data_items = [item for item in items if item['y'] > header_y + 5]
    if not data_items:
        print("  未找到数据行")
        return pd.DataFrame()

    rows = []
    current_row = []
    last_y = None
    y_threshold = 15  # 平安截图行间距较小

    for item in data_items:
        if last_y is None or abs(item['y'] - last_y) <= y_threshold:
            current_row.append(item)
        else:
            if current_row:
                rows.append(current_row)
            current_row = [item]
        last_y = item['y']

    if current_row:
        rows.append(current_row)

    # 第三步：从每行提取各列数据
    raw_records = []
    for row in rows:
        record = {}
        for col_name, (left, right) in column_ranges.items():
            col_texts = [r['text'] for r in row if left <= r['x'] < right]
            if col_texts:
                # 取该范围内置信度最高的文本
                best = max(
                    [r for r in row if left <= r['x'] < right],
                    key=lambda r: r['conf'],
                    default=None
                )
                record[col_name] = best['text'] if best else ' '.join(col_texts)
            else:
                record[col_name] = ''

        raw_records.append(record)

    # 第四步：数据清洗和转换
    trades = []
    for rec in raw_records:
        # 股票代码清洗：修正OCR常见误读
        code = rec.get('证券代码', '').strip()
        # 去除非数字字符
        code_digits = re.sub(r'[^0-9]', '', code)

        # 修正常见OCR错误模式
        if code_digits.startswith('633') and len(code_digits) == 6:
            code_digits = '688' + code_digits[3:]  # 633xxx → 688xxx
        elif code_digits.startswith('655') and len(code_digits) == 6:
            code_digits = '688' + code_digits[3:]  # 655xxx → 688xxx

        # 修正缺位：如 68775 → 688778（OCR漏了数字）
        if len(code_digits) == 5:
            code_match = re.search(r'(\d{6})', code)
            if not code_match:
                # 尝试常见的科创板补位模式
                if code_digits.startswith('688') or code_digits.startswith('68'):
                    # 可能漏了中间的数字，用已知代码表推断
                    pass  # 无法可靠推断，保留原始

        # 提取6位数字
        code_match = re.search(r'(\d{6})', code_digits if len(code_digits) >= 6 else code)
        if code_match:
            code = code_match.group(1)
        elif len(code_digits) == 6:
            code = code_digits
        else:
            # 代码不完整，尝试通过已识别到的股票名称推断
            code = _infer_code_by_name(code, rec.get('证券名称', ''))
            if not code:
                # 通过已知名称映射查找
                code = _lookup_code_by_name(rec.get('证券名称', ''))
            if not code:
                print(f"    [跳过] 无法识别股票代码: {rec.get('证券代码', '')}")
                continue

        # 股票名称清洗
        name = rec.get('证券名称', '').strip()
        if not name:
            name = '未知'

        # 操作方向清洗
        operation = rec.get('操作', '').strip()
        # 修正常见OCR误读
        if any(kw in operation for kw in ['卖', '觌出', '卖出']):
            direction = '证券卖出'
        elif any(kw in operation for kw in ['买', '买入']):
            direction = '证券买入'
        else:
            print(f"    [跳过] 无法判断买卖方向: {operation}")
            continue

        # 成交量清洗
        volume_text = rec.get('成交量', '').strip()
        volume = None
        if volume_text:
            vol_clean = re.sub(r'[^0-9]', '', volume_text)
            try:
                volume = int(vol_clean)
            except ValueError:
                pass

        # 成交金额清洗
        amount_text = rec.get('成交金额', '').strip()
        amount = None
        if amount_text:
            amt_clean = re.sub(r'[^0-9.]', '', amount_text)
            # 处理多个小数点的情况
            parts = amt_clean.split('.')
            if len(parts) > 2:
                amt_clean = parts[0] + '.' + ''.join(parts[1:])
            try:
                amount = float(amt_clean)
            except ValueError:
                pass

        # 成交均价清洗
        price_text = rec.get('成交均价', '').strip()
        price = None
        if price_text:
            price_clean = re.sub(r'[^0-9.]', '', price_text)
            parts = price_clean.split('.')
            if len(parts) > 2:
                price_clean = parts[0] + '.' + ''.join(parts[1:])
            try:
                price = float(price_clean)
            except ValueError:
                pass

        # 交叉验证与修正：如果 amount 存在，用它反推更可靠的 price/volume
        if amount and volume and price:
            calculated = round(volume * price, 2)
            if abs(calculated - amount) > max(amount * 0.02, 1):
                # 差异超过2%或1元，说明price可能误读，用amount反推
                price = round(amount / volume, 4)
        elif amount and volume and not price:
            price = round(amount / volume, 4)
        elif amount and price and not volume:
            volume = int(round(amount / price))
        elif volume and price and not amount:
            amount = round(volume * price, 2)

        # 跳过无效记录
        if not code or not name or volume is None:
            continue

        trades.append({
            '证券代码': code,
            '证券名称': name,
            '买卖类别': direction,
            '成交类型': '成交',
            '成交数量': volume,
            '成交价格': price or 0,
            '成交金额': amount or 0
        })

    # 去重：同代码+同方向+同数量+同金额视为重复
    seen = set()
    unique_records = []
    for t in trades:
        key = (t['证券代码'], t['买卖类别'], t['成交数量'], t['成交金额'])
        if key not in seen:
            seen.add(key)
            unique_records.append(t)

    df = pd.DataFrame(unique_records)
    print(f"  识别到 {len(df)} 条交易记录")
    if len(df) > 0:
        for _, row in df.iterrows():
            print(f"    {row['证券名称']}({row['证券代码']}) {row['买卖类别']} "
                  f"数量={row['成交数量']} 均价={row['成交价格']:.4f} 金额={row['成交金额']:.2f}")
    return df


# ==================== Excel处理 ====================

def parse_pingan_excel(file_path):
    """解析平安证券导出的成交记录
    平安导出文件特点：
    - 文件名含"平安"，扩展名可能是 .xls 但实际上可能是 TSV 格式（GBK编码）
    - 也可能是真正的 .xls 老格式
    - 第一行为列名，数据从第二行开始
    - 列：成交时间 | 证券代码 | 证券名称 | 操作 | 成交数量 | 成交均价 | 成交金额 | ...
    - 证券代码用格式：="688778"（需要剥离）
    - 操作用"买入"/"卖出"（需映射为标准格式）
    """
    # 先尝试作为TSV读取（GBK编码，平安常见导出格式）
    try:
        raw_df = pd.read_csv(file_path, sep='\t', encoding='gbk', dtype=str)
    except Exception:
        # 回退到xlrd引擎读真正的.xls
        engine = 'xlrd' if file_path.lower().endswith('.xls') else None
        raw_df = pd.read_excel(file_path, engine=engine)

    # 只取需要的列，映射到标准字段名
    col_map = {
        '证券代码': '证券代码', '证券名称': '证券名称',
        '操作': '买卖类别', '成交数量': '成交数量',
        '成交均价': '成交价格', '成交金额': '成交金额',
    }

    df = pd.DataFrame()
    name_col = None
    for src_col, dst_col in col_map.items():
        if src_col in raw_df.columns:
            val = raw_df[src_col]
            if dst_col == '证券代码':
                # 清洗 ="688778" 格式
                val = val.astype(str).str.replace('=', '').str.replace('"', '').str.strip()
            df[dst_col] = val
        elif dst_col == '证券名称' and '证券名称' in raw_df.columns:
            name_col = '证券名称'
            df['证券名称'] = raw_df['证券名称']

    # 添加统一的成交类型
    df['成交类型'] = '成交'

    # 映射操作字段：买入→证券买入，卖出→证券卖出
    df['买卖类别'] = df['买卖类别'].apply(
        lambda x: '证券买入' if '买' in str(x) else ('证券卖出' if '卖' in str(x) else str(x))
    )

    # 证券代码清洗（去除非数字字符）
    df['证券代码'] = df['证券代码'].astype(str).str.replace(r'[^0-9]', '', regex=True)
    df['成交数量'] = pd.to_numeric(df['成交数量'], errors='coerce')
    df['成交价格'] = pd.to_numeric(df['成交价格'], errors='coerce')
    df['成交金额'] = pd.to_numeric(df['成交金额'], errors='coerce')

    # 去空行，验证代码6位
    df = df.dropna(subset=['证券代码'])
    df = df[df['证券代码'].str.len() == 6]

    print(f"  识别到 {len(df)} 条交易记录")
    if len(df) > 0:
        name_col_ref = name_col if name_col and name_col in raw_df.columns else None
        for i, row in df.iterrows():
            name = raw_df.loc[i, name_col_ref] if name_col_ref and i in raw_df.index else row.get('证券名称', '')
            print(f"    {name}({row['证券代码']}) {row['买卖类别']} "
                  f"数量={int(row['成交数量'])} 均价={row['成交价格']:.4f} 金额={row['成交金额']:.2f}")

    return df


def process_excel_file(input_file):
    """处理单个Excel文件（根据来源路由到不同解析器）"""
    trading_date = extract_date_from_filename(input_file)
    source = get_source_from_filename(os.path.basename(input_file))

    print(f"\n{'='*80}")
    print(f"处理文件：{os.path.basename(input_file)}")
    print(f"交易日期：{trading_date}")
    print(f"数据来源：{source}")
    print('='*80)

    if source == '平安账户':
        df = parse_pingan_excel(input_file)
    else:
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
    """计算盈亏（通用逻辑，适用于Excel和图片输入）
    支持跨账户同股配对：当source包含多个账户时，按证券代码合并买卖后统一匹配
    支持单边交易标记：无法配对的买入/卖出记录标记为"未平仓"
    """
    profit_results = []
    all_stocks = set(df['证券代码'].unique())

    for stock_code in all_stocks:
        stock_name = df[df['证券代码'] == stock_code]['证券名称'].iloc[0]

        buys = buy_records[buy_records['证券代码'] == stock_code]
        sells = sell_records[sell_records['证券代码'] == stock_code]

        # 收集涉及的账户来源
        involved_sources = set()
        if '数据来源' in buys.columns:
            involved_sources.update(buys['数据来源'].unique())
        if '数据来源' in sells.columns:
            involved_sources.update(sells['数据来源'].unique())
        if not involved_sources:
            involved_sources = {source}

        # 确定显示的来源标签
        if len(involved_sources) > 1:
            display_source = '跨账户(' + '+'.join(sorted(involved_sources)) + ')'
        else:
            display_source = source

        total_buy_qty = buys['成交数量'].sum()
        total_sell_qty = sells['成交数量'].sum()

        if total_buy_qty == 0 and total_sell_qty == 0:
            continue

        # 无配对的情况：记录单边交易（未平仓）
        if total_buy_qty == 0 or total_sell_qty == 0:
            side = '仅买入' if total_buy_qty > 0 else '仅卖出'
            qty = total_buy_qty if total_buy_qty > 0 else total_sell_qty
            amt = buys['成交金额'].sum() if total_buy_qty > 0 else sells['成交金额'].sum()
            avg_price = amt / qty if qty > 0 else 0

            # 计算单边交易成本（只有佣金，没有印花税因为没有卖出；只有卖出时收印花税+佣金）
            if total_buy_qty > 0:
                commission = max(amt * COMMISSION_RATE, MIN_COMMISSION)
                stamp_duty = 0
            else:
                commission = max(amt * COMMISSION_RATE, MIN_COMMISSION)
                stamp_duty = round(amt * STAMP_DUTY_RATE, 2)
            total_cost = round(commission + stamp_duty, 2)

            profit_results.append({
                '日期': trading_date,
                '数据来源': display_source,
                '证券代码': stock_code,
                '证券名称': stock_name,
                '买入数量': int(total_buy_qty),
                '卖出数量': int(total_sell_qty),
                '匹配数量': 0,
                '买入均价': round(avg_price, 4) if total_buy_qty > 0 else 0,
                '卖出均价': round(avg_price, 4) if total_sell_qty > 0 else 0,
                '买入金额': round(buys['成交金额'].sum(), 2) if total_buy_qty > 0 else 0,
                '卖出金额': round(sells['成交金额'].sum(), 2) if total_sell_qty > 0 else 0,
                '毛盈亏': 0,
                '佣金': commission,
                '印花税': stamp_duty,
                '交易成本': total_cost,
                '盈亏金额': round(-total_cost, 2),
                '盈亏比例': f"⚠️{side}未平仓"
            })
            print(f"股票：{stock_name} ({stock_code}) → {side}未平仓，数量={int(qty)}")
            continue

        matched_qty = min(total_buy_qty, total_sell_qty)

        if matched_qty == 0:
            continue

        total_buy_amt = buys['成交金额'].sum()
        avg_buy_price = total_buy_amt / total_buy_qty
        matched_buy_amt = avg_buy_price * matched_qty

        total_sell_amt = sells['成交金额'].sum()
        avg_sell_price = total_sell_amt / total_sell_qty
        matched_sell_amt = avg_sell_price * matched_qty

        profit = matched_sell_amt - matched_buy_amt  # 毛盈亏

        # 交易成本计算
        buy_commission = max(matched_buy_amt * COMMISSION_RATE, MIN_COMMISSION)
        sell_commission = max(matched_sell_amt * COMMISSION_RATE, MIN_COMMISSION)
        commission = round(buy_commission + sell_commission, 2)
        stamp_duty = round(matched_sell_amt * STAMP_DUTY_RATE, 2)
        total_cost = round(commission + stamp_duty, 2)
        net_profit = round(profit - total_cost, 2)
        profit_pct = (net_profit / matched_buy_amt) * 100 if matched_buy_amt != 0 else 0

        # 如果有未匹配的部分，也记录
        unmatched_buy_qty = total_buy_qty - matched_qty
        unmatched_sell_qty = total_sell_qty - matched_qty

        profit_results.append({
            '日期': trading_date,
            '数据来源': display_source,
            '证券代码': stock_code,
            '证券名称': stock_name,
            '买入数量': int(total_buy_qty),
            '卖出数量': int(total_sell_qty),
            '匹配数量': int(matched_qty),
            '买入均价': round(avg_buy_price, 4),
            '卖出均价': round(avg_sell_price, 4),
            '买入金额': round(matched_buy_amt, 2),
            '卖出金额': round(matched_sell_amt, 2),
            '毛盈亏': round(profit, 2),
            '佣金': commission,
            '印花税': stamp_duty,
            '交易成本': total_cost,
            '盈亏金额': net_profit,
            '盈亏比例': f"{profit_pct:.2f}%"
        })

        print(f"股票：{stock_name} ({stock_code})" + (f" [跨账户]" if len(involved_sources) > 1 else ""))
        print(f"  买入：数量={total_buy_qty:.0f}, 均价={avg_buy_price:.4f}, 金额={matched_buy_amt:.2f}")
        print(f"  卖出：数量={total_sell_qty:.0f}, 均价={avg_sell_price:.4f}, 金额={matched_sell_amt:.2f}")
        print(f"  毛盈亏：{profit:.2f} 元")
        print(f"  交易成本：佣金={commission:.2f}, 印花税={stamp_duty:.2f}, 合计={total_cost:.2f}")
        print(f"  净盈亏：{net_profit:.2f} 元 ({profit_pct:.2f}%)")

        # 记录未匹配部分
        if unmatched_buy_qty > 0:
            unmatched_buy_amt = avg_buy_price * unmatched_buy_qty
            uc = max(unmatched_buy_amt * COMMISSION_RATE, MIN_COMMISSION)
            profit_results.append({
                '日期': trading_date,
                '数据来源': display_source,
                '证券代码': stock_code,
                '证券名称': stock_name,
                '买入数量': int(unmatched_buy_qty),
                '卖出数量': 0,
                '匹配数量': 0,
                '买入均价': round(avg_buy_price, 4),
                '卖出均价': 0,
                '买入金额': round(unmatched_buy_amt, 2),
                '卖出金额': 0,
                '毛盈亏': 0,
                '佣金': round(uc, 2),
                '印花税': 0,
                '交易成本': round(uc, 2),
                '盈亏金额': round(-uc, 2),
                '盈亏比例': "⚠️多买未平仓"
            })
            print(f"  ⚠️ 多买入 {int(unmatched_buy_qty)} 股未平仓")

        if unmatched_sell_qty > 0:
            unmatched_sell_amt = avg_sell_price * unmatched_sell_qty
            uc = max(unmatched_sell_amt * COMMISSION_RATE, MIN_COMMISSION)
            usd = round(unmatched_sell_amt * STAMP_DUTY_RATE, 2)
            profit_results.append({
                '日期': trading_date,
                '数据来源': display_source,
                '证券代码': stock_code,
                '证券名称': stock_name,
                '买入数量': 0,
                '卖出数量': int(unmatched_sell_qty),
                '匹配数量': 0,
                '买入均价': 0,
                '卖出均价': round(avg_sell_price, 4),
                '买入金额': 0,
                '卖出金额': round(unmatched_sell_amt, 2),
                '毛盈亏': 0,
                '佣金': round(uc, 2),
                '印花税': usd,
                '交易成本': round(uc + usd, 2),
                '盈亏金额': round(-(uc + usd), 2),
                '盈亏比例': "⚠️多卖未平仓"
            })
            print(f"  ⚠️ 多卖出 {int(unmatched_sell_qty)} 股未平仓")

    return profit_results


# ==================== Excel汇总 ====================

def append_to_excel(result_df, trading_date, source):
    """追加数据到Excel汇总文件（支持按日期去重，同一天的数据整体替换）"""
    if len(result_df) == 0:
        return

    if os.path.exists(EXCEL_OUTPUT):
        existing_df = pd.read_excel(EXCEL_OUTPUT)
        # 删除该日期的旧数据（整日替换，因为跨账户匹配可能改变所有记录）
        existing_df = existing_df[existing_df['日期'] != trading_date]
        combined_df = pd.concat([existing_df, result_df], ignore_index=True)
        combined_df = combined_df.sort_values(['日期', '数据来源']).reset_index(drop=True)
    else:
        combined_df = result_df

    wb = Workbook()
    ws = wb.active
    ws.title = "股票盈亏汇总"

    headers = ['日期', '数据来源', '证券代码', '证券名称', '买入数量', '卖出数量', '匹配数量',
               '买入均价', '卖出均价', '买入金额', '卖出金额',
               '毛盈亏', '佣金', '印花税', '交易成本', '盈亏金额', '盈亏比例']

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

        # 毛盈亏
        gross_cell = ws.cell(row=idx, column=12, value=row.毛盈亏)
        if row.毛盈亏 > 0:
            gross_cell.font = Font(color="FF0000", bold=True)
        else:
            gross_cell.font = Font(color="00B050", bold=True)

        ws.cell(row=idx, column=13, value=row.佣金)
        ws.cell(row=idx, column=14, value=row.印花税)
        ws.cell(row=idx, column=15, value=row.交易成本)

        profit_cell = ws.cell(row=idx, column=16, value=row.盈亏金额)
        if row.盈亏金额 > 0:
            profit_cell.font = Font(color="FF0000", bold=True)
        else:
            profit_cell.font = Font(color="00B050", bold=True)

        ws.cell(row=idx, column=17, value=row.盈亏比例)

        for col in range(1, 18):
            ws.cell(row=idx, column=col).alignment = Alignment(horizontal='center', vertical='center')

    column_widths = [12, 12, 12, 14, 10, 10, 10, 12, 12, 14, 14, 14, 12, 12, 12, 14, 12]
    for i, width in enumerate(column_widths, 1):
        ws.column_dimensions[chr(64+i)].width = width

    wb.save(EXCEL_OUTPUT)
    action = "更新" if os.path.exists(EXCEL_OUTPUT) else "创建"
    print(f"Excel汇总已{action}（按日期去重，整日替换）：{EXCEL_OUTPUT}")


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
    total_commission = day_df['佣金'].sum() if '佣金' in day_df.columns else 0
    total_stamp_duty = day_df['印花税'].sum() if '印花税' in day_df.columns else 0
    total_cost = total_commission + total_stamp_duty

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
        .source-平安 {{ background: #e8f5e9; color: #388e3c; }}
        /* 列宽设置 */
        th:nth-child(1), td:nth-child(1) {{ width: 10%; }} /* 数据来源 */
        th:nth-child(2), td:nth-child(2) {{ width: 14%; }} /* 证券名称 */
        th:nth-child(3), td:nth-child(3) {{ width: 8%; }} /* 买入数量 */
        th:nth-child(4), td:nth-child(4) {{ width: 8%; }} /* 卖出数量 */
        th:nth-child(5), td:nth-child(5) {{ width: 8%; }} /* 匹配数量 */
        th:nth-child(6), td:nth-child(6) {{ width: 10%; }} /* 买入均价 */
        th:nth-child(7), td:nth-child(7) {{ width: 10%; }} /* 卖出均价 */
        th:nth-child(8), td:nth-child(8) {{ width: 11%; }} /* 买入金额 */
        th:nth-child(9), td:nth-child(9) {{ width: 11%; }} /* 卖出金额 */
        th:nth-child(10), td:nth-child(10) {{ width: 8%; }} /* 佣金 */
        th:nth-child(11), td:nth-child(11) {{ width: 8%; }} /* 印花税 */
        th:nth-child(12), td:nth-child(12) {{ width: 12%; }} /* 盈亏金额 */
        th:nth-child(13), td:nth-child(13) {{ width: 10%; }} /* 盈亏比例 */
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
                <div class="title">佣金合计</div>
                <div class="value neutral">¥{total_commission:,.2f}</div>
            </div>
            <div class="card">
                <div class="title">印花税合计</div>
                <div class="value neutral">¥{total_stamp_duty:,.2f}</div>
            </div>
            <div class="card">
                <div class="title">交易成本合计</div>
                <div class="value neutral">¥{total_cost:,.2f}</div>
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
                        <th>佣金</th>
                        <th>印花税</th>
                        <th>盈亏金额</th>
                        <th>盈亏比例</th>
                    </tr>
                </thead>
                <tbody>
"""
        for _, row in day_df.iterrows():
            profit_class = 'profit-amount' if row['盈亏金额'] > 0 else 'loss-amount'
            profit_sign = '+' if row['盈亏金额'] > 0 else ''
            commission_val = row['佣金'] if '佣金' in row.index else 0
            stamp_val = row['印花税'] if '印花税' in row.index else 0
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
                        <td>¥{commission_val:,.2f}</td>
                        <td>¥{stamp_val:,.2f}</td>
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
            'commission': round(float(row['佣金']), 2) if '佣金' in row.index else 0,
            'stampDuty': round(float(row['印花税']), 2) if '印花税' in row.index else 0,
            'totalCost': round(float(row['交易成本']), 2) if '交易成本' in row.index else 0,
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
    """生成HTML模板（v4.0 — 新增个股总览+盈亏日历标签页）"""
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
/* 标签页导航 */
.tab-nav {{
    display: flex; background: #f0f2f5; border-bottom: 2px solid #e0e0e0; padding: 0 30px;
}}
.tab-btn {{
    padding: 14px 28px; font-size: 15px; font-weight: 600; color: #666;
    background: transparent; border: none; border-bottom: 3px solid transparent;
    cursor: pointer; transition: all 0.2s; position: relative; top: 2px;
}}
.tab-btn:hover {{ color: #667eea; }}
.tab-btn.active {{ color: #667eea; border-bottom-color: #667eea; background: #fff; }}
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
/* 个股总览表格 */
.stock-overview-table {{
    width: 100%; border-collapse: collapse; font-size: 14px;
    border-radius: 8px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.06);
}}
.stock-overview-table thead {{ background: linear-gradient(135deg,#667eea 0%,#764ba2 100%); }}
.stock-overview-table th {{ color: #fff; padding: 14px 10px; text-align: center; font-size: 13px; white-space: nowrap; }}
.stock-overview-table td {{ padding: 12px 10px; text-align: center; border-bottom: 1px solid #f0f0f0; }}
.stock-overview-table tbody tr {{ cursor: pointer; transition: background 0.15s; }}
.stock-overview-table tbody tr:hover {{ background: #f0f4ff; }}
.stock-overview-table .stock-name {{ font-weight: 700; font-size: 15px; }}
.stock-overview-table .rank {{ font-weight: 700; font-size: 16px; color: #667eea; }}
/* 胜率进度条 */
.winrate-bar {{
    display: inline-block; width: 60px; height: 8px; background: #e0e0e0;
    border-radius: 4px; overflow: hidden; vertical-align: middle; margin-right: 6px;
}}
.winrate-fill {{ height: 100%; border-radius: 4px; background: #667eea; }}
/* 盈亏热力条 */
.profit-bar {{
    display: inline-block; width: 80px; height: 10px; background: #f0f0f0;
    border-radius: 5px; overflow: hidden; vertical-align: middle; margin-right: 6px;
}}
.profit-bar-fill {{ height: 100%; border-radius: 5px; }}
.empty-state {{ text-align: center; padding: 60px 20px; color: #aaa; font-size: 16px; }}
.footer {{
    padding: 16px; text-align: center; color: #999; font-size: 13px;
    background: #f8f9fa; border-top: 1px solid #eee;
}}
/* 页面容器 */
.page-container {{ display: none; }}
.page-container.active {{ display: block; }}
@media (max-width: 768px) {{
    .two-charts {{ grid-template-columns: 1fr; }}
    .header {{ flex-direction: column; gap: 10px; text-align: center; }}
    .stats-grid {{ grid-template-columns: repeat(2, 1fr); }}
    .filter-bar {{ justify-content: center; }}
    .tab-nav {{ overflow-x: auto; }}
    .tab-btn {{ padding: 12px 16px; font-size: 14px; }}
}}
    </style>
</head>
<body>
<div class="container">
    <div class="header">
        <div>
            <h1>📊 股票交易汇总分析</h1>
            <div class="subtitle">点击图表数据项可钻取下级明细 · 个股总览 · 盈亏日历</div>
        </div>
    </div>
    <!-- 标签页导航 -->
    <div class="tab-nav">
        <button class="tab-btn active" data-tab="overview" onclick="switchTab('overview')">📈 汇总概览</button>
        <button class="tab-btn" data-tab="stocks" onclick="switchTab('stocks')">🏆 个股总览</button>
        <button class="tab-btn" data-tab="calendar" onclick="switchTab('calendar')">📅 盈亏日历</button>
    </div>
    <!-- 筛选条 -->
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

    <!-- 三个页面容器 -->
    <div class="page-container active" id="pageOverview">
        <div class="stats-grid" id="statsGrid"></div>
        <div class="charts-area" id="chartsArea"></div>
        <div class="detail-section" id="detailSection"></div>
    </div>
    <div class="page-container" id="pageStocks">
        <div id="stocksContent" style="padding:24px 30px;"></div>
    </div>
    <div class="page-container" id="pageCalendar">
        <div id="calendarContent" style="padding:24px 30px;"></div>
    </div>

    <div class="footer">
        <p>报告生成时间：{gen_time} | 数据来源：股票交易盈亏汇总.xlsx | v4.0</p>
    </div>
</div>
<script>
const ALL_DATA = {data_json};
const NOW = '{now_str}';
let state = {{ view: 'overview', filterType: 'mtd', customStart: '', customEnd: '', drillParam: null, activeTab: 'overview' }};
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
/* ====== 标签页切换 ====== */
function switchTab(tab) {{
    state.activeTab = tab;
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.toggle('active', b.dataset.tab === tab));
    document.querySelectorAll('.page-container').forEach(p => p.classList.remove('active'));
    if (tab === 'overview') {{
        document.getElementById('pageOverview').classList.add('active');
        state.view = 'overview'; state.drillParam = null; navStack = [];
        render();
    }} else if (tab === 'stocks') {{
        document.getElementById('pageStocks').classList.add('active');
        renderStocksOverview();
    }} else if (tab === 'calendar') {{
        document.getElementById('pageCalendar').classList.add('active');
        renderCalendarView();
    }}
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
    if (state.activeTab === 'stocks') renderStocksOverview();
    if (state.activeTab === 'calendar') renderCalendarView();
}}
function applyCustomFilter() {{
    state.customStart = document.getElementById('startDate').value;
    state.customEnd = document.getElementById('endDate').value;
    if (state.customStart && state.customEnd) {{
        state.view = 'overview'; state.drillParam = null; navStack = []; render();
        if (state.activeTab === 'stocks') renderStocksOverview();
        if (state.activeTab === 'calendar') renderCalendarView();
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
    const tComm = sumField(recs, 'commission');
    const tStamp = sumField(recs, 'stampDuty');
    const tCost = sumField(recs, 'totalCost');

    // 计算连续亏损和最大回撤
    const dayGroups = groupBy(recs, 'date');
    const days = Object.keys(dayGroups).sort();
    const dayProfits = days.map(d => sumField(dayGroups[d], 'profit'));

    // 最长连续亏损天数
    let maxConsecLoss = 0, curConsecLoss = 0;
    dayProfits.forEach(p => {{
        if (p < 0) {{ curConsecLoss++; maxConsecLoss = Math.max(maxConsecLoss, curConsecLoss); }}
        else {{ curConsecLoss = 0; }}
    }});

    // 最大回撤（从累计收益峰值到谷值的最大跌幅）
    let cumProfit = 0, maxCum = 0, maxDrawdown = 0;
    dayProfits.forEach(p => {{
        cumProfit += p;
        maxCum = Math.max(maxCum, cumProfit);
        maxDrawdown = Math.min(maxDrawdown, cumProfit - maxCum);
    }});

    // 最大单日亏损
    const maxDayLoss = dayProfits.length > 0 ? Math.min(...dayProfits) : 0;

    return [
        {{label:'交易天数', value:dates.length, cls:'neutral'}},
        {{label:'交易笔数', value:recs.length, cls:'neutral'}},
        {{label:'涉及股票', value:stocks.length, cls:'neutral'}},
        {{label:'盈利笔数', value:winCount, cls:'profit'}},
        {{label:'亏损笔数', value:loseCount, cls:'loss'}},
        {{label:'胜率', value: recs.length > 0 ? (winCount/recs.length*100).toFixed(1)+'%' : 'N/A', cls:'neutral'}},
        {{label:'总买入金额', value: fmtMoney(tb), cls:'neutral'}},
        {{label:'总卖出金额', value: fmtMoney(ts), cls:'neutral'}},
        {{label:'佣金合计', value: fmtMoney(tComm), cls:'neutral'}},
        {{label:'印花税合计', value: fmtMoney(tStamp), cls:'neutral'}},
        {{label:'交易成本合计', value: fmtMoney(tCost), cls:'neutral'}},
        {{label:'总盈亏', value: fmtMoney(tp,true), cls: profitCls(tp)}},
        {{label:'总收益率', value: fmtPct(tr,true), cls: profitCls(tr)}},
        {{label:'最长连续亏损', value: maxConsecLoss + '天', cls: maxConsecLoss >= 3 ? 'loss' : 'neutral'}},
        {{label:'最大回撤', value: fmtMoney(maxDrawdown,true), cls: 'loss'}},
        {{label:'最大单日亏损', value: fmtMoney(maxDayLoss,true), cls: 'loss'}},
    ];
}}
function renderDetailTable(recs, title) {{
    const el = document.getElementById('detailSection');
    if (recs.length === 0) {{ el.innerHTML = ''; return; }}
    el.innerHTML = '<h2 class="chart-title" style="margin-bottom:14px">' + title + '</h2>' +
        '<table class="detail-table"><thead><tr><th>日期</th><th>来源</th><th>证券名称</th><th>匹配数量</th><th>买入均价</th><th>卖出均价</th><th>买入金额</th><th>卖出金额</th><th>佣金</th><th>印花税</th><th>盈亏金额</th><th>盈亏比例</th></tr></thead><tbody>' +
        recs.map(r => {{
            const pc = r.profit >= 0 ? 'profit-cell' : 'loss-cell';
            return '<tr><td>' + r.date + '</td><td>' + r.source + '</td><td><b>' + r.name + '</b></td><td>' + r.matchQty + '</td><td>¥' + r.buyPrice.toFixed(4) + '</td><td>¥' + r.sellPrice.toFixed(4) + '</td><td>' + fmtMoney(r.buyAmount) + '</td><td>' + fmtMoney(r.sellAmount) + '</td><td>' + fmtMoney(r.commission) + '</td><td>' + fmtMoney(r.stampDuty) + '</td><td class="' + pc + '">' + fmtMoney(r.profit,true) + '</td><td class="' + pc + '">' + r.profitPct + '</td></tr>';
        }}).join('') + '</tbody></table>';
}}

/* ====== 个股总览页面 ====== */
function renderStocksOverview() {{
    const recs = getFilteredRecords();
    const el = document.getElementById('stocksContent');
    if (recs.length === 0) {{
        el.innerHTML = '<div class="empty-state">当前筛选条件下暂无数据</div>';
        return;
    }}
    const stockGroups = groupBy(recs, 'name');
    const stockList = Object.keys(stockGroups).map(n => {{
        const trades = stockGroups[n];
        const totalProfit = sumField(trades, 'profit');
        const totalBuy = sumField(trades, 'buyAmount');
        const totalSell = sumField(trades, 'sellAmount');
        const totalCost = sumField(trades, 'totalCost');
        const totalComm = sumField(trades, 'commission');
        const totalStamp = sumField(trades, 'stampDuty');
        const winCount = trades.filter(r => r.profit > 0).length;
        const loseCount = trades.filter(r => r.profit < 0).length;
        const winRate = trades.length > 0 ? winCount / trades.length * 100 : 0;
        const avgProfit = trades.length > 0 ? totalProfit / trades.length : 0;
        const returnRate = totalBuy > 0 ? totalProfit / totalBuy * 100 : 0;
        const tradeDays = unique(trades, 'date').length;
        const maxWin = Math.max(...trades.map(r => r.profit));
        const maxLoss = Math.min(...trades.map(r => r.profit));
        return {{
            name: n, code: trades[0].code, count: trades.length, tradeDays,
            totalProfit, totalBuy, totalSell, totalCost, totalComm, totalStamp,
            winCount, loseCount, winRate, avgProfit, returnRate, maxWin, maxLoss
        }};
    }}).sort((a,b) => b.totalProfit - a.totalProfit);

    const maxAbsProfit = Math.max(...stockList.map(s => Math.abs(s.totalProfit)), 1);

    // 顶部汇总卡片
    const totalProfit = stockList.reduce((s,x) => s + x.totalProfit, 0);
    const totalCost = stockList.reduce((s,x) => s + x.totalCost, 0);
    const avgWinRate = stockList.length > 0 ? stockList.reduce((s,x) => s + x.winRate, 0) / stockList.length : 0;

    let html = `
    <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:16px;margin-bottom:28px;">
        <div class="stat-card"><div class="label">交易股票数</div><div class="value neutral">${{stockList.length}}</div></div>
        <div class="stat-card"><div class="label">总盈亏</div><div class="value ${{profitCls(totalProfit)}}">${{fmtMoney(totalProfit,true)}}</div></div>
        <div class="stat-card"><div class="label">总交易成本</div><div class="value neutral">${{fmtMoney(totalCost)}}</div></div>
        <div class="stat-card"><div class="label">平均胜率</div><div class="value neutral">${{avgWinRate.toFixed(1)}}%</div></div>
    </div>
    <h2 class="chart-title">🏆 个股累计盈亏排行榜</h2>
    <div style="overflow-x:auto;">
    <table class="stock-overview-table">
        <thead><tr>
            <th>排名</th><th>证券名称</th><th>代码</th><th>交易次数</th><th>交易天数</th>
            <th>盈利次数</th><th>亏损次数</th><th>胜率</th>
            <th>总买入金额</th><th>总卖出金额</th>
            <th>佣金</th><th>印花税</th><th>交易成本</th>
            <th>平均每次盈亏</th><th>单笔最大盈利</th><th>单笔最大亏损</th>
            <th>总盈亏</th><th>累计收益率</th>
        </tr></thead><tbody>`;

    stockList.forEach((s, i) => {{
        const profitPct = s.totalProfit / maxAbsProfit * 100;
        const barColor = s.totalProfit >= 0 ? '#e74c3c' : '#27ae60';
        const barWidth = Math.abs(profitPct);
        html += `<tr onclick="switchTab('overview'); setTimeout(()=>drillTo('stock','${{s.name}}'),100);">
            <td class="rank">${{i+1}}</td>
            <td class="stock-name" style="color:${{barColor}}">${{s.name}}</td>
            <td>${{s.code}}</td>
            <td>${{s.count}}</td><td>${{s.tradeDays}}</td>
            <td style="color:#e74c3c">${{s.winCount}}</td><td style="color:#27ae60">${{s.loseCount}}</td>
            <td><span class="winrate-bar"><span class="winrate-fill" style="width:${{s.winRate}}%"></span></span>${{s.winRate.toFixed(1)}}%</td>
            <td>${{fmtMoney(s.totalBuy)}}</td><td>${{fmtMoney(s.totalSell)}}</td>
            <td>${{fmtMoney(s.totalComm)}}</td><td>${{fmtMoney(s.totalStamp)}}</td><td>${{fmtMoney(s.totalCost)}}</td>
            <td class="${{profitCls(s.avgProfit)}}">${{fmtMoney(s.avgProfit,true)}}</td>
            <td style="color:#e74c3c">${{fmtMoney(s.maxWin,true)}}</td>
            <td style="color:#27ae60">${{fmtMoney(s.maxLoss,true)}}</td>
            <td>
                <span class="profit-bar"><span class="profit-bar-fill" style="width:${{barWidth}}%;background:${{barColor}}"></span></span>
                <span class="${{profitCls(s.totalProfit)}}" style="font-weight:700;font-size:15px">${{fmtMoney(s.totalProfit,true)}}</span>
            </td>
            <td class="${{profitCls(s.returnRate)}}" style="font-weight:600">${{fmtPct(s.returnRate,true)}}</td>
        </tr>`;
    }});

    html += '</tbody></table></div>';

    // 个股盈亏柱状图
    html += `<div style="margin-top:32px;"><h2 class="chart-title">📊 个股累计盈亏对比图</h2><div id="chartStockOverview" class="chart-wrap"></div></div>`;
    // 个股胜率对比
    html += `<div style="margin-top:32px;"><h2 class="chart-title">🎯 个股胜率对比</h2><div id="chartStockWinrate" class="chart-wrap"></div></div>`;

    el.innerHTML = html;

    // 盈亏对比柱状图
    const sc = mkChart('chartStockOverview');
    sc.setOption({{
        tooltip: {{ trigger:'axis', formatter: p => p[0].name+'<br/>累计盈亏：'+fmtMoney(p[0].value,true) }},
        grid: {{ left:'3%', right:'4%', bottom:'15%', containLabel:true }},
        xAxis: {{ type:'category', data: stockList.map(s=>s.name), axisLabel:{{color:'#555', rotate:30}} }},
        yAxis: {{ type:'value', name:'盈亏（元）', axisLabel:{{formatter:'¥{{value}}'}} }},
        series: [{{ type:'bar', data: stockList.map(s=>Math.round(s.totalProfit*100)/100), itemStyle:{{ color: p => profitColor(p.value) }}, label:{{ show:true, position:'top', formatter: p => fmtMoney(p.value,true), color:'#333', fontSize:10 }} }}]
    }});
    sc.on('click', p => {{ switchTab('overview'); setTimeout(()=>drillTo('stock',p.name),100); }});

    // 胜率对比
    const wc = mkChart('chartStockWinrate');
    wc.setOption({{
        tooltip: {{ trigger:'axis', formatter: p => p[0].name+'<br/>胜率：'+p[0].value.toFixed(1)+'%' }},
        grid: {{ left:'3%', right:'4%', bottom:'15%', containLabel:true }},
        xAxis: {{ type:'category', data: stockList.map(s=>s.name), axisLabel:{{color:'#555', rotate:30}} }},
        yAxis: {{ type:'value', name:'胜率（%）', max:100, axisLabel:{{formatter:'{{value}}%'}} }},
        series: [{{ type:'bar', data: stockList.map(s=>Math.round(s.winRate*10)/10), itemStyle:{{ color: p => p.value >= 60 ? '#667eea' : '#f39c12' }}, label:{{ show:true, position:'top', formatter: p => p.value.toFixed(1)+'%', color:'#333', fontSize:11 }} }}]
    }});
}}

/* ====== 盈亏日历热力图 ====== */
function renderCalendarView() {{
    const recs = getFilteredRecords();
    const el = document.getElementById('calendarContent');
    if (recs.length === 0) {{
        el.innerHTML = '<div class="empty-state">当前筛选条件下暂无数据</div>';
        return;
    }}

    // 按日聚合盈亏
    const dayGroups = groupBy(recs, 'date');
    const dayData = Object.keys(dayGroups).sort().map(d => {{
        const dayProfit = sumField(dayGroups[d], 'profit');
        return [d, Math.round(dayProfit*100)/100];
    }});

    // 统计卡片
    const dayProfits = dayData.map(d => d[1]);
    const profitDays = dayProfits.filter(p => p > 0).length;
    const lossDays = dayProfits.filter(p => p < 0).length;
    const flatDays = dayProfits.filter(p => p === 0).length;
    const maxProfitDay = dayData.reduce((a,b) => b[1] > a[1] ? b : a, dayData[0]);
    const maxLossDay = dayData.reduce((a,b) => b[1] < a[1] ? b : a, dayData[0]);

    // 连续亏损计算
    let maxConsecLoss = 0, currentConsecLoss = 0;
    let consecLossEnd = '';
    let maxConsecLossStart = '', maxConsecLossEnd = '';
    let currentConsecStart = '';
    dayData.forEach(d => {{
        if (d[1] < 0) {{
            currentConsecLoss++;
            if (currentConsecLoss === 1) currentConsecStart = d[0];
            if (currentConsecLoss > maxConsecLoss) {{
                maxConsecLoss = currentConsecLoss;
                maxConsecLossStart = currentConsecStart;
                maxConsecLossEnd = d[0];
            }}
        }} else {{
            currentConsecLoss = 0;
        }}
    }});

    // 当前是否处于连续亏损中
    const lastDays = dayData.slice(-3);
    const isInConsecLoss = lastDays.length > 0 && lastDays.every(d => d[1] < 0);

    let statsHtml = `
    <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:16px;margin-bottom:28px;">
        <div class="stat-card"><div class="label">交易天数</div><div class="value neutral">${{dayData.length}}</div></div>
        <div class="stat-card"><div class="label">盈利天数</div><div class="value profit">${{profitDays}}</div></div>
        <div class="stat-card"><div class="label">亏损天数</div><div class="value loss">${{lossDays}}</div></div>
        <div class="stat-card"><div class="label">最长连续亏损</div><div class="value ${{maxConsecLoss >= 3 ? 'loss' : 'neutral'}}">${{maxConsecLoss}}天</div></div>
        <div class="stat-card"><div class="label">最大单日盈利</div><div class="value profit">${{maxProfitDay[0]}} ${{fmtMoney(maxProfitDay[1],true)}}</div></div>
        <div class="stat-card"><div class="label">最大单日亏损</div><div class="value loss">${{maxLossDay[0]}} ${{fmtMoney(maxLossDay[1],true)}}</div></div>
    </div>`;

    // 连续亏损警告
    if (isInConsecLoss) {{
        statsHtml += `<div style="background:#fff3cd;border:1px solid #ffc107;border-radius:10px;padding:16px 20px;margin-bottom:24px;color:#856404;font-size:15px;">
            ⚠️ <b>注意：</b>最近${{lastDays.length}}个交易日均为亏损，请注意控制风险，避免情绪化交易！
        </div>`;
    }}
    if (maxConsecLoss >= 3) {{
        statsHtml += `<div style="background:#f8d7da;border:1px solid #f5c6cb;border-radius:10px;padding:16px 20px;margin-bottom:24px;color:#721c24;font-size:14px;">
            🔴 历史最长连续亏损：${{maxConsecLoss}}天（${{maxConsecLossStart}} 至 ${{maxConsecLossEnd}}），复盘时重点关注该区间交易逻辑。
        </div>`;
    }}

    // 日历热力图
    let calHtml = `<h2 class="chart-title">📅 盈亏日历热力图</h2>
    <div style="margin-bottom:12px;font-size:13px;color:#888;">
        <span style="display:inline-block;width:14px;height:14px;background:#b71c1c;border-radius:2px;vertical-align:middle;margin-right:3px;"></span>大赚
        <span style="display:inline-block;width:14px;height:14px;background:#ef9a9a;border-radius:2px;vertical-align:middle;margin-left:12px;margin-right:3px;"></span>小赚
        <span style="display:inline-block;width:14px;height:14px;background:#e8e8e8;border-radius:2px;vertical-align:middle;margin-left:12px;margin-right:3px;"></span>无交易
        <span style="display:inline-block;width:14px;height:14px;background:#a5d6a7;border-radius:2px;vertical-align:middle;margin-left:12px;margin-right:3px;"></span>小亏
        <span style="display:inline-block;width:14px;height:14px;background:#2e7d32;border-radius:2px;vertical-align:middle;margin-left:12px;margin-right:3px;"></span>大亏
    </div>
    <div id="chartCalendar" style="width:100%;height:280px;"></div>`;

    // 日度盈亏柱状图（按时间排列，可点击查看当天明细）
    calHtml += `<div style="margin-top:32px;"><h2 class="chart-title">📈 日度盈亏详情 <span style="font-size:13px;color:#999;font-weight:normal">（点击柱体查看当日明细）</span></h2><div id="chartCalDaily" class="chart-wrap"></div></div>`;

    // 当日交易明细表
    calHtml += '<div id="calDetailSection"></div>';

    el.innerHTML = statsHtml + calHtml;

    // 渲染ECharts日历热力图
    const cc = mkChart('chartCalendar');
    // 确定日历范围
    const dateRange = dayData.length > 0 ? [dayData[0][0], dayData[dayData.length-1][0]] : [NOW, NOW];

    // 构建可视化数据（带颜色映射）
    const maxAbs = Math.max(...dayData.map(d => Math.abs(d[1])), 1);

    cc.setOption({{
        tooltip: {{
            formatter: function(p) {{
                const d = p.data || p.value;
                const date = Array.isArray(d) ? d[0] : d[0];
                const val = Array.isArray(d) ? d[1] : d[1];
                return date + '<br/>当日盈亏：' + fmtMoney(val, true);
            }}
        }},
        visualMap: {{
            min: -maxAbs,
            max: maxAbs,
            orient: 'horizontal',
            left: 'center',
            top: 0,
            itemWidth: 14,
            itemHeight: 120,
            inRange: {{
                color: ['#2e7d32', '#4caf50', '#a5d6a7', '#e8e8e8', '#ef9a9a', '#e57373', '#b71c1c']
            }},
            text: ['大赚', '大亏'],
            textStyle: {{ color: '#555', fontSize: 12 }},
            formatter: function(val) {{ return '¥' + val.toFixed(0); }}
        }},
        calendar: {{
            top: 60,
            left: 60,
            right: 60,
            bottom: 20,
            cellSize: ['auto', 28],
            range: dateRange[0].substring(0,4),
            itemStyle: {{ borderWidth: 2, borderColor: '#fff' }},
            yearLabel: {{ show: false }},
            dayLabel: {{ firstDay: 1, color: '#555', nameMap: 'ZH' }},
            monthLabel: {{ color: '#555', nameMap: 'ZH' }},
            splitLine: {{ show: true, lineStyle: {{ color: '#e0e0e0', width: 1 }} }}
        }},
        series: [{{
            type: 'heatmap',
            coordinateSystem: 'calendar',
            data: dayData,
            itemStyle: {{ borderWidth: 2, borderColor: '#fff', borderRadius: 3 }}
        }}]
    }});
    cc.on('click', p => {{
        const date = Array.isArray(p.data) ? p.data[0] : (p.value ? p.value[0] : null);
        if (date) showCalDayDetail(date);
    }});

    // 日度盈亏柱状图
    const dc = mkChart('chartCalDaily');
    dc.setOption({{
        tooltip: {{ trigger:'axis', formatter: p => p[0].name+'<br/>盈亏：'+fmtMoney(p[0].value,true) }},
        xAxis: {{ type:'category', data: dayData.map(d=>d[0]), axisLabel:{{color:'#555', rotate:45}} }},
        yAxis: {{ type:'value', name:'盈亏（元）', axisLabel:{{formatter:'¥{{value}}'}} }},
        series: [{{ type:'bar', data: dayData.map(d=>d[1]), itemStyle:{{ color: p => profitColor(p.value) }}, label:{{ show:true, position:'top', formatter: p => fmtMoney(p.value,true), color:'#333', fontSize:10 }} }}]
    }});
    dc.on('click', p => showCalDayDetail(p.name));
}}

function showCalDayDetail(date) {{
    const recs = getFilteredRecords().filter(r => r.date === date);
    if (recs.length === 0) return;
    const el = document.getElementById('calDetailSection');
    el.innerHTML = '<h2 class="chart-title" style="margin-top:24px;">📋 ' + date + ' 交易明细</h2>' +
        '<table class="detail-table"><thead><tr><th>来源</th><th>证券名称</th><th>匹配数量</th><th>买入均价</th><th>卖出均价</th><th>买入金额</th><th>卖出金额</th><th>佣金</th><th>印花税</th><th>盈亏金额</th><th>盈亏比例</th></tr></thead><tbody>' +
        recs.map(r => {{
            const pc = r.profit >= 0 ? 'profit-cell' : 'loss-cell';
            return '<tr><td>' + r.source + '</td><td><b>' + r.name + '</b></td><td>' + r.matchQty + '</td><td>¥' + r.buyPrice.toFixed(4) + '</td><td>¥' + r.sellPrice.toFixed(4) + '</td><td>' + fmtMoney(r.buyAmount) + '</td><td>' + fmtMoney(r.sellAmount) + '</td><td>' + fmtMoney(r.commission) + '</td><td>' + fmtMoney(r.stampDuty) + '</td><td class="' + pc + '">' + fmtMoney(r.profit,true) + '</td><td class="' + pc + '">' + r.profitPct + '</td></tr>';
        }}).join('') + '</tbody></table>';
    el.scrollIntoView({{ behavior: 'smooth', block: 'start' }});
}}

/* ====== 原有概览页面渲染 ====== */
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
        <div class="chart-section"><h2 class="chart-title">🎯 各股票盈亏排行 <span style="font-size:13px;color:#999;font-weight:normal">（点击柱体查看个股明细）</span></h2><div id="chartStockBar" class="chart-wrap"></div></div>
        <div class="chart-section"><h2 class="chart-title">🥧 盈亏构成分析 <span style="font-size:13px;color:#999;font-weight:normal">（盈利 vs 亏损分布）</span></h2><div id="chartStockPie" class="chart-wrap"></div></div>`;
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
    const winStocks = stockList.filter(s => s.profit > 0);
    const loseStocks = stockList.filter(s => s.profit < 0);
    let pieData = [];
    winStocks.forEach(s => {{
        pieData.push({{name: s.name + ' (盈)', value: s.profit, itemStyle:{{color: '#e74c3c'}}}});
    }});
    loseStocks.forEach(s => {{
        pieData.push({{name: s.name + ' (亏)', value: Math.abs(s.profit), itemStyle:{{color: '#27ae60'}}}});
    }});
    sp.setOption({{
        tooltip:{{ trigger:'item', formatter: function(p) {{
            const isWin = p.name.includes('(盈)');
            const rawName = p.name.replace(' (盈)', '').replace(' (亏)', '');
            const val = p.value;
            return rawName + '<br/>' + (isWin ? '盈利：' : '亏损：') + '¥' + val.toLocaleString('zh-CN', {{minimumFractionDigits:2}}) + '<br/>占比：' + p.percent + '%';
        }} }},
        series:[{{
            type:'pie',
            radius:['30%','60%'],
            center:['50%','55%'],
            itemStyle:{{ borderRadius:6, borderColor:'#fff', borderWidth:2 }},
            label:{{ show:true, formatter: function(p) {{
                const rawName = p.name.replace(' (盈)', '').replace(' (亏)', '');
                return rawName + '\\n¥' + p.value.toLocaleString('zh-CN', {{minimumFractionDigits:0}});
            }}, fontSize:11 }},
            emphasis:{{ label:{{ show:true, fontSize:13, fontWeight:'bold' }} }},
            data: pieData
        }}]
    }});
    sp.on('click', p => {{
        const rawName = p.name.replace(' (盈)', '').replace(' (亏)', '');
        drillTo('stock', rawName);
    }});
}}
function renderMonthView() {{
    const allFiltered = getFilteredRecords();
    const month = state.drillParam;
    const recs = allFiltered.filter(r => r.date.startsWith(month));
    if (recs.length === 0) {{ goBackOverview(); return; }}
    renderStats(computeStats(recs));
    const dayGroups = groupBy(recs, 'date');
    const days = Object.keys(dayGroups).sort();
    const dayProfits = days.map(d => Math.round(sumField(dayGroups[d],'profit')*100)/100);
    let cum = [], run = 0;
    dayProfits.forEach(p => {{ run += p; cum.push(Math.round(run*100)/100); }});
    const stockGroups = groupBy(recs, 'name');
    const stockList = Object.keys(stockGroups).map(n => ({{
        name: n, profit: Math.round(sumField(stockGroups[n],'profit')*100)/100,
        count: stockGroups[n].length
    }})).sort((a,b) => b.profit - a.profit);
    const chartsArea = document.getElementById('chartsArea');
    chartsArea.innerHTML = `
        <div class="chart-section"><h2 class="chart-title">📅 ${{month}} 每日盈亏 <span style="font-size:13px;color:#999;font-weight:normal">（点击柱体查看当日明细）</span></h2><div id="chartMonthDaily" class="chart-wrap"></div></div>
        <div class="chart-section"><h2 class="chart-title">📊 ${{month}} 各股票盈亏 <span style="font-size:13px;color:#999;font-weight:normal">（横轴：股票名称）</span></h2><div id="chartMonthStock" class="chart-wrap"></div></div>
        <div class="chart-section"><h2 class="chart-title">💰 ${{month}} 月内累计收益</h2><div id="chartMonthCum" class="chart-wrap"></div></div>`;
    const dc = mkChart('chartMonthDaily');
    dc.setOption({{
        tooltip:{{ trigger:'axis', formatter: p => p[0].name+'<br/>盈亏：'+fmtMoney(p[0].value,true) }},
        xAxis:{{ type:'category', data: days.map(d=>d.substring(8)), axisLabel:{{color:'#555'}} }},
        yAxis:{{ type:'value', name:'盈亏（元）', axisLabel:{{formatter:'¥{{value}}'}} }},
        series:[{{ type:'bar', data: dayProfits, itemStyle:{{ color: p => profitColor(p.value) }}, label:{{ show:true, position:'top', formatter: p => fmtMoney(p.value,true), color:'#333', fontSize:11 }} }}]
    }});
    dc.on('click', p => drillTo('day', month + '-' + p.name));
    const msc = mkChart('chartMonthStock');
    msc.setOption({{
        tooltip:{{ trigger:'axis', formatter: p => p[0].name+'<br/>盈亏：'+fmtMoney(p[0].value,true)+'<br/>交易次数：'+stockGroups[p[0].name].length }},
        grid:{{ left:'3%', right:'4%', bottom:'15%', containLabel:true }},
        xAxis:{{ type:'category', data: stockList.map(s=>s.name), axisLabel:{{color:'#555', rotate:30}} }},
        yAxis:{{ type:'value', name:'盈亏（元）', axisLabel:{{formatter:'¥{{value}}'}} }},
        series:[{{ type:'bar', data: stockList.map(s=>s.profit), itemStyle:{{ color: p => profitColor(p.value) }}, label:{{ show:true, position:'top', formatter: p => fmtMoney(p.value,true), color:'#333', fontSize:11 }} }}]
    }});
    msc.on('click', p => drillTo('stock', p.name));
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
    renderStats(computeStats(recs));
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
    const totalSell = sumField(recs, 'sellAmount');
    const totalRate = totalBuy > 0 ? totalProfit / totalBuy * 100 : 0;
    const tradeDays = unique(recs, 'date');
    const winCount = recs.filter(r => r.profit > 0).length;
    const tComm = sumField(recs, 'commission');
    const tStamp = sumField(recs, 'stampDuty');
    const tCost = sumField(recs, 'totalCost');
    renderStats([
        {{label:'股票名称', value: stockName, cls:'neutral'}},
        {{label:'证券代码', value: recs[0].code, cls:'neutral'}},
        {{label:'交易天数', value: tradeDays.length, cls:'neutral'}},
        {{label:'交易次数', value: recs.length, cls:'neutral'}},
        {{label:'盈利次数', value: winCount, cls:'profit'}},
        {{label:'亏损次数', value: recs.length - winCount, cls:'loss'}},
        {{label:'胜率', value: recs.length > 0 ? (winCount/recs.length*100).toFixed(1)+'%' : 'N/A', cls:'neutral'}},
        {{label:'总买入金额', value: fmtMoney(totalBuy), cls:'neutral'}},
        {{label:'总卖出金额', value: fmtMoney(totalSell), cls:'neutral'}},
        {{label:'佣金合计', value: fmtMoney(tComm), cls:'neutral'}},
        {{label:'印花税合计', value: fmtMoney(tStamp), cls:'neutral'}},
        {{label:'交易成本合计', value: fmtMoney(tCost), cls:'neutral'}},
        {{label:'平均每次盈亏', value: fmtMoney(totalProfit/recs.length,true), cls: profitCls(totalProfit)}},
        {{label:'总盈亏', value: fmtMoney(totalProfit,true), cls: profitCls(totalProfit)}},
        {{label:'总收益率', value: fmtPct(totalRate,true), cls: profitCls(totalRate)}},
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
    excel_files = [f for f in glob.glob('*.xlsx') + glob.glob('*.xls')
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
    print("股票交易盈亏分析系统 v4.5")
    print("支持输入：Excel文件（券商导出）+ 图片文件（手机App截图/平安证券截图）")
    print("="*80)

    excel_files, image_files = find_input_files()
    all_files = excel_files + image_files

    if not all_files:
        print("\n未找到待处理的文件")
        print("请将以下类型文件放到当前文件夹：")
        print("  - Excel文件：券商/平安导出的交易记录（如：2026-04-22-两融-当日成交汇总.xlsx 或 20260430_平安.xls）")
        print("  - 图片文件：手机App截图（如：2026-04-22-手机交易.png）")
        print("  - 图片文件：平安证券截图（如：20260430_平安.png）")
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

        # 收集该日期所有文件的买卖记录，用于跨账户合并匹配
        all_buy_records = []
        all_sell_records = []
        all_stock_info = {}  # code -> name
        file_sources = set()

        # 处理Excel文件：先解析，收集买卖记录
        for excel_file in date_files['excel']:
            try:
                source = get_source_from_filename(os.path.basename(excel_file))
                if source == '平安账户':
                    df = parse_pingan_excel(excel_file)
                else:
                    df = pd.read_excel(excel_file, skiprows=4, header=0)
                    df.columns = ['证券代码', '证券名称', '买卖类别', '成交类型', '成交数量', '成交价格', '成交金额']
                    df['证券代码'] = df['证券代码'].astype(str).str.replace('\t', '')
                    df = df.dropna(subset=['证券代码'])
                    df = df[df['证券代码'] != '证券代码']
                    df['成交数量'] = pd.to_numeric(df['成交数量'], errors='coerce')
                    df['成交价格'] = pd.to_numeric(df['成交价格'], errors='coerce')
                    df['成交金额'] = pd.to_numeric(df['成交金额'], errors='coerce')

                buy_records = df[df['买卖类别'].str.contains('证券买入', na=False)].copy()
                sell_records = df[df['买卖类别'].str.contains('证券卖出', na=False)].copy()

                # 标记来源
                buy_records['数据来源'] = source
                sell_records['数据来源'] = source
                file_sources.add(source)

                all_buy_records.append(buy_records)
                all_sell_records.append(sell_records)

                # 收集证券名称映射
                for _, row in df.iterrows():
                    code = str(row.get('证券代码', ''))
                    name = str(row.get('证券名称', ''))
                    if code and name and name != 'nan':
                        all_stock_info[code] = name

                print(f"  [{source}] Excel解析完成：买入{len(buy_records)}笔，卖出{len(sell_records)}笔")
                archive_file(excel_file)
            except Exception as e:
                print(f"\n[错误] 处理Excel文件 {excel_file} 时出错：{str(e)}")
                continue

        # 处理图片文件：先解析，收集买卖记录
        for image_file in date_files['image']:
            try:
                source = get_source_from_filename(os.path.basename(image_file))
                if source == '平安账户':
                    df = parse_pingan_image_trades(image_file)
                else:
                    df = parse_image_trades(image_file)

                if len(df) == 0:
                    print(f"  [{source}] 图片未识别到有效交易记录")
                    archive_file(image_file)
                    continue

                buy_records = df[df['买卖类别'].str.contains('证券买入', na=False)].copy()
                sell_records = df[df['买卖类别'].str.contains('证券卖出', na=False)].copy()

                # 标记来源
                buy_records['数据来源'] = source
                sell_records['数据来源'] = source
                file_sources.add(source)

                all_buy_records.append(buy_records)
                all_sell_records.append(sell_records)

                # 收集证券名称映射
                for _, row in df.iterrows():
                    code = str(row.get('证券代码', ''))
                    name = str(row.get('证券名称', ''))
                    if code and name and name != 'nan':
                        all_stock_info[code] = name

                print(f"  [{source}] 图片解析完成：买入{len(buy_records)}笔，卖出{len(sell_records)}笔")
                archive_file(image_file)
            except Exception as e:
                print(f"\n[错误] 处理图片文件 {image_file} 时出错：{str(e)}")
                continue

        # 合并所有买卖记录，统一跨账户匹配
        if all_buy_records or all_sell_records:
            merged_buys = pd.concat(all_buy_records, ignore_index=True) if all_buy_records else pd.DataFrame()
            merged_sells = pd.concat(all_sell_records, ignore_index=True) if all_sell_records else pd.DataFrame()

            # 构建统一的df用于calculate_profits
            all_dfs = []
            if len(merged_buys) > 0:
                all_dfs.append(merged_buys)
            if len(merged_sells) > 0:
                all_dfs.append(merged_sells)
            merged_df = pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame()

            # 使用主来源标签（如果只有一个来源就用那个，否则用"多账户合并"）
            primary_source = '+'.join(sorted(file_sources)) if len(file_sources) > 1 else (list(file_sources)[0] if file_sources else '未知')

            print(f"\n  --- 跨账户合并匹配 ---")
            print(f"  合并买入：{len(merged_buys)} 笔")
            print(f"  合并卖出：{len(merged_sells)} 笔")
            print(f"  涉及来源：{', '.join(sorted(file_sources))}")

            profit_results = calculate_profits(merged_df, merged_buys, merged_sells, trading_date, primary_source)

            if profit_results:
                result_df = pd.DataFrame(profit_results)
                total_profit = sum(r['盈亏金额'] for r in profit_results)
                print(f"\n  当日总盈亏：{total_profit:.2f} 元")
                append_to_excel(result_df, trading_date, primary_source)

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
