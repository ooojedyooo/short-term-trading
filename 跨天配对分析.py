# -*- coding: utf-8 -*-
"""
月度/年度跨天配对分析
- 数据来源：股票交易盈亏汇总.xlsx（主汇总表）
- 口径：按月或按年，聚合范围内每只股票的"未平仓剩余买卖"（多买/多卖/仅买/仅卖行），
       用加权均价(WAC, 与系统一致)做跨天/跨月配对，计算跨天释放的净盈亏。
- 不重算每日已配对部分；仅对"未匹配"的买卖做整合匹配。
- 用法：
    默认（无参数）      ：全周期逐月总览
    --month            ：算上月（每月初跑，自动取上个月）
    --month YYYY-MM    ：算指定月
    --year             ：算去年（每年初跑，自动取上一年）
    --year YYYY        ：算指定年（全年未平仓行跨月全局配对）
- 剩余未匹配持仓会导出为 Excel 文件，便于留档与下一步消费。
"""
import pandas as pd
import json
import argparse
from datetime import date

PATH = '股票交易盈亏汇总.xlsx'
COMMISSION_RATE = 0.0001  # 万一
MIN_COMMISSION = 5.0
STAMP_DUTY_RATE = 0.0005  # 万五，卖出单边


def load():
    df = pd.read_excel(PATH, sheet_name='股票盈亏汇总')
    df['ym'] = df['日期'].astype(str).str[:7]
    # 行类型：含 '%' = 配对行（已配对）；含 '⚠' = 未平仓行
    def tag(x):
        s = str(x)
        if '%' in s:
            return 'pair'
        if '⚠' in s:
            return 'unmatched'
        return 'other'
    df['type'] = df['盈亏比例'].apply(tag)
    cols = ['买入数量', '卖出数量', '匹配数量', '买入均价', '卖出均价',
            '买入金额', '卖出金额', '盈亏金额']
    for c in cols:
        df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
    return df


def analyze(sub):
    """对一段数据（某月 / 某年全量）做跨天/跨月配对。
    返回该范围的配对结果 dict。"""
    sys_total = sub['盈亏金额'].sum()
    pair_pnl = sub[sub['type'] == 'pair']['盈亏金额'].sum()
    unmatched_cost = sub[sub['type'] == 'unmatched']['盈亏金额'].sum()

    um = sub[sub['type'] == 'unmatched']
    # 按证券代码聚合（名称可能有截断/变体，如"中国稀土"vs"中国稀士"，
    # 必须用代码而非(代码,名称)分组，否则同票被拆成多只导致漏配）
    code_name = um.groupby('证券代码')['证券名称'].agg(
        lambda s: s[s.astype(str) != 'nan'].value_counts().index[0] if (s.astype(str) != 'nan').any() else str(s.iloc[0]))
    agg = um.groupby('证券代码').agg(
        buy_q=('买入数量', 'sum'),
        buy_amt=('买入金额', 'sum'),
        sell_q=('卖出数量', 'sum'),
        sell_amt=('卖出金额', 'sum'),
    ).reset_index()

    cross = []      # 跨天/跨月实际发生配对
    remain = []     # 配对后（或纯单边）仍完全未平仓
    cross_net = 0.0

    for _, r in agg.iterrows():
        bq, bamt, sq, samt = r['buy_q'], r['buy_amt'], r['sell_q'], r['sell_amt']
        code = r['证券代码']
        name = code_name.get(code, str(code))
        if bq > 0 and sq > 0:
            match = min(bq, sq)
            bavg = bamt / bq
            savg = samt / sq
            m_bamt = bamt * (match / bq)
            m_samt = samt * (match / sq)
            gross = m_samt - m_bamt
            comm = max(m_bamt * COMMISSION_RATE, MIN_COMMISSION) + \
                   max(m_samt * COMMISSION_RATE, MIN_COMMISSION)
            stamp = m_samt * STAMP_DUTY_RATE
            cost = comm + stamp
            net = gross - cost
            cross_net += net
            rbuy = bq - match
            rsell = sq - match
            cross.append({
                'code': code, 'name': name,
                'buy_q': int(bq), 'buy_amt': round(bamt, 2), 'buy_avg': round(bavg, 3),
                'sell_q': int(sq), 'sell_amt': round(samt, 2), 'sell_avg': round(savg, 3),
                'match': int(match), 'gross': round(gross, 2),
                'cost': round(cost, 2), 'net': round(net, 2),
            })
            if rbuy > 0 or rsell > 0:
                rba = round(bamt - m_bamt, 2) if rbuy > 0 else 0
                rsa = round(samt - m_samt, 2) if rsell > 0 else 0
                remain.append({
                    'code': code, 'name': name,
                    'remain_buy': int(rbuy), 'remain_sell': int(rsell),
                    'remain_buy_amt': rba, 'remain_sell_amt': rsa,
                    'note': ('仍买入未平仓 {}股'.format(int(rbuy)) if rbuy > 0 else '') +
                            ('；仍卖出未平仓 {}股'.format(int(rsell)) if rsell > 0 else ''),
                })
        else:
            rba = round(bamt, 2) if bq > 0 else 0
            rsa = round(samt, 2) if sq > 0 else 0
            remain.append({
                'code': code, 'name': name,
                'remain_buy': int(bq), 'remain_sell': int(sq),
                'remain_buy_amt': rba, 'remain_sell_amt': rsa,
                'note': ('仅买入未平仓 {}股（跨期持有底仓）'.format(int(bq)) if bq > 0
                         else '仅卖出未平仓 {}股（平旧仓，无对应买入）'.format(int(sq))),
            })

    corrected = pair_pnl + cross_net
    return {
        'sys_total': round(sys_total, 2),
        'pair_pnl': round(pair_pnl, 2),
        'unmatched_cost': round(unmatched_cost, 2),
        'cross_net': round(cross_net, 2),
        'corrected': round(corrected, 2),
        'cross': cross,
        'remain': remain,
    }


def export_remain(remain, path):
    rows = [{
        '证券代码': r['code'], '证券名称': r['name'],
        '剩余买入数量': r['remain_buy'], '剩余买入金额': r.get('remain_buy_amt', 0),
        '剩余卖出数量': r['remain_sell'], '剩余卖出金额': r.get('remain_sell_amt', 0),
        '说明': r['note'],
    } for r in remain]
    pd.DataFrame(rows).to_excel(path, index=False)


def fmt(x):
    return '{:,.2f}'.format(x)


def color_pnl(x):
    return '#d4380d' if x >= 0 else '#16a34a'  # 红涨绿跌


def build_html(results, months, title, out_path, show_chart, mode_label):
    tot_sys = tot_corrected = tot_cross = 0.0
    tot_remain = 0
    chart_months, chart_sys, chart_corr = [], [], []

    blocks = []
    for ym in months:
        r = results[ym]
        tot_sys += r['sys_total']
        tot_corrected += r['corrected']
        tot_cross += r['cross_net'] + r.get('monthly_cross', 0)
        tot_remain += len(r['remain'])
        chart_months.append(ym)
        chart_sys.append(r['sys_total'])
        chart_corr.append(r['corrected'])

        cross_rows = ''
        for c in r['cross']:
            cross_rows += (
                '<tr>'
                '<td>{}<span class="code">{}</span></td>'
                '<td>{}股<br><span class="sub">{}/股</span></td>'
                '<td>{}<br><span class="sub">{}/股</span></td>'
                '<td class="hl">{}股</td>'
                '<td style="color:{}">{}</td>'
                '<td>{}</td>'
                '<td style="color:{};font-weight:700">{}</td>'
                '</tr>'
            ).format(
                c['name'], c['code'],
                '{:,}'.format(c['buy_q']), fmt(c['buy_amt']),
                '{:,}'.format(c['sell_q']), fmt(c['sell_amt']),
                '{:,}'.format(c['match']),
                color_pnl(c['gross']), fmt(c['gross']),
                fmt(c['cost']),
                color_pnl(c['net']), fmt(c['net']),
            )
        if not cross_rows:
            cross_rows = '<tr><td colspan="7" class="empty">该范围无非当日配对的跨期买卖，无需整合</td></tr>'

        remain_rows = ''
        for rm in r['remain']:
            remain_rows += (
                '<tr><td>{}<span class="code">{}</span></td>'
                '<td>{}</td><td class="note">{}</td></tr>'
            ).format(rm['name'], rm['code'],
                     '{:,}'.format(rm['remain_buy']) if rm['remain_buy'] else '-',
                     rm['note'])
        if not remain_rows:
            remain_rows = '<tr><td colspan="3" class="empty">无跨期仍持有的未平仓持仓</td></tr>'

        delta = r['corrected'] - r['sys_total']
        delta_str = ('+' if delta >= 0 else '') + fmt(delta)
        if 'monthly_cross' in r:
            mc = r['monthly_cross']
            cards_html = '''
            <div class="card"><div class="t">系统现有口径盈亏</div><div class="v" style="color:{c1}">{v1}</div></div>
            <div class="card"><div class="t">月度跨天释放盈亏</div><div class="v" style="color:{c2}">{v2}</div></div>
            <div class="card"><div class="t">年度跨月释放盈亏</div><div class="v" style="color:{c3}">{v3}</div></div>
            <div class="card"><div class="t">修正后真实盈亏</div><div class="v" style="color:{c4}">{v4}</div></div>
            <div class="card"><div class="t">较现有偏差</div><div class="v" style="color:{c5}">{v5}</div></div>
            '''.format(
                c1=color_pnl(r['sys_total']), v1=fmt(r['sys_total']),
                c2=color_pnl(mc), v2=('+' if mc >= 0 else '') + fmt(mc),
                c3=color_pnl(r['cross_net']), v3=('+' if r['cross_net'] >= 0 else '') + fmt(r['cross_net']),
                c4=color_pnl(r['corrected']), v4=fmt(r['corrected']),
                c5=color_pnl(delta), v5=delta_str)
        else:
            cards_html = '''
            <div class="card"><div class="t">系统现有口径盈亏</div><div class="v" style="color:{c1}">{v1}</div></div>
            <div class="card"><div class="t">跨期配对释放盈亏</div><div class="v" style="color:{c2}">{v2}</div></div>
            <div class="card"><div class="t">修正后真实盈亏</div><div class="v" style="color:{c3}">{v3}</div></div>
            <div class="card"><div class="t">较现有偏差</div><div class="v" style="color:{c4}">{v4}</div></div>
            '''.format(
                c1=color_pnl(r['sys_total']), v1=fmt(r['sys_total']),
                c2=color_pnl(r['cross_net']), v2=('+' if r['cross_net'] >= 0 else '') + fmt(r['cross_net']),
                c3=color_pnl(r['corrected']), v3=fmt(r['corrected']),
                c4=color_pnl(delta), v4=delta_str)
        blocks.append('''
        <div class="block">
          <h2>{} {}</h2>
          <div class="cards">{}</div>
          <h3>跨期配对明细（未平仓买卖整合匹配）</h3>
          <table>
            <thead><tr><th>股票</th><th>跨期买入剩余</th><th>跨期卖出剩余</th><th>跨期匹配</th><th>毛盈亏</th><th>交易成本</th><th>跨期净盈亏</th></tr></thead>
            <tbody>{}</tbody>
          </table>
          <h3>跨期仍完全未平仓持仓（已导出Excel）</h3>
          <table class="small">
            <thead><tr><th>股票</th><th>剩余买入</th><th>说明</th></tr></thead>
            <tbody>{}</tbody>
          </table>
        </div>
        '''.format(mode_label, ym, cards_html, cross_rows, remain_rows))

    overview = '''
    <div class="overview">
      <h2>总览</h2>
      <div class="cards">
        <div class="card"><div class="t">覆盖范围</div><div class="v">{}</div></div>
        <div class="card"><div class="t">系统现有累计盈亏</div><div class="v" style="color:{}">{}</div></div>
        <div class="card"><div class="t">跨期配对释放盈亏</div><div class="v" style="color:{}">{}</div></div>
        <div class="card"><div class="t">修正后累计真实盈亏</div><div class="v" style="color:{}">{}</div></div>
        <div class="card"><div class="t">跨期仍持有持仓数</div><div class="v">{} 只</div></div>
      </div>
      <p class="hint">说明：跨期配对 = 将范围内每日配对后剩余的零散买入/卖出，按股票整合后用加权均价(WAC)配对。
      系统现有口径把"未平仓"记为当日交易成本（负数）；修正后口径将其替换为跨期配对的真实净盈亏。红涨绿跌。
      剩余未匹配持仓已导出为 Excel 文件（见同目录 剩余持仓_*.xlsx）。</p>
    </div>
    '''.format(
        title,
        color_pnl(tot_sys), fmt(tot_sys),
        color_pnl(tot_cross), (('+' if tot_cross >= 0 else '') + fmt(tot_cross)),
        color_pnl(tot_corrected), fmt(tot_corrected),
        tot_remain,
    )

    chart_div = ''
    chart_js = ''
    if show_chart and len(months) > 1:
        chart_div = '''
        <div class="block">
          <h2>各月盈亏对比</h2>
          <div id="chart" style="width:100%;height:420px;"></div>
        </div>
        '''
        chart_js = '''
        <script src="https://cdn.jsdelivr.net/npm/echarts@5/dist/echarts.min.js"></script>
        <script>
        var chart = echarts.init(document.getElementById('chart'));
        var months = %s;
        var sys = %s;
        var corr = %s;
        chart.setOption({
          tooltip: { trigger: 'axis', valueFormatter: function(v){ return (v>=0?'+':'')+v.toLocaleString('zh-CN',{minimumFractionDigits:2}); } },
          legend: { data: ['系统现有口径','修正后真实'] },
          grid: { left: 70, right: 30, bottom: 50, top: 40 },
          xAxis: { type: 'category', data: months },
          yAxis: { type: 'value', name: '盈亏(元)' },
          series: [
            { name:'系统现有口径', type:'bar', data: sys, itemStyle:{color:'#bfbfbf'} },
            { name:'修正后真实', type:'bar', data: corr, itemStyle:{color:'#d4380d'} }
          ]
        });
        </script>
        ''' % (json.dumps(chart_months, ensure_ascii=False),
               json.dumps(chart_sys), json.dumps(chart_corr))

    html = '''
    <!DOCTYPE html>
    <html lang="zh-CN"><head><meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>{}</title>
    <style>
      body{{font-family:-apple-system,"Microsoft YaHei",sans-serif;background:#f5f6f8;color:#222;margin:0;padding:24px;}}
      h1{{font-size:24px;margin:0 0 16px;color:#1a1a1a;}}
      h2{{font-size:19px;border-left:4px solid #d4380d;padding-left:10px;margin:24px 0 12px;}}
      h3{{font-size:15px;color:#555;margin:18px 0 8px;}}
      .overview,.block{{background:#fff;border-radius:10px;padding:18px 20px;margin-bottom:18px;box-shadow:0 1px 4px rgba(0,0,0,.06);}}
      .cards{{display:flex;flex-wrap:wrap;gap:12px;}}
      .card{{flex:1;min-width:150px;background:#fafafa;border:1px solid #eee;border-radius:8px;padding:12px 14px;}}
      .card .t{{font-size:12px;color:#888;margin-bottom:6px;}}
      .card .v{{font-size:20px;font-weight:700;}}
      table{{width:100%%;border-collapse:collapse;font-size:13px;}}
      th,td{{border:1px solid #eee;padding:8px 10px;text-align:center;}}
      th{{background:#fafafa;color:#666;font-weight:600;}}
      .hl{{color:#d4380d;font-weight:700;}}
      .code{{display:block;font-size:11px;color:#aaa;margin-top:2px;}}
      .sub{{font-size:11px;color:#999;}}
      .note{{color:#b06a00;font-size:12px;}}
      .empty{{color:#bbb;padding:14px;}}
      .hint{{font-size:12px;color:#888;line-height:1.7;margin-top:12px;}}
      table.small{{font-size:12px;}}
    </style></head>
    <body>
      <h1>{}</h1>
      {}
      {}
      {}
      {}
    </body></html>
    '''.format(title, title, overview, ''.join(blocks), chart_div, chart_js)

    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)


def resolved_month(arg):
    if arg is None:
        return None
    if arg == 'AUTO':
        t = date.today()
        if t.month == 1:
            return '{:04d}-{:02d}'.format(t.year - 1, 12)
        return '{:04d}-{:02d}'.format(t.year, t.month - 1)
    return arg


def resolved_year(arg):
    if arg is None:
        return None
    if arg == 'AUTO':
        return date.today().year - 1
    return int(arg)


def main():
    ap = argparse.ArgumentParser(description='月度/年度跨天配对分析')
    ap.add_argument('--month', nargs='?', const='AUTO', default=None,
                    help='算上月(--month)或指定月(--month YYYY-MM)')
    ap.add_argument('--year', nargs='?', const='AUTO', default=None,
                    help='算去年(--year)或指定年(--year YYYY)')
    args = ap.parse_args()

    df = load()
    all_months = sorted(df['ym'].unique())

    mode = 'all'
    months = all_months
    year = None
    if args.year is not None:
        year = resolved_year(args.year)
        months = [m for m in all_months if m.startswith(str(year))]
        if not months:
            print('该年无数据:', year)
            return
        mode = 'year'
    elif args.month is not None:
        mtag = resolved_month(args.month)
        if mtag not in all_months:
            print('该月无数据:', mtag)
            return
        months = [mtag]
        mode = 'month'

    # 计算
    if mode == 'year':
        # 忠实流水线：每月先配对取剩余（带金额），再把各月剩余跨月合并配对
        year_months = [m for m in all_months if m.startswith(str(year))]
        month_results = {m: analyze(df[df['ym'] == m]) for m in year_months}
        monthly_cross_sum = sum(r['cross_net'] for r in month_results.values())

        # 合并各月剩余（跨月持有池）
        merged = {}
        for m, r in month_results.items():
            for rem in r['remain']:
                d = merged.setdefault(rem['code'], {
                    'code': rem['code'], 'name': rem['name'],
                    'buy_q': 0, 'buy_amt': 0.0, 'sell_q': 0, 'sell_amt': 0.0})
                d['buy_q'] += rem['remain_buy']
                d['buy_amt'] += rem.get('remain_buy_amt', 0)
                d['sell_q'] += rem['remain_sell']
                d['sell_amt'] += rem.get('remain_sell_amt', 0)

        # 跨月配对（流水线年度层）
        cross_month = []
        remain_final = []
        cross_month_net = 0.0
        for code, d in merged.items():
            bq, bamt, sq, samt = d['buy_q'], d['buy_amt'], d['sell_q'], d['sell_amt']
            if bq > 0 and sq > 0:
                match = min(bq, sq)
                bavg = bamt / bq
                savg = samt / sq
                m_bamt = bamt * (match / bq)
                m_samt = samt * (match / sq)
                gross = m_samt - m_bamt
                comm = max(m_bamt * COMMISSION_RATE, MIN_COMMISSION) + \
                       max(m_samt * COMMISSION_RATE, MIN_COMMISSION)
                stamp = m_samt * STAMP_DUTY_RATE
                cost = comm + stamp
                net = gross - cost
                cross_month_net += net
                rbuy = bq - match
                rsell = sq - match
                cross_month.append({
                    'code': code, 'name': d['name'],
                    'buy_q': int(bq), 'buy_amt': round(bamt, 2), 'buy_avg': round(bavg, 3),
                    'sell_q': int(sq), 'sell_amt': round(samt, 2), 'sell_avg': round(savg, 3),
                    'match': int(match), 'gross': round(gross, 2),
                    'cost': round(cost, 2), 'net': round(net, 2),
                })
                if rbuy > 0 or rsell > 0:
                    remain_final.append({
                        'code': code, 'name': d['name'],
                        'remain_buy': int(rbuy), 'remain_sell': int(rsell),
                        'note': ('仍买入未平仓 {}股'.format(int(rbuy)) if rbuy > 0 else '') +
                                ('；仍卖出未平仓 {}股'.format(int(rsell)) if rsell > 0 else ''),
                    })
            else:
                remain_final.append({
                    'code': code, 'name': d['name'],
                    'remain_buy': int(bq), 'remain_sell': int(sq),
                    'note': ('仅买入未平仓 {}股（跨年持有底仓）'.format(int(bq)) if bq > 0
                             else '仅卖出未平仓 {}股（平旧仓，无对应买入）'.format(int(sq))),
                })

        # 全年系统现有口径
        sub_all = df[df['ym'].str.startswith(str(year))]
        sys_total = sub_all['盈亏金额'].sum()
        pair_pnl = sub_all[sub_all['type'] == 'pair']['盈亏金额'].sum()
        corrected = pair_pnl + monthly_cross_sum + cross_month_net

        results = {str(year): {
            'sys_total': round(sys_total, 2),
            'pair_pnl': round(pair_pnl, 2),
            'unmatched_cost': 0,
            'cross_net': round(cross_month_net, 2),
            'monthly_cross': round(monthly_cross_sum, 2),
            'corrected': round(corrected, 2),
            'cross': cross_month,
            'remain': remain_final,
        }}
        months = [str(year)]
        title = '{} 年度跨天配对分析（月度剩余跨月进一步匹配）'.format(year)
        out = 'reports/年度跨天配对分析_{}.html'.format(year)
        remain_path = 'reports/剩余持仓_{}年度.xlsx'.format(year)
    elif mode == 'month':
        mtag = months[0]
        results = {mtag: analyze(df[df['ym'] == mtag])}
        title = '{} 月度跨天配对分析'.format(mtag)
        out = 'reports/月度跨天配对分析_{}.html'.format(mtag)
        remain_path = 'reports/剩余持仓_{}.xlsx'.format(mtag)
    else:
        results = {m: analyze(df[df['ym'] == m]) for m in months}
        title = '全周期月度跨天配对分析总览'
        out = 'reports/月度跨天配对分析.html'
        remain_path = None

    show_chart = (mode == 'all')
    mode_label = {'month': '月度', 'year': '年度', 'all': '月度'}[mode]

    build_html(results, months, title, out, show_chart, mode_label)
    print('=' * 70)
    print(title)
    print('=' * 70)
    for ym in months:
        r = results[ym]
        if 'monthly_cross' in r:
            print('{}  现有={:>12,.2f}  月度跨天释放={:>12,.2f}  年度跨月释放={:>12,.2f}  修正后={:>12,.2f}  跨年仍持有={}'.format(
                ym, r['sys_total'], r['monthly_cross'], r['cross_net'], r['corrected'], len(r['remain'])))
        else:
            print('{}  现有={:>12,.2f}  跨期释放={:>12,.2f}  修正后={:>12,.2f}  跨期仍持有={}'.format(
                ym, r['sys_total'], r['cross_net'], r['corrected'], len(r['remain'])))
        # 抽样核对：跨期配对净盈亏 Top3
        for c in sorted(r['cross'], key=lambda x: x['net'], reverse=True)[:3]:
            print('   ↳ {} {}  买{}股@{} 卖{}股@{} 匹配{} 净{}'.format(
                c['name'], c['code'],
                '{:,}'.format(c['buy_q']), c['buy_avg'],
                '{:,}'.format(c['sell_q']), c['sell_avg'],
                '{:,}'.format(c['match']), fmt(c['net'])))
    tot_sys = sum(r['sys_total'] for r in results.values())
    tot_cross = sum(r['cross_net'] + r.get('monthly_cross', 0) for r in results.values())
    tot_corr = sum(r['corrected'] for r in results.values())
    tot_remain = sum(len(r['remain']) for r in results.values())
    print('-' * 70)
    print('合计  现有={:>12,.2f}  跨期释放={:>12,.2f}  修正后={:>12,.2f}  跨期仍持有={}'.format(
        tot_sys, tot_cross, tot_corr, tot_remain))
    print('报告已生成：', out)
    if remain_path is not None:
        export_remain(results[months[0]]['remain'], remain_path)
        print('剩余持仓已导出：', remain_path)


if __name__ == '__main__':
    main()
