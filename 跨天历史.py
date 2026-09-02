# -*- coding: utf-8 -*-
"""
跨天配对历史快照与趋势
================================================
解决的问题：跨天/跨月配对分析每次跑完只显示"当下这一刻"的数值，
下次跑就被覆盖，看不到数值是怎么一步步长上来的。

本模块做三件事：
  1) 快照持久化 —— 每次跑批把当次结果存成一条记录（按数据日期 upsert，同日重跑覆盖）
  2) 历史回填   —— 按交易日逐日重算，一次性把历史曲线补出来（不用等几个月攒数据）
  3) 趋势报告   —— 生成自包含 HTML，用图表展示数值变化过程

用法：
  python 跨天历史.py --backfill          回填全部（所有年份 + 所有月份）
  python 跨天历史.py --backfill 2026     只回填 2026 年度曲线
  python 跨天历史.py --backfill 2026-09  只回填 2026-09 月度曲线
  python 跨天历史.py --trend             生成历史趋势报告
  python 跨天历史.py --list              打印快照概况

存储文件：reports/snapshots/跨天配对历史.json
"""
import os
import json
import argparse
from datetime import date

import pandas as pd

from 跨天配对分析 import load, analyze, analyze_year, ECHARTS_LOCAL

SNAP_DIR = os.path.join('reports', 'snapshots')
SNAP_PATH = os.path.join(SNAP_DIR, '跨天配对历史.json')
TREND_OUT = os.path.join('reports', '跨天配对历史趋势.html')

RED = '#d4380d'    # 红涨
GREEN = '#16a34a'  # 绿跌
GREY = '#8c8c8c'
ORANGE = '#f2994a'
BLUE = '#2f80ed'
PURPLE = '#9b51e0'


# ---------------------------------------------------------------- 快照存储
def load_snap():
    if not os.path.exists(SNAP_PATH):
        return {'monthly': {}, 'yearly': {}}
    try:
        with open(SNAP_PATH, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except Exception:
        return {'monthly': {}, 'yearly': {}}
    data.setdefault('monthly', {})
    data.setdefault('yearly', {})
    return data


def save_snap(data):
    os.makedirs(SNAP_DIR, exist_ok=True)
    with open(SNAP_PATH, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=1)


def upsert(scope, key, rec):
    """按数据日期 upsert 一条快照（同一天重跑覆盖旧值）。"""
    data = load_snap()
    bucket = data[scope].setdefault(key, [])
    bucket = [r for r in bucket if r['d'] != rec['d']]
    bucket.append(rec)
    bucket.sort(key=lambda r: r['d'])
    data[scope][key] = bucket
    save_snap(data)
    return data


def upsert_many(scope, key, recs):
    """批量 upsert（回填用，避免逐条重复读写文件）。"""
    data = load_snap()
    bucket = {r['d']: r for r in data[scope].get(key, [])}
    for r in recs:
        bucket[r['d']] = r
    data[scope][key] = sorted(bucket.values(), key=lambda r: r['d'])
    save_snap(data)
    return data


def rec_month(d, r):
    return {
        'd': d,
        'sys': round(r['sys_total'], 2),
        'pair': round(r['pair_pnl'], 2),
        'unmatch': round(r['unmatched_cost'], 2),
        'cross': round(r['cross_net'], 2),
        'corr': round(r['corrected'], 2),
        'remain': len(r['remain']),
    }


def rec_year(d, r):
    return {
        'd': d,
        'sys': round(r['sys_total'], 2),
        'pair': round(r['pair_pnl'], 2),
        'mcross': round(r['monthly_cross'], 2),
        'ycross': round(r['cross_net'], 2),
        'corr': round(r['corrected'], 2),
        'remain': len(r['remain']),
    }


# ---------------------------------------------------------------- 历史回填
def _day_list(sub):
    """取一段数据里的交易日列表（升序）。"""
    days = sub['日期'].astype(str).str[:10].unique().tolist()
    return sorted(days)


def backfill_month(df, ym):
    """按交易日逐日重算某月的跨天配对，返回快照记录列表。"""
    df = df.copy()
    df['_d'] = df['日期'].astype(str).str[:10]
    sub = df[df['ym'] == ym]
    if sub.empty:
        return []
    recs = []
    for d in _day_list(sub):
        r = analyze(sub[sub['_d'] <= d])
        recs.append(rec_month(d, r))
    return recs


def backfill_year(df, year):
    """按交易日逐日重算某年的跨月配对，返回快照记录列表。"""
    df = df.copy()
    df['_d'] = df['日期'].astype(str).str[:10]
    sub = df[df['ym'].str.startswith(str(year))]
    if sub.empty:
        return []
    recs = []
    for d in _day_list(sub):
        r = analyze_year(sub[sub['_d'] <= d], year)
        recs.append(rec_year(d, r))
    return recs


# ---------------------------------------------------------------- 统计
def max_drawdown(series):
    """最大回撤：从历史峰值到后续谷底的最大跌幅（正数表示回撤金额）。"""
    if not series:
        return 0.0
    peak = series[0]
    mdd = 0.0
    for v in series:
        if v > peak:
            peak = v
        dd = peak - v
        if dd > mdd:
            mdd = dd
    return round(mdd, 2)


def fmt(x):
    return '{:,.2f}'.format(x)


def color_pnl(x):
    return RED if x >= 0 else GREEN


def _inline_echarts(html):
    """把本地 ECharts 库内联进 HTML，做成单文件自包含。"""
    tag = '<script src="echarts.min.js"></script>'
    if tag not in html or not os.path.exists(ECHARTS_LOCAL):
        return html
    with open(ECHARTS_LOCAL, 'r', encoding='utf-8') as f:
        lib = f.read().replace('</script', '<\\/script')
    return html.replace(tag, '<script>\n' + lib + '\n</script>')


JS_FMT = """
function fmtV(v){ if(v===null||v===undefined) return '-';
  var s=(v>=0?'+':'')+v.toLocaleString('zh-CN',{minimumFractionDigits:2,maximumFractionDigits:2});
  return s; }
function fmtAbs(v){ if(v===null||v===undefined) return '-';
  return v.toLocaleString('zh-CN',{minimumFractionDigits:2,maximumFractionDigits:2}); }
"""


# ---------------------------------------------------------------- 趋势报告
def _chart_block(div_id, title, dates, series, height=400, unit='元'):
    """生成一个 ECharts 折线图块。series = [{name,color,data,area}]"""
    ser_js = []
    for s in series:
        area = 'true' if s.get('area') else 'false'
        ser_js.append(
            "{name:'%s',type:'line',smooth:true,symbol:'circle',symbolSize:4,"
            "lineStyle:{width:2},itemStyle:{color:'%s'},"
            "areaStyle:%s,data:%s}" % (
                s['name'], s['color'],
                ("{color:'%s',opacity:0.08}" % s['color']) if s.get('area') else 'null',
                json.dumps(s['data'])))
    return '''
    <div class="block">
      <h2>%s</h2>
      <div id="%s" style="width:100%%;height:%dpx;"></div>
    </div>
    <script>
    (function(){
      function fmtV(v){ if(v===null||v===undefined) return '-';
        return (v>=0?'+':'')+v.toLocaleString('zh-CN',{minimumFractionDigits:2,maximumFractionDigits:2}); }
      var c = echarts.init(document.getElementById('%s'));
      c.setOption({
        tooltip:{trigger:'axis',axisPointer:{type:'cross'},
          valueFormatter:function(v){return fmtV(v);}},
        legend:{data:[%s],top:0},
        grid:{left:80,right:30,bottom:60,top:40},
        xAxis:{type:'category',data:%s,axisLabel:{rotate:45}},
        yAxis:{type:'value',name:'%s',axisLabel:{formatter:function(v){return v.toLocaleString('zh-CN');}}},
        dataZoom:[{type:'inside'},{type:'slider',height:18,bottom:12}],
        series:[%s]
      });
      window.addEventListener('resize',function(){c.resize();});
    })();
    </script>
    ''' % (title, div_id, height, div_id,
           ','.join("'%s'" % s['name'] for s in series),
           json.dumps(dates), unit, ','.join(ser_js))


def _bar_block(div_id, title, dates, data, height=320, unit='元'):
    return '''
    <div class="block">
      <h2>%s</h2>
      <div id="%s" style="width:100%%;height:%dpx;"></div>
    </div>
    <script>
    (function(){
      function fmtV(v){ if(v===null||v===undefined) return '-';
        return (v>=0?'+':'')+v.toLocaleString('zh-CN',{minimumFractionDigits:2,maximumFractionDigits:2}); }
      var c = echarts.init(document.getElementById('%s'));
      var vals = %s;
      c.setOption({
        tooltip:{trigger:'axis',valueFormatter:function(v){return fmtV(v);}},
        grid:{left:80,right:30,bottom:60,top:20},
        xAxis:{type:'category',data:%s,axisLabel:{rotate:45}},
        yAxis:{type:'value',name:'%s',axisLabel:{formatter:function(v){return v.toLocaleString('zh-CN');}}},
        dataZoom:[{type:'inside'},{type:'slider',height:18,bottom:12}],
        series:[{name:'日环比',type:'bar',data:vals,
          itemStyle:{color:function(p){return p.value>=0?'%s':'%s';}}}]
      });
      window.addEventListener('resize',function(){c.resize();});
    })();
    </script>
    ''' % (title, div_id, height, div_id, json.dumps(data),
           json.dumps(dates), unit, RED, GREEN)


def history_chart_html(kind, recs, prefix='h'):
    """给月度/年度报告用的「历史轨迹」图块（返回 HTML 字符串，含 div + script）。
    recs 少于 2 条时返回空串（一个点画不出趋势）。
    """
    if not recs or len(recs) < 2:
        return ''
    dates = [r['d'] for r in recs]
    if kind == 'year':
        series = [
            {'name': '系统现有', 'color': GREY, 'data': [r['sys'] for r in recs]},
            {'name': '月度跨天释放', 'color': ORANGE, 'data': [r['mcross'] for r in recs]},
            {'name': '年度跨月释放', 'color': BLUE, 'data': [r['ycross'] for r in recs]},
            {'name': '修正后真实', 'color': RED, 'data': [r['corr'] for r in recs], 'area': True},
        ]
        title = '历史轨迹 · 年度口径随交易日的变化（%d 个快照点）' % len(recs)
    else:
        series = [
            {'name': '系统现有', 'color': GREY, 'data': [r['sys'] for r in recs]},
            {'name': '跨天释放', 'color': ORANGE, 'data': [r['cross'] for r in recs]},
            {'name': '修正后真实', 'color': RED, 'data': [r['corr'] for r in recs], 'area': True},
        ]
        title = '历史轨迹 · 当月口径随交易日的变化（%d 个快照点）' % len(recs)
    return _chart_block(prefix + 'c', title, dates, series, height=380)


def build_trend_report(data):
    yearly = data['yearly']
    monthly = data['monthly']

    # ---------- 年度部分 ----------
    year_blocks = ''
    overview_cards = ''
    for yk in sorted(yearly.keys()):
        recs = yearly[yk]
        if not recs:
            continue
        dates = [r['d'] for r in recs]
        corr = [r['corr'] for r in recs]
        last = recs[-1]
        first = recs[0]
        delta = round(last['corr'] - first['corr'], 2)
        mdd = max_drawdown(corr)

        overview_cards += '''
        <div class="card"><div class="t">%s 年 · 快照点数</div><div class="v">%d 个交易日</div></div>
        <div class="card"><div class="t">%s 年 · 最新修正后</div><div class="v" style="color:%s">%s</div></div>
        <div class="card"><div class="t">%s 年 · 区间变化</div><div class="v" style="color:%s">%s</div></div>
        <div class="card"><div class="t">%s 年 · 最大回撤</div><div class="v" style="color:%s">-%s</div></div>
        ''' % (yk, len(recs),
               yk, color_pnl(last['corr']), fmt(last['corr']),
               yk, color_pnl(delta), ('+' if delta >= 0 else '') + fmt(delta),
               yk, GREEN, fmt(mdd))

        year_blocks += _chart_block(
            'yc_%s' % yk, '%s 年度口径变化轨迹' % yk, dates, [
                {'name': '系统现有', 'color': GREY, 'data': [r['sys'] for r in recs]},
                {'name': '月度跨天释放', 'color': ORANGE, 'data': [r['mcross'] for r in recs]},
                {'name': '年度跨月释放', 'color': BLUE, 'data': [r['ycross'] for r in recs]},
                {'name': '修正后真实', 'color': RED, 'data': corr, 'area': True},
            ], height=430)

        # 日环比增量
        diffs = [None] + [round(corr[i] - corr[i - 1], 2) for i in range(1, len(corr))]
        year_blocks += _bar_block('yd_%s' % yk,
                                  '%s 年 · 修正后盈亏日环比增量' % yk, dates, diffs, height=320)

    # ---------- 月度部分 ----------
    month_blocks = ''
    for mk in sorted(monthly.keys()):
        recs = monthly[mk]
        if len(recs) < 1:
            continue
        dates = [r['d'] for r in recs]
        month_blocks += _chart_block(
            'mc_%s' % mk.replace('-', ''), '%s 月度口径变化轨迹' % mk, dates, [
                {'name': '系统现有', 'color': GREY, 'data': [r['sys'] for r in recs]},
                {'name': '当月跨天释放', 'color': ORANGE, 'data': [r['cross'] for r in recs]},
                {'name': '修正后真实', 'color': RED, 'data': [r['corr'] for r in recs], 'area': True},
            ], height=340)

    # ---------- 明细表（年度最近 30 条） ----------
    table_rows = ''
    if yearly:
        yk = sorted(yearly.keys())[-1]
        recs = yearly[yk][-30:]
        for i, r in enumerate(recs):
            prev = yearly[yk][-31:][i - 1] if i > 0 else None
            dv = round(r['corr'] - prev['corr'], 2) if prev else None
            table_rows += (
                '<tr><td>%s</td><td>%s</td><td>%s</td><td>%s</td>'
                '<td style="color:%s;font-weight:700">%s</td>'
                '<td style="color:%s">%s</td><td>%s</td></tr>'
            ) % (r['d'],
                 fmt(r['sys']),
                 ('+' if r['mcross'] >= 0 else '') + fmt(r['mcross']),
                 ('+' if r['ycross'] >= 0 else '') + fmt(r['ycross']),
                 color_pnl(r['corr']), fmt(r['corr']),
                 (GREY if dv is None else color_pnl(dv)),
                 ('-' if dv is None else ('+' if dv >= 0 else '') + fmt(dv)),
                 r['remain'])
    table_html = ''
    if table_rows:
        table_html = '''
        <div class="block">
          <h2>年度快照明细（最近 30 个交易日）</h2>
          <table>
            <thead><tr><th>数据日期</th><th>系统现有</th><th>月度跨天</th><th>年度跨月</th>
            <th>修正后</th><th>日环比</th><th>跨期仍持有</th></tr></thead>
            <tbody>%s</tbody>
          </table>
        </div>
        ''' % table_rows

    total_pts = sum(len(v) for v in yearly.values()) + sum(len(v) for v in monthly.values())
    subtitle = '年度曲线 %d 条 · 月度曲线 %d 条 · 共 %d 个快照点' % (
        sum(len(v) for v in yearly.values()),
        sum(len(v) for v in monthly.values()), total_pts)

    html = '''
<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>跨天配对历史趋势</title>
<style>
  body{font-family:-apple-system,"Microsoft YaHei",sans-serif;background:#f5f6f8;color:#222;margin:0;padding:24px;}
  h1{font-size:24px;margin:0 0 6px;color:#1a1a1a;}
  .sub{font-size:13px;color:#888;margin-bottom:16px;}
  h2{font-size:19px;border-left:4px solid #d4380d;padding-left:10px;margin:0 0 12px;}
  .overview,.block{background:#fff;border-radius:10px;padding:18px 20px;margin-bottom:18px;box-shadow:0 1px 4px rgba(0,0,0,.06);}
  .cards{display:flex;flex-wrap:wrap;gap:12px;}
  .card{flex:1;min-width:170px;background:#fafafa;border:1px solid #eee;border-radius:8px;padding:12px 14px;}
  .card .t{font-size:12px;color:#888;margin-bottom:6px;}
  .card .v{font-size:20px;font-weight:700;}
  table{width:100%%;border-collapse:collapse;font-size:13px;}
  th,td{border:1px solid #eee;padding:7px 10px;text-align:center;}
  th{background:#fafafa;color:#666;font-weight:600;}
  .hint{font-size:12px;color:#888;line-height:1.8;margin-top:12px;}
  .empty{color:#bbb;padding:20px;text-align:center;}
</style></head>
<body>
  <script src="echarts.min.js"></script>
  <script>%s</script>
  <h1>跨天配对历史趋势</h1>
  <div class="sub">%s · 生成于 %s</div>
  <div class="overview">
    <div class="cards">%s</div>
    <p class="hint">
      <b>口径说明：</b>系统现有 = 当日配对 + 未平仓成本；修正后 = 月度跨天释放 + 年度跨月释放 + 当日配对。<br>
      <b>怎么看：</b>「修正后真实」曲线是累计值，随每天跑批更新；「日环比增量」柱子是当天新产生的变化量，
      柱子为正说明当天跨天配对释放了盈利（或减少了浮亏计提）。最大回撤 = 修正后曲线从历史峰值到谷底的最大跌幅。<br>
      <b>红涨绿跌。</b>快照按数据日期去重，同一天重复跑批会覆盖旧值。
    </p>
  </div>
  %s
  %s
  %s
</body></html>
''' % (JS_FMT, subtitle, date.today().isoformat(),
       overview_cards or '<div class="empty">暂无年度快照，先跑 --backfill</div>',
       year_blocks, month_blocks, table_html)

    html = _inline_echarts(html)
    os.makedirs('reports', exist_ok=True)
    with open(TREND_OUT, 'w', encoding='utf-8') as f:
        f.write(html)
    return TREND_OUT


# ---------------------------------------------------------------- CLI
def main():
    ap = argparse.ArgumentParser(description='跨天配对历史快照与趋势')
    ap.add_argument('--backfill', nargs='?', const='ALL', default=None,
                    help='回填历史：省略=全部，2026=该年，2026-09=该月')
    ap.add_argument('--trend', action='store_true', help='生成历史趋势报告')
    ap.add_argument('--list', action='store_true', help='打印快照概况')
    args = ap.parse_args()

    if args.list:
        data = load_snap()
        print('快照文件：', SNAP_PATH)
        print('年度曲线：')
        for k in sorted(data['yearly']):
            v = data['yearly'][k]
            last = v[-1] if v else None
            print('  %s  %3d 点  %s ~ %s  最新修正后=%s' % (
                k, len(v), v[0]['d'] if v else '-', v[-1]['d'] if v else '-',
                fmt(last['corr']) if last else '-'))
        print('月度曲线：')
        for k in sorted(data['monthly']):
            v = data['monthly'][k]
            last = v[-1] if v else None
            print('  %s  %3d 点  %s ~ %s  最新修正后=%s' % (
                k, len(v), v[0]['d'] if v else '-', v[-1]['d'] if v else '-',
                fmt(last['corr']) if last else '-'))
        return

    did = False
    if args.backfill:
        did = True
        df = load()
        target = args.backfill
        if target == 'ALL':
            years = sorted(set(m[:4] for m in df['ym'].unique()))
            months = sorted(df['ym'].unique())
        elif len(target) == 4 and target.isdigit():
            years = [target]
            months = [m for m in df['ym'].unique() if m.startswith(target)]
        else:
            years = []
            months = [target]

        for y in years:
            print('回填年度 %s ...' % y, end=' ', flush=True)
            recs = backfill_year(df, y)
            upsert_many('yearly', y, recs)
            print('%d 点  最新=%s' % (len(recs), fmt(recs[-1]['corr']) if recs else '-'))
        for m in months:
            print('回填月度 %s ...' % m, end=' ', flush=True)
            recs = backfill_month(df, m)
            upsert_many('monthly', m, recs)
            print('%d 点  最新=%s' % (len(recs), fmt(recs[-1]['corr']) if recs else '-'))

    if args.trend or (did and not args.list):
        data = load_snap()
        out = build_trend_report(data)
        print('趋势报告已生成：', out)


if __name__ == '__main__':
    main()
