import re, sys
mode = sys.argv[1]  # 'no-markpoint' or 'no-markarea'
html = open(r'D:/短线收益/reports/汇总可视化报告.html', encoding='utf-8').read()
if mode == 'no-markpoint':
    new = html.replace(
        ", markPoint:{ data: markPts }, markArea:{ itemStyle:{ color:'rgba(231,76,60,0.12)' }, data: ddArea }",
        ", markArea:{ itemStyle:{ color:'rgba(231,76,60,0.12)' }, data: ddArea }"
    ).replace(
        ", markPoint:{ data: mcmPts }, markArea:{ itemStyle:{ color:'rgba(231,76,60,0.12)' }, data: mcDdArea }",
        ", markArea:{ itemStyle:{ color:'rgba(231,76,60,0.12)' }, data: mcDdArea }"
    )
elif mode == 'no-markarea':
    new = html.replace(
        ", markPoint:{ data: markPts }, markArea:{ itemStyle:{ color:'rgba(231,76,60,0.12)' }, data: ddArea }",
        ", markPoint:{ data: markPts }"
    ).replace(
        ", markPoint:{ data: mcmPts }, markArea:{ itemStyle:{ color:'rgba(231,76,60,0.12)' }, data: mcDdArea }",
        ", markPoint:{ data: mcmPts }"
    )
else:
    raise ValueError('bad mode')
print(f'{mode} 替换:', html != new)
open(r'D:/短线收益/reports/汇总可视化报告_debug.html', 'w', encoding='utf-8').write(new)
