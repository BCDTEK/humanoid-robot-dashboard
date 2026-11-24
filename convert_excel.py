# -*- coding: utf-8 -*-
"""
人形机器人看板数据源 → dashboard.json 转换脚本（2026-11 版）
读取 5 张分表（①企业 / ②产品 / ③融资 / ④市场规模 / ⑤应用场景），
生成 index.html 直接 fetch 的 dashboard.json。

升级点（相对旧版）：
1. 新增 ⑤ 应用场景聚合（场景卡 / 案例 / 时间轴），此前该板块为硬编码；
2. meta 增加 sections（分板块更新时间），updated 由数据源最后更新日自动取；
3. 市场规模趋势：同一机构同一年份出现多版数据时，取“机构发布时间”最新的一版；
   单季（Q1-Q4）数据只进明细表，不进入年度趋势线；上半年(H1)单独标注。
"""
import openpyxl, json, os, re
from collections import Counter

HERE = os.path.dirname(os.path.abspath(__file__))
_CANDS = ['人形机器人看板数据源.xlsx', 'data.xlsx']
SRC = next((os.path.join(HERE, f) for f in _CANDS if os.path.exists(os.path.join(HERE, f))),
           os.path.join(HERE, _CANDS[0]))
OUT = os.path.join(HERE, 'dashboard.json')


def num(v):
    try:
        f = float(v)
        return f if f == int(f) else round(f, 4)
    except Exception:
        return None


def clean(v):
    if v is None:
        return None
    s = str(v).strip()
    if s == '' or s.upper() == 'NA':
        return None
    return v


def dt(v):
    s = str(v or '').strip()
    return s


wb = openpyxl.load_workbook(SRC, data_only=True)


def rows_of(name):
    ws = wb[name]
    hdr = [c.value for c in ws[1]]
    out = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r is None or all(v is None or str(v).strip() == '' for v in r):
            continue
        out.append({hdr[i]: r[i] for i in range(min(len(hdr), len(r)))})
    return out


companies = rows_of('① 企业名单')
products = rows_of('② 产品名单')
financing = rows_of('③ 融资情况')
market = rows_of('④ 市场规模')
scene_rows = rows_of('⑤ 应用场景')

# 数据源最后更新日（各表“更新时间”列的最大值）
all_upd = [dt(r.get('更新时间')) for r in companies + products + financing + market + scene_rows]
all_upd = [x for x in all_upd if x]
UPDATED = max(all_upd) if all_upd else ''
SECTIONS = {k: UPDATED for k in ('market', 'compete', 'scene', 'finance', 'policy', 'news')}

# ===== 市场规模聚合：KPI + 多机构趋势 series =====
def mk_market_series():
    series = {}
    pub_of = {}

    def add(scope, metric, org, y, v, u, f, pub):
        if metric not in ('出货量', '市场规模'):
            return
        key = f'{scope}/{metric}'
        series.setdefault(key, {}).setdefault(org, {'y': [], 'v': [], 'f': [], 'u': u})
        d = series[key][org]
        idx = None
        for i, yy in enumerate(d['y']):
            if str(yy) == str(y):
                idx = i
                break
        if idx is None:
            d['y'].append(y); d['v'].append(v); d['f'].append(f)
            pub_of[(key, org, str(y))] = pub
        else:
            k2 = (key, org, str(y))
            if pub > pub_of.get(k2, ''):
                print(f'  [趋势替换] {key} {org} {y}: {d["v"][idx]} -> {v} (发布时间 {pub_of.get(k2)} -> {pub})')
                d['v'][idx] = v; d['f'][idx] = f; d['u'] = u
                pub_of[k2] = pub

    for row in market:
        ent = clean(row.get('企业/全体市场'))
        ent = str(ent).strip() if ent else '全体市场'
        if ent != '全体市场':
            continue
        scope = clean(row.get('范围')) or '全体'
        metric = clean(row.get('指标类型'))
        if metric not in ('出货量', '市场规模'):
            continue
        y = clean(row.get('数据所属年份'))
        try:
            y = int(float(y))
        except Exception:
            continue
        bm = str(clean(row.get('备注')) or '')
        if re.search(r'Q[1-4]|单季', bm):
            continue  # 单季数据不进年度趋势
        if re.search(r'上半年|H1|上半年度', bm):
            y = f'{y}H1'
        org = clean(row.get('报告机构'))
        v = num(clean(row.get('数值')))
        u = clean(row.get('量纲'))
        f = clean(row.get('预测标记'))
        f = 'F' if f == 'F' else None
        pub = dt(row.get('机构发布时间'))
        if metric == '出货量' and u == '台' and v is not None:
            v = round(v / 10000, 4); u = '万台'
        if v is None:
            continue
        add(scope, metric, org, y, v, u, f, pub)

    for orgs in series.values():
        for o, s in orgs.items():
            z = sorted(zip([str(x) for x in s['y']], s['v'], s['f']))
            s['y'] = [x[0] for x in z]; s['v'] = [x[1] for x in z]; s['f'] = [x[2] for x in z]
    return series


def range_kpi(series, scope, metric, year, note):
    orgs = series.get(f'{scope}/{metric}', {})
    groups = {}
    for o, d in orgs.items():
        for i, yy in enumerate(d['y']):
            if str(yy) == str(year) and d['f'][i] == 'F':
                groups.setdefault(d['u'], []).append((d['v'][i], o))
    if not groups:
        return {'t': f'{year}{scope}{metric}预测', 'min': None, 'max': None,
                'unit': '', 'orgs': [], 'note': note + '（暂无机构预测数据）'}
    main_u = max(groups, key=lambda u: len(groups[u]))
    vals = [v for v, _ in groups[main_u]]
    orgs_list = [o for _, o in groups[main_u]]
    others = []
    for u, lst in groups.items():
        if u != main_u:
            others.append(u + '：' + '、'.join(o for _, o in lst))
    extra = ('；另见 ' + '；'.join(others)) if others else ''
    return {'t': f'{year}{scope}{metric}预测', 'min': min(vals), 'max': max(vals),
            'unit': main_u, 'orgs': orgs_list, 'note': note + extra}


series = mk_market_series()
kpis = [
    dict(range_kpi(series, '全球', '出货量', 2026, '全球人形机器人出货量 2026年预测'), t='2026全球出货量预测'),
    dict(range_kpi(series, '中国', '出货量', 2026, '中国人形机器人出货量 2026年预测'), t='2026中国出货量预测'),
    dict(range_kpi(series, '全球', '市场规模', 2026, '全球人形机器人市场规模 2026年预测'), t='2026全球市场规模预测'),
    dict(range_kpi(series, '中国', '市场规模', 2026, '中国人形机器人市场规模 2026年预测'), t='2026中国市场规模预测'),
]


# ===== 产品竞争格局 =====
def mk_compete():
    prodrows = []
    for p in products:
        pr = clean(p.get('售价数值')); pu = clean(p.get('售价单位'))
        try:
            prtxt = f'{float(pr):g}{pu}' if (pr is not None and pu) else None
        except Exception:
            prtxt = None
        ship = clean(p.get('出货数值'))
        ship_txt = f'{ship:g}' if isinstance(ship, (int, float)) else None
        spec = clean(p.get('屏幕规格'))
        if spec and len(str(spec)) > 34:
            spec = str(spec)[:34] + '…'
        co = clean(p.get('企业名称'))
        if co:
            co = str(co).split('（')[0]
        prodrows.append({'co': co, 'name': clean(p.get('产品名称')), 'stage': clean(p.get('阶段')),
                         'price': prtxt, 'time': clean(p.get('发布时间')), 'ship': ship_txt,
                         'screen': clean(p.get('是否带屏')), 'spec': spec,
                         'upd': dt(p.get('更新时间'))})
    stage = Counter(); screen = Counter()
    for p in products:
        st = clean(p.get('阶段'))
        stage[str(st) if st else '未知'] += 1
        sc = clean(p.get('是否带屏'))
        screen[str(sc) if sc else '未知'] += 1
    country = Counter()
    for c in companies:
        ct = clean(c.get('国家'))
        if ct:
            country[str(ct)] += 1
    # 产品所属国别（按企业国别映射，口径与“企业数”区分开，避免 46/51 混算）
    co_country = {}
    for c in companies:
        nm = clean(c.get('企业名称'))
        if nm:
            co_country[str(nm).split('（')[0]] = clean(c.get('国家'))
    prod_country = Counter()
    for p in products:
        nm = clean(p.get('企业名称'))
        if nm:
            prod_country[str(co_country.get(str(nm).split('（')[0], '未知'))] += 1
    return {'prodrows': prodrows, 'stage_dist': dict(stage),
            'screen_dist': dict(screen), 'country_dist': dict(country),
            'product_country_dist': dict(prod_country)}


# ===== 应用场景聚合 =====
def mk_scenes():
    scenes = []
    idx = {}
    for row in scene_rows:
        name = clean(row.get('场景'))
        if not name:
            continue
        if clean(row.get('当前阶段')) and str(name) not in idx:
            shA = clean(row.get('出货占比(凤凰网)'))
            shB = clean(row.get('出货占比(IDC)'))
            mat = []
            for k in ('概念', 'Demo', 'Pilot', '小批量', '批量', '规模化'):
                mat.append(float(row.get(k) or 0))
            obj = {'scene': str(name),
                   'shareA': str(shA) if shA is not None else None,
                   'scopeA': clean(row.get('占比口径A')),
                   'shareB': str(shB) if shB is not None else None,
                   'scopeB': clean(row.get('占比口径B')),
                   'stage': clean(row.get('当前阶段')),
                   'maturity': mat, 'cases': []}
            scenes.append(obj)
            idx[str(name)] = obj
    for row in scene_rows:
        co = clean(row.get('企业名称'))
        if not co:
            continue
        sname = str(clean(row.get('场景')) or '')
        if sname not in idx:
            continue
        idx[sname]['cases'].append({
            'co': str(co), 'prod': clean(row.get('产品名称')) or '',
            'task': clean(row.get('典型任务')) or '',
            'stage': clean(row.get('部署阶段')) or '',
            'time': clean(row.get('案例时间')) or '',
            'desc': clean(row.get('案例描述')) or '',
            'src': clean(row.get('数据来源')),
            'tl': clean(row.get('时间轴标志')),
            'upd': dt(row.get('更新时间')),
        })
    for s in scenes:
        s['case_count'] = len(s['cases'])
    timeline = []
    for row in scene_rows:
        tl = clean(row.get('时间轴标志'))
        if not tl:
            continue
        timeline.append({'time': clean(row.get('案例时间')) or '',
                         'tl': str(tl), 'scene': clean(row.get('场景')) or ''})
    timeline.sort(key=lambda x: str(x['time']))
    return {'scenes': scenes, 'timeline': timeline}


data = {
    'meta': {'title': '人形机器人市场情报看板', 'badge': '芯视佳专用',
             'updated': UPDATED, 'src_file': '人形机器人看板数据源.xlsx',
             'sections': SECTIONS},
    'companies': companies,
    'products': products,
    'financing': financing,
    'market': market,
    'kpis': kpis,
    'market_series': series,
    'compete': mk_compete(),
    'scenes': mk_scenes(),
}

with open(OUT, 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=1)

print('OK', OUT, 'updated', UPDATED)
print('companies', len(companies), 'products', len(products), 'financing', len(financing), 'market', len(market))
print('scenes', len(data['scenes']['scenes']), 'cases', sum(s['case_count'] for s in data['scenes']['scenes']),
      'timeline', len(data['scenes']['timeline']))
print('kpis', [(k['t'], k['min'], k['max'], k['unit']) for k in kpis])
