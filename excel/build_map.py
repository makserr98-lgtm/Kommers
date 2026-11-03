# -*- coding: utf-8 -*-
"""КП = один файл Excel. Первый лист — карта: «облака» вопросов и действий с выпадающими
списками, связи нарисованы ячейками и красятся формулами (зелёная / красная / бледная).
Без макросов. Запуск: python3 build_map.py out.xlsx [example]"""
import sys, datetime
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from kp_data import STAGES, QUESTIONS, TASKS, SHOPS

OUT = sys.argv[1] if len(sys.argv) > 1 else 'kp-map.xlsx'
EXAMPLE = len(sys.argv) > 2 and sys.argv[2] == 'example'
FONT = 'Arial'
def F(**k):
    k.setdefault('name', FONT); k.setdefault('size', 8); return Font(**k)
def fill(c): return PatternFill('solid', start_color=c, end_color=c)
UNLOCK = Protection(locked=False)
STAGE = {i: s for i, s in enumerate(STAGES)}
STATUSES = ['Сделано', 'Не требуется', 'Хвост']

# ------------------------------------------------------------------ узлы и связи
class Node:
    def __init__(s, key, kind, title, stage, opts=None, multi=False, shops=False):
        s.key, s.kind, s.title, s.stage = key, kind, title, stage
        s.opts, s.multi, s.shops = opts or [], multi, shops
        s.when, s.any, s.after = [], [], []
        s.hint = ''

nodes = {}
for qq in QUESTIONS:
    qid, st, typ, text, hint, after, when, any_, opts = qq
    n = Node('q:' + qid, 'q', text, st, opts=opts, multi=(typ == 'несколько'))
    n.qid = qid; n.hint = hint
    n.when, n.any = when, any_
    n.after = [a for a in after.split(',') if a]
    nodes[n.key] = n
for t in TASKS:
    tid, grp, st, title, hint, after, when, any_ = t
    if grp == 'Труд':
        if 'a:Труд' in nodes: continue
        n = Node('a:Труд', 'a', 'Трудоёмкость и заявки цехов на материалы', st, shops=True)
        n.hint = hint
        n.when = ['Итог=да,частично', 'Цеха=' + ','.join(c for c, _ in SHOPS)]
        n.after = [a for a in after.split(',') if a]
        nodes[n.key] = n
        continue
    n = Node('a:' + tid, 'a', title, st)
    n.hint = hint; n.when, n.any = when, any_
    n.after = [a for a in after.split(',') if a]
    nodes[n.key] = n
QKEY = {n.qid: n.key for n in nodes.values() if n.kind == 'q'}

edges = []  # (src, dst, kind, info)
for n in nodes.values():
    for a in n.after:
        edges.append(('a:' + a, n.key, 'after', None))
    seen = set()
    for grp_name, cls in (('when', n.when), ('any', n.any)):
        for c in cls:
            qid = c.split('=')[0]
            src = QKEY[qid]
            if src in seen: continue
            seen.add(src)
            edges.append((src, n.key, 'clause', c))
for s, d, _, _ in edges:
    assert s in nodes and d in nodes, (s, d)

succ, pred = defaultdict(list), defaultdict(list)
for i, (s, d, k, _) in enumerate(edges):
    succ[s].append(i); pred[d].append(i)

# --- что рисовать: убираем дублирующие линии (логика считает по всем связям)
aft = defaultdict(set)
for s, d, k, _ in edges:
    if k == 'after': aft[d].add(s)
def reach(a, b, skip):  # есть ли путь a -> ... -> b по «нужно до», не через прямую связь
    st = [x for x in aft[b] if x != skip or True]
    seen = set(); st = [x for x in aft[b] if not (x == a)]
    while st:
        x = st.pop()
        if x == a: return True
        if x in seen: continue
        seen.add(x); st.extend(aft[x])
    return False
clause_of = defaultdict(dict)
for s, d, k, info in edges:
    if k == 'clause': clause_of[d][s] = info
draw = []
for i, (s, d, k, info) in enumerate(edges):
    if k == 'after' and s.startswith('a:') and d.startswith('a:') and reach(s, d, None):
        continue
    if k == 'clause' and any(clause_of[p].get(s) == info for p in aft[d]):
        continue
    draw.append(i)
DRAW = set(draw)
dsucc, dpred = defaultdict(list), defaultdict(list)
for i in draw:
    s, d = edges[i][0], edges[i][1]; dsucc[s].append(i); dpred[d].append(i)

# ------------------------------------------------------------------ слои
order = []
indeg = {k: len(pred[k]) for k in nodes}
stack = sorted([k for k in nodes if indeg[k] == 0], key=lambda k: (nodes[k].stage, k))
while stack:
    k = stack.pop(0); order.append(k)
    for i in succ[k]:
        d = edges[i][1]; indeg[d] -= 1
        if indeg[d] == 0: stack.append(d)
assert len(order) == len(nodes), 'цикл в дереве'
layer = {}
for k in order:
    layer[k] = max([layer[edges[i][0]] + 1 for i in pred[k]] or [0])
for k in reversed(order):  # вопросы — как можно правее (ближе к действиям, которые включают)
    if nodes[k].kind == 'q' and succ[k]:
        layer[k] = max(layer[k], min(layer[edges[i][1]] for i in succ[k]) - 1)
# балансировка: каждый узел — в слой, ближайший к медиане соседей по рисуемым связям
for _ in range(12):
    for k in order:
        lo = max([layer[edges[i][0]] + 1 for i in pred[k]] or [0])
        hi = min([layer[edges[i][1]] - 1 for i in succ[k]] or [max(layer.values())])
        nb = sorted([layer[edges[i][0]] + 1 for i in dpred[k]] + [layer[edges[i][1]] - 1 for i in dsucc[k]])
        if not nb or lo > hi: continue
        want = nb[(len(nb) - 1) // 2] if len(nb) % 2 else (nb[len(nb) // 2 - 1] + nb[len(nb) // 2]) / 2
        layer[k] = int(min(hi, max(lo, round(want))))
mn = min(layer.values())
for k in layer: layer[k] -= mn
NL = max(layer.values()) + 1

import math
SLOT = 4
SZ_NODES = {'a:ОГТ', 'a:ОГК', 'a:ОСТ', 'a:ОПП', 'a:ОВК', 'a:ОВК2', 'a:КартаРазр'}
for n in nodes.values():
    lines = math.ceil((len(n.title) + 8) / 27) + 2
    n.trows = math.ceil(lines * 14 / 6) + 1
    n.sz = n.key in SZ_NODES
items = {}
chains = {}
def node_h(n):
    nin = len(dpred[n.key]); nout = len(dsucc[n.key])
    body = n.trows + SLOT * (len(SHOPS) if n.shops else (len(n.opts) if n.multi else 1)) + (SLOT if n.sz else 0)
    return max(body, 3 * max(nin, nout) + 1)
for k, n in nodes.items():
    items[k] = dict(layer=layer[k], kind='node', h=node_h(n))
for i in draw:
    s, d = edges[i][0], edges[i][1]
    ch = [s]
    for L in range(layer[s] + 1, layer[d]):
        did = f'd{i}_{L}'; items[did] = dict(layer=L, kind='dummy', h=1); ch.append(did)
    ch.append(d); chains[i] = ch
links = []
for i, ch in chains.items():
    for a, b in zip(ch, ch[1:]): links.append((a, b, i))
left_of, right_of = defaultdict(list), defaultdict(list)
for a, b, i in links:
    right_of[a].append(b); left_of[b].append(a)

layers = [[] for _ in range(NL)]
for k, it in items.items(): layers[it['layer']].append(k)
for L in range(NL):
    layers[L].sort(key=lambda k: (nodes[k].stage if k in nodes else 9, 0 if k in nodes else 1, k))
pos = {}
def renum():
    for L in range(NL):
        for i, k in enumerate(layers[L]): pos[k] = i
renum()
def crossings():
    c = 0
    for L in range(NL - 1):
        ls = [(pos[a], pos[b]) for a, b, i in links if items[a]['layer'] == L]
        for x in range(len(ls)):
            for z in range(x + 1, len(ls)):
                if (ls[x][0] - ls[z][0]) * (ls[x][1] - ls[z][1]) < 0: c += 1
    return c
best_c, best_layers = crossings(), [l[:] for l in layers]
for it in range(24):
    rng_ = range(1, NL) if it % 2 == 0 else range(NL - 2, -1, -1)
    for L in rng_:
        nb = left_of if it % 2 == 0 else right_of
        def bary(k):
            ns = nb[k]
            return sum(pos[x] for x in ns) / len(ns) if ns else pos[k]
        layers[L].sort(key=lambda k: (bary(k), pos[k]))
        renum()
    c = crossings()
    if c < best_c: best_c, best_layers = c, [l[:] for l in layers]
layers = best_layers; renum()

# вертикальные позиции: изотоническая регрессия (PAV) к желаемым центрам
GAP_N, GAP_D = 5, 2
PACK = False
y = {}
def gap(a, b):
    return GAP_N if items[a]['kind'] == 'node' or items[b]['kind'] == 'node' else GAP_D
def place(L, desired):
    ks = layers[L]
    off = []; acc = 0
    for j, k in enumerate(ks):
        if j: acc += items[ks[j - 1]]['h'] + gap(ks[j - 1], k)
        off.append(acc)
    vals = [desired.get(k, y.get(k, off[j])) - off[j] for j, k in enumerate(ks)]
    blocks = []  # [sum, count]
    for v in vals:
        blocks.append([v, 1])
        while len(blocks) > 1 and blocks[-2][0] / blocks[-2][1] > blocks[-1][0] / blocks[-1][1]:
            s2, c2 = blocks.pop(); blocks[-1][0] += s2; blocks[-1][1] += c2
    q = []
    for s_, c_ in blocks: q += [s_ / c_] * c_
    for j, k in enumerate(ks): y[k] = int(round(q[j] + off[j]))
    # округление могло сблизить — поправим
    for j in range(1, len(ks)):
        mn = y[ks[j - 1]] + items[ks[j - 1]]['h'] + gap(ks[j - 1], ks[j])
        if y[ks[j]] < mn: y[ks[j]] = mn
for L in range(NL): place(L, {})
def center(k): return y[k] + items[k]['h'] / 2
for it in range(16):
    seq = range(NL) if it % 2 == 0 else range(NL - 1, -1, -1)
    for L in seq:
        des = {}
        for k in layers[L]:
            ns = (left_of[k] if it % 2 == 0 else right_of[k]) or (left_of[k] + right_of[k])
            if ns: des[k] = sum(center(x) for x in ns) / len(ns) - items[k]['h'] / 2
        place(L, des)
miny = min(y.values())
for k in y: y[k] -= miny

# порты: строки входа (слева) и выхода (справа)
port_in, port_out = {}, {}   # (item, link idx) -> row
def assign_ports(k, lst, getter, store):
    it = items[k]
    if it['kind'] == 'dummy':
        for li in lst: store[(k, li)] = y[k]
        return
    cands = list(range(y[k] + 1, y[k] + it['h'] - 1, 3)) or [y[k]]
    lst = sorted(lst, key=getter)
    n = len(lst); prev = -1
    for j, li in enumerate(lst):
        want = getter(li)
        lo = prev + 1; hi = len(cands) - (n - j)
        best = min(range(lo, hi + 1), key=lambda c: abs(cands[c] - want))
        store[(k, li)] = cands[best]; prev = best
def gkey(j):
    a, b, i = links[j]
    if a in nodes:
        return (a, 'after') if edges[i][2] == 'after' else (a, 'clause', edges[i][3])
    return ('link', j)
groups_of = defaultdict(list)
for j in range(len(links)): groups_of[gkey(j)].append(j)
for k in items:
    ins = [j for j, (a, b, i) in enumerate(links) if b == k]
    assign_ports(k, ins, lambda j: center(links[j][0]), port_in)
    gs = defaultdict(list)
    for j, (a, b, i) in enumerate(links):
        if a == k: gs[gkey(j)].append(j)
    reps = list(gs.keys())
    tmp = {}
    assign_ports(k, reps, lambda g: sum(center(links[j][1]) for j in gs[g]) / len(gs[g]), tmp)
    for g in reps:
        for j in gs[g]: port_out[(k, j)] = tmp[(k, g)]

# ------------------------------------------------------------------ колонки
HEAD_ROWS = 9            # строки шапки над картой
R0 = HEAD_ROWS + 2       # первая строка карты
W_LABEL, W_INPUT, W_TRACK, W_MARG = 22, 14, 0.7, 1.4
col = 2
layer_col = {}
chan = {}
for L_ in range(NL):
    layer_col[L_] = col; col += 2
    if L_ < NL - 1:
        units = [g for g, js in groups_of.items() if items[links[js[0]][0]]['layer'] == L_]
        def y1(g): j = groups_of[g][0]; return port_out[(links[j][0], j)]
        def y2s(g): return [port_in[(links[j][1], j)] for j in groups_of[g]]
        need = [g for g in units if any(v != y1(g) for v in y2s(g))]
        def span(g): v = [y1(g)] + y2s(g); return min(v), max(v)
        def cost(seq):
            c = 0; idx = {g: t for t, g in enumerate(seq)}
            for a in seq:
                lo, hi = span(a); a1 = y1(a)
                for b in seq:
                    if a == b: continue
                    b1 = y1(b)
                    if lo < b1 < hi and idx[b] > idx[a]: c += 1
                    for b2 in y2s(b):
                        if lo < b2 < hi and idx[b] < idx[a]: c += 1
                        if b2 == a1 and idx[b] < idx[a]: c += 1000
            return c
        order_t = sorted(need, key=lambda g: (sum(span(g)) / 2))
        best = cost(order_t); improved = True
        while improved:
            improved = False
            for t in range(len(order_t) - 1):
                cand = order_t[:]; cand[t], cand[t + 1] = cand[t + 1], cand[t]
                cc = cost(cand)
                if cc < best: best, order_t, improved = cc, cand, True
        # упаковка: несколько связей на одной дорожке, если их вертикали не пересекаются
        tl = []  # список дорожек: [ [группы] ]
        where = {}
        for g in order_t:
            lo_g, hi_g = span(g)
            lb, ub = 0, None
            for bgrp, tobj in where.items():
                pos_b = next(i_ for i_, tt in enumerate(tl) if tt is tobj)
                if y1(bgrp) in y2s(g): lb = max(lb, pos_b + 1)
                if y1(g) in y2s(bgrp): ub = pos_b if ub is None else min(ub, pos_b)
            placed = False
            hi_lim = len(tl) - 1 if ub is None else ub - 1
            for i_ in range(lb, hi_lim + 1) if PACK else []:
                if all(hi_g + 2 < span(o)[0] or span(o)[1] + 2 < lo_g for o in tl[i_]):
                    tl[i_].append(g); where[g] = tl[i_]; placed = True; break
            if not placed:
                ins = len(tl) if ub is None else min(max(lb, 0), ub)
                if ub is None: ins = max(ins, lb)
                tl.insert(ins, [g]); where[g] = tl[ins]
        first = col; col += 1
        tracks = {}
        for tt in tl:
            for g in tt: tracks[g] = col
            col += 2   # между дорожками пустой столбец
        last = col; col += 1
        chan[L_] = (first, tracks, last, units)
NCOL = col

# ------------------------------------------------------------------ книга
wb = Workbook()
mp = wb.active; mp.title = 'Карта'
jr = wb.create_sheet('Журнал')
calc = wb.create_sheet('Расчёт')
lists = wb.create_sheet('Списки')
helpws = wb.create_sheet('Как пользоваться')

def mref(r, c, sheet='Карта'): return f"'{sheet}'!${CL(c)}${r}"

# --- раскладка ячеек узлов на карте
node_cells = {}   # key -> dict(title=(r,c), input=(r,c) | slots=[(r,c,code,label)])
for k, n in nodes.items():
    L = layer[k]; c0 = layer_col[L]; r0 = R0 + y[k]; h = items[k]['h']
    d = dict(top=r0, bottom=r0 + h - 1, c0=c0, c1=c0 + 1, title=(r0, c0))
    rr = r0 + n.trows
    if n.sz: d['sz'] = (r0 + h - SLOT, c0)
    if n.kind == 'q' and n.multi:
        d['slots'] = [(rr + SLOT * i, c0, o[0], o[1]) for i, o in enumerate(n.opts)]
    elif n.shops:
        d['slots'] = [(rr + SLOT * i, c0, c, l) for i, (c, l) in enumerate(SHOPS)]
    else:
        d['input'] = (rr, c0)
    node_cells[k] = d

# --- лист Списки: варианты одиночных вопросов
lists['A1'] = 'Списки для выпадающих меню'
lc = 1
list_ref = {}
for k, n in nodes.items():
    if n.kind == 'q' and not n.multi:
        lc += 1
        lists.cell(row=1, column=lc, value=n.qid)
        for i, o in enumerate(n.opts): lists.cell(row=2 + i, column=lc, value=o[1])
        list_ref[k] = f"='Списки'!${CL(lc)}$2:${CL(lc)}${1 + len(n.opts)}"
lists.sheet_state = 'hidden'

# --- Журнал: строки действий (ввод деталей)
J0 = 5
jrow = {}
jr_rows = []
for k in order:
    n = nodes[k]
    if n.kind != 'a': continue
    if n.shops:
        for c, l in SHOPS: jr_rows.append((k, c, l))
    else:
        jr_rows.append((k, None, None))
jr_rows.sort(key=lambda x: (nodes[x[0]].stage, order.index(x[0])))
for i, (k, c, l) in enumerate(jr_rows): jrow[(k, c)] = J0 + i

# --- Расчёт
calc['A1'] = 'Служебные расчёты карты. Не редактировать.'
CR = {}   # (key, field) -> cell ref string
row = 3
calc.append([])
hdr = ['узел', 'ответ дан', 'применимость', 'закрыт', 'не закрыто до', 'первое незакрытое', 'символ', 'состояние', 'надпись']
for j, h in enumerate(hdr): calc.cell(row=2, column=1 + j, value=h)
node_row = {}
for k in order: node_row[k] = row; row += 1
clause_row = {}
for i, (s, d, kd, info) in enumerate(edges):
    clause_row[i] = row; row += 1
shop_row = {}
for c, l in SHOPS: shop_row[c] = row; row += 1
def C(k, f):
    col_ = {'ans': 'B', 'app': 'C', 'sat': 'D', 'unmet': 'E', 'first': 'F', 'sym': 'G', 'state': 'H', 'label': 'I'}[f]
    return f"'Расчёт'!${col_}${node_row[k]}"
def LC(k, f):  # локально на листе Расчёт
    col_ = {'ans': 'B', 'app': 'C', 'sat': 'D', 'unmet': 'E', 'first': 'F', 'sym': 'G', 'state': 'H', 'label': 'I'}[f]
    return f'${col_}${node_row[k]}'

def q_input(qk):
    d = node_cells[qk]
    return mref(*d['input'])
def q_slot(qk, code):
    for (r, c, cd, lb) in node_cells[qk]['slots']:
        if cd == code: return mref(r, c + 1)
    raise KeyError((qk, code))
def match_expr(clause):
    qid, codes = clause.split('=')
    qk = QKEY[qid]; n = nodes[qk]; codes = codes.split(',')
    if n.multi:
        return 'OR(' + ','.join(f'{q_slot(qk, c)}="да"' for c in codes) + ')'
    lab = {o[0]: o[1] for o in n.opts}
    return 'OR(' + ','.join(f'{q_input(qk)}="{lab[c]}"' for c in codes) + ')'
def clause_value(clause):
    qk = QKEY[clause.split('=')[0]]
    return f'IF({match_expr(clause)},"да",IF(OR({LC(qk, "ans")}=1,{LC(qk, "app")}="нет"),"нет","?"))'
def comb(w, y):
    if not w and not y: return '"да"'
    W = w or ['""']; Y = y
    no_w = 'OR(' + ','.join(f'{v}="нет"' for v in W) + ')'
    und_w = 'OR(' + ','.join(f'{v}="?"' for v in W) + ')'
    if Y:
        any_yes = 'OR(' + ','.join(f'{v}="да"' for v in Y) + ')'
        any_und = 'OR(' + ','.join(f'{v}="?"' for v in Y) + ')'
        return (f'IF({no_w},"нет",IF(AND(NOT({any_yes}),NOT({any_und})),"нет",'
                f'IF(OR({und_w},AND(NOT({any_yes}),{any_und})),"?","да")))')
    return f'IF({no_w},"нет",IF({und_w},"?","да"))'

# значения условий для каждого узла (отдельные ячейки J.. в строке узла)
clause_cells = {}
for k in order:
    n = nodes[k]; r = node_row[k]
    calc[f'A{r}'] = k
    w, yv = [], []
    for j, cl in enumerate(n.when + n.any):
        cc = CL(10 + j)
        calc[f'{cc}{r}'] = '=' + clause_value(cl)
        (w if j < len(n.when) else yv).append(f'${cc}${r}')
        clause_cells[(k, cl.split('=')[0])] = f'${cc}${r}'
    calc[f'C{r}'] = '=' + comb(w, yv)
    # ответ / статус
    if n.kind == 'q':
        if n.multi:
            calc[f'B{r}'] = '=IF(OR(' + ','.join(f'{q_slot(k, o[0])}<>""' for o in n.opts) + '),1,0)'
        else:
            calc[f'B{r}'] = f'=IF({q_input(k)}="",0,1)'
    # предшественники
    preds = [edges[i][0] for i in pred[k] if edges[i][2] == 'after']
    if preds:
        calc[f'E{r}'] = '=' + '+'.join(f'(1-{LC(p, "sat")})' for p in preds)
        f_ = '""'
        for p in reversed(preds):
            f_ = f'IF({LC(p, "sat")}=0,"{nodes[p].title}",{f_})'
        calc[f'F{r}'] = '=' + f_
    else:
        calc[f'E{r}'] = 0; calc[f'F{r}'] = ''
    # «зависит от»
    und = '""'
    for j, cl in reversed(list(enumerate(n.when + n.any))):
        qk = QKEY[cl.split('=')[0]]
        und = f'IF(${CL(10 + j)}${r}="?","{nodes[qk].title}",{und})'
    hd = f'«{STAGE[n.stage][0]} · {n.title}»'
    if n.kind == 'q':
        calc[f'D{r}'] = f'=IF({LC(k, "ans")}=1,1,IF({LC(k, "app")}="нет",1,0))'
        calc[f'G{r}'] = f'=IF({LC(k, "ans")}=1,"✓",IF({LC(k, "app")}="нет","×",IF({LC(k, "app")}="?","?",IF({LC(k, "unmet")}>0,"⏳","●"))))'
        ans_txt = '"ответ дан"' if n.multi else f'"ответ: "&{q_input(k)}'
        calc[f'H{r}'] = (f'=IF(G{r}="✓",{ans_txt},IF(G{r}="×","не нужен",IF(G{r}="?","зависит: "&{und},'
                         f'IF(G{r}="⏳","после: "&F{r},"ответить ▼"))))')
    elif n.shops:
        chosen = [f"{q_slot(QKEY['Цеха'], c)}=\"да\"" for c, _ in SHOPS]
        stc = [mref(sl[0], sl[1] + 1) for sl in node_cells[k]['slots']]
        for (c, l), ch, s_ in zip(SHOPS, chosen, stc):
            sr = shop_row[c]
            calc[f'A{sr}'] = 'цех ' + c
            calc[f'B{sr}'] = f'=IF({ch},1,0)'
            calc[f'C{sr}'] = f'=IF(AND(B{sr}=1,OR({s_}="Сделано",{s_}="Не требуется",{s_}="Хвост")),1,0)'
            calc[f'D{sr}'] = f'=IF(AND(B{sr}=1,{s_}="Хвост"),1,0)'
        rs = f'{shop_row[SHOPS[0][0]]}'; re_ = f'{shop_row[SHOPS[-1][0]]}'
        nch, ncl, ndb = f'SUM($B${rs}:$B${re_})', f'SUM($C${rs}:$C${re_})', f'SUM($D${rs}:$D${re_})'
        calc[f'B{r}'] = f'={ncl}'
        calc[f'D{r}'] = f'=IF({LC(k, "app")}="нет",1,IF(AND({nch}>0,{ncl}={nch}),1,0))'
        calc[f'G{r}'] = (f'=IF({LC(k, "app")}="нет","×",IF({LC(k, "app")}="?","?",IF(AND({nch}>0,{ncl}={nch}),IF({ndb}>0,"⚑","✓"),'
                         f'IF({LC(k, "unmet")}>0,"⏳","▶"))))')
        calc[f'H{r}'] = (f'=IF(G{r}="×","не нужно",IF(G{r}="?","зависит: "&{und},IF(G{r}="⏳","ждёт: "&F{r},'
                         f'"цехов закрыто "&{ncl}&" из "&{nch})))')
    else:
        stc = mref(*node_cells[k]['input'])
        jr_r = jrow[(k, None)]
        calc[f'D{r}'] = f'=IF(OR({stc}="Сделано",{stc}="Не требуется",{stc}="Хвост",{LC(k, "app")}="нет"),1,0)'
        calc[f'G{r}'] = (f'=IF({stc}="Сделано","✓",IF({stc}="Не требуется","—",IF({stc}="Хвост","⚑",IF({LC(k, "app")}="нет","×",'
                         f'IF({LC(k, "app")}="?","?",IF({LC(k, "unmet")}>0,"⏳","▶"))))))')
        due = f"'Журнал'!$H${jr_r}"
        szt = '""'
        if n.sz:
            szc = mref(node_cells[k]['sz'][0], node_cells[k]['sz'][1] + 1)
            szt = f'IF({szc}="",", впишите № СЗ",", СЗ № "&{szc})'
        calc[f'H{r}'] = (f'=IF(G{r}="✓","сделано"&{szt}&IF(E{r}>0," (до него не закрыто: "&E{r}&")",""),IF(G{r}="—","не требуется",'
                         f'IF(G{r}="⚑","хвост"&IF(ISNUMBER({due})," до "&DAY({due})&"."&RIGHT("0"&MONTH({due}),2),""),'
                         f'IF(G{r}="×","не нужно",IF(G{r}="?","зависит: "&{und},IF(G{r}="⏳","ждёт: "&F{r},"можно делать ▼"))))))')
    calc[f'I{r}'] = f'=G{r}&" {STAGE[n.stage][0]} · {n.title}"&CHAR(10)&H{r}'

# коды связей
edge_code = {}
for i, (s, d, kd, info) in enumerate(edges):
    r = clause_row[i]
    calc[f'A{r}'] = f'{s} → {d}'
    if kd == 'clause':
        v = clause_cells[(d, info.split('=')[0])]
        calc[f'B{r}'] = f'=IF({v}="да","G",IF({v}="нет","R","P"))'
    else:
        sym = LC(s, 'sym')
        calc[f'B{r}'] = f'=IF(OR({sym}="✓",{sym}="—"),"G",IF({sym}="⚑","O",IF({sym}="×","X","P")))'
    edge_code[i] = f"'Расчёт'!$B${r}"
SUMR = row + 1
acts = [k for k in order if nodes[k].kind == 'a']
calc[f'A{SUMR}'] = 'действий нужно'; calc[f'B{SUMR}'] = '=' + '+'.join(f'IF({LC(k, "sym")}<>"×",1,0)' for k in acts)
calc[f'A{SUMR+1}'] = 'закрыто'; calc[f'B{SUMR+1}'] = '=' + '+'.join(f'IF(OR({LC(k, "sym")}="✓",{LC(k, "sym")}="—"),1,0)' for k in acts)
calc[f'A{SUMR+2}'] = 'хвостов'; calc[f'B{SUMR+2}'] = '=' + '+'.join(f'IF({LC(k, "sym")}="⚑",1,0)' for k in acts)
calc[f'A{SUMR+3}'] = 'можно делать'
f_ = '"—"'
for k in reversed(acts): f_ = f'IF({LC(k, "sym")}="▶","{nodes[k].title}",{f_})'
calc[f'B{SUMR+3}'] = '=' + f_
calc[f'A{SUMR+4}'] = 'ответить'
f_ = '"—"'
for k in reversed([k for k in order if nodes[k].kind == 'q']): f_ = f'IF({LC(k, "sym")}="●","{nodes[k].title}",{f_})'
calc[f'B{SUMR+4}'] = '=' + f_
calc[f'A{SUMR+5}'] = 'вопросов к ответу'; calc[f'B{SUMR+5}'] = '=' + '+'.join(f'IF({LC(k, "sym")}="●",1,0)' for k in order if nodes[k].kind == 'q')
calc.sheet_state = 'hidden'

# ------------------------------------------------------------------ рисуем карту
mp.sheet_view.showGridLines = False
mp.sheet_view.zoomScale = 85
mp.column_dimensions['A'].width = 2
for L in range(NL):
    c = layer_col[L]
    mp.column_dimensions[CL(c)].width = W_LABEL; mp.column_dimensions[CL(c + 1)].width = W_INPUT
    if L in chan:
        first, tracks, last, segs = chan[L]
        mp.column_dimensions[CL(first)].width = W_MARG; mp.column_dimensions[CL(last)].width = W_MARG
        for cc in range(first + 1, last): mp.column_dimensions[CL(cc)].width = W_TRACK
MAXR = R0 + max(y[k] + items[k]['h'] for k in items) + 2
for r in range(R0, MAXR + 1): mp.row_dimensions[r].height = 6

line = {}   # (r,c) -> code ref ; вертикали перекрывают горизонтали
def hline(r, c1, c2, code):
    for c in range(min(c1, c2), max(c1, c2) + 1): line.setdefault((r, c), code)
vert = {}
for j, (a, b, i) in enumerate(links):
    L = items[a]['layer']; first, tracks, last, segs = chan[L]
    y1 = R0 + port_out[(a, j)] - 0; y2 = R0 + port_in[(b, j)]
    code = edge_code[i]
    if a not in nodes:  # проходная точка: линия через колонки узла
        hline(y1, layer_col[L], layer_col[L] + 1, code)
    if y1 == y2:
        hline(y1, first, last, code)
    else:
        t = tracks[gkey(j)]
        hline(y1, first, t, code); hline(y2, t, last, code)
        for r in range(min(y1, y2), max(y1, y2) + 1): vert[(r, t)] = code
line.update(vert)
for (r, c), code in line.items():
    cell = mp.cell(row=r, column=c, value='=' + code)
    cell.number_format = ';;;'
    cell.font = F(size=1)
# стрелки: последняя ячейка канала перед узлом
EDGE_COLORS = {'G': '2E9E4F', 'R': 'D64545', 'O': 'F0A030', 'P': 'D5DAE0', 'X': 'EEF0F3'}
map_rng = f'B{R0}:{CL(NCOL)}{MAXR}'
for code, colr in EDGE_COLORS.items():
    mp.conditional_formatting.add(map_rng, CellIsRule(operator='equal', formula=[f'"{code}"'], fill=fill(colr), font=Font(color=colr)))

thin = Side(style='thin', color='8A96A8'); med = Side(style='medium', color='1F3A5F')
NODE_FILL = {'✓': ('DFF2E1', '1E6B34'), '—': ('EEF1F4', '6B7480'), '⚑': ('FFE2BF', '8A4200'), '▶': ('CFE3FF', '0B4FA8'),
             '●': ('CFE3FF', '0B4FA8'), '⏳': ('F7F8FA', '3A4350'), '?': ('FFF5BF', '6E5A00'), '×': ('F7F8FA', 'B8BEC6')}
dv_status = DataValidation(type='list', formula1='"' + ','.join(STATUSES) + '"', allow_blank=True,
                           promptTitle='Статус действия', prompt='Сделано / Не требуется / Хвост. Пусто — открыто.', showInputMessage=True)
dv_yesno = DataValidation(type='list', formula1='"да,нет"', allow_blank=True, promptTitle='Вариант', prompt='да — выбран, нет или пусто — не выбран', showInputMessage=True)
mp.add_data_validation(dv_status); mp.add_data_validation(dv_yesno)
title_cells = []
input_style = dict(font=F(size=8, bold=True), fill=fill('FFFFFF'))
ex = {}
if EXAMPLE:
    ex = {'q:ГОЗ': 'Да, ГОЗ', 'q:Работы': ['ремонт', 'кабель', 'щиты'], 'q:Док': 'Частично', 'q:Демонтаж': 'Да, демонтаж и доставка',
          'q:Вес': 'Да', 'q:Условия': ['годкабеля'],
          'q:Закупка': 'Да, сразу понятно', 'a:Рег': 'Сделано', 'a:ЗапросДок': 'Сделано',
          'a:Груз15': 'Хвост', 'a:ГодКабеля': 'Сделано', 'a:ОГТ': 'Сделано'}
for k, n in nodes.items():
    d = node_cells[k]; r0, r1, c0, c1 = d['top'], d['bottom'], d['c0'], d['c1']
    for r in range(r0, r1 + 1):
        for c in (c0, c1):
            cell = mp.cell(row=r, column=c)
            cell.border = Border(left=med if c == c0 else None, right=med if c == c1 else None,
                                 top=med if r == r0 else None, bottom=med if r == r1 else None)
    tr, tc = d['title']
    t = mp.cell(row=tr, column=tc, value=f"='Расчёт'!$I${node_row[k]}")
    t.font = F(size=10, bold=True); t.alignment = Alignment(wrap_text=True, vertical='top')
    mp.merge_cells(start_row=tr, start_column=tc, end_row=tr + n.trows - 1, end_column=tc + 1)
    if n.sz:
        r, c = d['sz']
        a = mp.cell(row=r, column=c, value='№ служебной записки:')
        mp.merge_cells(start_row=r, start_column=c, end_row=r + SLOT - 1, end_column=c)
        a.font = F(size=9, color='1F3A5F'); a.alignment = Alignment(horizontal='right', vertical='center')
        b = mp.cell(row=r, column=c + 1)
        mp.merge_cells(start_row=r, start_column=c + 1, end_row=r + SLOT - 1, end_column=c + 1)
        b.font = F(size=10, bold=True, color='1F3A5F'); b.fill = fill('FFF6D5'); b.protection = UNLOCK; b.number_format = '@'
        b.alignment = Alignment(horizontal='center', vertical='center', shrink_to_fit=True)
        b.border = Border(left=thin, right=med, top=thin, bottom=med)
        if EXAMPLE and k == 'a:ОГТ': b.value = '112-45'
    title_cells.append(f'{CL(tc)}{tr}')
    if n.hint:
        from openpyxl.comments import Comment
        t.comment = Comment(n.hint, 'Хаб КП'); t.comment.width = 260; t.comment.height = 90
    if 'input' in d:
        r, c = d['input']
        lab = mp.cell(row=r, column=c)
        mp.merge_cells(start_row=r, start_column=c, end_row=r + SLOT - 1, end_column=c + 1)
        lab.alignment = Alignment(horizontal='center', vertical='center', shrink_to_fit=True)
        lab.font = F(size=10, bold=True, color='1F3A5F'); lab.fill = fill('FFF6D5'); lab.protection = UNLOCK
        lab.border = Border(left=med, right=med, top=thin, bottom=med)
        if n.kind == 'q':
            dv = DataValidation(type='list', formula1=list_ref[k], allow_blank=True, promptTitle='Ответ', prompt=n.title[:250], showInputMessage=True)
            mp.add_data_validation(dv); dv.add(lab.coordinate)
        else:
            dv_status.add(lab.coordinate)
        if k in ex: lab.value = ex[k]
    else:
        for (r, c, code, label) in d['slots']:
            a = mp.cell(row=r, column=c, value=label)
            mp.merge_cells(start_row=r, start_column=c, end_row=r + SLOT - 1, end_column=c)
            a.font = F(size=9); a.alignment = Alignment(wrap_text=True, vertical='center', shrink_to_fit=False)
            b = mp.cell(row=r, column=c + 1)
            mp.merge_cells(start_row=r, start_column=c + 1, end_row=r + SLOT - 1, end_column=c + 1)
            b.font = F(size=10, bold=True, color='1F3A5F'); b.fill = fill('FFF6D5'); b.protection = UNLOCK
            b.alignment = Alignment(horizontal='center', vertical='center', shrink_to_fit=True)
            b.border = Border(left=thin, right=med, top=thin, bottom=thin)
            if n.shops:
                dv_status.add(b.coordinate)
                ch = q_slot(QKEY['Цеха'], code)
                mp.conditional_formatting.add(f'{a.coordinate}:{b.coordinate}', FormulaRule(formula=[f'{ch.split("!")[1]}<>"да"'], font=Font(color='C3C8CF')))
            else:
                dv_yesno.add(b.coordinate)
                if k in ex and code in ex[k]: b.value = 'да'
for sym, (bg, fg) in NODE_FILL.items():
    mp.conditional_formatting.add(' '.join(title_cells), FormulaRule(formula=[f'LEFT({title_cells[0]},1)="{sym}"'], fill=fill(bg), font=Font(color=fg, bold=True)))
# Правило с относительной ссылкой должно применяться к каждой ячейке: задаём по одной
mp.conditional_formatting = type(mp.conditional_formatting)()
for code, colr in EDGE_COLORS.items():
    mp.conditional_formatting.add(map_rng, CellIsRule(operator='equal', formula=[f'"{code}"'], fill=fill(colr), font=Font(color=colr)))
for tcell in title_cells:
    for sym, (bg, fg) in NODE_FILL.items():
        mp.conditional_formatting.add(tcell, FormulaRule(formula=[f'LEFT({tcell},1)="{sym}"'], fill=fill(bg), font=Font(color=fg, bold=True)))
for k, n in nodes.items():
    if n.shops:
        for (r, c, code, label) in node_cells[k]['slots']:
            ch = q_slot(QKEY['Цеха'], code).split('!')[1]
            mp.conditional_formatting.add(f'{CL(c)}{r}', FormulaRule(formula=[f'{ch}<>"да"'], font=Font(color='C3C8CF')))

# заголовки слоёв
for L in range(NL):
    c = layer_col[L]
    h = mp.cell(row=R0 - 1, column=c, value=f'Шаг {L + 1}')
    h.font = F(size=8, bold=True, color='5B6573'); mp.merge_cells(start_row=R0 - 1, start_column=c, end_row=R0 - 1, end_column=c + 1)
    h.alignment = Alignment(horizontal='center')

# ------------------------------------------------------------------ шапка карты
mp['B1'] = 'Карта КП'; mp['B1'].font = F(size=16, bold=True, color='1F3A5F')
for r in range(1, HEAD_ROWS + 1): mp.row_dimensions[r].height = 15
mp.row_dimensions[R0 - 1].height = 13
fields = [('Номер КП', 'КП № 0347/26'), ('Заказчик', 'Заказчик А'), ('Предмет', 'Ремонт электродвигателя насоса и прокладка кабеля'),
          ('Поступила', datetime.date(2026, 10, 31)), ('Срок ответа', datetime.date(2026, 11, 9)), ('Исполнитель', '')]
HC = layer_col[0]
for i, (lab, v) in enumerate(fields):
    r = 2 + i
    a = mp.cell(row=r, column=HC, value=lab + ':'); a.font = F(size=9, bold=True); a.alignment = Alignment(horizontal='right')
    b = mp.cell(row=r, column=HC + 1, value=(v if EXAMPLE else None))
    b.font = F(size=9); b.fill = fill('FFF6D5'); b.protection = UNLOCK
    end = layer_col[min(2, NL - 1)] + 1
    mp.merge_cells(start_row=r, start_column=HC + 1, end_row=r, end_column=end)
    if lab in ('Поступила', 'Срок ответа'): b.number_format = 'DD.MM.YYYY'
FLD = {lab: f"'Карта'!${CL(HC + 1)}${2 + i}" for i, (lab, _) in enumerate(fields)}
SC = layer_col[min(3, NL - 1)]
info = [('Готовность', f"=IF('Расчёт'!$B${SUMR}=0,0,'Расчёт'!$B${SUMR+1}/'Расчёт'!$B${SUMR})", '0%'),
        ('До срока', f'=IF({FLD["Срок ответа"]}="","",IF({FLD["Срок ответа"]}-TODAY()<0,"просрочено на "&(TODAY()-{FLD["Срок ответа"]})&" дн.",({FLD["Срок ответа"]}-TODAY())&" дн."))', None),
        ('Можно делать', f"='Расчёт'!$B${SUMR+3}", None), ('Ответить', f"='Расчёт'!$B${SUMR+4}&IF('Расчёт'!$B${SUMR+5}>1,\"  (всего вопросов: \"&'Расчёт'!$B${SUMR+5}&\")\",\"\")", None),
        ('Хвостов', f"='Расчёт'!$B${SUMR+2}", None)]
for i, (lab, f_, nf) in enumerate(info):
    r = 2 + i
    a = mp.cell(row=r, column=SC, value=lab + ':'); a.font = F(size=9, bold=True, color='5B6573'); a.alignment = Alignment(horizontal='right')
    b = mp.cell(row=r, column=SC + 1, value=f_); b.font = F(size=9, bold=True)
    if nf: b.number_format = nf
    b.alignment = Alignment(horizontal='left')
# легенда
LG = layer_col[min(5, NL - 1)]
leg = [('G', 'зелёная — условие выполнено / действие закрыто'), ('R', 'красная — ответ закрыл эту дорожку'),
       ('O', 'оранжевая — закрыто хвостом'), ('P', 'бледная — ещё не открыта')]
mp.cell(row=2, column=LG, value='Связи:').font = F(size=9, bold=True)
for i, (code, txt) in enumerate(leg):
    r = 3 + i
    sw = mp.cell(row=r, column=LG); sw.fill = fill(EDGE_COLORS[code])
    t = mp.cell(row=r, column=LG + 1, value=txt); t.font = F(size=8)
mp.cell(row=2, column=LG + 3, value='Облака:').font = F(size=9, bold=True)
leg2 = [('▶', 'можно делать'), ('●', 'вопрос: ответить'), ('✓', 'сделано / ответ дан'), ('⚑', 'хвост'), ('?', 'зависит от ответа'), ('⏳', 'ждёт предыдущее'), ('×', 'не нужно')]
for i, (sym, txt) in enumerate(leg2):
    r = 3 + i
    c_ = mp.cell(row=r, column=LG + 3, value=f'{sym} {txt}')
    c_.font = F(size=8, color=NODE_FILL[sym][1], bold=True); c_.fill = fill(NODE_FILL[sym][0])
mp.cell(row=HEAD_ROWS, column=HC, value='Нажмите на жёлтую ячейку под облаком — откроется меню ответов или статусов. Подсказка — в примечании к облаку (наведите мышь).').font = F(size=8, italic=True, color='5B6573')
mp.freeze_panes = mp.cell(row=R0, column=1)
mp.protection.sheet = True
mp.protection.formatColumns = False; mp.protection.formatRows = False; mp.protection.formatCells = False
mp.protection.selectLockedCells = False; mp.protection.selectUnlockedCells = False

# ------------------------------------------------------------------ Журнал
jr['A1'] = 'Журнал действий'; jr['A1'].font = F(size=14, bold=True, color='1F3A5F')
jr['A2'] = 'Статус ставится на карте. Здесь — кто, когда, срок хвоста и комментарии.'; jr['A2'].font = F(size=8, italic=True, color='5B6573')
hdr = ['Этап', 'Действие', 'Состояние', 'Статус (с карты)', 'Кто / от кого', 'Дата', 'Срок хвоста', 'Нужно к', 'Комментарий / причина', '№ служебной записки (с карты)']
for j, h in enumerate(hdr):
    c = jr.cell(row=J0 - 1, column=1 + j, value=h); c.font = F(size=9, bold=True, color='FFFFFF'); c.fill = fill('1F3A5F')
    c.alignment = Alignment(horizontal='center', wrap_text=True)
for col_, w in zip('ABCDEFGHIJ', [7, 46, 34, 14, 18, 11, 11, 18, 44, 18]): jr.column_dimensions[col_].width = w
for (k, c), r in sorted(jrow.items(), key=lambda x: x[1]):
    n = nodes[k]
    jr[f'A{r}'] = STAGE[n.stage][0]
    jr[f'B{r}'] = n.title + (': ' + dict(SHOPS)[c] if c else '')
    jr[f'C{r}'] = f"='Расчёт'!$G${node_row[k]}&\" \"&'Расчёт'!$H${node_row[k]}"
    if c:
        sl = [s_ for s_ in node_cells[k]['slots'] if s_[2] == c][0]
        jr[f'D{r}'] = f'={mref(sl[0], sl[1] + 1)}&""'
    else:
        jr[f'D{r}'] = f'={mref(*node_cells[k]["input"])}&""'
    if 'sz' in node_cells[k]:
        jr[f'J{r}'] = f'={mref(node_cells[k]["sz"][0], node_cells[k]["sz"][1] + 1)}&""'
    for col_ in 'EFGHI':
        cell = jr[f'{col_}{r}']; cell.fill = fill('FFF6D5'); cell.protection = UNLOCK
    jr[f'F{r}'].number_format = jr[f'G{r}'].number_format = 'DD.MM.YYYY'
    for col_ in 'ABCDEFGHIJ':
        jr[f'{col_}{r}'].font = F(size=9); jr[f'{col_}{r}'].alignment = Alignment(wrap_text=True, vertical='top')
    if EXAMPLE and k == 'a:Груз15':
        jr[f'G{r}'] = datetime.date(2026, 11, 5); jr[f'H{r}'] = 'до демонтажа'
        jr[f'I{r}'] = 'Двигатель около 60 кг. Уточнить у заказчика, кто демонтирует и грузит.'
JL = J0 + len(jr_rows) - 1
dvd = DataValidation(type='date', operator='greaterThan', formula1='36526', allow_blank=True); jr.add_data_validation(dvd); dvd.add(f'F{J0}:G{JL}')
for sym, (bg, fg) in NODE_FILL.items():
    jr.conditional_formatting.add(f'A{J0}:I{JL}', FormulaRule(formula=[f'LEFT($C{J0},1)="{sym}"'], fill=fill(bg), font=Font(color=fg)))
jr.conditional_formatting.add(f'A{J0}:I{JL}', FormulaRule(formula=[f'AND(LEFT($C{J0},1)="⚑",ISNUMBER($G{J0}),$G{J0}<TODAY())'], fill=fill('FFD6D6'), font=Font(color='A40000', bold=True)))
jr.freeze_panes = f'C{J0}'
jr.protection.sheet = True; jr.protection.formatColumns = False; jr.protection.formatRows = False

# ------------------------------------------------------------------ Как пользоваться
helpws.column_dimensions['A'].width = 120
HELP = [
 'Один файл = одно КП. Новое КП: откройте чистый «КП шаблон.xlsx» и сразу «Файл → Сохранить как» с номером КП в имени.',
 'Лист «Карта»: облака — вопросы и действия, линии — связи между ними. Нажмите на жёлтую ячейку под облаком: у вопроса откроется список вариантов ответа, у действия — статус (Сделано / Не требуется / Хвост). У вопросов с несколькими ответами напротив каждого варианта ставится «да».',
 'Цвет связи: зелёная — ответ открыл дорожку или предыдущее действие закрыто; красная — ответ закрыл дорожку; оранжевая — предыдущее закрыто хвостом; бледная — ещё не открыта.',
 'Цвет облака: синий — можно делать или пора ответить; зелёный — сделано; оранжевый — хвост; жёлтый — зависит от ещё не данного ответа; белый — ждёт предыдущее; бледный — не нужно.',
 'Подсказка к облаку — в примечании (наведите мышь на облако).',
 'Лист «Журнал»: кто, дата, срок хвоста, к чему нужно, комментарий. Просроченный хвост краснеет.',
 'Листы защищены без пароля (Рецензирование → Снять защиту листа), чтобы случайно не стереть формулы.',
 'Ограничения без макросов: связи нарисованы закрашенными ячейками (в местах пересечения видна только вертикальная линия); карта рисуется генератором, поэтому новые вопросы и действия добавляю я, пересобрав шаблон; облака не скрываются, а бледнеют.',
]
helpws['A1'] = 'Как пользоваться'; helpws['A1'].font = F(size=14, bold=True, color='1F3A5F')
for i, t in enumerate(HELP):
    c = helpws.cell(row=3 + i, column=1, value=t); c.font = F(size=10); c.alignment = Alignment(wrap_text=True, vertical='top')

# все ячейки объединённых полей ввода тоже разблокируем (иначе защищённый лист не даст выбрать)
for ws_ in (mp,):
    for mr in ws_.merged_cells.ranges:
        tl = ws_.cell(row=mr.min_row, column=mr.min_col)
        if tl.protection.locked is False:
            for r_ in range(mr.min_row, mr.max_row + 1):
                for c_ in range(mr.min_col, mr.max_col + 1):
                    ws_.cell(row=r_, column=c_).protection = UNLOCK
wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print('saved', OUT, 'layers', NL, 'cols', NCOL, 'rows', MAXR, 'links', len(links), 'edges', len(edges))
