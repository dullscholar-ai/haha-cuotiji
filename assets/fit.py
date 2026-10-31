# -*- coding: utf-8 -*-
"""fit.py —— 依据 measure.ps1 的实测结果，反解每题该压多少，使"一题一页"。

两条规则，按题独立生效：
  1) 通栏截图：块高超出单页 → 先压截图高度（下限 4.5cm），压到底还超 → 再压字号（下限 0.78）；
  2) 并列截图：截图在解析右侧独立一栏 → 在"一页放得下"的前提下，把截图栏放到最宽。

用法：
    python fit.py --measure measure.txt --content content.json --figs figs \
                  --fit fit.json --side 01,02,04

    # 第一轮（还没 fit.json 时）会产生 fig/fs；第二轮针对仍跨页的题再压一次；
    # 一般 2~3 轮收敛，每轮：build_docx.py → measure.ps1 → fit.py → build_docx.py

字段口径：
    fig[no]   通栏截图的最大高度（cm），对应 build_docx.py 的 FIG_MAX_H 覆盖值
    fs[no]    该题字号缩放系数（1.0 = 原始字号）
    sidew[no] 并列版式的截图栏宽度（cm）
"""
import argparse
import io
import json
import os
import re

from PIL import Image

TOPC, BOTTOM = 1.55, 1.35
USABLE = 29.7 - TOPC - BOTTOM     # 26.8cm
FIG_FLOOR = 4.5                   # 通栏截图最小高度
FS_FLOOR = 0.78                   # 字号最小缩放
WMAX_R = 9.0                      # 并列版式截图栏最宽
WNARROW = 17.04                   # 笔记区可用宽度
FIG_MAX_W = 11.9
AVAIL = WNARROW - 0.36


def parse_measure(path):
    rows = {}
    for line in io.open(path, encoding='utf-8'):
        line = line.strip()
        m = re.match(r'Q(\d+) Tq=([\d.]+) Ta=([\d.]+) yans=([\d.]+) p=(\d+) pa=(\d+)', line)
        if m:
            rows[m.group(1).zfill(2)] = dict(tq=float(m.group(2)), ta=float(m.group(3)),
                                             yans=float(m.group(4)), page=int(m.group(5)),
                                             pa=int(m.group(6)))
    return rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--measure', required=True)
    ap.add_argument('--content', required=True)
    ap.add_argument('--figs', default='figs')
    ap.add_argument('--fit', default='fit.json')
    ap.add_argument('--side', default='', help='并列版式题号，如 01,02')
    ap.add_argument('--limit', type=float, default=26.15, help='允许的最大块高（cm）')
    args = ap.parse_args()

    side = {x.strip() for x in args.side.split(',') if x.strip()}
    meas = parse_measure(args.measure)
    content = json.load(io.open(args.content, encoding='utf-8'))
    if isinstance(content, dict):
        content = content.get('questions', [])
    meta = {}
    for it in content:
        no = it['no'].zfill(2)
        meta[no] = dict(n=len(it.get('summ', [])), fig=it['fig'][0])

    fit = {'fig': {}, 'fs': {}, 'sidew': {}}
    if os.path.exists(args.fit):
        fit.update(json.load(io.open(args.fit, encoding='utf-8')))
    for k in ('fig', 'fs', 'sidew'):
        fit.setdefault(k, {})

    def ratio(no):
        p = os.path.join(args.figs, meta[no]['fig'])
        w, h = Image.open(p).size
        return h / w

    def band_h(no, cap):
        r = ratio(no)
        return min(FIG_MAX_W, AVAIL, cap / r) * r

    touched = []
    print('%-4s %-6s %-7s %-7s %-8s %-6s %s' % ('no', 'yans', 'over', 'fig_h', 'fs', 'wR', 'action'))
    for no in sorted(meas):
        d = meas[no]
        if no not in meta:
            continue
        ans_h = 0.85 + 0.47 * meta[no]['n']
        H = (d['pa'] - d['page']) * USABLE + (d['yans'] - TOPC) + ans_h
        over = H - args.limit
        cur_fs = float(fit['fs'].get(no, 1.0))
        note = []
        wR = ''

        if no in side:
            fixed = H - d['ta']
            r = ratio(no)
            best = float(fit['sidew'].get(no, 6.6))
            for w in range(60, int(WMAX_R * 10) + 1, 2):
                ww = w / 10.0
                wl_in = (WNARROW - ww) - 0.30
                if wl_in < 5.0:
                    continue
                ta2 = d['ta'] * (10.15 / wl_in)
                if fixed + max(ta2, (ww - 0.45) * r) <= args.limit:
                    best = ww
            fit['sidew'][no] = round(best, 2)
            wR = '%.1f' % best
            note.append('sidew->%.1f' % best)
            if over > 0.10:                       # 并列版式仍超：减小字号
                t = max(2.0, d['tq'] + d['ta'])
                f = (max(6.0, t - over) / t) ** 0.5
                fit['fs'][no] = round(max(FS_FLOOR, cur_fs * f), 3)
                note.append('fs->%.3f' % fit['fs'][no])
        else:
            cur = float(fit['fig'].get(no, 10.8))
            curh = band_h(no, cur)
            if over > 0.10:
                newh = max(FIG_FLOOR, curh - over)
                deficit = over - (curh - newh)
                fit['fig'][no] = round(newh, 2)
                note.append('fig->%.2f' % newh)
                if deficit > 0.05:
                    text_h = H - 1.60 - curh - 0.62 - ans_h
                    if text_h > 2:
                        f = (max(6.0, text_h - deficit) / text_h) ** 0.5
                        fit['fs'][no] = round(max(FS_FLOOR, cur_fs * f), 3)
                        note.append('fs->%.3f' % fit['fs'][no])
        fit['fig'].setdefault(no, float(fit['fig'].get(no, 10.8)))
        fit['fs'].setdefault(no, cur_fs)
        figh = (float(fit['sidew'].get(no, 6.6)) - 0.45) * ratio(no) if no in side             else band_h(no, float(fit['fig'][no]))
        print('%-4s %-6.2f %-7.2f %-7.2f %-8.3f %-6s %s'
              % (no, d['yans'], over, figh, fit['fs'][no], wR, ' '.join(note)))
        if over > 0.10:
            touched.append(no)

    json.dump(fit, io.open(args.fit, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print('saved', args.fit)
    if touched:
        print('本轮调过的题:', ' '.join(touched), '→ 重新 build + measure 复测')


if __name__ == '__main__':
    main()
