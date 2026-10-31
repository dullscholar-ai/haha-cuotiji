# -*- coding: utf-8 -*-
"""prep_figs.py —— 卷面截图的准备工作（切分 / 拼接 / 去白边 / 量尺寸）

一道题一张图是"图题一一对应"的前提。原始素材常见三种情况：
  1) 一张长截图里叠了好几道题 → 按空行切开（auto 或 手动 cuts）；
  2) 竖长条太窄、缩放后字太小 → 切成左右两栏并排（join）；
  3) 四周大片留白 → 先 trim 再去排版。

用法：
    python prep_figs.py auto  --src scan.jpg --outdir figs --prefix fig --gap 40
    python prep_figs.py split --src scan.jpg --cuts 1200,2400 --outdir figs --prefix fig
    python prep_figs.py join  --a figs/l.jpg --b figs/r.jpg --out figs/fig01.jpg
    python prep_figs.py trim  --src scan.jpg --out figs/fig01.jpg --pad 8
    python prep_figs.py info  --glob "figs/*.jpg"

只依赖 Pillow。
"""
import argparse
import glob
import io
import os

from PIL import Image

WHITE = 245          # 认为"接近白"的灰度阈值
EDGE = 12            # 判定整行/整列是空白的容差（像素行内非白像素数）


def _is_blank_row(img_gray, y):
    row = img_gray.crop((0, y, img_gray.width, y + 1))
    return min(row.getdata()) > WHITE


def auto_cuts(path, gap=40):
    """返回空白带的中点 y 列表：在一道题与下一道题之间下刀。"""
    im = Image.open(path).convert('L')
    blank = [_is_blank_row(im, y) for y in range(im.height)]
    cuts, run_start = [], None
    for y, b in enumerate(blank):
        if b and run_start is None:
            run_start = y
        elif not b and run_start is not None:
            if y - run_start >= gap and run_start > 0:
                cuts.append((run_start + y) // 2)
            run_start = None
    return cuts


def cmd_auto(a):
    cuts = auto_cuts(a.src, a.gap)
    print('检测到空白带 %d 处，切点 y = %s' % (len(cuts), cuts))
    cmd_split(argparse.Namespace(src=a.src, cuts=','.join(str(c) for c in cuts),
                                 outdir=a.outdir, prefix=a.prefix, start=a.start,
                                 trim=a.trim, ext=a.ext))


def cmd_split(a):
    src = Image.open(a.src)
    im = src.crop((0, 0, src.width, src.height))
    ys = [0] + [int(float(x)) for x in a.cuts.split(',') if x.strip()] + [im.height]
    os.makedirs(a.outdir, exist_ok=True)
    for i in range(len(ys) - 1):
        piece = im.crop((0, ys[i], im.width, ys[i + 1]))
        name = os.path.join(a.outdir, '%s%02d.%s' % (a.prefix, a.start + i, a.ext))
        _save(piece, name, a.trim)
        print(name, piece.width, 'x', piece.height)


def cmd_join(a):
    top, bot = a.src, a.dst
    l = Image.open(top)
    r = Image.open(bot)
    h = max(l.height, r.height)
    out = Image.new('RGB', (l.width + r.width + a.gap, h), 'white')
    out.paste(l, (0, 0))
    out.paste(r, (l.width + a.gap, 0))
    _save(out, a.out, a.pad)
    print(a.out, out.width, 'x', out.height)


def cmd_trim(a):
    im = Image.open(a.src).convert('RGB')
    gray = im.convert('L')
    bbox = _content_bbox(gray, a.tol)
    if bbox is None:
        print('整幅都是白的，跳过'); return
    box = (max(0, bbox[0] - a.pad), max(0, bbox[1] - a.pad),
           min(im.width, bbox[2] + a.pad), min(im.height, bbox[3] + a.pad))
    out = im.crop(box)
    os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
    _save(out, a.out, 0)
    print(a.out, out.width, 'x', out.height)


def cmd_info(a):
    for p in sorted(glob.glob(a.glob)):
        w, h = Image.open(p).size
        r = h / w
        # 在 17cm 通栏里的高度；在 6.2cm 侧栏里的高度
        print('%-28s %5dx%-5d 比例 %.2f  通栏高 %.1fcm  侧栏高 %.1fcm'
              % (os.path.basename(p), w, h, r, min(10.8, 11.9 * r), 6.2 * r))


def _content_bbox(gray, tol):
    from PIL import ImageOps
    inv = ImageOps.invert(gray)
    bbox = inv.point(lambda v: 255 if v > tol else 0).getbbox()
    return bbox


def _save(img, path, trim_pad):
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    if trim_pad:
        gray = img.convert('L')
        bbox = _content_bbox(gray, 6)
        if bbox:
            img = img.crop((max(0, bbox[0] - trim_pad), max(0, bbox[1] - trim_pad),
                            min(img.width, bbox[2] + trim_pad), min(img.height, bbox[3] + trim_pad)))
    img.save(path, quality=92)


def main():
    ap = argparse.ArgumentParser(description='卷面截图准备工具')
    sub = ap.add_subparsers(dest='cmd', required=True)

    p = sub.add_parser('auto', help='按空白带自动切分长截图')
    p.add_argument('--src', required=True)
    p.add_argument('--outdir', default='figs')
    p.add_argument('--prefix', default='fig')
    p.add_argument('--start', type=int, default=1)
    p.add_argument('--gap', type=int, default=40, help='空白带最小高度（像素）')
    p.add_argument('--trim', type=int, default=0, help='每块再去白边（像素）')
    p.add_argument('--ext', default='jpg')
    p.set_defaults(func=cmd_auto)

    p = sub.add_parser('split', help='按给定 y 坐标切分')
    p.add_argument('--src', required=True)
    p.add_argument('--cuts', required=True, help='如 1200,2400')
    p.add_argument('--outdir', default='figs')
    p.add_argument('--prefix', default='fig')
    p.add_argument('--start', type=int, default=1)
    p.add_argument('--trim', type=int, default=0)
    p.add_argument('--ext', default='jpg')
    p.set_defaults(func=cmd_split)

    p = sub.add_parser('join', help='左右两张拼成一张（拯救超窄竖条）')
    p.add_argument('--a', required=True, dest='src')
    p.add_argument('--b', required=True, dest='dst')
    p.add_argument('--out', required=True)
    p.add_argument('--gap', type=int, default=16)
    p.add_argument('--pad', type=int, default=0)
    p.set_defaults(func=cmd_join)

    p = sub.add_parser('trim', help='裁掉四周白边')
    p.add_argument('--src', required=True)
    p.add_argument('--out', required=True)
    p.add_argument('--pad', type=int, default=8)
    p.add_argument('--tol', type=int, default=6)
    p.set_defaults(func=cmd_trim)

    p = sub.add_parser('info', help='列出尺寸与建议图高')
    p.add_argument('--glob', default='figs/*.jpg')
    p.set_defaults(func=cmd_info)

    a = ap.parse_args()
    a.func(a)


if __name__ == '__main__':
    main()
