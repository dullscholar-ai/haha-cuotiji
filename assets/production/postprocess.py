# -*- coding: utf-8 -*-
"""在 build_notebook 产物上做「精简版」改造（生产脚本，改路径即可复用）：

1. 删除封面页（含「错题纲要」）
2. 删除「题目索引」页与「编排说明」页
3. 删除「错因归纳」页（含核心问题、学习建议）
4. 删除每题「线索栏」「总结栏」及其分区标签（不再用康奈尔三栏）
5. 卷面原件照片并入「② 解题过程」右侧
6. Markdown 按精简结构重写
"""
import io
import json
import os
import re

from bs4 import BeautifulSoup

DIR = os.environ.get("CJT_DIR") or os.path.dirname(os.path.abspath(__file__))
CODE = os.environ.get("CJT_CODE") or _detect_code(DIR)


def _detect_code(d):
    """目录里唯一一个 <code>.html 就当作本册代号；多了就设 CJT_CODE。"""
    hits = [f[:-5] for f in os.listdir(d) if f.endswith(".html")]
    if len(hits) == 1:
        return hits[0]
    raise SystemExit("请用环境变量 CJT_CODE 指定本册代号（%s）" % (hits or "目录里没有 html"))

ROWS = [
    ["温度／℃", "15", "20", "25", "30", "35", "40"],
    ["产卵数／个", "0", "1500", "2200", "1800", "1000", "0"],
    ["孵化率／%", "0", "60", "85", "83", "75", "0"],
]

EXTRA_CSS = """
  /* 单栏正文：卷面原件与「② 解题过程」并排，图大幅放大 */
  .page.qpage .cornell-grid { display: block; }
  .main-column .sec-row { display: flex; gap: 3mm; align-items: flex-start; margin: 1mm 0 0; }
  .main-column .sec-row .sec-text { flex: 0 0 55mm; min-width: 0; }
  .main-column .sec-row .sec-fig { flex: 1 1 auto; margin: 0; text-align: center; }
  .main-column .sec-row .sec-fig img {
    width: auto !important; height: auto !important;
    max-width: 100% !important; max-height: 160mm !important;
    border: 1px solid #8c8c8c; margin: 0;
  }
  .figpage { text-align: center; margin: 3mm 0 0; }
  .figpage img { width: 48% !important; height: auto !important; border: 1px solid #8c8c8c; margin: 0 1%; vertical-align: top; }
  .figpage .figcap { font-size: 9.5px; color: #666; margin-top: 1.5mm; }
  .main-column .sec-row .sec-fig .figcap { font-size: 9.5px; color: #666; margin-top: 1.5mm; line-height: 1.35; }
"""


def strip_q(q):
    q = q.strip()
    m = re.match(r"^(典例|变式)\s*([0-9]+\S*)$", q)
    if m:
        return "%s %s" % (m.group(1), m.group(2))
    return q


def html_table():
    out = ['<table class="data-table">']
    for i, row in enumerate(ROWS):
        tag = "th" if i == 0 else "td"
        out.append("<tr>" + "".join("<%s>%s</%s>" % (tag, c, tag) for c in row) + "</tr>")
    out.append("</table>")
    return "\n".join(out)


def md_inline(s):
    return re.sub(r"@@(.+?)@@", r"`\1`", s)


def clean_questions(C):
    """题目数据瘦身：去掉出处、星标；③ 只留答案（删掉错因）。"""
    for x in C["questions"]:
        x.pop("src", None)
        new_body = []
        for title, lines in x["body"]:
            if title.startswith("③"):
                lines = [ln for ln in lines if not ln.strip().startswith("**错因**")]
                title = "③ 答案"
            new_body.append([title, lines])
        x["body"] = new_body
    return C


def build_md(C):
    out = ["# 初中科学典型习题 · 详细解析（%d 题）" % len(C["questions"]), ""]
    for x in C["questions"]:
        out.append("---")
        out.append("")
        out.append("## 第 %s 题 ｜ %s ｜ %s" % (
            strip_q(x["q"]), x["sec"], x["topic"]))
        out.append("")
        for title, lines in x["body"]:
            out.append("### %s" % title)
            out.append("")
            for ln in lines:
                if ln == "————":
                    out.append("---")
                    out.append("")
                    continue
                if ln.strip() == "[[TABLE:MOSQUITO]]":
                    out.append("| " + " | ".join(ROWS[0]) + " |")
                    out.append("|" + "|".join([":---:"] * len(ROWS[0])) + "|")
                    for row in ROWS[1:]:
                        out.append("| " + " | ".join(row) + " |")
                    out.append("")
                    continue
                out.append(md_inline(ln))
                out.append("")
            if title.startswith("②") and x.get("img"):
                if x.get("imgcap"):
                    out.append("*%s*" % x["imgcap"])
                    out.append("")
                out.append("![%s](%s)" % (x.get("imgcap", "卷面原件"), x["img"]))
                out.append("")
    return "\n".join(out)


def main():
    C = clean_questions(json.load(io.open(os.path.join(DIR, "content.json"), encoding="utf-8")))
    html_path = os.path.join(DIR, CODE + ".html")
    html = io.open(html_path, encoding="utf-8").read()
    soup = BeautifulSoup(html, "lxml")

    # 1) 删封面、索引、编排说明、错因归纳（整页）
    for page in soup.select("div.page"):
        cls = page.get("class") or []
        title = page.find(class_="page-title")
        t = title.get_text().strip() if title else ""
        if "cover" in cls or (t and ("索引" in t or "归纳" in t or "编排" in t)):
            page.decompose()
    for box in soup.select("div.callout"):
        box.decompose()

    # 2) 每题：删线索栏、总结栏、分区标签、出处、日期页码、星标；③ 只留答案
    for blk in soup.select("div.question-block"):
        for node in blk.select("div.cue-column, div.summary-column, div.zone-label, div.q-src, span.q-meta"):
            node.decompose()
        h2 = blk.find("h2")
        if h2 is not None:
            s = re.sub(r"\s*｜\s*★+\s*$", "", h2.get_text().strip())
            h2.string = s
        for sec in blk.select("div.section"):
            st = sec.find("span", class_="section-title")
            if st is None:
                continue
            if st.get_text().strip().startswith("③"):
                st.string = "③ 答案"
                for div in sec.find_all("div"):
                    if div.get_text().strip().startswith("错因"):
                        div.decompose()

    # 3) 数据表插回原题
    for div in soup.find_all("div"):
        if div.string == "[[TABLE:MOSQUITO]]":
            div.replace_with(BeautifulSoup(html_table(), "html.parser"))

    # 4) 卷面原件与「② 解题过程」并排（图在右、放大）
    for i, blk in enumerate(soup.select("div.question-block")):
        q = C["questions"][i]
        if q.get("imgpage"):
            # 整页长截图：作为独立图页跟在题目页之后（两栏并排）
            fp = BeautifulSoup('<div class="figpage"></div>', "html.parser").div
            for key in ("a", "b"):
                im = soup.new_tag("img", src=q["imgpage"][key], **{"class": "q-img"})
                fp.append(im)
            cap = soup.new_tag("div", **{"class": "figcap"})
            cap.string = q["imgpage"].get("cap", "")
            fp.append(cap)
            blk.insert_after(fp)
            continue
        if not q.get("img"):
            continue
        sec = None
        for s in blk.select("div.section"):
            st = s.find("span", class_="section-title")
            if st and st.get_text().strip().startswith("②"):
                sec = s
                break
        if sec is None:
            continue
        row = BeautifulSoup('<div class="sec-row"></div>', "html.parser").div
        text = BeautifulSoup('<div class="sec-text"></div>', "html.parser").div
        fig = BeautifulSoup('<div class="sec-fig"></div>', "html.parser").div
        img = soup.new_tag("img", src=q["img"], **{"class": "q-img", "alt": q.get("imgcap", "")})
        cap = soup.new_tag("div", **{"class": "figcap"})
        cap.string = q.get("imgcap", "")
        fig.append(img)
        fig.append(cap)
        for child in list(sec.children):
            text.append(child.extract())
        row.append(text)
        row.append(fig)
        sec.replace_with(row)

    # 5) 输出
    html = re.sub(r"<h2>第 (典例|变式)\s*([0-9][^｜]*?)\s*题 ｜", r"<h2>\1 \2 ｜", str(soup))
    html = html.replace("康奈尔错题本", "典型习题 · 详细解析")
    html = html.replace("</style>", EXTRA_CSS + "\n</style>", 1)

    io.open(html_path, "w", encoding="utf-8").write(html)
    io.open(os.path.join(DIR, CODE + ".md"), "w", encoding="utf-8").write(build_md(C))
    print("postprocess ok | pages:", html.count('class="page'), "| q:", html.count("question-block"))


if __name__ == "__main__":
    main()



# ---------------------------------------------------------------------------
# 复用提示
# 1) 本脚本吃的是「康奈尔错题集」skill 的 build_notebook.py 产物 HTML；
# 2) 目录与代号：CJT_DIR / CJT_CODE 环境变量，或直接改上面的常量；
# 3) [[TABLE:KEY]] 在这里由 html_table() 展开，换成自己的数据表即可；
# 4) 想恢复被删的区块（封面/索引/错因/订正区），删掉 main() 里对应那一段。
# ---------------------------------------------------------------------------
