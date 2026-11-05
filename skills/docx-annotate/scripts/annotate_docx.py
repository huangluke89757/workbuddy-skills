#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
docx-annotate: 往 docx 注入 Word 批注(comments)并对问题内容 run 级标红，输出新文档。

用法:
  python annotate_docx.py --src SRC.docx --out OUT.docx --findings findings.json \
      [--author "评审"] [--keep-source] [--no-keep-source]

findings.json 格式:
  [
    {
      "tokens": ["子串1", "子串2"],
      "comment": "【类别】一句话问题 + 应/建议…修正",
      "exact_tokens": ["备用"]      # 可选：仅匹配表格单元精确相等值
    }
  ]

实现要点（来自 AMHS 评审任务三轮收敛 v2→v3→v4）:
- 用 lxml 直写 word/document.xml，绕开 python-docx 段落对象别名导致同 id 错挂。
- 先扫源文档 commentRangeStart 的 id 最大值 base；保留源批注则新 id 从 base 起，否则剥离并 base=0。
- 标红 = run 级：遍历 run 算累计偏移，只给与命中子串重叠的 run 设 w:color=C00000。
- 匹配避开 TOC / PAGEREF 段落；表格精确值用 require_table+exact 区分正文与表格。
"""
import argparse
import datetime
import json
import os
import zipfile
import shutil
from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
WURI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XMLNS = "{http://www.w3.org/XML/1998/namespace}"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


def get_t(r):
    t = r.find(W + "t")
    return t.text if (t is not None and t.text) else ""


def get_text(p):
    return "".join(get_t(r) for r in p.findall(W + "r"))


def pstyle(p):
    ppr = p.find(W + "pPr")
    if ppr is None:
        return None
    ps = ppr.find(W + "pStyle")
    return ps.get(W + "val") if ps is not None else None


def find_p(all_p, substr, require_table=False, exact=False):
    res = []
    for p in all_p:
        txt = get_text(p)
        st = pstyle(p)
        if st and st.upper().startswith("TOC"):
            continue
        if "PAGEREF" in txt:
            continue
        in_table = len(list(p.iterancestors(W + "tc"))) > 0
        if require_table and not in_table:
            continue
        if exact:
            if txt.strip() == substr:
                res.append(p)
        elif substr in txt:
            res.append(p)
    return res


def add_red(r):
    rpr = r.find(W + "rPr")
    if rpr is None:
        rpr = etree.Element(W + "rPr")
        r.insert(0, rpr)
    color = rpr.find(W + "color")
    if color is None:
        color = etree.SubElement(rpr, W + "color")
    color.set(W + "val", "C00000")


def redden_span(p, substr):
    runs = p.findall(W + "r")
    full = "".join(get_t(r) for r in runs)
    start = full.find(substr)
    if start < 0:
        for r in runs:
            add_red(r)
        return
    end = start + len(substr)
    pos = 0
    for r in runs:
        rt = get_t(r)
        rs, re_ = pos, pos + len(rt)
        if not (re_ <= start or rs >= end):
            add_red(r)
        pos = re_


def add_comment_markers(p, cid):
    crs = etree.Element(W + "commentRangeStart")
    crs.set(W + "id", str(cid))
    p.insert(0, crs)
    cre = etree.Element(W + "commentRangeEnd")
    cre.set(W + "id", str(cid))
    ref_r = etree.Element(W + "r")
    etree.SubElement(ref_r, W + "rPr")
    cref = etree.SubElement(ref_r, W + "commentReference")
    cref.set(W + "id", str(cid))
    p.append(cre)
    p.append(ref_r)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--findings", required=True)
    ap.add_argument("--author", default="评审")
    ap.add_argument("--keep-source", dest="keep_source", action="store_true", default=True)
    ap.add_argument("--no-keep-source", dest="keep_source", action="store_false")
    args = ap.parse_args()

    with open(args.findings, encoding="utf-8") as f:
        findings = json.load(f)

    NOW = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")
    src_zip = zipfile.ZipFile(args.src)
    names = src_zip.namelist()
    doc_bytes = src_zip.read("word/document.xml")
    root = etree.fromstring(doc_bytes)
    body = root.find(W + "body")
    all_p = list(root.iter(W + "p"))

    # ---- 1) 扫源批注 id 范围，定 base ----
    src_markers = root.findall(".//" + W + "commentRangeStart")
    src_ids = [int(s.get(W + "id")) for s in src_markers if s.get(W + "id", "").isdigit()]
    src_cmt_bytes = src_zip.read("word/comments.xml") if "word/comments.xml" in names else None
    if src_cmt_bytes is not None:
        scm = etree.fromstring(src_cmt_bytes)
        for c in scm.findall(W + "comment"):
            try:
                src_ids.append(int(c.get(W + "id")))
            except (TypeError, ValueError):
                pass

    if args.keep_source:
        base = (max(src_ids) + 1) if src_ids else 0
    else:
        # 剥离源批注标记
        for s in root.findall(".//" + W + "commentRangeStart"):
            s.getparent().remove(s)
        for s in root.findall(".//" + W + "commentRangeEnd"):
            s.getparent().remove(s)
        for s in root.findall(".//" + W + "commentReference"):
            # commentReference 在 w:r 内
            parent = s.getparent()
            if parent is not None:
                parent.getparent().remove(parent)
        base = 0
    print("保留源批注=%s, base id=%d" % (args.keep_source, base))

    # ---- 2) 注入新批注 ----
    new_comments = []  # (cid, text)

    def reg(text):
        cid = base + len(new_comments)
        new_comments.append((cid, text))
        return cid

    for fi in findings:
        tokens = fi.get("tokens", [])
        comment = fi.get("comment", "")
        exact_tokens = set(fi.get("exact_tokens", []))
        cid = reg(comment)
        for s in tokens:
            if s in exact_tokens:
                ps = find_p(all_p, s, require_table=True, exact=True)
            else:
                ps = find_p(all_p, s)
                if not ps:
                    ps = find_p(all_p, s, require_table=True)
            if not ps:
                print("  [WARN] cid=%d 未匹配到: %s" % (cid, s[:24]))
            for p in ps:
                redden_span(p, s)
                add_comment_markers(p, cid)

    # ---- 3) 序列化 document.xml ----
    new_doc = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    data = {n: src_zip.read(n) for n in names}
    data["word/document.xml"] = new_doc

    # ---- 4) comments.xml：合并或新建 ----
    if args.keep_source and src_cmt_bytes is not None:
        cmt_root = etree.fromstring(src_cmt_bytes)
    else:
        cmt_root = etree.Element(W + "comments")
        cmt_root.set("xmlns", WURI)
    for cid, text in new_comments:
        c = etree.SubElement(cmt_root, W + "comment")
        c.set(W + "id", str(cid))
        c.set(W + "author", args.author)
        c.set(W + "date", NOW)
        c.set(W + "initials", "评")
        cp = etree.SubElement(c, W + "p")
        cr = etree.SubElement(cp, W + "r")
        ct = etree.SubElement(cr, W + "t")
        ct.set(XMLNS + "space", "preserve")
        ct.text = text
    comments_xml = etree.tostring(cmt_root, xml_declaration=True, encoding="UTF-8", standalone=True)
    data["word/comments.xml"] = comments_xml
    if "word/comments.xml" not in names:
        names.append("word/comments.xml")

    # ---- 5) 补 rels + content types（源无批注时必须）----
    rels_name = "word/_rels/document.xml.rels"
    if rels_name in data:
        rels_root = etree.fromstring(data[rels_name])
        if not any(r.get("Type", "").endswith("/comments") for r in rels_root):
            existing = [r.get("Id") for r in rels_root]
            rid = "rIdComments"
            n = 1
            while rid in existing:
                rid = "rIdComments%d" % n
                n += 1
            el = etree.SubElement(rels_root, "{" + REL_NS + "}Relationship")
            el.set("Id", rid)
            el.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments")
            el.set("Target", "comments.xml")
            data[rels_name] = etree.tostring(rels_root, xml_declaration=True, encoding="UTF-8", standalone=True)

    ct_name = "[Content_Types].xml"
    if ct_name in data:
        ct_root = etree.fromstring(data[ct_name])
        has_ct = any(o.get("PartName") == "/word/comments.xml"
                     for o in ct_root if o.tag.endswith("Override"))
        if not has_ct:
            o = etree.SubElement(ct_root, "{" + CT_NS + "}Override")
            o.set("PartName", "/word/comments.xml")
            o.set("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml")
            data[ct_name] = etree.tostring(ct_root, xml_declaration=True, encoding="UTF-8", standalone=True)

    # ---- 6) 重打包 ----
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    tmp = args.out + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.writestr(n, data[n])
    shutil.move(tmp, args.out)

    # ---- 验收统计 ----
    red_runs = 0
    for r in root.iter(W + "r"):
        rpr = r.find(W + "rPr")
        if rpr is not None:
            col = rpr.find(W + "color")
            if col is not None and col.get(W + "val") == "C00000":
                red_runs += 1
    starts = len(root.findall(".//" + W + "commentRangeStart"))
    print("DONE. 新增批注=%d (id %d..%d), 标红run=%d, 标记start=%d"
          % (len(new_comments), base, base + len(new_comments) - 1, red_runs, starts))


if __name__ == "__main__":
    main()
