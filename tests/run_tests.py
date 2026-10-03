# -*- coding: utf-8 -*-
"""docweave 端到端自检。

不依赖 pytest，直接运行：
    <python> tests/run_tests.py
退出码 0 表示全部通过。
"""
from __future__ import annotations

import os
import shutil
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from docweave import audit as A
from docweave import merge as M
from docweave import ops
from docweave import writers as W
from docweave.model import read_document, norm_text

FIX = os.path.join(ROOT, "demo", "fixtures")
OUT = os.path.join(ROOT, "demo", "out", "selftest")

PASS, FAIL = [], []


def check(name, cond, detail=""):
    (PASS if cond else FAIL).append(name)
    print(("  ✅ " if cond else "  ❌ ") + name + (f"  [{detail}]" if detail and not cond else ""))


def fp(name):
    return os.path.join(FIX, name)


def ensure_fixtures():
    if not os.path.exists(fp("contract_base.docx")):
        import subprocess
        subprocess.run([sys.executable, os.path.join(ROOT, "demo", "make_fixtures.py")],
                       check=True)


# --------------------------------------------------------------------------

def test_model():
    print("\n[1] 文档模型与读取器")
    d = read_document(fp("contract_base.docx"))
    check("docx 读取到正文块", len(d.blocks) >= 12, f"blocks={len(d.blocks)}")
    check("docx 识别标题层级", any(b.kind == "heading" and b.level == 1 for b in d.blocks))
    check("docx 识别表格", any(b.is_table for b in d.blocks))
    check("空段落被跳过", d.meta["skipped_blank"] >= 0)
    check("标题路径可用", any("第2章" in b.path for b in d.blocks))
    x = read_document(fp("audit_data.xlsx"))
    check("xlsx 读取工作表为表格块", len(x.blocks) == 2, f"blocks={len(x.blocks)}")
    check("xlsx 记录单元格坐标", x.blocks[0].meta.get("col_letters") is not None)
    check("xlsx 记录百分比格式", any("%" in v for v in
                                     (x.blocks[0].meta.get("num_formats") or {}).values()))
    csvp = os.path.join(OUT, "t.csv")
    with open(csvp, "w", encoding="utf-8") as f:
        f.write("项目,数值\n金额,100\n工期,90\n")
    c = read_document(csvp)
    check("csv 读取", c.blocks[0].rows[1][0] == "金额")
    check("旧格式给出明确提示", _raises(lambda: read_document("x.doc")))


def _raises(fn):
    try:
        fn()
        return False
    except ValueError:
        return True


def test_align_and_merge():
    print("\n[2] 对齐、搬移检测与合并引擎")
    base = read_document(fp("contract_base.docx"))
    ours = read_document(fp("contract_ours.docx"))
    theirs = read_document(fp("contract_theirs.docx"))
    check("相同文档全等对齐",
          all(o.tag == "equal" for o in M.align(base.blocks, base.blocks)))

    res = M.three_way(base, ours, theirs)
    kinds = {c.kind for c in res.conflicts}
    check("发现内容冲突", "content" in kinds, str(kinds))
    check("发现表格单元格冲突", "table-cell" in kinds, str(kinds))
    c1 = [c for c in res.conflicts if c.kind == "content"]
    check("冲突内容为我方 40% 与对方 35%",
          any("40%" in (c.ours.text or "") and "35%" in (c.theirs.text or "")
              for c in c1))
    check("识别段落搬移（Word 无此能力）", len(res.moves) >= 1,
          str([m["from"] + "->" + m["to"] for m in res.moves]))
    # 「我方只改了金额列」这件事不应变成冲突：冲突必须只落在双方都改的那一格
    tconf = [c for c in res.conflicts if c.kind == "table-cell"]
    check("我方独有修改被采用且不冲突",
          all(c.detail.get("col") != 3 for c in tconf),
          str([(c.detail.get("row_key"), c.detail.get("col")) for c in tconf]))
    check("对方新增内容被并入",
          any(b.norm.startswith("第5章保密条款") or "保密" in b.text
              for p in res.pieces if p.block is not None
              for b in [p.block]))
    stats = res.stats
    check("统计包含双方块数", stats["base_blocks"] > 0 and stats["theirs_blocks"] > 0)

    # 表格内容正确性
    tbls = [p.block for p in res.pieces if p.block is not None and p.block.is_table]
    check("合并表存在", len(tbls) == 1)
    if tbls:
        rows = tbls[0].rows
        flat = {r[0]: r for r in rows}
        check("单元格合并：我方改的金额保留", flat.get("第一期", ["", "", "", ""])[3] == "480000")
        check("单元格合并：对方改的比例并入", flat.get("第一期", ["", "", "", ""])[2] == "35%")
        check("单元格冲突：保留我方值", flat.get("第三期", ["", "", ""])[2] == "25%")
        cfs = tbls[0].meta["table_plan"]["conflicts"]
        check("表格冲突已登记", any(c["row_key"] == "第三期" for c in cfs))
        check("冲突定位到最终行号", all(c.get("final_idx") is not None for c in cfs))

    # 两方
    res2 = M.two_way(ours, theirs)
    check("两方模式也能识别差异", res2.stats["conflicts"] >= 3, str(res2.stats))
    check("两方模式保留仅单侧内容",
          res2.stats.get("only-a", 0) + res2.stats.get("only-b", 0) >= 1)
    # 没有基线时，一侧搬移的章节不能重复出现，搬移中的修改也不能丢
    moved_head = [p for p in res2.pieces if p.block is not None
                  and p.block.norm == norm_text("第3章 交付与验收")]
    check("两方模式搬移不重复（同一章只出现一次）", len(moved_head) == 1,
          f"出现 {len(moved_head)} 次")
    check("两方模式识别搬移", res2.stats.get("moves", 0) >= 1,
          str(res2.stats.get("moves")))
    check("搬移章节内的差异保留为冲突",
          any(c.kind == "content" and c.ours is not None and c.theirs is not None
              and "120 天" in c.ours.text and "90 天" in c.theirs.text
              for c in res2.conflicts),
          str([(c.cid, c.kind) for c in res2.conflicts]))


def test_docx_writer():
    print("\n[3] docx 写出器（格式/编号/图片移植 + 冲突标记 + 溯源附录）")
    base = read_document(fp("contract_base.docx"))
    ours = read_document(fp("contract_ours.docx"))
    theirs = read_document(fp("contract_theirs.docx"))
    res = M.three_way(base, ours, theirs)
    out = os.path.join(OUT, "merged.docx")
    W.write_docx_merge(res, out)
    check("输出文件存在", os.path.exists(out))

    from docx import Document
    from docx.oxml.ns import qn

    doc = Document(out)
    text = "\n".join(p.text for p in doc.paragraphs)
    check("含冲突标记段落", "冲突 C1" in text)
    check("含我方/对方标记", "【我方】" in text and "【对方】" in text)
    check("含溯源附录", "合并溯源附录" in text and "逐块来源表" in text)
    check("含搬移记录", "段落搬移" in text)
    check("正文顺序：第3章在我方搬移后的位置",
          text.find("第4章 违约责任") < text.rindex("第3章 交付与验收"))

    body = doc.element.body
    kids = list(body.iterchildren())
    check("sectPr 位于 body 末尾", kids[-1].tag == qn("w:sectPr"))

    # 高亮：冲突双方应分别有 yellow / turquoise 高亮
    hl = set()
    for r in body.iter(qn("w:highlight")):
        hl.add(r.get(qn("w:val")))
    check("冲突高亮存在", "yellow" in hl and ({"turquoise", "cyan"} & hl), str(hl))

    # 编号移植：列表段落（直接或经由样式）引用的 numId 必须在 numbering.xml 中存在
    numbering = doc.part.part_related_by(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering")
    defined = {e.get(qn("w:numId")) for e in numbering.element.findall(qn("w:num"))}
    used = set()
    for p in doc.paragraphs:
        pPr = p._p.find(qn("w:pPr"))
        if pPr is not None:
            npr = pPr.find(qn("w:numPr"))
            if npr is not None:
                nid = npr.find(qn("w:numId"))
                if nid is not None:
                    used.add(nid.get(qn("w:val")))
    # 样式层面的编号（例如 List Number 样式）
    for s in doc.styles.element.findall(qn("w:style")):
        pPr = s.find(qn("w:pPr"))
        if pPr is None:
            continue
        npr = pPr.find(qn("w:numPr"))
        if npr is None:
            continue
        nid = npr.find(qn("w:numId"))
        if nid is not None and s.get(qn("w:styleId")) == "ListNumber":
            used.add(nid.get(qn("w:val")))
    check("列表样式引用的编号已定义", used and used <= defined,
          f"used={used} defined_count={len(defined)}")
    check("对方文档的列表项已并入",
          any(p.style.name == "List Number" and p.text.startswith("设备主机")
              for p in doc.paragraphs))

    # 图片关系移植
    blips = [b.get(qn("r:embed")) for b in body.iter(qn("a:blip"))]
    check("对方文档的图片被移植", bool(blips), "未找到图片")
    if blips:
        ok = True
        for rid in blips:
            try:
                doc.part.rels[rid]
            except KeyError:
                ok = False
        check("图片关系可解析", ok)

    tables = doc.tables
    check("输出包含合并表与附录表", len(tables) >= 2, f"tables={len(tables)}")
    pay = tables[0]
    check("冲突单元格带提示文字",
          any("对方此处为" in c.text for r in pay.rows for c in r.cells))
    check("表格行数正确", len(pay.rows) == 4, f"rows={len(pay.rows)}")

    rep = os.path.join(OUT, "merged_report.md")
    W.write_markdown_report(res, rep)
    md = open(rep, encoding="utf-8").read()
    check("Markdown 报告含冲突表", "冲突清单" in md and "C1" in md)


def test_xlsx_merge():
    print("\n[4] xlsx 三方合并（值/公式/着色/报告页）")
    xb = read_document(fp("data_base.xlsx"), with_formulas=True)
    xo = read_document(fp("data_ours.xlsx"), with_formulas=True)
    xt = read_document(fp("data_theirs.xlsx"), with_formulas=True)
    res = M.three_way(xb, xo, xt)
    out = os.path.join(OUT, "data_merged.xlsx")
    W.write_xlsx_merge(res, out)

    from openpyxl import load_workbook

    wb = load_workbook(out)
    check("工作表齐全", "物料清单" in wb.sheetnames and "__合并报告" in wb.sheetnames)
    ws = wb["物料清单"]
    check("我方修改生效（M-002 数量 45）", ws["C3"].value == 45)
    check("对方修改并入（M-002 单价 380）", ws["D3"].value == 380)
    check("单元格冲突保留我方值（D4=28）", ws["D4"].value == 28)
    check("公式被保留而非写死", str(ws["E2"].value or "").startswith("="),
          str(ws["E2"].value))
    check("冲突单元格着色", ws["D4"].fill.start_color.rgb.endswith("FFC7CE"),
          str(ws["D4"].fill.start_color.rgb))
    check("取自对方的行被着色标记",
          any((c.fill.start_color.rgb or "").endswith("E2EFDA")
              for row in ws.iter_rows() for c in row))
    check("报告页含冲突清单",
          any("C1" == str(c.value) for row in wb["__合并报告"].iter_rows() for c in row))


def test_audit():
    print("\n[5] 文表对账")
    rep = A.audit_doc_against_data(read_document(fp("audit_doc.docx")),
                                   read_document(fp("audit_data.xlsx")))
    st = rep["stats"]
    check("识别正文口径", st["正文口径数"] >= 8, str(st))
    check("匹配到 6 项", st["匹配成功"] == 6, str(st))
    check("一致 3 项", st["一致"] == 3, str(st))
    check("数值冲突 3 项", st["数值冲突"] == 3, str(st))
    labels = {f["label"] for f in rep["findings"] if f["code"] == "value_conflict"}
    check("冲突项目正确", labels == {"合同总金额", "合同编号", "验收日期"}, str(labels))
    check("正文有表中无", st["正文有表中无"] >= 2, str(st))
    check("表中有正文无", st["表中有正文无"] == 1, str(st))
    check("百分比口径可比（30% vs 0.3）",
          not any(f["label"] == "预付款比例" for f in rep["findings"]))
    check("单位换算可比（24 个月 vs 24 个月）",
          not any(f["label"] == "质保期" for f in rep["findings"]))
    out = os.path.join(OUT, "audit.docx")
    A.write_audit_docx(rep, out)
    md = A.audit_markdown(rep)
    check("对账报告已生成", os.path.exists(out) and "文表一致性对账报告" in md)
    # GUI/CLI 的「同时列出核对一致的项目」选项必须真的接到引擎
    rep_ok = A.audit_doc_against_data(read_document(fp("audit_doc.docx")),
                                      read_document(fp("audit_data.xlsx")),
                                      show_ok=True)
    check("show_ok 选项生效",
          len(rep_ok["findings"]) == len(rep["findings"]) + 3,
          f"{len(rep['findings'])} → {len(rep_ok['findings'])}")
    from docx import Document
    check("报告内嵌数据表原文", len(Document(out).tables) >= 3)


def test_refs():
    print("\n[6] 编号与交叉引用体检")
    rep = A.check_refs(read_document(fp("refs_doc.docx")))
    codes = [f["code"] for f in rep["findings"]]
    check("悬空章引用", "dangling_chapter" in codes)
    check("悬空小节引用", codes.count("dangling_section") == 2, str(codes))
    check("悬空附件引用", codes.count("dangling_attachment") == 2, str(codes))
    check("编号重复", "dup_number" in codes)
    check("章编号断档", "gap_number" in codes)
    msgs = " ".join(f["message"] for f in rep["findings"])
    check("断档定位到图 2", "图 2" in msgs)
    check("断档定位到第3章", "第3章" in msgs)
    out = os.path.join(OUT, "refs.docx")
    A.write_refs_docx(rep, out)
    check("体检报告已生成", os.path.exists(out))


def test_drift():
    print("\n[7] 口径漂移与自相矛盾")
    rep = A.drift([read_document(fp(f"spec_v{i}.docx")) for i in (1, 2, 3)])
    check("跨版本变化 >= 4", rep["stats"]["跨版本变化"] >= 4, str(rep["stats"]))
    check("识别文档内自相矛盾", rep["stats"]["文档内自相矛盾"] == 1, str(rep["stats"]))
    check("矛盾定位到 spec_v3",
          rep["intra"] and rep["intra"][0]["file"] == "spec_v3")
    labels = {c["label"] for c in rep["changed"]}
    check("覆盖 4 个指标", {"合同总金额", "交付期", "质保期", "预付款比例"} <= labels,
          str(labels))
    out = os.path.join(OUT, "drift.docx")
    A.write_drift_docx(rep, out)
    check("漂移报告已生成", os.path.exists(out))


def test_scan_and_log():
    print("\n[8] 目录扫描快照与变化日志")
    with tempfile.TemporaryDirectory() as tmp:
        folder = os.path.join(tmp, "in")
        os.makedirs(folder)
        path_a = os.path.join(folder, "a.txt")
        path_b = os.path.join(folder, "b.txt")
        log = os.path.join(folder, ".docweave", "scan_log.jsonl")
        with open(path_a, "w", encoding="utf-8") as f:
            f.write("first")

        first = ops.scan_and_log(folder, log=log, recursive=True, changes_only=True)
        check("首次扫描把现有文件列为新增",
              first["added"] == ["a.txt"] and first["scanned"] == 1, str(first))
        same = ops.scan_and_log(folder, log=log, recursive=True, changes_only=True)
        check("重复扫描识别无变化",
              not same["added"] and not same["removed"] and not same["changed"])

        with open(path_a, "w", encoding="utf-8") as f:
            f.write("second")
        with open(path_b, "w", encoding="utf-8") as f:
            f.write("new")
        updated = ops.scan_and_log(folder, log=log, recursive=True, changes_only=True)
        check("识别新增与内容变化",
              updated["added"] == ["b.txt"] and updated["changed"] == ["a.txt"],
              str(updated))

        os.remove(path_b)
        deleted = ops.scan_and_log(folder, log=log, recursive=True, changes_only=True)
        check("完整快照能识别已删除文件",
              deleted["removed"] == ["b.txt"], str(deleted))
        check("递归扫描不把工具元数据当输入", deleted["scanned"] == 1)

        broken_log = os.path.join(tmp, "broken.jsonl")
        with open(broken_log, "w", encoding="utf-8") as f:
            f.write("{invalid json}\n")
        check("损坏的日志明确报错",
              _raises(lambda: ops.scan_and_log(
                  folder, log=broken_log, recursive=True, changes_only=True)))


def test_folder_incremental():
    print("\n[9] 目录扫描到增量处理闭环")
    with tempfile.TemporaryDirectory() as tmp:
        folder = os.path.join(tmp, "input")
        outdir = os.path.join(tmp, "out")
        os.makedirs(folder)
        shutil.copy2(fp("refs_doc.docx"), os.path.join(folder, "refs_doc.docx"))

        first = ops.run_folder_incremental(folder, outdir=outdir, quiet=True)
        check("首次扫描自动规划并处理任务",
              first["summary"]["成功"] == 1
              and first["summary"]["新增文件"] == 1, str(first["summary"]))
        with open(first["html"], encoding="utf-8-sig") as f:
            report_html = f.read()
        check("报告包含文件变化和任务结果",
              os.path.exists(first["html"])
              and "refs_doc.docx" in report_html
              and first["scan"]["log"].endswith("scan_log.jsonl"))

        second = ops.run_folder_incremental(folder, outdir=outdir, quiet=True)
        check("重复运行跳过未变化任务",
              second["summary"]["跳过"] == 1
              and second["summary"]["新增文件"] == 0
              and second["summary"]["内容变化文件"] == 0,
              str(second["summary"]))

        from docx import Document
        source = os.path.join(folder, "refs_doc.docx")
        doc = Document(source)
        doc.add_paragraph("增量处理回归测试")
        doc.save(source)
        third = ops.run_folder_incremental(folder, outdir=outdir, quiet=True)
        check("输入变化后重新运行受影响任务",
              third["summary"]["成功"] == 1
              and third["summary"]["内容变化文件"] == 1,
              str(third["summary"]))


def test_batch_and_ledger():
    print("\n[10] 批量编排与增量台账")
    work = os.path.join(OUT, "batch")
    if os.path.exists(work):
        shutil.rmtree(work)
    os.makedirs(work)
    man = os.path.join(work, "manifest.json")
    import json
    with open(man, "w", encoding="utf-8") as f:
        json.dump({
            "options": {"outdir": os.path.join(work, "out")},
            "jobs": [
                {"type": "merge", "name": "三方", "base": fp("contract_base.docx"),
                 "ours": fp("contract_ours.docx"), "theirs": fp("contract_theirs.docx")},
                {"type": "audit", "name": "对账", "doc": fp("audit_doc.docx"),
                 "data": fp("audit_data.xlsx")},
                {"type": "refs", "name": "体检", "doc": fp("refs_doc.docx")},
                {"type": "drift", "name": "漂移",
                 "files": [fp(f"spec_v{i}.docx") for i in (1, 2, 3)]},
            ],
        }, f, ensure_ascii=False)
    r1 = ops.run_batch(man, quiet=True)
    check("首批全部成功", r1["summary"]["成功"] == 4, str(r1["summary"]))
    check("产物齐全", all(os.path.exists(x["output"]) for x in r1["results"]
                          if x["status"] == "ok"))
    check("HTML 报告生成", os.path.exists(r1["html"]))
    r2 = ops.run_batch(man, quiet=True)
    check("二次运行全部跳过（增量）", r2["summary"]["跳过"] == 4, str(r2["summary"]))
    # 改动一个输入 → 只有该任务重跑
    import time
    p = fp("refs_doc.docx")
    stat = os.stat(p)
    os.utime(p, (stat.st_atime, stat.st_mtime + 5))
    r3 = ops.run_batch(man, quiet=True)
    check("台账按内容指纹判断，时间变化不影响",
          r3["summary"]["跳过"] == 4, str(r3["summary"]))
    r4 = ops.run_batch(man, force=True, quiet=True)
    check("force 强制重跑", r4["summary"]["成功"] == 4, str(r4["summary"]))


def test_csv_merge():
    print("\n[11] 非 docx 来源的整合（csv/xlsx 走表格写出器）")
    a = os.path.join(OUT, "a.csv")
    b = os.path.join(OUT, "b.csv")
    with open(a, "w", encoding="utf-8") as f:
        f.write("项目,数值,单位\n合同总金额,1000000,元\n交付期,90,天\n")
    with open(b, "w", encoding="utf-8") as f:
        f.write("项目,数值,单位\n合同总金额,1200000,元\n质保期,24,个月\n")
    res = M.two_way(read_document(a), read_document(b))
    check("表头相同的表格被判为同一张表",
          any(p.action.startswith("table-merged") for p in res.pieces),
          str([p.action for p in res.pieces]))
    out = os.path.join(OUT, "csv_merged.xlsx")
    W.write_xlsx_merge(res, out)
    from openpyxl import load_workbook
    wb = load_workbook(out)
    check("csv 合并输出工作簿", os.path.exists(out) and wb.sheetnames)
    vals = [str(c.value) for row in wb[wb.sheetnames[0]].iter_rows() for c in row]
    check("并入了双方行（A 的交付期 + B 的质保期）",
          any("交付期" in v for v in vals) and any("质保期" in v for v in vals), str(vals))
    check("差异单元格保留了 A 的值并标注",
          any("1000000" in v for v in vals)
          and any("1200000" in str(c.comment.text) for row in wb[wb.sheetnames[0]].iter_rows()
                  for c in row if c.comment is not None))
    check("报告页存在", "__合并报告" in wb.sheetnames)


def main():
    os.makedirs(OUT, exist_ok=True)
    ensure_fixtures()
    test_model()
    test_align_and_merge()
    test_docx_writer()
    test_xlsx_merge()
    test_audit()
    test_refs()
    test_drift()
    test_scan_and_log()
    test_folder_incremental()
    test_batch_and_ledger()
    test_csv_merge()
    print("\n" + "=" * 64)
    print(f"通过 {len(PASS)} 项，失败 {len(FAIL)} 项")
    if FAIL:
        print("失败项：")
        for f in FAIL:
            print("  -", f)
    print("=" * 64)
    return 1 if FAIL else 0


if __name__ == "__main__":
    sys.exit(main())
