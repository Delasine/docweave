# -*- coding: utf-8 -*-
"""冒烟测试：三方合并 / 两方合并 / 表格合并 / 关系移植。"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from docweave import merge as M
from docweave.model import read_document
from docweave import writers as W

FIX = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                   "demo", "fixtures")
OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "demo", "out")
os.makedirs(OUT, exist_ok=True)


def show(tag, res):
    print("=" * 70)
    print(tag, res.stats)
    print("冲突：")
    for c in res.conflicts:
        print(f"  {c.cid} [{c.kind}] {c.path} | 我方={c.ours.preview(34) if c.ours else None}"
              f" | 对方={c.theirs.preview(34) if c.theirs else None}")
    print("搬移：")
    for m in res.moves:
        print(f"  {m['side']} {m['from']} -> {m['to']} :: {m['text']}")
    print("警告：", res.warnings)
    print("前 18 个输出块：")
    for p in res.pieces[:18]:
        b = p.block
        print(f"  {p.action:14s} {p.origin:12s} {b.kind:7s} {b.preview(44)}")


def main():
    base = read_document(os.path.join(FIX, "contract_base.docx"))
    ours = read_document(os.path.join(FIX, "contract_ours.docx"))
    theirs = read_document(os.path.join(FIX, "contract_theirs.docx"))
    for d in (base, ours, theirs):
        print(f"{d.label}: 块={len(d.blocks)} 表={d.meta.get('tables')} "
              f"标题={d.meta.get('headings')} 修订={d.meta.get('tracked_changes')}")
    res = M.three_way(base, ours, theirs)
    show("三方合并", res)
    docx_out = os.path.join(OUT, "contract_merged.docx")
    W.write_docx_merge(res, docx_out)
    W.write_markdown_report(res, os.path.join(OUT, "contract_merged_report.md"))

    doc2 = read_document(docx_out)
    print(f"\n输出文档：块={len(doc2.blocks)} 表格={doc2.meta.get('tables')}")

    # 两方
    res2 = M.two_way(ours, theirs)
    show("两方合并", res2)
    W.write_docx_merge(res2, os.path.join(OUT, "contract_twoway.docx"))

    # xlsx 三方
    xb = read_document(os.path.join(FIX, "data_base.xlsx"), with_formulas=True)
    xo = read_document(os.path.join(FIX, "data_ours.xlsx"), with_formulas=True)
    xt = read_document(os.path.join(FIX, "data_theirs.xlsx"), with_formulas=True)
    res3 = M.three_way(xb, xo, xt)
    show("xlsx 三方", res3)
    W.write_xlsx_merge(res3, os.path.join(OUT, "data_merged.xlsx"))
    W.write_markdown_report(res3, os.path.join(OUT, "data_merged_report.md"))
    print("\n全部冒烟步骤完成")


if __name__ == "__main__":
    main()
