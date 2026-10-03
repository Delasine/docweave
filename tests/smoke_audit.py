# -*- coding: utf-8 -*-
"""冒烟测试：文表对账 / 引用体检 / 口径漂移。"""
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from docweave.audit import audit_doc_against_data, check_refs, drift
from docweave.model import read_document

FIX = os.path.join(ROOT, "demo", "fixtures")


def main():
    doc = read_document(os.path.join(FIX, "audit_doc.docx"))
    data = read_document(os.path.join(FIX, "audit_data.xlsx"))
    rep = audit_doc_against_data(doc, data)
    print("=== 文表对账 ===")
    print(rep["stats"])
    for f in rep["findings"]:
        print(f"  [{f['severity']}] {f['code']}: {f['message']}")

    print("\n=== 引用体检 ===")
    refs = check_refs(read_document(os.path.join(FIX, "refs_doc.docx")))
    print(refs["stats"])
    for f in refs["findings"]:
        print(f"  [{f['severity']}] {f['code']}: {f['message']}")

    print("\n=== 口径漂移 ===")
    docs = [read_document(os.path.join(FIX, f"spec_v{i}.docx")) for i in (1, 2, 3)]
    dr = drift(docs)
    print(dr["stats"])
    for c in dr["changed"]:
        print("  变化:", c["message"])
    for c in dr["intra"]:
        print("  矛盾:", c["message"])


if __name__ == "__main__":
    main()
