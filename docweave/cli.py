# -*- coding: utf-8 -*-
"""docweave 命令行入口。"""
from __future__ import annotations

import argparse
import glob
import json
import os
import subprocess
import sys

from . import audit as A
from . import ops
from . import writers as W
from .model import read_document


def _p(msg=""):
    print(msg)


# --------------------------------------------------------------------------
# 子命令
# --------------------------------------------------------------------------

def cmd_merge(args) -> int:
    paths = args.paths
    if len(paths) not in (2, 3):
        _p("用法：docweave merge [基线] 我方 对方 -o 输出.docx")
        return 2
    base = paths[0] if len(paths) == 3 else None
    ours, theirs = (paths[1], paths[2]) if len(paths) == 3 else (paths[0], paths[1])
    out = args.out
    if not out:
        stem = os.path.splitext(os.path.abspath(theirs))[0]
        out = stem + "_merged" + os.path.splitext(ours)[1]
    os.makedirs(os.path.dirname(os.path.abspath(out)) or ".", exist_ok=True)
    r = ops.run_merge(base, ours, theirs, out,
                      report=args.report,
                      on_conflict=args.on_conflict,
                      table_key=(None if args.table_key.lower() == "none"
                                 else int(args.table_key)),
                      with_appendix=not args.no_appendix,
                      prefer_template=args.template,
                      sidecars=args.json)
    _p(f"合并模式：{'三方（基线+我方+对方）' if base else '两方对比整合'}")
    _p(f"输出：{out}")
    _p("统计：" + json.dumps(r["stats"], ensure_ascii=False))
    if r["conflicts"]:
        from .labels import CONFLICT, zh

        _p(f"冲突 {r['conflicts']} 处：")
        for c in r["conflict_list"][:20]:
            _p(f"  {c['id']} [{zh(CONFLICT, c['kind'])}] {c.get('ours')} ⚔ {c.get('theirs')}")
        if r["conflicts"] > 20:
            _p(f"  …… 其余 {r['conflicts'] - 20} 处见报告")
    else:
        _p("未发现冲突。")
    if args.report:
        _p(f"冲突报告：{args.report}")
    if args.json:
        _p(f"机读结果：{args.json}")
    return 0


def cmd_audit(args) -> int:
    out = args.out
    if not out and not args.report:
        out = os.path.splitext(args.data)[0] + "_audit.docx"
    r = ops.run_audit(args.doc, args.data, out=out, report=args.report,
                      tol=args.tol, embed=not args.no_embed, recalc=args.recalc,
                      sidecars=args.json, show_ok=args.show_ok)
    _p("对账统计：" + json.dumps(r["stats"], ensure_ascii=False))
    for f in r["findings"]:
        _p(f"  [{f['severity']}] {f['message']}")
    if out:
        _p(f"对账报告（docx，已把正文与数据表整合）：{out}")
    if args.report:
        _p(f"对账报告（Markdown）：{args.report}")
    return 0 if not any(f["code"] == "value_conflict" for f in r["findings"]) else 1


def cmd_refs(args) -> int:
    out = args.out
    if not out and not args.report:
        out = os.path.splitext(args.doc)[0] + "_refs.docx"
    r = ops.run_refs(args.doc, out=out, report=args.report, sidecars=args.json)
    _p("体检统计：" + json.dumps(r["stats"], ensure_ascii=False))
    for f in r["findings"]:
        _p(f"  [{f['severity']}] {f['message']}")
    if out:
        _p(f"体检报告：{out}")
    if args.report:
        _p(f"体检报告（Markdown）：{args.report}")
    return 0 if not any(f["severity"] == A.SEV_HIGH for f in r["findings"]) else 1


def cmd_drift(args) -> int:
    files = []
    for pat in args.files:
        if os.path.isdir(pat):
            files.extend(sorted(glob.glob(os.path.join(pat, "*.docx"))))
        else:
            files.extend(sorted(glob.glob(pat)) or [pat])
    files = [f for f in files if os.path.exists(f)]
    if len(files) < 1:
        _p("未找到可用文件。")
        return 2
    if len(files) == 1:
        _p("只给了 1 个文件：将只检查该文件内部是否存在自相矛盾。")
    out = args.out
    if not out and not args.report:
        out = os.path.join(os.path.dirname(os.path.abspath(files[0])),
                          "drift_report.docx")
    r = ops.run_drift(files, out=out, report=args.report, tol=args.tol,
                      sidecars=args.json)
    _p("漂移统计：" + json.dumps(r["stats"], ensure_ascii=False))
    for f in r["findings"]:
        _p(f"  [{f['severity']}] {f['message']}")
    if out:
        _p(f"漂移报告：{out}")
    if args.report:
        _p(f"漂移报告（Markdown）：{args.report}")
    return 0


def cmd_batch(args) -> int:
    r = ops.run_batch(args.manifest, outdir=args.outdir, force=args.force,
                      ledger_root=args.ledger_root, quiet=args.quiet)
    _p("")
    _p("汇总：" + json.dumps(r["summary"], ensure_ascii=False))
    _p(f"HTML 报告：{r['html']}")
    _p(f"机读汇总：{r['json']}")
    _p(f"增量台账：{r['ledger']}")
    return 0 if r["summary"]["失败"] == 0 else 1


def cmd_process_dir(args) -> int:
    r = ops.run_folder_incremental(args.folder, outdir=args.outdir,
                                   force=args.force, quiet=args.quiet)
    _p("增量处理汇总：" + json.dumps(r["summary"], ensure_ascii=False))
    _p(f"自动清单：{r['manifest']}")
    _p(f"HTML 报告：{r['html']}")
    _p(f"机读汇总：{r['json']}")
    _p(f"增量台账：{r['ledger']}")
    return 0 if r["summary"]["失败"] == 0 else 1


def cmd_ledger(args) -> int:
    led = ops.Ledger(args.root)
    if args.reset:
        if os.path.exists(led.path):
            os.remove(led.path)
            _p(f"已清除台账：{led.path}")
        else:
            _p("台账不存在，无需清除。")
        return 0
    _p(led.summary())
    return 0


def cmd_recalc(args) -> int:
    out = ops.recalc_workbook(args.file, outdir=args.outdir)
    _p(f"已重算公式：{out}")
    return 0


def cmd_plan(args) -> int:
    r = ops.build_manifest(args.folder, outdir=getattr(args, "outdir", None),
                           out_path=args.out)
    _p(f"已生成清单：{r['path']}（{len(r['jobs'])} 个任务）")
    for j in r["jobs"]:
        _p(f"  · {j.get('type')} {j.get('name', '')}")
    _p(f"运行：docweave batch \"{r['path']}\"")
    return 0


def cmd_demo(args) -> int:
    here = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    fixtures = os.path.join(here, "demo", "fixtures")
    out_root = os.path.join(here, "demo", "out")
    os.makedirs(out_root, exist_ok=True)
    maker = os.path.join(here, "demo", "make_fixtures.py")
    if not os.path.exists(os.path.join(fixtures, "contract_base.docx")):
        _p("正在生成演示样例 ……")
        subprocess.run([sys.executable, maker], check=True)
    man = os.path.join(out_root, "manifest.json")
    cmd_plan(argparse.Namespace(folder=fixtures, out=man, outdir=out_root))
    _p("")
    ns2 = argparse.Namespace(manifest=man, outdir=None, force=True,
                             ledger_root=out_root, quiet=False)
    return cmd_batch(ns2)


# --------------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="docweave",
        description="docweave —— docx/xlsx 批量合并与一致性审计工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""示例：
  docweave merge 基线.docx 我方.docx 对方.docx -o 合并.docx --report 冲突.md
  docweave merge A.docx B.docx -o 对比整合.docx            # 没有基线时做两方整合
  docweave audit 合同.docx 指标表.xlsx -o 对账.docx --report 对账.md
  docweave refs 方案.docx -o 引用体检.docx
  docweave drift spec_v1.docx spec_v2.docx spec_v3.docx -o 漂移.docx
  docweave plan ./文档 -o manifest.json && docweave batch manifest.json
""")
    p.add_argument("--version", action="version", version="docweave 1.0.0")
    sub = p.add_subparsers(dest="cmd", required=True)

    m = sub.add_parser("merge", help="合并两个/三个文档（两方对比整合 或 三方基线合并）")
    m.add_argument("paths", nargs="+", help="[基线] 我方 对方")
    m.add_argument("-o", "--out", help="输出文件（.docx 或 .xlsx）")
    m.add_argument("--report", help="Markdown 冲突报告路径")
    m.add_argument("--json", help="机读结果 JSON 路径")
    m.add_argument("--on-conflict", choices=["mark", "ours", "theirs"], default="mark",
                   help="冲突处置：mark=保留双方并标记（默认）")
    m.add_argument("--table-key", default="0",
                   help="表格行对齐用第几列作主键，none=不做单元格级合并")
    m.add_argument("--template", choices=["ours", "base", "theirs"],
                   help="docx 版式模板来源，默认优先用我方")
    m.add_argument("--no-appendix", action="store_true", help="不输出溯源附录")
    m.set_defaults(func=cmd_merge)

    a = sub.add_parser("audit", help="文表对账：正文数字/日期/编号 与 表格数据 核对")
    a.add_argument("doc", help="正文文档（docx/md/txt）")
    a.add_argument("data", help="数据文件（xlsx/xlsm/csv）")
    a.add_argument("-o", "--out", help="输出 docx 报告（把两边整合成一份）")
    a.add_argument("--report", help="Markdown 报告路径")
    a.add_argument("--json", help="机读结果 JSON 路径")
    a.add_argument("--tol", type=float, default=0.005, help="数值容差（默认 0.5%%）")
    a.add_argument("--show-ok", action="store_true", help="同时列出核对一致的项目")
    a.add_argument("--no-embed", action="store_true", help="报告里不附数据表原文")
    a.add_argument("--recalc", action="store_true",
                   help="先用内置 LibreOffice 重算 xlsx 公式再对账")
    a.set_defaults(func=cmd_audit)

    r = sub.add_parser("refs", help="编号与交叉引用体检（悬空/断档/重复）")
    r.add_argument("doc")
    r.add_argument("-o", "--out", help="输出 docx 报告")
    r.add_argument("--report", help="Markdown 报告路径")
    r.add_argument("--json", help="机读结果 JSON 路径")
    r.set_defaults(func=cmd_refs)

    d = sub.add_parser("drift", help="口径漂移：跨版本变化 + 文档内自相矛盾")
    d.add_argument("files", nargs="+", help="多个版本文件，或一个目录")
    d.add_argument("-o", "--out", help="输出 docx 报告")
    d.add_argument("--report", help="Markdown 报告路径")
    d.add_argument("--json", help="机读结果 JSON 路径")
    d.add_argument("--tol", type=float, default=0.005)
    d.set_defaults(func=cmd_drift)

    b = sub.add_parser("batch", help="按清单批量处理（带增量台账）")
    b.add_argument("manifest", help="清单 JSON")
    b.add_argument("--outdir", help="覆盖输出目录")
    b.add_argument("--force", action="store_true", help="忽略台账，全部重跑")
    b.add_argument("--ledger-root", help="台账存放根目录")
    b.add_argument("--quiet", action="store_true")
    b.set_defaults(func=cmd_batch)

    pd = sub.add_parser("process-dir", help="扫描目录并仅增量处理变化任务")
    pd.add_argument("folder", help="待处理目录")
    pd.add_argument("--outdir", help="产物目录（默认 <目录>/out）")
    pd.add_argument("--force", action="store_true", help="忽略台账，强制重跑所有任务")
    pd.add_argument("--quiet", action="store_true", help="不输出逐任务进度")
    pd.set_defaults(func=cmd_process_dir)

    l = sub.add_parser("ledger", help="查看/清除增量台账")
    l.add_argument("--root", default=".")
    l.add_argument("--reset", action="store_true")
    l.set_defaults(func=cmd_ledger)

    pl = sub.add_parser("plan", help="扫描目录自动生成批量清单")
    pl.add_argument("folder")
    pl.add_argument("-o", "--out", help="清单输出路径")
    pl.add_argument("--outdir", help="产物输出根目录（默认 <目录>/out）")
    pl.set_defaults(func=cmd_plan)

    rc = sub.add_parser("recalc", help="用内置 LibreOffice 重算 xlsx 公式缓存")
    rc.add_argument("file")
    rc.add_argument("--outdir")
    rc.set_defaults(func=cmd_recalc)

    dm = sub.add_parser("demo", help="一键跑通演示（自动生成样例并批量处理）")
    dm.set_defaults(func=cmd_demo)
    return p


def main(argv=None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        return args.func(args)
    except (ValueError, RuntimeError, FileNotFoundError) as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
