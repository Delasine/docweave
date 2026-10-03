# -*- coding: utf-8 -*-
"""docweave.ops —— 批量编排、增量台账、HTML 报告、公式重算。

批量处理最怕「跑一遍就得重跑一遍」。这里给每个任务算输入指纹，
写进 .docweave/ledger.json：输入没变、产物还在，就直接跳过；
每次运行的结论也留在 history 里，便于回溯「上周那批到底改了什么」。
"""
from __future__ import annotations

import datetime as _dt
import glob as _glob
import hashlib
import html
import json
import os
import shutil
import subprocess
import tempfile
import uuid
from typing import Any, Optional

from . import audit as A
from . import merge as M
from . import writers as W
from .model import read_document

LEDGER_DIR = ".docweave"
LEDGER_NAME = "ledger.json"


# --------------------------------------------------------------------------
# 台账
# --------------------------------------------------------------------------

def sha256_file(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


class Ledger:
    def __init__(self, root: str = "."):
        self.dir = os.path.join(root, LEDGER_DIR)
        self.path = os.path.join(self.dir, LEDGER_NAME)
        self.data = {"version": 1, "entries": {}, "history": []}
        if os.path.exists(self.path):
            try:
                # utf-8-sig：容忍手工编辑过的台账带 BOM
                with open(self.path, encoding="utf-8-sig") as f:
                    loaded = json.load(f)
                if isinstance(loaded, dict):
                    self.data.update(loaded)
            except (json.JSONDecodeError, OSError):
                pass

    def task_key(self, kind: str, inputs: list, out: str) -> str:
        raw = "|".join([kind] + sorted(os.path.abspath(p) for p in inputs)
                       + [os.path.abspath(out)])
        return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:16]

    def fingerprint(self, inputs: list) -> str:
        parts = []
        for p in sorted(inputs):
            try:
                parts.append(f"{p}:{sha256_file(p)}")
            except OSError:
                parts.append(f"{p}:MISSING")
        return hashlib.sha1("|".join(parts).encode("utf-8")).hexdigest()

    def is_fresh(self, key: str, inputs: list, out: str) -> bool:
        e = self.data["entries"].get(key)
        if not e:
            return False
        if not out or not os.path.exists(out):
            return False
        return e.get("fingerprint") == self.fingerprint(inputs)

    def record(self, key: str, kind: str, inputs: list, out: str,
               stats: dict, status: str = "ok", note: str = ""):
        self.data["entries"][key] = {
            "kind": kind, "inputs": inputs, "output": out,
            "fingerprint": self.fingerprint(inputs),
            "at": _dt.datetime.now().isoformat(timespec="seconds"),
            "stats": stats, "status": status,
        }
        self.data["history"].append({
            "at": _dt.datetime.now().isoformat(timespec="seconds"),
            "kind": kind, "output": out, "status": status, "note": note,
        })
        self.data["history"] = self.data["history"][-500:]

    def save(self):
        os.makedirs(self.dir, exist_ok=True)
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(self.data, f, ensure_ascii=False, indent=1)

    def summary(self) -> str:
        e = self.data.get("entries", {})
        lines = [f"台账文件：{self.path}", f"已记录任务：{len(e)} 个"]
        for k, v in list(e.items())[-20:]:
            lines.append(f"  · [{v.get('kind')}] {v.get('status')} "
                         f"{os.path.basename(v.get('output') or '')} @ {v.get('at')}")
        h = self.data.get("history", [])
        if h:
            lines.append(f"最近运行：{h[-1]['at']} {h[-1]['kind']} {h[-1]['status']}")
        return "\n".join(lines)


# --------------------------------------------------------------------------
# LibreOffice 公式重算（可选）
# --------------------------------------------------------------------------

# --------------------------------------------------------------------------
# 目录扫描与增量处理（供 GUI / CLI 调用）
# --------------------------------------------------------------------------

def scan_and_log(folder: str, log: Optional[str] = None, pattern: str = "*",
                 recursive: bool = False, changes_only: bool = False) -> dict:
    """扫描目录并可选地把结果追加写入 JSONL 日志。

    返回字典：{"scanned": n, "added": [], "removed": [], "changed": [], "log": logfile}
    比较只使用同一目录、匹配模式和递归设置下最近一次完成的扫描。
    """
    folder = os.path.abspath(folder)
    if not os.path.isdir(folder):
        raise FileNotFoundError(folder)

    log_path = os.path.abspath(log) if log else None

    def raise_walk_error(error):
        raise error

    def iter_files_local():
        if recursive:
            for root, dirs, files in os.walk(folder, onerror=raise_walk_error):
                dirs[:] = [d for d in dirs if d != LEDGER_DIR]
                for name in files:
                    if _glob.fnmatch.fnmatch(name, pattern):
                        path = os.path.abspath(os.path.join(root, name))
                        if path != log_path:
                            yield path
        else:
            for name in os.listdir(folder):
                path = os.path.abspath(os.path.join(folder, name))
                if (path != log_path and os.path.isfile(path)
                        and _glob.fnmatch.fnmatch(name, pattern)):
                    yield path

    cur = {}
    for p in iter_files_local():
        rel = os.path.relpath(p, folder)
        try:
            st = os.stat(p)
            sha = sha256_file(p)
            cur[rel] = {"path": p, "rel_path": rel, "size": st.st_size,
                        "mtime_ns": st.st_mtime_ns, "sha256": sha}
        except OSError as e:
            raise OSError(f"无法读取扫描文件 {p}: {e}") from e

    previous = None
    if log_path and os.path.exists(log_path):
        previous = _load_scan_snapshot(log_path, folder, pattern, recursive)

    added, removed, changed = [], [], []
    if changes_only:
        prev_files = previous or {}
        prev_keys = set(prev_files)
        cur_keys = set(cur.keys())
        for k in sorted(cur_keys - prev_keys):
            added.append(k)
        for k in sorted(prev_keys - cur_keys):
            removed.append(k)
        for k in sorted(cur_keys & prev_keys):
            a = prev_files[k]
            b = cur[k]
            if a.get("sha256") != b.get("sha256"):
                changed.append(k)

    if log_path:
        os.makedirs(os.path.dirname(log_path), exist_ok=True)
        scan_id = uuid.uuid4().hex
        scanned_at = _dt.datetime.now(_dt.timezone.utc).isoformat()
        metadata = {"scan_id": scan_id, "folder": folder, "pattern": pattern,
                    "recursive": recursive, "at": scanned_at}
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(json.dumps({"event": "scan_start", **metadata},
                               ensure_ascii=False) + "\n")
            for rel, meta in sorted(cur.items()):
                f.write(json.dumps({"event": "file", "scan_id": scan_id, **meta},
                                   ensure_ascii=False) + "\n")
            f.write(json.dumps({"event": "scan_complete", **metadata,
                                "scanned": len(cur)}, ensure_ascii=False) + "\n")

    return {"scanned": len(cur), "added": added, "removed": removed,
            "changed": changed, "log": log_path}


def _load_scan_snapshot(log_path: str, folder: str, pattern: str,
                        recursive: bool) -> Optional[dict]:
    """Load the newest complete matching snapshot; ignore only a torn final line."""
    snapshots = {}
    latest = None
    with open(log_path, encoding="utf-8") as f:
        lines = f.readlines()
    for index, line in enumerate(lines):
        try:
            record = json.loads(line)
        except json.JSONDecodeError as e:
            if index == len(lines) - 1 and not line.endswith("\n"):
                break
            raise ValueError(f"扫描日志格式错误（第 {index + 1} 行）：{e}") from e

        event = record.get("event")
        scan_id = record.get("scan_id")
        if event == "scan_start":
            snapshots[scan_id] = {"metadata": record, "files": {}}
        elif event == "file" and scan_id in snapshots:
            snapshots[scan_id]["files"][record["rel_path"]] = record
        elif event == "scan_complete" and scan_id in snapshots:
            metadata = snapshots[scan_id]["metadata"]
            if (metadata.get("folder") == folder
                    and metadata.get("pattern") == pattern
                    and metadata.get("recursive") == recursive):
                latest = snapshots[scan_id]["files"]
            del snapshots[scan_id]

    # Read legacy per-file JSONL entries as a one-time baseline.
    if latest is None:
        legacy = {}
        for line in lines:
            try:
                record = json.loads(line)
            except json.JSONDecodeError:
                continue
            if record.get("event") is not None:
                continue
            path = record.get("path")
            rel = record.get("rel_path")
            if path and rel:
                try:
                    same_folder = (os.path.commonpath([
                        os.path.normcase(folder),
                        os.path.normcase(os.path.abspath(path)),
                    ]) == os.path.normcase(folder))
                except ValueError:
                    same_folder = False
                if same_folder:
                    legacy[rel] = record
        if legacy:
            latest = legacy
    return latest


LO_NODE = os.environ.get(
    "DSH_LO_NODE",
    r"D:\deepseek\resources\runtime\primary-runtime\dependencies\node\bin\node.exe")
LO_CLI = os.environ.get(
    "DSH_LO_CLI",
    r"D:\deepseek\resources\app.asar.unpacked\dsh\node_modules\@deepseek-ai"
    r"\libreoffice-kit\lib\cli.js")


def recalc_workbook(path: str, outdir: Optional[str] = None) -> str:
    """用内置 LibreOffice 重算公式，返回带最新缓存值的副本路径。

    Excel 文件里的公式结果可能是过期缓存（甚至从未计算过），
    对账/合并前先重算一次，可以避免把过期数字当成事实。
    """
    if not (os.path.exists(LO_NODE) and os.path.exists(LO_CLI)):
        raise RuntimeError(
            "未找到内置 LibreOffice 套件，无法重算公式；"
            "可设置 DSH_LO_NODE / DSH_LO_CLI 环境变量。")
    outdir = outdir or tempfile.mkdtemp(prefix="docweave_recalc_")
    os.makedirs(outdir, exist_ok=True)
    out = os.path.join(outdir, os.path.basename(path))
    cmd = [LO_NODE, LO_CLI, "recalculate", "--input", path, "--output", out]
    proc = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
    if proc.returncode != 0 or not os.path.exists(out):
        raise RuntimeError(f"公式重算失败：{proc.stdout[-500:]} {proc.stderr[-500:]}")
    return out


# --------------------------------------------------------------------------
# 单个任务
# --------------------------------------------------------------------------

def _read_doc(path: str):
    """按扩展名选择合适的读取选项：xlsx 需要额外读出公式字符串，
    否则公式单元格（data_only 下值为空）会在合并时丢失。"""
    if path.lower().endswith((".xlsx", ".xlsm")):
        return read_document(path, with_formulas=True)
    return read_document(path)


def run_merge(base: Optional[str], ours: str, theirs: str, out: str,
              report: Optional[str] = None, on_conflict: str = "mark",
              table_key: Optional[int] = 0, with_appendix: bool = True,
              prefer_template: Optional[str] = None,
              sidecars: Optional[str] = None) -> dict:
    docs = {}
    if base:
        docs["base"] = _read_doc(base)
    docs["ours"] = _read_doc(ours)
    docs["theirs"] = _read_doc(theirs)
    if base:
        res = M.three_way(docs["base"], docs["ours"], docs["theirs"],
                          on_conflict=on_conflict, table_key=table_key)
    else:
        res = M.two_way(docs["ours"], docs["theirs"],
                        on_conflict=on_conflict, table_key=table_key)

    fmts = {d.fmt for d in docs.values()}
    if "docx" in fmts:
        info = W.write_docx_merge(res, out, with_appendix=with_appendix,
                                  template=prefer_template)
    else:
        info = W.write_xlsx_merge(res, out)
    if report:
        W.write_markdown_report(res, report)
    if sidecars:
        with open(sidecars, "w", encoding="utf-8") as f:
            json.dump({
                "mode": res.mode, "stats": res.stats,
                "conflicts": [{"id": c.cid, "kind": c.kind, "path": c.path,
                               "ours": c.ours.preview(120) if c.ours else None,
                               "theirs": c.theirs.preview(120) if c.theirs else None}
                              for c in res.conflicts],
                "moves": res.moves, "warnings": res.warnings,
                "provenance": res.provenance_rows(),
            }, f, ensure_ascii=False, indent=1)
    return {"kind": "merge", "mode": res.mode, "stats": res.stats,
            "conflicts": len(res.conflicts), "output": out,
            "report": report, "write_info": info,
            "conflict_list": [
                {"id": c.cid, "kind": c.kind, "path": c.path,
                 "ours": c.ours.preview(80) if c.ours else None,
                 "theirs": c.theirs.preview(80) if c.theirs else None}
                for c in res.conflicts]}


def run_audit(doc: str, data: str, out: Optional[str] = None,
              report: Optional[str] = None, tol: float = 0.005,
              embed: bool = True, recalc: bool = False,
              sidecars: Optional[str] = None, show_ok: bool = False) -> dict:
    data_path = data
    if recalc:
        data_path = recalc_workbook(data)
    d = read_document(doc)
    x = _read_doc(data_path)
    rep = A.audit_doc_against_data(d, x, tol=tol, show_ok=show_ok)
    rep["data_source"] = data
    if recalc:
        rep["recalculated"] = True
    if out:
        A.write_audit_docx(rep, out, embed_data=embed)
    if report:
        with open(report, "w", encoding="utf-8-sig", newline="\r\n") as f:
            f.write(A.audit_markdown(rep))
    if sidecars:
        with open(sidecars, "w", encoding="utf-8") as f:
            json.dump({"stats": rep["stats"],
                       "findings": [{k: v for k, v in fd.items()
                                     if k in ("code", "severity", "label", "message")}
                                    for fd in rep["findings"]]},
                      f, ensure_ascii=False, indent=1)
    return {"kind": "audit", "stats": rep["stats"], "output": out,
            "report": report,
            "findings": [{"code": fd["code"], "severity": fd["severity"],
                          "label": fd["label"], "message": fd["message"]}
                         for fd in rep["findings"]]}


def run_refs(doc: str, out: Optional[str] = None, report: Optional[str] = None,
             sidecars: Optional[str] = None) -> dict:
    rep = A.check_refs(read_document(doc))
    if out:
        A.write_refs_docx(rep, out)
    if report:
        with open(report, "w", encoding="utf-8-sig", newline="\r\n") as f:
            f.write(A.refs_markdown(rep))
    if sidecars:
        with open(sidecars, "w", encoding="utf-8") as f:
            json.dump({"stats": rep["stats"], "findings": rep["findings"],
                       "defined": rep["defined"]}, f, ensure_ascii=False, indent=1)
    return {"kind": "refs", "stats": rep["stats"], "output": out, "report": report,
            "findings": [{"code": fd["code"], "severity": fd["severity"],
                          "label": fd.get("locator", ""), "message": fd["message"]}
                         for fd in rep["findings"]]}


def run_drift(files: list, out: Optional[str] = None,
              report: Optional[str] = None, tol: float = 0.005,
              sidecars: Optional[str] = None) -> dict:
    rep = A.drift([read_document(p) for p in files], tol=tol)
    if out:
        A.write_drift_docx(rep, out)
    if report:
        with open(report, "w", encoding="utf-8-sig", newline="\r\n") as f:
            f.write(A.drift_markdown(rep))
    if sidecars:
        with open(sidecars, "w", encoding="utf-8") as f:
            json.dump(rep, f, ensure_ascii=False, indent=1)
    return {"kind": "drift", "stats": rep["stats"], "output": out, "report": report,
            "findings": ([{"code": "intra_conflict", "severity": A.SEV_HIGH,
                           "label": c["label"], "message": c["message"]}
                          for c in rep["intra"]] +
                         [{"code": "version_drift", "severity": A.SEV_MID,
                           "label": c["label"], "message": c["message"]}
                          for c in rep["changed"]])}


# --------------------------------------------------------------------------
# 批量清单
# --------------------------------------------------------------------------

def _iter_dir(folder: str, pattern: str = "*") -> list:
    if pattern in ("*", "*.*"):
        pats = ["*.docx", "*.xlsx", "*.csv", "*.md", "*.txt"]
    else:
        pats = [pattern]
    out = []
    for p in pats:
        out.extend(_glob.glob(os.path.join(folder, p)))
    return sorted(f for f in out if not os.path.basename(f).startswith("~$"))


def expand_jobs(manifest: dict, base_dir: str) -> list:
    """把清单里的目录型任务展开成单文件任务。"""
    jobs = []
    for job in manifest.get("jobs", []):
        jtype = (job.get("type") or "merge").lower()
        if jtype == "merge":
            jobs.append(job)
        elif jtype == "merge-dir":
            folder = os.path.join(base_dir, job["dir"])
            suf = job.get("suffixes") or {"base": "_base", "ours": "_ours",
                                          "theirs": "_theirs"}
            outdir = os.path.join(base_dir, job.get("outdir", "merged"))
            groups: dict = {}
            for f in _iter_dir(folder, job.get("pattern", "*.docx")):
                stem = os.path.splitext(os.path.basename(f))[0]
                for role in ("base", "ours", "theirs"):
                    s = suf.get(role)
                    if s and stem.endswith(s):
                        groups.setdefault(stem[: -len(s)], {})[role] = f
            for name, g in sorted(groups.items()):
                if "ours" not in g or "theirs" not in g:
                    continue
                srcs = [g.get("base"), g["ours"], g["theirs"]]
                ext = ".xlsx" if all(
                    (s or "").lower().endswith((".xlsx", ".xlsm", ".csv", ".tsv"))
                    for s in srcs if s) else ".docx"
                jobs.append({
                    "type": "merge", "name": "合并-" + name,
                    "base": g.get("base"), "ours": g["ours"], "theirs": g["theirs"],
                    "out": os.path.join(outdir, f"{name}_merged{ext}"),
                    "on_conflict": job.get("on_conflict", "mark"),
                    "table_key": job.get("table_key", 0),
                })
        elif jtype in ("refs-dir", "refs_dir"):
            folder = os.path.join(base_dir, job["dir"])
            outdir = os.path.join(base_dir, job.get("outdir", "refs"))
            for f in _iter_dir(folder, job.get("pattern", "*.docx")):
                stem = os.path.splitext(os.path.basename(f))[0]
                jobs.append({"type": "refs", "doc": f, "name": "引用体检-" + stem,
                             "out": os.path.join(outdir, stem + "_refs.docx")})
        elif jtype == "drift-dir":
            folder = os.path.join(base_dir, job["dir"])
            outdir = os.path.join(base_dir, job.get("outdir", "drift"))
            files = _iter_dir(folder, job.get("pattern", "*.docx"))
            if files:
                jobs.append({"type": "drift", "files": files,
                             "out": os.path.join(outdir, "drift_report.docx")})
        else:
            jobs.append(job)
    return jobs


def build_manifest(folder: str, outdir: Optional[str] = None,
                   out_path: Optional[str] = None) -> dict:
    """扫描目录，自动发现可批量处理的任务并生成清单。

    识别三类模式：
      * `*_base/_ours/_theirs` 组合（docx 或 xlsx）→ 三方合并
      * 共享文件名词根的 `正文 + 数据表` → 文表对账
      * `*_v1/_v2/...` 版本序列 → 口径漂移
    """
    folder = os.path.abspath(folder)
    files = sorted(f for f in _glob.glob(os.path.join(folder, "*"))
                   if not os.path.basename(f).startswith("~$"))
    docx = [f for f in files if f.lower().endswith(".docx")]
    sheets = [f for f in files if f.lower().endswith((".xlsx", ".xlsm", ".csv"))]
    outdir = os.path.abspath(outdir) if outdir else os.path.join(folder, "out")
    roles = {"_base": "base", "_ours": "ours", "_theirs": "theirs"}

    def group(paths):
        g: dict = {}
        for f in paths:
            stem = os.path.splitext(os.path.basename(f))[0]
            for suf, role in roles.items():
                if stem.endswith(suf):
                    g.setdefault(stem[: -len(suf)], {})[role] = f
        return g

    jobs: list = []
    if group(docx):
        jobs.append({"type": "merge-dir", "dir": folder,
                     "suffixes": {"base": "_base", "ours": "_ours", "theirs": "_theirs"},
                     "outdir": os.path.join(outdir, "merged"), "on_conflict": "mark"})
    if group(sheets):
        jobs.append({"type": "merge-dir", "dir": folder, "pattern": "*.xlsx",
                     "suffixes": {"base": "_base", "ours": "_ours", "theirs": "_theirs"},
                     "outdir": os.path.join(outdir, "merged_sheets")})

    def common_prefix(a, b):
        n = 0
        for x, y in zip(a, b):
            if x != y:
                break
            n += 1
        return n

    used = set()
    for x in sheets:
        xs = os.path.splitext(os.path.basename(x))[0]
        best, score = None, 0
        for d in docx:
            if d in used:
                continue
            ds = os.path.splitext(os.path.basename(d))[0]
            n = common_prefix(xs, ds)
            if n > score:
                best, score = d, n
        if best and score >= 3:
            used.add(best)
            stem = os.path.splitext(os.path.basename(best))[0]
            jobs.append({"type": "audit", "name": "文表对账-" + stem, "doc": best,
                         "data": x,
                         "out": os.path.join(outdir, "audit", stem + "_audit.docx")})
    if docx:
        jobs.append({"type": "refs-dir", "dir": folder, "pattern": "*.docx",
                     "outdir": os.path.join(outdir, "refs")})
    vers = [d for d in docx if "_v" in os.path.basename(d).lower()]
    if len(vers) >= 2:
        jobs.append({"type": "drift", "name": "版本口径漂移", "files": vers,
                     "out": os.path.join(outdir, "drift", "drift_report.docx")})

    manifest = {"options": {"tol": 0.005, "outdir": outdir}, "jobs": jobs}
    out_path = out_path or os.path.join(outdir, "manifest.json")
    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=1)
    return {"path": out_path, "manifest": manifest, "jobs": jobs, "folder": folder,
            "outdir": outdir}


def run_folder_incremental(folder: str, outdir: Optional[str] = None,
                           force: bool = False, quiet: bool = False,
                           log=None) -> dict:
    """Scan a folder, plan supported tasks, and run only stale tasks."""
    folder = os.path.abspath(folder)
    if not os.path.isdir(folder):
        raise FileNotFoundError(folder)
    outdir = os.path.abspath(outdir) if outdir else os.path.join(folder, "out")
    state_dir = os.path.join(outdir, LEDGER_DIR)
    scan_log = os.path.join(state_dir, "scan_log.jsonl")
    manifest_path = os.path.join(state_dir, "manifest.json")

    scan = scan_and_log(folder, log=scan_log, pattern="*", recursive=False,
                        changes_only=True)
    plan = build_manifest(folder, outdir=outdir, out_path=manifest_path)
    if not plan["jobs"]:
        raise ValueError(
            "目录中没有识别到可处理任务。请放入支持的 Word/Excel/CSV 文件，"
            "或 *_base、*_ours、*_theirs / *_v1、*_v2 文件组合。")

    scan_changes = {
        "scanned": scan["scanned"],
        "added": scan["added"],
        "removed": scan["removed"],
        "changed": scan["changed"],
    }
    if log:
        log(f"扫描到 {scan['scanned']} 个文件：新增 {len(scan['added'])}，"
            f"删除 {len(scan['removed'])}，内容变化 {len(scan['changed'])}")
        for label, key in (("新增", "added"), ("删除", "removed"),
                           ("内容变化", "changed")):
            for path in scan[key]:
                log(f"  {label}：{path}")
    result = run_batch(
        manifest_path, outdir=outdir, force=force, ledger_root=outdir,
        quiet=quiet, log=log, scan_changes=scan_changes)
    result.update({"scan": scan, "manifest": manifest_path, "planned": len(plan["jobs"])})
    return result


def run_batch(manifest_path: str, outdir: Optional[str] = None, force: bool = False,
              ledger_root: Optional[str] = None, quiet: bool = False,
              log=None, scan_changes: Optional[dict] = None) -> dict:
    def say(msg):
        if log:
            log(msg)
        elif not quiet:
            print(msg)

    with open(manifest_path, encoding="utf-8") as f:
        manifest = json.load(f)
    base_dir = os.path.dirname(os.path.abspath(manifest_path))
    opts = manifest.get("options", {}) or {}
    tol = float(opts.get("tol", 0.005))
    ledger = Ledger(ledger_root or base_dir)
    jobs = expand_jobs(manifest, base_dir)
    outdir = outdir or os.path.join(base_dir, opts.get("outdir", "out"))
    os.makedirs(outdir, exist_ok=True)

    results = []
    for i, job in enumerate(jobs, start=1):
        jtype = (job.get("type") or "merge").lower()
        name = job.get("name") or f"{jtype}-{i}"
        try:
            if jtype == "merge":
                inputs = [p for p in (job.get("base"), job.get("ours"),
                                      job.get("theirs")) if p]
                out = job.get("out") or os.path.join(outdir, f"{name}_merged.docx")
                report = job.get("report")
                side = job.get("sidecars")
            elif jtype == "audit":
                inputs = [job["doc"], job["data"]]
                out = job.get("out") or os.path.join(outdir, f"{name}_audit.docx")
                report = job.get("report")
                side = job.get("sidecars")
            elif jtype == "refs":
                inputs = [job["doc"]]
                out = job.get("out") or os.path.join(outdir, f"{name}_refs.docx")
                report = job.get("report")
                side = job.get("sidecars")
            elif jtype == "drift":
                inputs = list(job["files"])
                out = job.get("out") or os.path.join(outdir, f"{name}_drift.docx")
                report = job.get("report")
                side = job.get("sidecars")
            else:
                raise ValueError(f"未知任务类型：{jtype}")

            os.makedirs(os.path.dirname(os.path.abspath(out)) or ".", exist_ok=True)
            key = ledger.task_key(jtype, inputs, out)
            tag = f"{name}@{jtype}"
            if not force and ledger.is_fresh(key, inputs, out):
                say(f"[{i}/{len(jobs)}] 跳过（无变化）：{tag}")
                ledger.record(key, jtype, inputs, out,
                              ledger.data["entries"].get(key, {}).get("stats", {}),
                              status="skip", note="输入未变化")
                results.append({"name": name, "kind": jtype, "status": "skip",
                                "output": out, "stats": {}, "findings": []})
                continue

            say(f"[{i}/{len(jobs)}] 处理：{tag}")
            if jtype == "merge":
                r = run_merge(job.get("base"), job["ours"], job["theirs"], out,
                              report=report,
                              on_conflict=job.get("on_conflict", "mark"),
                              table_key=job.get("table_key", 0),
                              with_appendix=job.get("appendix", True),
                              sidecars=side)
            elif jtype == "audit":
                r = run_audit(job["doc"], job["data"], out=out, report=report,
                              tol=float(job.get("tol", tol)),
                              embed=job.get("embed", True),
                              recalc=bool(job.get("recalc", False)),
                              sidecars=side)
            elif jtype == "refs":
                r = run_refs(job["doc"], out=out, report=report, sidecars=side)
            else:
                r = run_drift(job["files"], out=out, report=report,
                              tol=float(job.get("tol", tol)), sidecars=side)
            r.update({"name": name, "status": "ok"})
            ledger.record(key, jtype, inputs, out, r.get("stats", {}), status="ok")
            results.append(r)
        except Exception as exc:                       # noqa: BLE001
            say(f"[{i}/{len(jobs)}] 失败：{name} → {exc}")
            results.append({"name": name, "kind": jtype, "status": "fail",
                            "error": str(exc), "output": None, "stats": {},
                            "findings": []})

    ledger.save()
    summary = summarise(results)
    if scan_changes is not None:
        summary.update({
            "扫描文件数": scan_changes["scanned"],
            "新增文件": len(scan_changes["added"]),
            "删除文件": len(scan_changes["removed"]),
            "内容变化文件": len(scan_changes["changed"]),
        })
    html_path = os.path.join(outdir, "batch_report.html")
    write_batch_html(results, summary, html_path, manifest_path,
                     scan_changes=scan_changes)
    json_path = os.path.join(outdir, "batch_summary.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump({"generated": _dt.datetime.now().isoformat(timespec="seconds"),
                   "manifest": manifest_path, "summary": summary,
                   "results": results, "ledger": ledger.path,
                   "scan_changes": scan_changes},
                  f, ensure_ascii=False, indent=1)
    return {"summary": summary, "results": results, "html": html_path,
            "json": json_path, "ledger": ledger.path}


def summarise(results: list) -> dict:
    s = {"任务总数": len(results), "成功": 0, "跳过": 0, "失败": 0,
         "冲突总数": 0, "高严重度问题": 0}
    for r in results:
        st = r.get("status")
        if st == "ok":
            s["成功"] += 1
        elif st == "skip":
            s["跳过"] += 1
        else:
            s["失败"] += 1
        s["冲突总数"] += int(r.get("conflicts") or 0)
        for f in r.get("findings") or []:
            if f.get("severity") == A.SEV_HIGH:
                s["高严重度问题"] += 1
    return s


# --------------------------------------------------------------------------
# HTML 报告
# --------------------------------------------------------------------------

_CSS = """
body{font-family:"Microsoft YaHei",-apple-system,Segoe UI,sans-serif;margin:0;
background:#f5f6f8;color:#1f2329}
.wrap{max-width:1180px;margin:0 auto;padding:28px 20px 60px}
h1{font-size:24px;margin:0 0 6px}
h2{font-size:17px;margin:26px 0 10px;padding-left:9px;border-left:4px solid #2f6fd0}
.meta{color:#646a73;font-size:13px;margin-bottom:18px}
.cards{display:flex;flex-wrap:wrap;gap:12px;margin:14px 0 6px}
.card{background:#fff;border:1px solid #e3e6eb;border-radius:8px;padding:12px 16px;
min-width:132px}
.card .k{font-size:12px;color:#646a73}
.card .v{font-size:22px;font-weight:600;margin-top:2px}
table{width:100%;border-collapse:collapse;background:#fff;font-size:13px;
border:1px solid #e3e6eb;border-radius:8px;overflow:hidden}
th,td{border-bottom:1px solid #eef0f3;padding:8px 10px;text-align:left;
vertical-align:top}
th{background:#f0f3f8;font-weight:600}
tr:last-child td{border-bottom:none}
.sev-高{color:#c00000;font-weight:600}
.sev-中{color:#b26a00}
.sev-低{color:#4a5568}
.ok{color:#1f7a3d;font-weight:600}
.skip{color:#8a8f99}
.fail{color:#c00000;font-weight:600}
code{background:#f0f3f8;padding:1px 5px;border-radius:4px;font-size:12px}
details{background:#fff;border:1px solid #e3e6eb;border-radius:8px;margin:10px 0;
padding:10px 14px}
summary{cursor:pointer;font-weight:600}
pre{white-space:pre-wrap;font-size:12px;color:#3d444d}
.legend{font-size:12px;color:#646a73;margin-top:6px}
"""


def write_batch_html(results: list, summary: dict, out_path: str,
                     manifest_path: str = "", scan_changes: Optional[dict] = None) -> str:
    from .labels import FINDING, STATUS, TASK, zh

    e = html.escape
    H = ["<!doctype html><html lang='zh-CN'><head><meta charset='utf-8'>",
         "<title>docweave 批量处理报告</title>",
         f"<style>{_CSS}</style></head><body><div class='wrap'>",
         "<h1>docweave 批量处理报告</h1>",
         f"<div class='meta'>生成时间：{_dt.datetime.now():%Y-%m-%d %H:%M:%S}"
         f"　清单：{e(os.path.basename(manifest_path)) if manifest_path else '-'}</div>",
         "<div class='cards'>"]
    for k, v in summary.items():
        H.append(f"<div class='card'><div class='k'>{e(k)}</div>"
                 f"<div class='v'>{e(str(v))}</div></div>")
    H.append("</div>")

    if scan_changes is not None:
        H.append("<h2>目录变化</h2>")
        labels = (("新增", "added"), ("删除", "removed"), ("内容变化", "changed"))
        if any(scan_changes[key] for _, key in labels):
            H.append("<table><tr><th>变化</th><th>相对路径</th></tr>")
            for label, key in labels:
                for path in scan_changes[key]:
                    H.append(f"<tr><td>{e(label)}</td><td><code>{e(path)}</code></td></tr>")
            H.append("</table>")
        else:
            H.append("<p class='ok'>与上次扫描相比没有文件变化。</p>")

    H.append("<h2>任务概览</h2><table><tr><th>任务</th><th>类型</th><th>状态</th>"
             "<th>结论</th><th>产物</th></tr>")
    for r in results:
        st = r.get("status", "")
        cls = {"ok": "ok", "skip": "skip", "fail": "fail"}.get(st, "")
        if st == "fail":
            concl = e(r.get("error", ""))
        elif st == "skip":
            concl = "输入未变化，沿用上次结果"
        else:
            bits = []
            if r.get("conflicts"):
                bits.append(f"冲突 {r['conflicts']} 处")
            stats = r.get("stats") or {}
            for key in ("数值冲突", "悬空引用", "重复编号", "编号断档",
                        "文档内自相矛盾", "跨版本变化"):
                if stats.get(key):
                    bits.append(f"{key} {stats[key]}")
            concl = "；".join(bits) or "无异常"
        out = r.get("output") or ""
        H.append(f"<tr><td>{e(r.get('name', ''))}</td>"
                 f"<td>{e(zh(TASK, r.get('kind', '')))}</td>"
                 f"<td class='{cls}'>{e(zh(STATUS, st))}</td><td>{concl}</td>"
                 f"<td><code>{e(os.path.basename(out)) if out else '—'}</code></td></tr>")
    H.append("</table>")

    H.append("<h2>问题明细</h2>")
    any_finding = False
    for r in results:
        fs = r.get("findings") or []
        if not fs:
            continue
        any_finding = True
        H.append(f"<h3>{e(r.get('name', ''))} "
                 f"<span class='meta'>（{e(zh(TASK, r.get('kind', '')))}）</span></h3>")
        H.append("<table><tr><th>严重度</th><th>类型</th><th>项目</th><th>说明</th></tr>")
        for f in fs:
            H.append(f"<tr><td class='sev-{e(f.get('severity', ''))}'>"
                     f"{e(f.get('severity', ''))}</td>"
                     f"<td>{e(zh(FINDING, f.get('code'), '其他'))}</td>"
                     f"<td>{e(str(f.get('label', '')))}</td>"
                     f"<td>{e(f.get('message', ''))}</td></tr>")
        H.append("</table>")
    if not any_finding:
        H.append("<p class='ok'>所有任务均未发现问题。</p>")

    H.append("<h2>图例</h2><div class='legend'>"
             "合并冲突 = 同一处双方改成不同内容；高/中/低 = 建议处理优先级。<br>"
             "绿色单元格 = 该行/列取自对方；红色单元格 = 冲突，注释里写着双方取值。</div>")
    H.append("</div></body></html>")
    text = "\n".join(H)
    with open(out_path, "w", encoding="utf-8-sig", newline="\r\n") as f:
        f.write(text)
    return text
