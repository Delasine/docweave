# -*- coding: utf-8 -*-
"""docweave.merge —— 块级对齐 + 两方/三方语义合并。

与 Word 自带「比较/合并」的区别：
  1. 三方合并需要共同基线（base + 我方 + 对方），Word 只能两两比较，噪音大；
  2. 支持「移动检测」：整段被搬走位置不会被当成删除+新增（Word 不识别移动）；
  3. 表格支持到单元格级别的合并，格式尽量沿用原文；
  4. 每个输出来源都写进溯源附录，冲突有编号，可回查。
"""
from __future__ import annotations

import difflib
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from typing import Optional

from .model import Block, Doc, block_similarity, norm_text

DEFAULT_THR = 0.72
_DP_LIMIT = 200_000          # DP 单元格上限，超过则退化到贪心对齐


# --------------------------------------------------------------------------
# 对齐
# --------------------------------------------------------------------------

@dataclass
class Op:
    tag: str                 # equal | replace | delete | insert
    bi: list
    si: list
    b0: int = 0              # 该 op 在 base 中的起始位置（插入位置参照）


def _keyify(base: list, side: list):
    """把两侧都唯一出现的相同文本抽成锚点，其余位置用唯一哨兵替换。

    这样 SequenceMatcher 的 equal 区间一定是「可信锚点」，
    不会被重复内容（如多个空小节标题）带偏。
    """
    cb = Counter(b.norm for b in base if b.norm)
    cs = Counter(s.norm for s in side if s.norm)
    common = {k for k in cb if k in cs and cb[k] == 1 and cs[k] == 1}
    A = [b.norm if (b.norm in common) else f"\x00B{i}" for i, b in enumerate(base)]
    B = [s.norm if (s.norm in common) else f"\x00S{j}" for j, s in enumerate(side)]
    return A, B


def align(base: list, side: list, thr: float = DEFAULT_THR) -> list:
    ops: list = []
    A, B = _keyify(base, side)
    sm = difflib.SequenceMatcher(None, A, B, autojunk=False)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            ops.append(Op("equal", list(range(i1, i2)), list(range(j1, j2)), i1))
        else:
            ops.extend(_align_region(base, side, i1, i2, j1, j2, thr))
    return ops


def _emit_grouped(steps: list, base: list, side: list, b0: int) -> list:
    """把 traceback 的逐步结果合并成 op 区间。

    insert 的 b0 必须是「插入点之前的 base 游标」，
    否则文档中段的插入会被误放到区间开头。
    """
    ops: list = []
    cur = {"tag": None, "bi": [], "si": [], "b0": b0}
    cursor = b0

    def flush():
        if cur["tag"]:
            ops.append(Op(cur["tag"], list(cur["bi"]), list(cur["si"]), cur["b0"]))

    for kind, i, j in steps:
        if kind == "diag":
            tag = "equal" if base[i].norm == side[j].norm else "replace"
        elif kind == "up":
            tag = "delete"
        else:
            tag = "insert"
        if tag != cur["tag"]:
            flush()
            cur = {"tag": tag, "bi": [], "si": [], "b0": cursor}
        if kind in ("diag", "up"):
            cur["bi"].append(i)
        if kind in ("diag", "ins"):
            cur["si"].append(j)
        if kind in ("diag", "up"):
            cursor = i + 1
    flush()
    return ops


def _align_region(base: list, side: list, i1: int, i2: int, j1: int, j2: int,
                  thr: float) -> list:
    n, m = i2 - i1, j2 - j1
    if n == 0:
        return [Op("insert", [], list(range(j1, j2)), i1)] if m else []
    if m == 0:
        return [Op("delete", list(range(i1, i2)), [], i1)]
    if n * m > _DP_LIMIT:
        return _greedy_align(base, side, i1, i2, j1, j2, thr)

    NEG = -1e9
    dp = [[0.0] * (m + 1) for _ in range(n + 1)]
    bt = [[""] * (m + 1) for _ in range(n + 1)]
    for i in range(1, n + 1):
        dp[i][0] = dp[i - 1][0] - 1.0
        bt[i][0] = "up"
    for j in range(1, m + 1):
        dp[0][j] = dp[0][j - 1] - 1.0
        bt[0][j] = "left"
    for i in range(1, n + 1):
        bb = base[i1 + i - 1]
        for j in range(1, m + 1):
            r = block_similarity(bb, side[j1 + j - 1])
            best = dp[i - 1][j - 1] + (r * 4.0 - 1.0 if r >= thr else NEG)
            ch = "diag"
            up = dp[i - 1][j] - 1.0
            if up > best:
                best, ch = up, "up"
            left = dp[i][j - 1] - 1.0
            if left > best:
                best, ch = left, "left"
            dp[i][j] = best
            bt[i][j] = ch

    steps = []
    i, j = n, m
    while i > 0 or j > 0:
        ch = bt[i][j]
        if ch == "diag":
            steps.append(("diag", i1 + i - 1, j1 + j - 1))
            i, j = i - 1, j - 1
        elif ch == "up" or j == 0:
            steps.append(("up", i1 + i - 1, -1))
            i -= 1
        else:
            steps.append(("ins", -1, j1 + j - 1))
            j -= 1
    steps.reverse()
    return _emit_grouped(steps, base, side, i1)


def _greedy_align(base: list, side: list, i1: int, i2: int, j1: int, j2: int,
                  thr: float) -> list:
    n, m = i2 - i1, j2 - j1
    steps = []
    i = j = 0
    while i < n and j < m:
        bb, sb = base[i1 + i], side[j1 + j]
        r = block_similarity(bb, sb)
        if r >= thr:
            steps.append(("diag", i1 + i, j1 + j))
            i += 1
            j += 1
            continue
        ahead_b = block_similarity(base[i1 + i + 1], sb) if i + 1 < n else 0.0
        ahead_s = block_similarity(bb, side[j1 + j + 1]) if j + 1 < m else 0.0
        if ahead_b >= ahead_s and ahead_b >= thr:
            steps.append(("up", i1 + i, -1))
            i += 1
        elif ahead_s >= thr:
            steps.append(("ins", -1, j1 + j))
            j += 1
        else:
            steps.append(("diag", i1 + i, j1 + j))
            i += 1
            j += 1
    while i < n:
        steps.append(("up", i1 + i, -1))
        i += 1
    while j < m:
        steps.append(("ins", -1, j1 + j))
        j += 1
    return _emit_grouped(steps, base, side, i1)


# --------------------------------------------------------------------------
# 单侧映射与移动检测
# --------------------------------------------------------------------------

@dataclass
class SideMap:
    to_side: dict = field(default_factory=dict)     # base idx -> side idx
    changed: dict = field(default_factory=dict)     # base idx -> bool
    inserts: dict = field(default_factory=dict)     # gap -> [side idx]
    deleted: set = field(default_factory=set)
    moved: dict = field(default_factory=dict)       # side idx -> base idx
    ops: list = field(default_factory=list)


def side_map(base_blocks: list, side_blocks: list, thr: float = DEFAULT_THR) -> SideMap:
    ops = align(base_blocks, side_blocks, thr)
    sm = SideMap(ops=ops)
    for op in ops:
        if op.tag in ("equal", "replace"):
            for bi, si in zip(op.bi, op.si):
                sm.to_side[bi] = si
                sm.changed[bi] = op.tag == "replace"
        elif op.tag == "delete":
            sm.deleted.update(op.bi)
        elif op.tag == "insert":
            sm.inserts.setdefault(op.b0, []).extend(op.si)

    # 移动检测：某侧「新增」的内容与同一侧「删除」的内容一致 → 判定为搬移。
    # 先做精确匹配；再做「搬移并顺手改了内容」的近似匹配（相似度 ≥ 0.9）。
    ins_list: list = []
    for _gap, sis in sorted(sm.inserts.items()):
        ins_list.extend(sis)
    del_by_norm = defaultdict(list)
    for bi in sorted(sm.deleted):
        del_by_norm[base_blocks[bi].norm].append(bi)
    consumed, rest_ins = set(), []
    for si in ins_list:
        n = side_blocks[si].norm
        lst = del_by_norm.get(n) or []
        while lst and lst[0] in consumed:
            lst.pop(0)
        if n and lst:
            bi = lst.pop(0)
            consumed.add(bi)
            sm.moved[si] = (bi, True)
        else:
            rest_ins.append(si)
    rest_del = [bi for bi in sorted(sm.deleted) if bi not in consumed]
    if rest_ins and rest_del and len(rest_ins) * len(rest_del) <= 20000:
        for si in rest_ins:
            sb = side_blocks[si]
            best, best_bi = 0.0, None
            for bi in rest_del:
                r = block_similarity(base_blocks[bi], sb)
                if r > best:
                    best, best_bi = r, bi
            if best_bi is not None and best >= 0.90:
                sm.moved[si] = (best_bi, False)
                rest_del.remove(best_bi)
    return sm


def block_equal(a: Block, b: Block) -> bool:
    if a is None or b is None:
        return False
    if a.is_table or b.is_table:
        return a.is_table and b.is_table and a.rows == b.rows
    return a.norm == b.norm


# --------------------------------------------------------------------------
# 合并结果模型
# --------------------------------------------------------------------------

ACTIONS = {
    "common": "双方一致",
    "ours": "采用我方",
    "theirs": "采用对方",
    "both": "双方一致修改",
    "moved": "搬移（两侧都有，位置不同）",
    "insert-ours": "我方新增",
    "insert-theirs": "对方新增",
    "insert-both": "双方一致新增",
    "drop-both": "双方均删除",
    "drop-ours": "我方删除",
    "drop-theirs": "对方删除",
    "conflict": "冲突（保留双方）",
    "resolved-ours": "冲突已按我方定稿",
    "resolved-theirs": "冲突已按对方定稿",
    "table-merged": "表格已单元格级合并",
    "only-a": "仅 A 有",
    "only-b": "仅 B 有",
}


@dataclass
class Piece:
    action: str
    block: Optional[Block] = None
    other: Optional[Block] = None
    origin: str = ""
    side: str = ""
    conflict_ids: list = field(default_factory=list)
    reason: str = ""
    meta: dict = field(default_factory=dict)


@dataclass
class Conflict:
    cid: str
    kind: str                # content | delete-modify | table-cell | table-shape | side-only
    reason: str
    path: str
    ours: Optional[Block] = None
    theirs: Optional[Block] = None
    base: Optional[Block] = None
    detail: dict = field(default_factory=dict)


@dataclass
class MergeResult:
    mode: str
    pieces: list = field(default_factory=list)
    conflicts: list = field(default_factory=list)
    stats: dict = field(default_factory=dict)
    moves: list = field(default_factory=list)
    warnings: list = field(default_factory=list)
    sources: dict = field(default_factory=dict)      # label -> Doc
    template: Optional[Doc] = None
    timeline: list = field(default_factory=list)

    def provenance_rows(self) -> list:
        from .labels import KIND, zh

        out = []
        for i, p in enumerate(self.pieces, start=1):
            b = p.block
            out.append({
                "no": i,
                "action": ACTIONS.get(p.action, p.action),
                "origin": p.origin,
                "locator": b.locator if b is not None else "",
                "locator_short": b.short_locator if b is not None else "",
                "kind": zh(KIND, b.kind if b is not None else ""),
                "kind_raw": (b.kind if b is not None else ""),
                "preview": b.preview(60) if b is not None else "",
                "conflict": ",".join(p.conflict_ids),
            })
        return out


# --------------------------------------------------------------------------
# 表格单元格级合并
# --------------------------------------------------------------------------

def _row_index(rows: list, key_col: int):
    """返回 (键列表, 键->行号)。键为空时退化为整行文本；重复键加序号后缀。"""
    keys, pos = [], {}
    for i, row in enumerate(rows):
        k = norm_text(row[key_col]) if key_col is not None and key_col < len(row) else ""
        if not k:
            k = norm_text("|".join(row))
        if k in pos:
            n = 2
            while f"{k}#{n}" in pos:
                n += 1
            k = f"{k}#{n}"
        keys.append(k)
        pos[k] = i
    return keys, pos


def _formula_at(block: Block, row_idx, col: int):
    """取某表格块在 (row_idx, col) 处的公式字符串（若有）。"""
    if block is None or row_idx is None:
        return None
    meta = block.meta or {}
    f = meta.get("formulas") or {}
    if not f:
        return None
    cols = meta.get("col_letters") or []
    rows = meta.get("row_numbers") or []
    if col >= len(cols) or row_idx >= len(rows):
        return None
    return f.get(f"{cols[col]}{rows[row_idx]}")


def _tbl_guard(blocks: list) -> bool:
    """只在不含横向/纵向合并单元格时做单元格级合并（避免 tc 复制错位）。"""
    try:
        from docx.oxml.ns import qn
    except ImportError:
        return True
    for b in blocks:
        if b is None or b.xml is None:
            continue
        root = b.xml
        if not hasattr(root, "iter"):
            continue
        for tc in root.iter(qn("w:tc")):
            tcPr = tc.find(qn("w:tcPr"))
            if tcPr is None:
                continue
            if tcPr.find(qn("w:gridSpan")) is not None or tcPr.find(qn("w:vMerge")) is not None:
                return False
    return True


def _tc_element(block: Block, doc: Optional[Doc], row_idx: int, col_idx: int):
    if doc is None or block.xml is None:
        return None
    try:
        import docx
        tbl = docx.table.Table(block.xml, doc.handle)
        return tbl.rows[row_idx].cells[col_idx]._tc
    except Exception:
        return None


def _tr_element(block: Block, doc: Optional[Doc], row_idx: int):
    if doc is None or block.xml is None:
        return None
    try:
        from docx.oxml.ns import qn
        trs = block.xml.findall(qn("w:tr"))
        return trs[row_idx] if row_idx < len(trs) else None
    except Exception:
        return None


def merge_tables(base_b: Optional[Block], ours_b: Block, theirs_b: Block,
                 key_col: int = 0, base_doc: Optional[Doc] = None,
                 ours_doc: Optional[Doc] = None,
                 theirs_doc: Optional[Doc] = None):
    """单元格级三方/两方表格合并。

    返回 (merged_block, cell_conflicts) ；无法处理时返回 (None, None)。
    """
    rows_o, rows_t = ours_b.rows, theirs_b.rows
    rows_b = base_b.rows if base_b is not None else None
    width = max((len(r) for r in rows_o), default=0)
    if width == 0:
        return None, None
    for r in rows_o + rows_t + (rows_b or []):
        if len(r) != width:
            return None, None
    if not _tbl_guard([base_b, ours_b, theirs_b]):
        return None, None

    keys_o, idx_o = _row_index(rows_o, key_col)
    keys_t, idx_t = _row_index(rows_t, key_col)
    keys_b, idx_b = _row_index(rows_b, key_col) if rows_b is not None else ([], {})

    # 目标行顺序：基线顺序 → 我方新增 → 对方新增。
    # 行级删除冲突遵循「不丢数据 + 显式告警」原则：删除方若无修改则删除生效，
    # 删除方有修改或被删除行被另一方改过，则保留内容并写进行级备注。
    order, seen, row_notes = [], set(), []

    def push(src, k):
        if k in seen:
            return
        seen.add(k)
        order.append((src, k))

    def row_mod(row, ref):
        return norm_text("|".join(row)) != norm_text("|".join(ref))

    if rows_b is not None:
        for k in keys_b:
            if k in idx_o:
                if k in idx_t:
                    push("o", k)
                elif row_mod(rows_o[idx_o[k]], rows_b[idx_b[k]]):
                    push("o", k)
                    row_notes.append(
                        f"行「{k}」：对方删除该行、我方有修改，已保留我方内容，请人工确认。")
                else:
                    row_notes.append(f"行「{k}」：对方删除该行、我方未修改，删除生效。")
            elif k in idx_t:
                if row_mod(rows_t[idx_t[k]], rows_b[idx_b[k]]):
                    push("t", k)
                    row_notes.append(
                        f"行「{k}」：我方删除该行、对方有修改，已保留对方内容，请人工确认。")
                else:
                    row_notes.append(f"行「{k}」：我方删除该行、对方未修改，删除生效。")
        for k in keys_o:
            if k not in idx_b:
                push("o", k)
        for k in keys_t:
            if k not in idx_b:
                push("t", k)
    else:
        for k in keys_o:
            push("o", k)
        for k in keys_t:
            push("t", k)
    if not order:
        return None, None

    final_rows, cell_over, conflicts, append_rows = [], {}, [], []
    final_meta = []
    for src, k in order:
        if src == "t":
            r_i = idx_t[k]
            final_rows.append(list(rows_t[r_i]))
            append_rows.append(r_i)
            final_meta.append(("t", r_i, k))
            continue
        r_o = idx_o[k]
        row = list(rows_o[r_o])
        r_t = idx_t.get(k)
        r_b = idx_b.get(k)
        for c in range(width):
            o_val = rows_o[r_o][c]
            t_val = rows_t[r_t][c] if r_t is not None else None
            b_val = rows_b[r_b][c] if (rows_b is not None and r_b is not None) else None
            if r_t is None:
                continue                        # 对方没有这一行：全部沿用我方
            if rows_b is None:
                o_ch, t_ch = True, True         # 两方模式：无基线，差异即差异
            else:
                o_ch = b_val is None or norm_text(o_val) != norm_text(b_val)
                t_ch = b_val is None or norm_text(t_val) != norm_text(b_val)
            if norm_text(o_val or "") == norm_text(t_val or ""):
                continue
            if not o_ch and t_ch:
                cell_over[(r_o, c)] = ("t", r_t)
                row[c] = t_val
            elif o_ch and t_ch:
                conflicts.append({
                    "row_key": k, "col": c, "row_idx": r_o,
                    "ours": o_val, "theirs": t_val, "base": b_val,
                })
        final_rows.append(row)
        final_meta.append(("o", r_o, k))

    # 后处理：把冲突定位到最终行号，并记录「该单元格取自对方」与存活公式
    key_to_final = {}
    for fi, (_src, _ri, k) in enumerate(final_meta):
        key_to_final.setdefault(k, fi)
    for cc in conflicts:
        cc["final_idx"] = key_to_final.get(cc["row_key"])

    prov, formulas = {}, {}
    for fi, (src, ri, k) in enumerate(final_meta):
        if src == "t":
            for c in range(width):
                prov[(fi, c)] = "theirs"
                f = _formula_at(theirs_b, ri, c)
                if f:
                    formulas[(fi, c)] = f
            continue
        r_t = idx_t.get(k)
        for c in range(width):
            from_theirs = isinstance(cell_over.get((ri, c)), tuple)
            if from_theirs:
                prov[(fi, c)] = "theirs"
            f = _formula_at(theirs_b, r_t, c) if (from_theirs and r_t is not None) else None
            if not f:
                f = _formula_at(ours_b, ri, c)
            if f:
                formulas[(fi, c)] = f

    merged = Block(
        kind="table", rows=final_rows,
        text="\n".join("\t".join(r) for r in final_rows),
        anchor=(ours_b.anchor or "T") + "⊕",
        path=ours_b.path or theirs_b.path,
        xml=ours_b.xml,          # 以我方表格为底稿，写出时按 plan 打补丁
        meta={
            "elem_tag": "table_merged",
            "table_plan": {
                "kind": "docx" if ours_b.xml is not None else "rows",
                "key_col": key_col,
                "cell_over": cell_over,
                "conflicts": conflicts,
                "append_rows": append_rows,
                "row_notes": row_notes,
                "prov": prov,
                "formulas": formulas,
            },
            "row_notes": row_notes,
            "sheet": ours_b.meta.get("sheet") or theirs_b.meta.get("sheet"),
            "header": ours_b.meta.get("header"),
            "col_letters": ours_b.meta.get("col_letters"),
            "row_numbers": ours_b.meta.get("row_numbers"),
            "formulas": ours_b.meta.get("formulas", {}),
        },
    )
    # 记录需要从对方取用的 OOXML 元素（docx 专用，写出时深拷贝并移植关系）
    if ours_b.xml is not None:
        elems = {}
        for key, val in cell_over.items():
            if not isinstance(key, tuple) or not isinstance(val, tuple):
                continue
            r_o, c = key
            el = _tc_element(theirs_b, theirs_doc, val[1], c)
            if el is not None:
                elems[(r_o, c)] = el
        trs = []
        for r_i in append_rows:
            el = _tr_element(theirs_b, theirs_doc, r_i)
            if el is not None:
                trs.append(el)
        merged.meta["table_plan"]["cell_elems"] = elems
        merged.meta["table_plan"]["append_trs"] = trs
    return merged, conflicts


# --------------------------------------------------------------------------
# 三方合并
# --------------------------------------------------------------------------

def three_way(base: Doc, ours: Doc, theirs: Doc, thr: float = DEFAULT_THR,
              on_conflict: str = "mark", prefer: str = "ours",
              table_key: Optional[int] = 0) -> MergeResult:
    res = MergeResult(mode="three-way")
    res.sources = {"base": base, "ours": ours, "theirs": theirs}
    res.template = ours
    for d in (base, ours, theirs):
        if d.meta.get("tracked_changes"):
            res.warnings.append(
                f"{d.label} 含有未接受的修订痕迹；合并针对的是「已显示内容」，"
                f"建议先在 Word 中接受/拒绝修订再合并。")
    if base.meta.get("fmt") not in (None, "docx") and ours.meta.get("fmt") == "docx":
        pass

    sm_o = side_map(base.blocks, ours.blocks, thr)
    sm_t = side_map(base.blocks, theirs.blocks, thr)
    stats = Counter()
    cid = [0]

    def new_conflict(kind, reason, path, ob, tb, bb, detail=None):
        cid[0] += 1
        c = Conflict(cid=f"C{cid[0]}", kind=kind, reason=reason, path=path,
                     ours=ob, theirs=tb, base=bb, detail=detail or {})
        res.conflicts.append(c)
        return c

    def resolve(action_base, ob, tb, bb, reason, path, conflict):
        """冲突处置：mark=保留双方并编号；ours/theirs=按指定一方定稿（仍记入冲突清单）。"""
        if on_conflict == "theirs":
            return Piece("resolved-theirs", tb, ob, origin="冲突已按对方定稿",
                         side="theirs", conflict_ids=[conflict.cid], reason=reason)
        if on_conflict == "ours":
            return Piece("resolved-ours", ob, tb, origin="冲突已按我方定稿",
                         side="ours", conflict_ids=[conflict.cid], reason=reason)
        return Piece("conflict", ob, tb, origin=f"冲突 {conflict.cid}",
                     side="ours" if ob is not None else "theirs",
                     conflict_ids=[conflict.cid], reason=reason)

    def emit_gap(g: int):
        o_ins = [ours.blocks[i] for i in sm_o.inserts.get(g, [])]
        t_ins = [theirs.blocks[i] for i in sm_t.inserts.get(g, [])]
        used_t = set()
        for ob in o_ins:
            match = None
            for j, tb in enumerate(t_ins):
                if j in used_t:
                    continue
                if ob.norm and block_equal(ob, tb):
                    match = j
                    break
            if match is not None:
                used_t.add(match)
                tb = t_ins[match]
                res.pieces.append(Piece("insert-both", ob, tb, origin="双方一致新增",
                                        side="both"))
                stats["insert-both"] += 1
            else:
                ob.meta["src"] = "ours"
                res.pieces.append(Piece("insert-ours", ob, origin="我方新增", side="ours"))
                stats["insert-ours"] += 1
        for j, tb in enumerate(t_ins):
            if j in used_t:
                continue
            tb.meta["src"] = "theirs"
            res.pieces.append(Piece("insert-theirs", tb, origin="对方新增", side="theirs"))
            stats["insert-theirs"] += 1

    def emit_base(i: int):
        bb = base.blocks[i]
        oi = sm_o.to_side.get(i)
        ti = sm_t.to_side.get(i)
        ob = ours.blocks[oi] if oi is not None else None
        tb = theirs.blocks[ti] if ti is not None else None
        o_ch = bool(sm_o.changed.get(i)) if oi is not None else None
        t_ch = bool(sm_t.changed.get(i)) if ti is not None else None

        if ob is None and tb is None:
            res.pieces.append(Piece("drop-both", bb, origin="双方均删除", side="both"))
            stats["drop-both"] += 1
            return
        if ob is None:
            if not t_ch:
                res.pieces.append(Piece("drop-ours", bb, tb, origin="我方删除（对方未改）",
                                        side="ours"))
                stats["drop-ours"] += 1
                return
            c = new_conflict("delete-modify", "我方删除，对方修改", bb.path, None, tb, bb)
            res.pieces.append(resolve(None, None, tb, bb,
                                      f"我方删除但对方修改：{tb.preview()}", bb.path, c))
            stats["conflict"] += 1
            return
        if tb is None:
            if not o_ch:
                res.pieces.append(Piece("drop-theirs", bb, ob, origin="对方删除（我方未改）",
                                        side="theirs"))
                stats["drop-theirs"] += 1
                return
            c = new_conflict("delete-modify", "对方删除，我方修改", bb.path, ob, None, bb)
            res.pieces.append(resolve(None, ob, None, bb,
                                      f"对方删除但我方修改：{ob.preview()}", bb.path, c))
            stats["conflict"] += 1
            return

        if not o_ch and not t_ch:
            ob.meta["src"] = "base"
            res.pieces.append(Piece("common", ob, tb, origin="双方一致", side="both"))
            stats["common"] += 1
            return
        if o_ch and not t_ch:
            ob.meta["src"] = "ours"
            res.pieces.append(Piece("ours", ob, bb, origin="我方修改", side="ours"))
            stats["ours"] += 1
            return
        if t_ch and not o_ch:
            tb.meta["src"] = "theirs"
            res.pieces.append(Piece("theirs", tb, bb, origin="对方修改", side="theirs"))
            stats["theirs"] += 1
            return

        # 双方都改
        if block_equal(ob, tb):
            ob.meta["src"] = "ours"
            res.pieces.append(Piece("both", ob, tb, origin="双方一致修改", side="both"))
            stats["both"] += 1
            return
        if ob.is_table and tb.is_table and table_key is not None:
            merged, cell_conf = merge_tables(bb, ob, tb, table_key, base, ours, theirs)
            if merged is not None:
                ids = []
                for cc in (cell_conf or []):
                    c = new_conflict(
                        "table-cell",
                        f"单元格 [{cc['row_key']} / 第{cc['col']+1}列] 双方都修改",
                        bb.path, ob, tb, bb, detail=cc)
                    ids.append(c.cid)
                if not ids:
                    merged.meta["src"] = "ours"
                    res.pieces.append(Piece("table-merged", merged, tb,
                                            origin="表格已单元格级合并", side="both"))
                    stats["table-merged"] += 1
                else:
                    merged.meta["src"] = "ours"
                    res.pieces.append(Piece("table-merged", merged, tb,
                                            origin="表格已合并（含单元格冲突）", side="both",
                                            conflict_ids=ids))
                    stats["table-merged"] += 1
                    stats["conflict"] += len(ids)
                return
            c = new_conflict("table-shape", "表格结构不同，无法单元格级合并",
                             bb.path, ob, tb, bb)
            res.pieces.append(resolve(None, ob, tb, bb, "表格结构冲突", bb.path, c))
            stats["conflict"] += 1
            return
        c = new_conflict("content", "同一处双方修改为不同内容", bb.path, ob, tb, bb)
        res.pieces.append(resolve(None, ob, tb, bb,
                                  f"我方：{ob.preview(40)} / 对方：{tb.preview(40)}",
                                  bb.path, c))
        stats["conflict"] += 1

    n_base = len(base.blocks)
    for g in range(n_base + 1):
        emit_gap(g)
        if g < n_base:
            emit_base(g)

    # 移动记录
    for side_name, sm, doc in (("ours", sm_o, ours), ("theirs", sm_t, theirs)):
        for si, (bi, exact) in sm.moved.items():
            res.moves.append({
                "side": side_name,
                "exact": exact,
                "text": doc.blocks[si].preview(50),
                "from": base.blocks[bi].locator,
                "to": doc.blocks[si].locator,
                "path": doc.blocks[si].path,
            })
    res.stats = dict(stats)
    res.stats["pieces"] = len(res.pieces)
    res.stats["conflicts"] = len(res.conflicts)
    res.stats["moves"] = len(res.moves)
    res.stats["base_blocks"] = len(base.blocks)
    res.stats["ours_blocks"] = len(ours.blocks)
    res.stats["theirs_blocks"] = len(theirs.blocks)
    return res


# --------------------------------------------------------------------------
# 两方合并（没有共同基线时：两个不同文档的对比 + 整合）
# --------------------------------------------------------------------------

def two_way(a: Doc, b: Doc, thr: float = DEFAULT_THR, on_conflict: str = "mark",
            prefer: str = "ours", table_key: Optional[int] = 0) -> MergeResult:
    res = MergeResult(mode="two-way")
    res.sources = {"ours": a, "theirs": b}
    res.template = a
    for d in (a, b):
        if d.meta.get("tracked_changes"):
            res.warnings.append(
                f"{d.label} 含有未接受的修订痕迹，建议先处理再对比。")

    ops = align(a.blocks, b.blocks, thr)
    stats = Counter()
    cid = [0]

    def new_conflict(kind, reason, path, ob, tb, detail=None):
        cid[0] += 1
        c = Conflict(cid=f"C{cid[0]}", kind=kind, reason=reason, path=path,
                     ours=ob, theirs=tb, base=None, detail=detail or {})
        res.conflicts.append(c)
        return c

    for op in ops:
        if op.tag in ("equal", "replace"):
            for bi, si in zip(op.bi, op.si):
                ob, tb = a.blocks[bi], b.blocks[si]
                if block_equal(ob, tb):
                    ob.meta["src"] = "ours"
                    res.pieces.append(Piece("common", ob, tb, origin="两文档一致", side="both"))
                    stats["common"] += 1
                    continue
                if ob.is_table and tb.is_table and table_key is not None:
                    merged, cell_conf = merge_tables(None, ob, tb, table_key, None, a, b)
                    if merged is not None:
                        ids = []
                        for cc in (cell_conf or []):
                            c = new_conflict(
                                "table-cell",
                                f"单元格 [{cc['row_key']} / 第{cc['col']+1}列] 内容不同",
                                ob.path, ob, tb, detail=cc)
                            ids.append(c.cid)
                        merged.meta["src"] = "ours"
                        res.pieces.append(Piece("table-merged", merged, tb,
                                                origin="表格差异已逐单元格标注",
                                                side="both", conflict_ids=ids))
                        stats["table-merged"] += 1
                        stats["conflict"] += len(ids)
                        continue
                c = new_conflict("content", "同一位置内容不同", ob.path, ob, tb)
                if on_conflict == "theirs":
                    res.pieces.append(Piece("resolved-theirs", tb, ob,
                                            origin="差异已按 B 定稿", side="theirs",
                                            conflict_ids=[c.cid]))
                elif on_conflict == "ours":
                    res.pieces.append(Piece("resolved-ours", ob, tb,
                                            origin="差异已按 A 定稿", side="ours",
                                            conflict_ids=[c.cid]))
                else:
                    res.pieces.append(Piece("conflict", ob, tb,
                                            origin=f"差异 {c.cid}", side="ours",
                                            conflict_ids=[c.cid],
                                            reason=f"A：{ob.preview(40)} / B：{tb.preview(40)}"))
                stats["conflict"] += 1
        elif op.tag == "delete":
            for bi in op.bi:
                ob = a.blocks[bi]
                ob.meta["src"] = "ours"
                res.pieces.append(Piece("only-a", ob, origin="仅 A 有", side="ours"))
                stats["only-a"] += 1
        elif op.tag == "insert":
            for si in op.si:
                tb = b.blocks[si]
                tb.meta["src"] = "theirs"
                res.pieces.append(Piece("only-b", tb, origin="仅 B 有", side="theirs"))
                stats["only-b"] += 1

    # 两方模式没有基线：一侧把整段/整章挪了位置时，同一内容会同时出现在
    # 「仅 A 有」和「仅 B 有」里，直接输出就会重复。这里做搬移消重：
    #   * 内容完全相同 → 判为搬移，只保留一次（保留 A 侧位置）
    #   * 内容不同但同属被搬移章节（同一标题之后、下一个标题之前）→ 判为冲突，
    #     在同一处保留双方，避免把对方的值悄悄丢掉
    pieces = res.pieces
    only_a = [i for i, p in enumerate(pieces)
              if p.action == "only-a" and p.block is not None]
    only_b = [i for i, p in enumerate(pieces)
              if p.action == "only-b" and p.block is not None]
    used_b: set = set()
    drop: set = set()

    def _run_after(idx: int) -> list:
        out = []
        for k in range(idx + 1, len(pieces)):
            p = pieces[k]
            if p.block is None or p.block.kind == "heading" or p.block.is_table:
                break
            out.append(k)
        return out

    for ia in only_a:
        if ia in drop:
            continue
        pa = pieces[ia]
        for ib in only_b:
            if ib in used_b:
                continue
            pb = pieces[ib]
            if pa.block.is_table != pb.block.is_table:
                continue
            if not block_equal(pa.block, pb.block):
                continue
            used_b.add(ib)
            drop.add(ib)
            pa.action = "moved"
            pa.origin = "两侧都有，位置不同（识别为搬移）"
            pa.other = pb.block
            res.moves.append({
                "side": "ours", "exact": True,
                "text": pa.block.preview(50),
                "from": pb.block.locator, "to": pa.block.locator,
                "path": pa.block.path,
            })
            # 标题后面跟着的正文一起挪动的情况：按位配对，避免留下孤儿段落
            if pa.block.kind == "heading":
                a_run = [k for k in _run_after(ia) if pieces[k].action == "only-a"]
                b_run = [k for k in _run_after(ib)
                         if pieces[k].action == "only-b" and k not in used_b]
                for ka, kb in zip(a_run, b_run):
                    qa, qb = pieces[ka], pieces[kb]
                    if qa.block.kind != qb.block.kind:
                        continue
                    used_b.add(kb)
                    drop.add(kb)
                    if qa.block.norm == qb.block.norm:
                        qa.action = "moved"
                        qa.origin = "随章节搬移，内容一致"
                    else:
                        c = new_conflict("content", "被搬移章节内的同一处内容不同",
                                         qa.block.path, qa.block, qb.block)
                        if on_conflict == "theirs":
                            qa.action, qa.block, qa.other = "resolved-theirs", qb.block, qa.block
                        elif on_conflict == "ours":
                            qa.action, qa.other = "resolved-ours", qb.block
                        else:
                            qa.action, qa.other = "conflict", qb.block
                        qa.conflict_ids = [c.cid]
                        qa.origin = f"冲突 {c.cid}（章节搬移后同一处不同）"
                    res.moves.append({
                        "side": "ours", "exact": qa.block.norm == qb.block.norm,
                        "text": qa.block.preview(50),
                        "from": qb.block.locator, "to": qa.block.locator,
                        "path": qa.block.path,
                    })
            break

    if drop:
        res.pieces = [p for i, p in enumerate(pieces) if i not in drop]

    counts = Counter(p.action for p in res.pieces)
    res.stats = dict(counts)
    res.stats["pieces"] = len(res.pieces)
    res.stats["conflicts"] = len(res.conflicts)
    res.stats["moves"] = len(res.moves)
    res.stats["a_blocks"] = len(a.blocks)
    res.stats["b_blocks"] = len(b.blocks)
    return res
