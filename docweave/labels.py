# -*- coding: utf-8 -*-
"""docweave.labels —— 面向用户的中文标签。

报告和界面里一律显示中文，内部代号只保留在 JSON sidecar 里（便于机器读取）。
"""

KIND = {
    "heading": "标题",
    "para": "段落",
    "list": "列表",
    "image": "图片",
    "caption": "题注",
    "table": "表格",
    "table_merged": "表格（已合并）",
    "xlsx_table": "工作表",
    "csv_table": "数据表",
    "empty": "空段",
}

CONFLICT = {
    "content": "内容分歧",
    "delete-modify": "删除/修改冲突",
    "table-cell": "表格单元格",
    "table-shape": "表格结构",
    "side-only": "单侧内容",
}

FINDING = {
    # 文表对账
    "value_conflict": "数值冲突",
    "unit_mismatch": "单位不可比",
    "missing_in_table": "正文有表中无",
    "missing_in_doc": "表中有正文无",
    "ok": "一致",
    # 编号与引用体检
    "dangling_chapter": "章引用悬空",
    "dangling_section": "小节引用悬空",
    "dangling_figure": "图引用悬空",
    "dangling_table": "表引用悬空",
    "dangling_attachment": "附件引用悬空",
    "dangling_clause": "条款引用悬空",
    "dup_number": "编号重复",
    "gap_number": "编号断档",
    "gap_section": "小节断档",
    "vague_ref": "模糊引用",
    # 口径漂移
    "intra_conflict": "文档内自相矛盾",
    "version_drift": "跨版本变化",
    # 合并
    "conflict": "内容冲突",
}

TASK = {
    "merge": "文档合并",
    "merge-dir": "批量合并",
    "audit": "文表对账",
    "refs": "引用体检",
    "refs-dir": "批量体检",
    "drift": "口径漂移",
}

STATUS = {"ok": "完成", "skip": "跳过（无变化）", "fail": "失败"}


def zh(mapping: dict, key, default: str = "") -> str:
    """查表给中文；查不到时返回 default，default 为空则回退到原始代号。"""
    if key is None:
        return default
    return mapping.get(key) or default or str(key)
