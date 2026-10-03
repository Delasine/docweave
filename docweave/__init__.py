# -*- coding: utf-8 -*-
"""docweave —— 面向 docx/xlsx 的批量文档合并与一致性审计工具。

核心能力（刻意避开 Word/Excel 自带的大众功能）：
  * merge  三方/两方语义合并：共同基线、段落搬移检测、单元格级表格合并、溯源附录
  * audit  文表对账：正文里的数字/日期/编号 与 表格数据 的跨格式一致性核对
  * refs   编号与交叉引用体检：悬空引用、编号断档、编号重复
  * drift  口径漂移：跨版本指标漂移 + 单文档自相矛盾
  * batch  批量增量处理 + 台账 + HTML 报告
"""
__version__ = "1.0.0"
__all__ = ["model", "merge", "writers", "audit", "ops"]
