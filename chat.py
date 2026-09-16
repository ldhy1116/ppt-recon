"""AI PPT 结构重组 - 自然语言交互入口。

完整链路：用户自然语言 → 智能体(chat.py) → Skill → CLI/核心库 → 开源项目 → 结果。

本模块是这条路线的"轻量智能体"层：用关键词与正则解析用户自然语言意图，
调用 pptx_reorganize 核心库完成操作，并把结果格式化为人类可读文字输出。

两种使用方式：
1) 交互式 REPL（推荐用于测试）：
    python chat.py
    > 分析一下 xxx.pptx 的结构
    < ...人类可读的结构分析结果...

2) 单次执行：
    python chat.py "把 xxx.pptx 按投资人汇报重组，输出到 data\\out.pptx"

设计原则：
- 不依赖大模型解析意图，纯规则解析，确定性可测
- 调用核心库时默认启用 --smart --use-llm --trim（全功能）
- 输出为人类可读文字（不只是 JSON），便于用户直接阅读
- 写操作默认先预览，输入 y 确认后才写盘（安全要求）
- 无法解析时给出清晰提示和示例

支持的自然语言示例：
- "分析 test1.pptx 的结构" / "看看 test2.pptx 有几页、有哪些章节"
- "把 test1.pptx 按技术汇报重组" / "把 test3.pptx 面向招生综合宣讲重组"
- "把 test3.pptx 的第3页移到最前面"
- "给 test2.pptx 加一页目录"
- "把 test1.pptx 按 3,1,2,4,5,6,7 的顺序重排"
REPL 中输入「格式」可查看全部 12 种汇报场景与所有重排格式。
"""
from __future__ import annotations

import json
import os
import re
import sys
from pathlib import Path

# 大模型配置：统一由 model_config 加载（.env 文件 → 内置 Ollama 默认值）
# 外部已设置的环境变量优先；nanobot 等 Agent 可通过环境变量注入任意模型。
import model_config
model_config.apply_model_env()

import pptx_reorganize as core


# --------------------------------------------------------------------------- #
# 自然语言意图解析（轻量规则，不依赖大模型）
# --------------------------------------------------------------------------- #
# 意图关键词（按优先级：reorder > reorganize > toc > analyze）
INTENT_KEYWORDS = {
    "reorder": ["重排", "重排顺序", "按顺序", "移到最前", "移到最后", "移到前面",
                "移到后面", "调换", "交换", "第.*页移到", "把第.*页"],
    "reorganize": ["重组", "重新组织", "重新排列", "按.*汇报", "面向", "按目的",
                   "投资人", "技术分享", "技术汇报", "教学", "产品评审", "按.*目的",
                   "招生", "宣讲", "路演", "课堂", "上课"],
    "toc": ["目录", "加一页目录", "插入目录", "生成目录", "toc", "table of contents"],
    "analyze": ["分析", "看看", "看一下", "结构", "有几页", "有哪些章节", "概览",
                "概要", "章节", "主题"],
}

# 汇报目的关键词（用于 reorganize，仅用于粗判“是否提到某目的”）
PURPOSE_KEYWORDS = ["投资人", "投资", "融资", "路演", "创业", "技术分享", "技术",
                    "教学", "授课", "课堂", "上课", "课程", "培训", "讲座",
                    "产品", "评审", "发布", "新品", "招生", "宣讲", "科普", "普及",
                    "答辩", "论文", "年终", "年度", "总结", "述职", "复盘",
                    "项目", "课题", "结题", "立项", "汇报", "综合汇报"]

# 目的识别有序规则：(标准目的, 关键词元组)。先具体后宽泛，命中即返回。
# 模型自由文本归一(_norm_purpose)与规则解析(_find_purpose)共用，保证口径一致。
PURPOSE_RULES: list[tuple[str, tuple[str, ...]]] = [
    ("科普宣讲", ("科普", "普及", "大众科学", "科学宣传")),
    ("学术答辩", ("答辩", "毕业论文", "学位论文", "开题报告", "论文汇报", "科研报告")),
    ("工作总结", ("年终", "年度总结", "工作总结", "述职", "复盘", "半年总结", "年终总结")),
    ("项目汇报", ("项目汇报", "课题汇报", "项目", "课题", "结题", "立项")),
    ("产品发布", ("发布会", "新品发布", "产品发布", "产品介绍", "新品", "新产品发布")),
    ("投资人路演", ("投资人", "融资", "路演", "商业计划", "创业融资", "投资")),
    ("招生综合宣讲", ("招生", "宣讲会", "招生简章", "招生宣讲")),
    ("培训讲座", ("培训", "讲座", "研修", "训练营")),
    ("课堂教学", ("教学", "授课", "课堂", "上课", "课程", "讲解", "讲课")),
    ("技术汇报", ("技术分享", "技术汇报", "技术")),
    ("产品评审", ("产品评审", "评审会", "评审")),
    ("综合汇报", ("工作汇报", "综合汇报", "汇报", "综合")),
]


def _find_pptx_path(text: str, cwd: Path) -> str | None:
    """从自然语言中提取 .pptx 文件名（只取文件名，不含路径）。

    产品约定：PPT 必须放在 ppts/ 文件夹下，用户只需输入文件名。
    如输入 "test3.pptx" → 返回 "test3.pptx"（自动在 ppts/ 下查找）
    """
    m = re.search(r'([A-Za-z0-9_\-\\\.\:/\u4e00-\u9fa5]+\.pptx)', text)
    if not m:
        return None
    path = m.group(1).strip('，。、的')
    path = re.sub(r'^(把|将|对|让|请|分析|重组|重排|给)\s*', '', path)
    # 只取文件名，去掉任何路径前缀（强制 ppts/ 查找）
    path = Path(path).name
    return path


def _find_output_path(text: str) -> str | None:
    """提取输出文件名，支持多种自然语言表达。

    支持：输出到/输出为/保存到/保存为/存到/存为/命名为/取名为/文件名叫/叫 xxx
    文件名可不写 .pptx 后缀，自动补全。
    """
    patterns = [
        r'(?:输出到|输出为|保存到|保存为|存到|存为|命名为|取名为|文件名叫|命名|叫)\s*'
        r'([A-Za-z0-9_\-\\\.\:/\u4e00-\u9fa5]+(?:\.pptx)?)',
    ]
    for pat in patterns:
        m = re.search(pat, text)
        if m:
            name = m.group(1).strip('，。、的 ')
            if not name.lower().endswith('.pptx'):
                name = name + '.pptx'
            return name
    return None


def _find_order(text: str) -> list[int] | None:
    """提取显式顺序，如"3,1,2,4,5,6,7"或"3 1 2 4"。"""
    m = re.search(r'(\d+\s*[,，、\s]\s*\d+(?:\s*[,，、\s]\s*\d+)*)', text)
    if not m:
        return None
    parts = re.split(r'[,，、\s]+', m.group(1).strip())
    nums: list[int] = []
    for p in parts:
        if p:
            try:
                nums.append(int(p))
            except ValueError:
                return None
    return nums or None


def _find_single_page_move(text: str) -> tuple[int, str] | None:
    """解析"把第3页/第三页移到最前/最后" → (页码, 'front'|'end')。

    支持阿拉伯数字（第3页）和中文数字（第三页）。
    """
    m = re.search(r'第\s*(\d+|[一二三四五六七八九十百]+)\s*页', text)
    if not m:
        return None
    page = _cn_to_int(m.group(1))
    if page is None:
        return None
    if "最前" in text or "前面" in text or "开头" in text or "第一位" in text:
        return (page, "front")
    if "最后" in text or "末尾" in text or "结尾" in text:
        return (page, "end")
    return None


def _cn_to_int(s: str) -> int | None:
    """中文数字转整数，支持一到百。"""
    if s.isdigit():
        return int(s)
    cn_map = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5,
              "六": 6, "七": 7, "八": 8, "九": 9, "十": 10}
    if s in cn_map:
        return cn_map[s]
    if len(s) == 2 and s[0] == "十":
        return 10 + cn_map.get(s[1], 0)
    if len(s) == 2 and s[1] == "十":
        return cn_map.get(s[0], 0) * 10
    if len(s) == 3 and s[1] == "十":
        return cn_map.get(s[0], 0) * 10 + cn_map.get(s[2], 0)
    if s == "百":
        return 100
    return None


def _purpose_from_text(text: str) -> str:
    """按 PURPOSE_RULES 有序匹配目的（模型归一与规则解析共用）。"""
    for purpose, kws in PURPOSE_RULES:
        if any(kw in text for kw in kws):
            return purpose
    return ""


def _find_purpose(text: str) -> str:
    """规则层：从自然语言中识别汇报目的。"""
    return _purpose_from_text(text)


def _find_sort_level(text: str) -> str:
    """从自然语言中识别排序粒度。

    提到"单页排序/按页排序/打乱/每页独立/不按章节"等 → slide
    否则默认 chapter
    """
    slide_kw = ("单页", "按页", "每页", "打乱", "独立排序", "不按章节", "不分章节")
    if any(kw in text for kw in slide_kw):
        return "slide"
    return "chapter"


def _find_use_llm(text: str) -> bool:
    """从自然语言中识别是否用大模型排序。

    提到"不用AI/不用大模型/不用LLM/关键词排序/快速"等 → False
    否则默认 True（final 版默认启用 LLM）
    """
    no_llm_kw = ("不用ai", "不用智能", "不用大模型", "不用llm", "不用模型",
                 "关键词排序", "快速", "不用模型排序")
    if any(kw in text.lower() for kw in no_llm_kw):
        return False
    return True


def _find_smart(text: str) -> bool:
    """从自然语言中识别是否启用智能级联。

    提到"简单/快速/基础/不用智能"等 → False
    否则默认 True（final 版默认启用智能级联）
    """
    no_smart_kw = ("简单", "基础", "不用智能", "不智能", "无智能")
    return not any(kw in text for kw in no_smart_kw)


def _find_trim(text: str) -> bool:
    """从自然语言中识别是否清理冗余分隔页。

    提到"不清理/保留分隔页/保留所有页"等 → False
    否则默认 True（final 版默认清理）
    """
    no_trim_kw = ("不清理", "保留分隔", "保留所有页", "不删", "全部保留")
    return not any(kw in text for kw in no_trim_kw)


def _find_toc_ops(text: str) -> tuple[bool, bool]:
    """识别"删除旧目录/增加新目录"组合操作 → (删旧目录页, 插新目录页)。

    支持表达：删去/删除/删掉/去掉/移除 + （旧）目录 → 删除；
    增加/新增/生成/插入/添加 + （新）目录、加目录、加一页目录 → 插入。
    两个操作可在同一句指令中组合，重组时一次完成。
    """
    del_kw = ("删除旧目录", "删去旧目录", "删掉旧目录", "删除目录", "删去目录",
              "删掉目录", "去掉目录", "移除目录", "不要目录", "不保留目录")
    add_kw = ("增加新目录", "新增目录", "生成新目录", "增加目录", "生成目录",
              "插入目录", "添加目录", "加目录", "加一页目录", "新目录")
    del_toc = any(kw in text for kw in del_kw)
    add_toc = any(kw in text for kw in add_kw)
    return del_toc, add_toc


def _parse_intent_by_rules(text: str, cwd: Path) -> dict:
    """规则解析自然语言意图（确定性兜底），返回完整参数字典。"""
    result: dict = {"intent": None, "pptx": None, "output": None,
                    "purpose": "", "order": None, "page_move": None,
                    "sort_level": "chapter", "use_llm": True,
                    "smart": True, "trim": True,
                    "del_toc": False, "add_toc": False}
    for intent, keywords in INTENT_KEYWORDS.items():
        for kw in keywords:
            if re.search(kw, text):
                result["intent"] = intent
                break
        if result["intent"]:
            break
    if result["intent"] is None:
        result["intent"] = "unknown"
    result["pptx"] = _find_pptx_path(text, cwd)
    result["output"] = _find_output_path(text)
    result["purpose"] = _find_purpose(text)
    result["order"] = _find_order(text)
    result["page_move"] = _find_single_page_move(text)
    result["sort_level"] = _find_sort_level(text)
    result["use_llm"] = _find_use_llm(text)
    result["smart"] = _find_smart(text)
    result["trim"] = _find_trim(text)
    result["del_toc"], result["add_toc"] = _find_toc_ops(text)
    return result


# --------------------------------------------------------------------------- #
# 大模型意图解析（3b 本地）：模型出语义，Python 做严格校验/枚举归一，
# 任一环节失败则整字段回退规则解析，保证下游拿到的结构永远合法。
# --------------------------------------------------------------------------- #
_INTENT_ENUM = ("reorder", "reorganize", "toc", "analyze")
_PURPOSE_ENUM = ("投资人路演", "技术汇报", "课堂教学", "招生综合宣讲",
                 "产品评审", "综合汇报", "科普宣讲", "学术答辩", "工作总结",
                 "项目汇报", "产品发布", "培训讲座")
_INTENT_PROMPT = """你是 PPT 智能重排工具的指令解析器。把用户中文口语指令转成 JSON，只输出 JSON。
字段：
- intent: reorder(按顺序排/移动页) | reorganize(面向某场景重组) | toc(只操作目录) | analyze(看结构/页数/章节)
- pptx: 输入文件名，必须是用户明确给出的 .pptx；“生成到/存到/输出到 X.pptx”里的 X 是 output 不是 pptx；没给输入文件填 null
- output: “存到/保存到/输出到/生成到/命名为/叫/存成”之后的文件名，可省略 .pptx；没有填 null
- purpose: 按场景从下列词里选一个，不要自造：
  投资人路演(融资/投资/创业)、科普宣讲(面向大众科普/普及知识)、学术答辩(毕业/论文/开题答辩)、
  工作总结(年终/年度/述职/复盘)、项目汇报(项目/课题/结题)、产品发布(发布会/新品介绍)、
  招生综合宣讲(招生/招新)、培训讲座(培训/研修/讲座)、课堂教学(授课/上课/课程/教学)、
  技术汇报(技术分享/评审)、产品评审(产品方案评审)、综合汇报(一般工作汇报)；无场景填 ""
- sort_level: chapter（默认）；只有明确说单页/按页/每页打乱/不按章节才填 slide
- order: 显式页码数组如 [3,1,2]，否则 null
- page_move: “把第N页移到最前/最后” → {"page":N,"pos":"front"或"end"}，否则 null
- use_llm/smart/trim: 布尔，默认 true；明确否定（不用AI/不用智能/保留分隔页）才 false
- del_toc/add_toc: 布尔，删旧目录/加新目录"""


def _norm_purpose(v) -> str:
    """模型 purpose 归一：枚举值直接收；自由文本走与规则一致的关键词映射。"""
    if not isinstance(v, str) or not v.strip():
        return ""
    v = v.strip()
    if v in _PURPOSE_ENUM:
        return v
    return _purpose_from_text(v)


def _as_bool(v) -> bool | None:
    return bool(v) if isinstance(v, (bool, int)) else None


def _norm_page_move(v) -> list | None:
    """容忍模型给成字符串 "[3,'front']"、数组、{page,pos} 三种形态。"""
    if isinstance(v, str):
        try:
            import ast
            v = ast.literal_eval(v)
        except (ValueError, SyntaxError):
            return None
    if isinstance(v, dict):
        page, pos = v.get("page"), v.get("pos", v.get("position"))
    elif isinstance(v, (list, tuple)) and len(v) == 2:
        page, pos = v
    else:
        return None
    if not isinstance(page, int) or str(pos) not in ("front", "end"):
        return None
    return [page, str(pos)]


def _norm_order(v) -> list[int] | None:
    if not isinstance(v, list) or not v:
        return None
    nums = [x for x in v if isinstance(x, int)]
    return nums if len(nums) == len(v) and sorted(nums) == list(
        range(1, len(nums) + 1)) else None


def _norm_filename(v, *, output: bool = False) -> str | None:
    if not isinstance(v, str) or not v.strip():
        return None
    name = Path(v.strip("，。、的 ")).name  # 只取文件名，去路径前缀
    if not name.lower().endswith(".pptx"):
        name += ".pptx" if output else ""
    if not name.lower().endswith(".pptx"):
        return None
    return name


def _parse_intent_by_llm(text: str) -> dict | None:
    """模型解析并逐字段校验归一；任一字段非法即丢弃该字段（交规则补）。"""
    data = core._llm_chat_json(
        f"{_INTENT_PROMPT}\n\n用户指令：{text}",
        temperature=0.0)
    if not isinstance(data, dict):
        return None
    out: dict = {}
    if data.get("intent") in _INTENT_ENUM:
        out["intent"] = data["intent"]
    p = _norm_filename(data.get("pptx"))
    if p is not None:
        out["pptx"] = p
    elif "pptx" in data and data.get("pptx") in (None, ""):
        out["pptx"] = None  # 模型显式判空优先（修规则把输出名当输入名的问题）
    o = _norm_filename(data.get("output"), output=True)
    if o is not None:
        out["output"] = o
    purpose = _norm_purpose(data.get("purpose"))
    if purpose:
        out["purpose"] = purpose
    if data.get("sort_level") in ("chapter", "slide"):
        out["sort_level"] = data["sort_level"]
    order = _norm_order(data.get("order"))
    if order is not None:
        out["order"] = order
    pm = _norm_page_move(data.get("page_move"))
    if pm is not None:
        out["page_move"] = pm
    for k in ("use_llm", "smart", "trim", "del_toc", "add_toc"):
        b = _as_bool(data.get(k))
        if b is not None:
            out[k] = b
    return out or None


def parse_intent(text: str, cwd: Path, *, use_llm: bool = True) -> dict:
    """自然语言 → 结构化指令。模型出语义，规则做确定性约束。

    purpose 特殊：用户原话出现明确目的词（答辩/科普/路演…）时规则命中最可靠，
    优先采用，避免模型被“期末→总结”这类联想带偏；规则识别不出（纯隐性语义）
    才用模型。其余字段模型优先、非法回退规则。
    返回字段额外带 parser: "llm" | "rule"。
    """
    base = _parse_intent_by_rules(text, cwd)
    if not use_llm:
        base["parser"] = "rule"
        return base
    llm = _parse_intent_by_llm(text)
    if llm is None:
        base["parser"] = "rule"
        return base
    merged = {**base, **llm, "parser": "llm"}
    if base.get("purpose"):
        merged["purpose"] = base["purpose"]  # 原话显式目的词优先于模型联想
    return merged



# --------------------------------------------------------------------------- #
# 人类可读的结果格式化
# --------------------------------------------------------------------------- #
def _fmt_analyze(result: core.AnalysisResult) -> str:
    """把分析结果格式化为人类可读文字。"""
    lines: list[str] = []
    lines.append(f"文件: {result.file}")
    lines.append(f"共 {result.slide_count} 页")
    lines.append("")
    lines.append("【页面摘要】")
    for s in result.slides:
        tag = f" [{s.special}]" if s.special else ""
        extra = []
        if s.has_image:
            extra.append("含图")
        if s.has_table:
            extra.append("含表")
        extra_str = f" ({', '.join(extra)})" if extra else ""
        snippet = s.text_snippet[:40] + ("…" if len(s.text_snippet) > 40 else "")
        lines.append(f"  P{s.index:>2} {s.title}{tag}{extra_str}")
        if snippet:
            lines.append(f"       └─ {snippet}")
    lines.append("")
    lines.append("【章节结构】")
    for i, ch in enumerate(result.chapters, 1):
        lines.append(f"  {i}. {ch.title}  (P{ch.start}–P{ch.end})")
    return "\n".join(lines)


def _fmt_preview(preview: dict, action: str, trace: dict | None = None) -> str:
    """把重组预览格式化为人类可读文字，含章节检定路径。"""
    lines: list[str] = []
    lines.append(f"[预览-未写盘] {action}:")
    lines.append(f"   源文件: {preview['file']}")
    if "purpose" in preview:
        lines.append(f"   汇报目的: {preview['purpose']}")
    if trace:
        source = trace.get("chapter_source", "none")
        ch_count = trace.get("induced_chapter_count", 0)
        if source == "structural_sections":
            lines.append(f"   章节检定: ⓪ 原PPT分标题切章（{ch_count} 块，章块不拆散，章内按逻辑重排）")
        elif source == "llm_induced":
            wc = trace.get("within_chapter_calls", 0)
            wok = trace.get("within_chapter_ok", 0)
            inner = f"，章内模型精排 {wok}/{wc} 章" if wc else ""
            lines.append(f"   章节检定: ② 大模型单次聚类（{ch_count} 章{inner}）")
        elif source == "rule_titles":
            lines.append(f"   章节检定: ① 页标题全局归组（{ch_count} 章）")
        elif source == "anchor_fallback":
            lines.append(f"   章节检定: ③ 特殊页锚定兜底")
        induced = trace.get("induced_chapters", [])
        if induced:
            titles = [c["title"] for c in induced]
            lines.append(f"   发现章节: {' | '.join(titles)}")
    lines.append("")
    lines.append("   新顺序（原始页码 → 标题）:")
    for p in preview["preview"]:
        lines.append(f"   新位置 {p['new_position']:>2}  <-  P{p['original_index']}  {p['title']}")
    return "\n".join(lines)


def _fmt_done(out_path: str, order: list[int] | None = None) -> str:
    """写盘完成提示。"""
    lines = [f"已保存到: {out_path}"]
    if order:
        lines.append(f"   新顺序（原始页码）: {order}")
    lines.append("   可用 PowerPoint 打开查看。")
    return "\n".join(lines)


# --------------------------------------------------------------------------- #
# 执行单条自然语言指令
# --------------------------------------------------------------------------- #
def execute_command(text: str, cwd: Path, *, auto_yes: bool = False) -> str:
    """执行一条自然语言指令，返回人类可读的结果字符串。

    auto_yes=True 时跳过写操作的交互确认（用于单次执行模式）。
    """
    parsed = parse_intent(text, cwd)

    if parsed["pptx"] is None:
        return ("未在指令中找到 .pptx 文件名。\n"
                "   示例: 分析 test3.pptx 的结构\n"
                "   提示：请把 PPT 放入 ppts/ 文件夹，只输入文件名即可")
    filename = parsed["pptx"]
    pptx = str((cwd / "ppts" / filename).resolve())
    if not Path(pptx).exists():
        return (f"文件不存在: {filename}\n"
                f"   请将 PPT 放入 ppts/ 文件夹后再试。\n"
                f"   查找位置: {cwd / 'ppts'}")

    intent = parsed["intent"]

    # 1. 分析
    if intent == "analyze":
        try:
            r = core.analyze_pptx(pptx)
            return _fmt_analyze(r)
        except Exception as e:
            return f"分析失败: {e}"

    # 2. 重组
    if intent == "reorganize":
        purpose = parsed["purpose"] or "综合汇报"
        sort_level = parsed["sort_level"]
        use_llm = parsed["use_llm"]
        smart = parsed["smart"]
        trim = parsed["trim"]
        del_toc = parsed["del_toc"]
        add_toc = parsed["add_toc"]
        out = parsed["output"]
        if not out:
            stem = Path(pptx).stem
            tag = _purpose_tag(purpose)
            out = str((cwd / "data" / f"{stem}_{tag}.pptx").resolve())
        out = str((cwd / "data" / Path(out).name).resolve())

        try:
            preview = core.preview_reorganize(
                pptx, purpose, sort_level=sort_level, use_llm=use_llm,
                smart=smart,
            )
        except Exception as e:
            return f"预览失败: {e}"
        mode_label = "按章节排序" if sort_level == "chapter" else "单页独立排序"
        llm_label = "，大模型排序" if use_llm else ""
        smart_label = "，智能级联" if smart else ""
        trim_label = "，清理分隔页" if trim else ""
        trace = preview.get("trace")
        analysis = None
        # trim 预览：与 CLI 一致，在预览阶段也应用 trim
        if trim:
            induced = trace.get("induced_chapters") if trace else None
            analysis = core.analyze_pptx(pptx)
            trimmed_order, removed_divs = core.trim_internal_dividers(
                preview["order"], induced, analysis.slides)
            preview = dict(preview)
            preview["order"] = trimmed_order
            title_by_idx = {s.index: s.title for s in analysis.slides}
            preview["preview"] = [
                {"new_position": pos + 1, "original_index": orig,
                 "title": title_by_idx[orig]}
                for pos, orig in enumerate(trimmed_order)
            ]
            if removed_divs:
                trim_label = f"，清理 {len(removed_divs)} 页分隔页"
        # 删除旧目录页：识别目录特殊页并从最终顺序中剔除
        toc_note = ""
        if del_toc:
            if analysis is None:
                analysis = core.analyze_pptx(pptx)
            toc_pages = [s.index for s in analysis.slides if s.special == "目录"]
            title_by_idx = {s.index: s.title for s in analysis.slides}
            if toc_pages:
                new_order = [p for p in preview["order"] if p not in toc_pages]
                preview = dict(preview)
                preview["order"] = new_order
                preview["preview"] = [
                    {"new_position": pos + 1, "original_index": orig,
                     "title": title_by_idx[orig]}
                    for pos, orig in enumerate(new_order)
                ]
                removed = "、".join(f"P{i} {title_by_idx[i]}" for i in toc_pages)
                toc_note = f"，删除旧目录页（{removed}）"
            else:
                toc_note = "，未发现旧目录页"
        # 结构硬规则：分标题页必须绑定自己的内容，禁止两个分标题紧挨。
        # trim/删目录可能造成新的相邻或悬空，在最终顺序上再绑定一次（幂等）。
        if analysis is None:
            analysis = core.analyze_pptx(pptx)
        bound_order, hb_moves, hb_drops = core.bind_section_headings(
            preview["order"], analysis.slides)
        # resolve_order 内部已做的绑定记录在 trace；二次绑定只报告 trim 后新增动作
        n_bound = len((trace or {}).get("heading_bound", [])) + len(hb_moves)
        n_dropped = len((trace or {}).get("heading_dropped", [])) + len(hb_drops)
        bind_note = ""
        if hb_moves or hb_drops:
            preview = dict(preview)
            preview["order"] = bound_order
            title_by_idx = {s.index: s.title for s in analysis.slides}
            preview["preview"] = [
                {"new_position": pos + 1, "original_index": orig,
                 "title": title_by_idx[orig]}
                for pos, orig in enumerate(bound_order)
            ]
        if n_bound or n_dropped:
            bind_parts = []
            if n_bound:
                bind_parts.append(f"分标题绑定内容 {n_bound} 页")
            if n_dropped:
                bind_parts.append(f"删除空分标题 {n_dropped} 页")
            bind_note = "，" + "、".join(bind_parts)
        if add_toc:
            toc_note += "，重组后自动插入新目录页（封面后）"
        action = (f"按「{purpose}」重组（{mode_label}{llm_label}{smart_label}"
                  f"{trim_label}{toc_note}{bind_note}）")
        preview_text = _fmt_preview(preview, action, trace)

        if not auto_yes:
            print(preview_text)
            print()
            ans = input("是否写盘保存？(y/N): ").strip().lower()
            if ans != "y":
                return "已取消写盘。"
        else:
            print(preview_text)
            print()

        try:
            # 写盘直接复用预览阶段已确定的顺序（含 trim/删目录调整），
            # 不再二次调用 LLM，保证"所见即所得"（预览=最终结果）
            final_order = preview["order"]
            extra = ""
            if add_toc:
                # 重排结果先写临时文件，再由 insert_toc 读取临时文件并直接写到最终 dst。
                # 避免 os.replace 在 Windows 上因目标文件被占用（PowerPoint 打开/句柄未释放）
                # 而报 "拒绝访问"。
                tmp_reordered = str(
                    Path(out).with_name(Path(out).stem + "_reorder_tmp.pptx"))
                core.reorder_slides(pptx, final_order, tmp_reordered, force=True)
                dst = core.insert_toc(tmp_reordered, out, force=True)
                try:
                    os.remove(tmp_reordered)
                except OSError:
                    pass
                extra = "\n   已插入新目录页（封面后，按重组后的章节生成，页码为新位置）"
            else:
                dst = core.reorder_slides(pptx, final_order, out, force=True)
            # 生成自然语言变动说明
            trace = preview.get("trace", {})
            explanation = core.explain_reorganization(
                trace, final_order, core.analyze_pptx(pptx).slides)
            if explanation:
                extra += f"\n\n--- 变动说明 ---\n{explanation}"
            return _fmt_done(dst, final_order) + extra
        except Exception as e:
            return f"重组失败: {e}"

    # 3. 显式重排
    if intent == "reorder":
        if parsed["page_move"] is not None and parsed["order"] is None:
            page, pos = parsed["page_move"]
            try:
                analysis = core.analyze_pptx(pptx)
                n = analysis.slide_count
                others = [i for i in range(1, n + 1) if i != page]
                order = [page] + others if pos == "front" else others + [page]
            except Exception as e:
                return f"解析页数失败: {e}"
        elif parsed["order"] is not None:
            order = parsed["order"]
        else:
            return ("未识别到重排顺序。\n"
                    "   示例1: 把 xxx.pptx 的第3页移到最前面\n"
                    "   示例2: 把 xxx.pptx 按 3,1,2,4,5,6,7 的顺序重排")

        out = parsed["output"]
        if not out:
            stem = Path(pptx).stem
            out = str((cwd / "data" / f"{stem}_重排.pptx").resolve())
        out = str((cwd / "data" / Path(out).name).resolve())

        try:
            preview = core.preview_reorder(pptx, order)
        except Exception as e:
            return f"预览失败: {e}"
        preview_text = _fmt_preview(preview, "显式重排")

        if not auto_yes:
            print(preview_text)
            print()
            ans = input("是否写盘保存？(y/N): ").strip().lower()
            if ans != "y":
                return "已取消写盘。"
        else:
            print(preview_text)
            print()

        try:
            dst = core.reorder_slides(pptx, order, out, force=True)
            return _fmt_done(dst, order)
        except Exception as e:
            return f"重排失败: {e}"

    # 4. 目录页
    if intent == "toc":
        out = parsed["output"]
        if not out:
            stem = Path(pptx).stem
            out = str((cwd / "data" / f"{stem}_含目录.pptx").resolve())
        out = str((cwd / "data" / Path(out).name).resolve())

        try:
            r = core.analyze_pptx(pptx)
            ch_list = "\n".join(
                f"   {i}. {ch.title}  (P{ch.start}-P{ch.end})"
                for i, ch in enumerate(r.chapters, 1)
            )
            preview_text = (f"[预览-未写盘] 目录页:\n"
                            f"   将在首页前插入目录页，列出 {len(r.chapters)} 个章节:\n"
                            f"{ch_list}")
        except Exception as e:
            return f"预览失败: {e}"

        if not auto_yes:
            print(preview_text)
            print()
            ans = input("是否写盘保存？(y/N): ").strip().lower()
            if ans != "y":
                return "已取消写盘。"
        else:
            print(preview_text)
            print()

        try:
            dst = core.insert_toc(pptx, out, force=True)
            return _fmt_done(dst)
        except Exception as e:
            return f"目录页生成失败: {e}"

    # 未知意图
    return ("无法理解指令意图。支持的操作:\n"
            "   - 分析: '分析 test1.pptx 的结构'\n"
            "   - 重组: '把 test3.pptx 按招生综合宣讲重组'\n"
            "   - 重排: '把 test3.pptx 的第3页移到最前面'\n"
            "   - 目录: '给 test2.pptx 加一页目录'\n"
            "   输入「格式」可查看全部 12 种汇报场景与重排格式")


def _purpose_tag(purpose: str) -> str:
    """从目的生成文件名标签。"""
    mapping = {
        "投资人路演": "investor",
        "技术汇报": "tech",
        "课堂教学": "classroom",
        "招生综合宣讲": "intro",
        "产品评审": "product",
        "综合汇报": "summary",
        "科普宣讲": "science",
        "学术答辩": "defense",
        "工作总结": "work",
        "项目汇报": "project",
        "产品发布": "launch",
        "培训讲座": "training",
    }
    return mapping.get(purpose, "reorg")


# --------------------------------------------------------------------------- #
# 交互式 REPL
# --------------------------------------------------------------------------- #
# 12 种重排场景的说明数据源（帮助窗口的唯一入口；触发词以 PURPOSE_RULES
# 为准，这里只维护场景简介与代表示例，避免多处文案各自漂移）。
# 目的 → (场景简介, 示例指令)
PURPOSE_INFO: dict[str, tuple[str, str]] = {
    "投资人路演": ("融资路演 / 商业计划书：痛点与市场机会先行，产品、商业模式、"
                "财务预测依次推进，团队与愿景收尾",
                "把 test1.pptx 面向投资人路演重组"),
    "技术汇报": ("技术分享 / 方案汇报：背景问题先行，架构设计、关键实现、"
              "测试验证依次展开",
              "把 test1.pptx 按技术汇报重组"),
    "课堂教学": ("授课课件：按认知规律循序渐进，概念先行、实例巩固、"
              "小结收尾，章节块不拆散",
              "把 test2.pptx 按课堂教学重组"),
    "招生综合宣讲": ("招生 / 招新宣讲：学校与专业亮点前置，培养、升学就业、"
                  "报考信息依次展开，问答板块殿后",
                  "把 test3.pptx 按招生综合宣讲重组"),
    "产品评审": ("内部产品方案评审：目标与方案先行，功能细节、风险与排期"
              "随后，便于评审决策",
              "把 my.pptx 按产品评审重组"),
    "综合汇报": ("一般工作汇报：总—分—总结构，总体情况先行，分项展开、"
              "总结收尾（未识别出具体场景时的默认格式）",
              "把 my.pptx 按综合汇报重组"),
    "科普宣讲": ("面向大众的科普 / 普及讲座：趣味现象引入，原理逐层解密，"
              "应用与展望收尾",
              "把 my.pptx 面向科普宣讲重组"),
    "学术答辩": ("毕业论文 / 学位论文 / 开题答辩：研究背景与意义先行，"
              "方法、结果、结论与展望依次展开",
              "把 my.pptx 按学术答辩重组"),
    "工作总结": ("年终 / 年度总结 / 述职复盘：工作成果先行，问题不足、"
              "下一步计划随后",
              "把 my.pptx 按工作总结重组"),
    "项目汇报": ("项目 / 课题汇报：进展与成果先行，问题风险、下一阶段"
              "计划随后",
              "把 my.pptx 按项目汇报重组"),
    "产品发布": ("新品发布会：最大亮点前置制造高潮，核心功能、使用场景、"
              "上市信息依次推进",
              "把 my.pptx 面向产品发布重组"),
    "培训讲座": ("培训 / 研修 / 训练营：学习目标先行，知识讲解、实操演练、"
              "答疑总结依次展开",
              "把 my.pptx 按培训讲座重组"),
}

_RULE_INDEX = dict(PURPOSE_RULES)


def render_formats_window() -> str:
    """渲染「全部重排格式」一览窗口（REPL 输入 格式/formats 时展示）。"""
    bar = "=" * 66
    thin = "-" * 66
    lines = [
        bar,
        "  支持的全部重排格式（12 种汇报场景 + 2 种直接重排 + 目录操作/分析）",
        bar,
        "",
        "【一】按汇报场景重组：自动检定章节、按演讲目的重排章序与章内页序",
        "  指令格式：把 <文件名> 按 <场景说法> 重组 [输出到 <名字>]",
        "  示例：把 test2.pptx 按课堂教学重组，输出到 my_class.pptx",
        "",
    ]
    for i, purpose in enumerate(_PURPOSE_ENUM, 1):
        desc, example = PURPOSE_INFO[purpose]
        kws = "、".join(_RULE_INDEX[purpose])
        lines.append(f"  {i:>2}. {purpose}")
        lines.append(f"      说法：{kws}")
        lines.append(f"      适用：{desc}")
        lines.append(f"      示例：{example}")
        lines.append("")
    lines += [
        thin,
        "【二】显式页序重排：不调模型，完全按你给的页码顺序",
        "  示例：把 test1.pptx 按 3,1,2,4,5,6,7 的顺序重排",
        "  补充：加「单页/按页/不按章节」= 每页独立排序，不按章块",
        "",
        "【三】单页移动：把指定页移到最前或最后",
        "  示例：把 test3.pptx 的第3页移到最前面",
        "        把 test3.pptx 的第5页移到最后",
        "",
        "【四】目录操作",
        "  加新目录：给 test2.pptx 加一页目录",
        "  删旧加新：把 test3.pptx 按招生综合宣讲重组，删去旧目录增加新目录",
        "",
        "【附】结构分析（不改动文件）：分析 test1.pptx 的结构",
        "",
        "控制选项（写在任意重组指令中）：不用AI（0 token 关键词模式）/ "
        "简单基础（关智能级联）/ 保留分隔页（不清理）",
        bar,
    ]
    return "\n".join(lines)


HELP_TEXT = f"""
========================================
  AI PPT 结构重组 - 自然语言交互入口
  （默认启用 LLM + 智能级联 + 清理分隔页）
========================================
PPT 文件请放入 ppts/ 文件夹，只需输入文件名。
内置 3 个测试用例：test1（研究报告 18 页）/ test2（课堂教学 57 页）/ test3（招生宣讲 60 页）。

常用指令:
  - 分析 test1.pptx 的结构
  - 把 test1.pptx 按技术汇报重组，输出到 my_tech.pptx
  - 把 test2.pptx 按课堂教学重组
  - 把 test3.pptx 按招生综合宣讲重组
  - 把 test1.pptx 面向投资人路演重组，删去旧目录增加新目录
  - 把 test1.pptx 按 3,1,2,4,5,6,7 的顺序重排
  - 把 test3.pptx 的第3页移到最前面
  - 给 test2.pptx 加一页目录

命令:
  格式 / formats  查看支持的全部重排格式（12 种汇报场景）
  help / ?        显示本帮助
  exit / quit     退出
========================================
"""


def run_repl(cwd: Path) -> None:
    """交互式 REPL：用户输入自然语言，系统输出结果。"""
    print(HELP_TEXT)
    while True:
        try:
            user_input = input("> ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\n再见！")
            break
        if not user_input:
            continue
        if user_input.lower() in ("exit", "quit", "q", "退出", "再见"):
            print("再见！")
            break
        if user_input.lower() in ("help", "?", "帮助"):
            print(HELP_TEXT)
            continue
        if user_input.lower() in ("格式", "formats", "format", "场景", "目的", "重排格式"):
            print(render_formats_window())
            continue
        try:
            result = execute_command(user_input, cwd)
            print()
            print(result)
            print()
        except Exception as e:
            print(f"\n执行出错: {e}\n")


def main() -> int:
    """入口：支持 REPL、单次执行，以及自然语言解析自测。

    解析自测（只看自然语言→结构化指令，不碰 PPT，便于 PowerShell 测试）：
      python chat.py --parse "指令"        # 模型解析 + 规则解析并排
      python chat.py --parse-rule "指令"   # 只看规则解析（不调模型）
    """
    cwd = Path(__file__).resolve().parent
    core.setup_logging(str(cwd / "logs"))

    # 解析自测模式
    if len(sys.argv) > 2 and sys.argv[1] in ("--parse", "--parse-rule"):
        text = " ".join(sys.argv[2:])
        rule = _parse_intent_by_rules(text, cwd)
        if sys.argv[1] == "--parse-rule":
            print(json.dumps(rule, ensure_ascii=False, indent=2))
            return 0
        merged = parse_intent(text, cwd, use_llm=True)
        print("【模型(生效)】")
        print(json.dumps(merged, ensure_ascii=False, indent=2))
        print("【规则(兜底)】")
        print(json.dumps(rule, ensure_ascii=False, indent=2))
        return 0

    # 单次执行: python chat.py "自然语言指令"
    if len(sys.argv) > 1:
        text = " ".join(sys.argv[1:])
        result = execute_command(text, cwd, auto_yes=True)
        print(result)
        return 0

    # 交互式 REPL
    run_repl(cwd)
    return 0


if __name__ == "__main__":
    sys.exit(main())
