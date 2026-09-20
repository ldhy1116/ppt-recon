"""AI PPT 结构重组核心库：结构分析、按目的重组、显式重排、目录页生成。

低 Token（analyze 只回标题/短片段/章节结构）、安全（路径与扩展名校验、默认
不覆盖源文件、写操作可 dry-run）、可独立运行（CLI 直调，无需 Agent/大模型）。
"""
from __future__ import annotations

import json
import logging
import os
import re
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, Sequence

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.opc.packuri import PackURI
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

try:
    import jieba
    import jieba.analyse
    _HAS_JIEBA = True
except ImportError:  # 未装 jieba 时退化到子串匹配
    _HAS_JIEBA = False

logger = logging.getLogger("ppt_reorganize")

# 章节标题编号：中英文常见方式
CHAPTER_PATTERNS = [
    re.compile(r"^\s*第[一二三四五六七八九十百零\d]+[章节部分篇]"),
    re.compile(r"^\s*[0-9]+(\.[0-9]+)*[\.\、]?\s+\S"),
    re.compile(r"^\s*Chapter\s+\d+", re.IGNORECASE),
    re.compile(r"^\s*Section\s+\d+", re.IGNORECASE),
]

# 特殊页面（标题匹配即识别，宽松以兼容真实 PPT）
SPECIAL_PATTERNS = {
    "目录": re.compile(
        r"^\s*(目\s*录|内容提要|contents|目\s*次|主要内容|本章主要内容|"
        r"本章内容|内容概要|内容概述|提\s*纲|议\s*程|大\s*纲|概\s*览|"
        r"agenda|outline)\s*$", re.IGNORECASE),
    "封面": re.compile(r"^\s*(封\s*面|cover)\s*$", re.IGNORECASE),
    "总结": re.compile(
        r"^\s*(?:(?:内容|本章|全文|全书|课程|学期|项目)\s*)?"
        r"(?:总\s*结|小\s*结|结\s*语|结\s*论|总结与展望|结语与展望|"
        r"summary|conclusion|研究结论).*$",
        re.IGNORECASE
    ),
    "致谢": re.compile(
        r"^\s*(谢\s*谢|致\s*谢|鸣\s*谢|感谢聆听|感谢观看|感谢倾听|"
        r"thanks|thank\s*you|the\s*end|q\s*&?\s*a|道德经|完)\s*$",
        re.IGNORECASE),
    "参考文献": re.compile(
        r"^\s*(参\s*考\s*文\s*献|参\s*考\s*资\s*料|引\s*文|references?)\s*$",
        re.IGNORECASE),
}
SPECIAL_KINDS = ("封面", "目录", "总结", "参考文献", "致谢")

# 封面内容识别（与位置无关，对乱序健壮）：短标题+稀疏正文为门控，机构/事件
# 关键词为正信号；纯数字编号（“01”“第三章”）的标题或副标题是分隔页，排除。
COVER_TITLE_MAX_LEN = 20        # 真实 PPT 标题常含主题副标题，放宽到 20 字
COVER_MAX_PARAS = 12            # 封面无项目符号；校名/题目/作者/日期可有多短行
COVER_FALLBACK_MAX_PARAS = 2    # 零候选时首页先验的严格门控
COVER_FALLBACK_BODY_MAX = 40
COVER_ORG_KEYWORDS = (
    "大学", "学院", "学校", "中学", "小学", "公司", "集团", "研究院", "研究所",
    "医院", "银行", "有限", "股份", "科技", "宣讲", "招聘", "发布会",
    "实验室", "中心", "教研室", "教研组", "课题组", "系", "部",
    "政府", "局", "厅", "委", "院", "署",
)
# 日期/时间模式（封面常见）：YYYY年、YYYY.MM、YYYY-MM、YYYY/MM 等
COVER_DATE_RE = re.compile(
    r"(?:19|20)\d{2}\s*年|"
    r"(?:19|20)\d{2}[.\-/年]\s*\d{1,2}\s*月?|"
    r"\d{4}\s*年\s*\d{1,2}\s*月"
)
DIVIDER_TITLE_RE = re.compile(r"^\s*(?:\d{1,2}|第[0-9一二三四五六七八九十]+[章节部分篇]?)\s*$")
# 分标题页：第N节/章（可带后续节名，如「第2节 匿名化」）、Chapter/Section N。
# 只认标题不认副标题——正文页副标题常带「01」角标（如两级编号稿），不能误判。
SECTION_HEADING_RE = re.compile(
    r"^\s*(?:第\s*[0-9一二三四五六七八九十百零]+\s*[章节部分篇]"
    r"|chapter\s+\d+|section\s+\d+)", re.IGNORECASE)
# 纯编号标题（「01」「第三章」）须页面稀疏才算分标题：带实质正文/图表的是
# 编号内容页（如两级编号稿每页标题都带同章角标，章首页极稀疏，内容页正文很长）。
SECTION_NUM_BODY_MAX = 30
SECTION_NUM_PARAS_MAX = 5
COVER_NEGATIVE_KEYWORDS = (
    # 特殊页
    "目录", "内容", "参考", "文献", "致谢", "谢谢", "鸣谢", "感谢",
    # 章节/正文性词汇
    "研究", "分析", "案例", "成果", "结论", "过程", "问题", "实践",
    "主题", "论证", "内涵", "建议", "精典", "经典", "背景", "动机",
    "方法", "实验", "评估", "总结", "展望", "概述", "简介", "章节",
    "汇报", "技术", "架构", "系统", "市场", "商业", "价值", "需求",
    "产品", "功能", "设计", "报告", "第一", "第二", "第三", "第四",
    "第五", "第六", "第七", "第八", "第九", "第十",
    "chapter", "section",
    # 防误判：正文常见动词/名词
    "应用", "推广", "部署", "实现", "原理", "结构", "发展", "趋势",
    "战略", "政策", "管理", "运营", "优化", "提升", "创新", "挑战",
    "安全", "防护", "检测", "防御", "攻击", "验证", "认证", "加密",
)

# 汇报阶段标签（策略④）：只作信息增强喂给大模型，不做 Python 强排。
STAGE_NAMES = {0: "背景动机", 1: "相关基础", 2: "方案设计", 3: "实现方法",
               4: "过程结果", 5: "结论建议"}
# 每阶段 (关键词, 权重)；强信号词（设计/结论等）权重 2
STAGE_KEYWORDS: dict[int, list[tuple[str, int]]] = {
    5: [("结论", 2), ("总结", 2), ("建议", 2), ("展望", 2)],
    2: [("架构设计", 2), ("设计", 2), ("方案", 1)],
    4: [("实验", 1), ("测试", 1), ("过程", 1), ("结果", 1), ("评估", 1),
        ("实践", 1), ("展示", 1), ("界面", 1), ("分析", 1)],
    3: [("实现", 2), ("模块", 1), ("功能", 1), ("步骤", 1), ("方法", 1), ("开发", 1)],
    1: [("相关工作", 2), ("研究基础", 2), ("现有研究", 2), ("原理", 1),
        ("理论基础", 1), ("预备", 1)],
    0: [("背景", 2), ("动机", 2), ("矛盾", 2), ("内涵", 1), ("意义", 1),
        ("概述", 1), ("简介", 1), ("问题提出", 2)],
}

# 组内序号（策略③）：段首 "1." "2、" "3)" 等显式编号
ORDINAL_LEAD = re.compile(r"^\s*(\d{1,2})\s*[.、)]")
ORDINAL_MAX = 99
ORDINAL_INLINE_MIN = 3  # 同页出现 ≥3 个不同段首编号 = 页内列举，不取序号

TEXT_SNIPPET_LEN = 80


@dataclass
class SlideInfo:
    """单页 PPT 的结构化摘要。"""

    index: int  # 1-based 原始页码
    title: str
    text_snippet: str
    n_shapes: int
    has_image: bool
    has_table: bool
    layout: str
    has_chart: bool = False
    special: str | None = None  # 封面/目录/总结/致谢/参考文献/None
    body_len: int = 0           # 正文字数（不含标题）
    paragraph_count: int = 0
    bullet_count: int = 0
    subtitle: str = ""          # 正文首个非空行，同标题组内区分页用
    ordinal: int | None = None  # 组内序号，如 "2.漏洞判别模块"
    stage: int | None = None    # 汇报阶段 0..5（背景→结论）
    org_signal: bool = False    # 正文任意位置含机构/事件词（封面归属行常在第2+行）

    def to_compact(self) -> dict:
        return {
            "index": self.index,
            "title": self.title,
            "snippet": self.text_snippet,
            "shapes": self.n_shapes,
            "img": self.has_image,
            "table": self.has_table,
            "chart": self.has_chart,
            "layout": self.layout,
            "special": self.special,
            "body_len": self.body_len,
            "paragraphs": self.paragraph_count,
            "bullets": self.bullet_count,
            "subtitle": self.subtitle,
            "ordinal": self.ordinal,
            "stage": STAGE_NAMES.get(self.stage) if self.stage is not None else None,
        }


@dataclass
class Chapter:
    """识别出的章节（页码 1-based，含端点）。"""

    title: str
    start: int
    end: int

    def to_compact(self) -> dict:
        return {"title": self.title, "start": self.start, "end": self.end}


@dataclass
class AnalysisResult:
    file: str
    slide_count: int
    slides: list[SlideInfo] = field(default_factory=list)
    chapters: list[Chapter] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "file": self.file,
            "slide_count": self.slide_count,
            "slides": [s.to_compact() for s in self.slides],
            "chapters": [c.to_compact() for c in self.chapters],
        }


# --------------------------------------------------------------------------- #
# 路径与安全校验
# --------------------------------------------------------------------------- #
def validate_pptx_path(path: str, *, must_exist: bool = True) -> Path:
    """校验扩展名、路径穿越与存在性。"""
    p = Path(path).expanduser().resolve()
    if p.suffix.lower() != ".pptx":
        raise ValueError(f"仅支持 .pptx 文件，收到: {p.suffix or '(无扩展名)'}")
    if ".." in Path(path).parts:
        raise ValueError("路径中不允许出现 '..'")
    if must_exist and not p.exists():
        raise FileNotFoundError(f"文件不存在: {p}")
    return p


def check_output_path(out_path: str, src_path: Path, *, force: bool = False) -> Path:
    """输出校验：默认禁止覆盖源文件/已存在文件。"""
    p = validate_pptx_path(out_path, must_exist=False)
    if p == src_path:
        raise ValueError("输出路径与源文件相同，拒绝覆盖源文件（请使用 --force 或更换输出名）")
    if p.exists() and not force:
        raise FileExistsError(f"输出文件已存在: {p}（使用 --force 覆盖）")
    p.parent.mkdir(parents=True, exist_ok=True)
    return p


# --------------------------------------------------------------------------- #
# 结构分析
# --------------------------------------------------------------------------- #
def _shape_first_line(shape) -> str:
    if not shape.has_text_frame:
        return ""
    t = shape.text_frame.text.strip()
    if not t:
        return ""
    return t.splitlines()[0][:TEXT_SNIPPET_LEN]


def _shape_font_pt(shape) -> float:
    """形状首个带显式字号 run 的字号（磅）；无显式字号返回 0。"""
    if not shape.has_text_frame:
        return 0.0
    for para in shape.text_frame.paragraphs:
        for r in para.runs:
            if r.font.size is not None:
                return r.font.size.pt
    return 0.0


# 标题带：页面顶部 22%（标准 7.5 英寸高 ≈ 1.65 英寸）内才算标题位置
_TITLE_BAND_EMU = int(6858000 * 0.22)
# 旧候选落到页面垂直中线以下，判定为图中标签/注释等噪声（如地图上的 Africa）
_TITLE_NOISE_TOP = int(6858000 * 0.5)
_TITLE_MAX_LEN = 30
# 无意义占位短词（封面署名占位等），重选时降权
_TITLE_JUNK = {"xxx", "xx", "n/a", "none", "title"}


def _fallback_title(slide) -> str:
    """无标题占位符时选标题。默认沿用形状树首个非空文本（成熟行为，覆盖面最广）；

    仅当旧候选位于页面垂直中线以下（图中标签/注释，如地图上 9pt 的 Africa）
    时，才在顶部标题带内按 字号大→纯数字降权→字数少→靠上靠左 重选，
    且占位词降权。其余位置的第一文本一律不动，避免大号装饰标语/数字
    反向抢占内容页标题。
    """
    cands = []
    first = ""
    for sh in slide.shapes:
        line = _shape_first_line(sh)
        if not line:
            continue
        if not first:
            first = line
        top = sh.top if sh.top is not None else 10**9
        left = sh.left if sh.left is not None else 10**9
        cands.append({"line": line, "top": int(top), "left": int(left),
                      "size": _shape_font_pt(sh),
                      "num": 1 if re.fullmatch(r"[0-9０-９]{1,3}", line) else 0,
                      "junk": 1 if line.strip().lower() in _TITLE_JUNK else 0})
    if not cands:
        return ""
    c0 = next((c for c in cands if c["line"] == first), cands[0])
    if c0["top"] < _TITLE_NOISE_TOP:
        return first  # 旧候选不在下半页，维持原行为

    pool = [c for c in cands
            if c["top"] <= _TITLE_BAND_EMU and len(c["line"]) <= _TITLE_MAX_LEN
            and c["size"] > c0["size"]]  # 顶部候选必须严格更大才替换，
    # 防止设计者刻意放在下方的大字号标题（44pt）被顶部更小的口号取代
    if not pool:
        return first

    def rank(c: dict) -> tuple:
        return (c["junk"], -c["size"], c["num"], len(c["line"]), c["top"], c["left"])

    return sorted(pool, key=rank)[0]["line"]


def _slide_title(slide) -> str:
    """优先标题占位符，否则启发式从全部文本形状中选标题。"""
    try:
        if slide.shapes.title is not None and slide.shapes.title.has_text_frame:
            t = slide.shapes.title.text_frame.text.strip()
            if t:
                return t
    except Exception:
        pass
    return _fallback_title(slide)


def _slide_text(slide) -> str:
    parts = [sh.text_frame.text for sh in slide.shapes if sh.has_text_frame]
    return "\n".join(parts).strip()


def _has_image(slide) -> bool:
    return any(s.shape_type == MSO_SHAPE_TYPE.PICTURE for s in slide.shapes)


def _has_table(slide) -> bool:
    return any(getattr(s, "has_table", False) and s.has_table for s in slide.shapes)


def _has_chart(slide) -> bool:
    return any(getattr(s, "has_chart", False) and s.has_chart for s in slide.shapes)


def _classify_special(title: str) -> str | None:
    if not title:
        return None
    for name, pat in SPECIAL_PATTERNS.items():
        if pat.match(title):
            return name
    return None


def _looks_like_cover(s: "SlideInfo") -> bool:
    """只看标题：短且不含章节词/特殊页词；唯一性等整库门控在 _pick_cover。
    日期模式（YYYY年/月）也放行——真实封面标题常为日期。"""
    if s.special is not None:
        return False
    title = re.sub(r"\s+", "", s.title or "")
    if not title or title == "(无标题)" or len(title) > COVER_TITLE_MAX_LEN:
        return False
    low = title.lower()
    if any(w.lower() in low for w in COVER_NEGATIVE_KEYWORDS):
        return False
    return True


def _is_divider(s: "SlideInfo") -> bool:
    """章节分隔页：标题或副标题为纯数字/章节编号（如"01""第三章"）。"""
    return bool(DIVIDER_TITLE_RE.match((s.title or "").strip())
                or DIVIDER_TITLE_RE.match((s.subtitle or "").strip()))


def _is_section_heading(s: "SlideInfo | None") -> bool:
    """分标题页（仅看标题，避免副标题「01」角标误伤正文页）：
    - 标题带编号+名称（「第2节 匿名化」「Chapter 3 …」）：即分标题；
    - 纯编号标题（「01」「第三章」）：须稀疏（少字少段无图表），
      带实质正文/图表的是编号内容页，不算分标题。"""
    if s is None or s.special is not None:
        return False
    title = (s.title or "").strip()
    if not DIVIDER_TITLE_RE.match(title):
        return bool(SECTION_HEADING_RE.match(title))
    if s.has_image or s.has_table or s.has_chart:
        return False
    return (s.body_len <= SECTION_NUM_BODY_MAX
            and s.paragraph_count <= SECTION_NUM_PARAS_MAX)


_CN_DIGITS = "零一二三四五六七八九十"


def _divider_section_num(s: "SlideInfo") -> int | None:
    """从分隔页提取节号（01→1, 第三章→3）；无法提取返回 None。"""
    for text in ((s.title or "").strip(), (s.subtitle or "").strip()):
        m = re.match(r"^\s*0*(\d{1,2})\s*$", text)
        if m:
            return int(m.group(1))
        m = re.match(r"^\s*第([0-9一二三四五六七八九十]+)", text)
        if m:
            v = m.group(1)
            if v.isdigit():
                return int(v)
            if len(v) == 1:
                return _CN_DIGITS.index(v)
            if v[0] == "十":
                return 10 + (_CN_DIGITS.index(v[1]) if len(v) > 1 else 0)
            return _CN_DIGITS.index(v[0]) * 10 + (
                _CN_DIGITS.index(v[1]) if len(v) > 1 else 0)
    return None


def _is_reverse_numbered_heading(s: "SlideInfo | None") -> bool:
    """反向形态分标题（仅用于目录构建）：标题为名称、正文首行为纯编号
    （「招生政策问答」+「03」，与编号章「编号标题+名称正文」相反），
    且整页除名称+编号外无任何其他内容（无图表、恰 2 段、字数恰为二者之和）。
    严格门控防误伤编号内容页（如「招生政策问答 03 常见问题1：报考条件」
    还带一段正文，被排除）。"""
    if s is None or s.special is not None:
        return False
    title = (s.title or "").strip()
    sub = (s.subtitle or "").strip()
    if (not title or title == "(无标题)"
            or not DIVIDER_TITLE_RE.match(sub)
            or DIVIDER_TITLE_RE.match(title)):
        return False
    return (not s.has_image and not s.has_table and not s.has_chart
            and s.paragraph_count == 2
            and s.body_len == len(title) + len(sub))


def _heading_display_name(s: "SlideInfo") -> str:
    """分标题展示名：纯数字标题拼上副标题章名（「02」+「升学与就业支持」
    →「02 升学与就业支持」）；反向形态（名称标题+编号正文）拼为
    「03 招生政策问答」；「第N节 名称」型直接用标题。"""
    title = (s.title or "").replace("\x0b", " ").replace("\n", " ").strip()
    sub = (s.subtitle or "").replace("\x0b", " ").replace("\n", " ").strip()
    if DIVIDER_TITLE_RE.match(title) and sub and not DIVIDER_TITLE_RE.match(sub):
        return f"{title} {sub}"[:24]
    if (sub and DIVIDER_TITLE_RE.match(sub) and title
            and title != "(无标题)" and not DIVIDER_TITLE_RE.match(title)):
        return f"{sub} {title}"[:24]
    return title[:24]


def _cn_to_int(v: str) -> int | None:
    """中文数字（一二…十、十一、二十三）转 int；阿拉伯数字串直接转。"""
    v = (v or "").strip()
    if not v:
        return None
    if v.isdigit():
        return int(v)
    if len(v) == 1:
        return _CN_DIGITS.index(v) if v in _CN_DIGITS else None
    if v[0] == "十":
        return 10 + (_CN_DIGITS.index(v[1]) if len(v) > 1 else 0)
    if v[0] in _CN_DIGITS:
        return _CN_DIGITS.index(v[0]) * 10 + (
            _CN_DIGITS.index(v[1]) if len(v) > 1 and v[1] in _CN_DIGITS else 0)
    return None


# 页编号路径：单级「01」/「1.1」/「1-1」/「1.1.2」（点分/连字符多级自带层级）
_NUMPATH_RE = re.compile(
    r"^\s*0*(\d{1,2})(?:\s*[.·\-－—]\s*0*(\d{1,2})"
    r"(?:\s*[.·\-－—]\s*0*(\d{1,2}))?)?\s*$")


def _page_numbering(s: "SlideInfo") -> tuple[int, ...] | None:
    """正文页编号路径（两级编号体系的归属证据）：副标题/标题上的
    「01」「1.1」「1-1」「第一章」「（一）」等 → tuple（如 (1,)、(1,2)、(1,1,2)）。
    一级分标题页自身返回 None（其编号是章锚，走 _divider_section_num）。
    孤立的正文数字（如某页正文首行出现的「12」）虽会被本函数识别，但由
    split_top_sections 的连续段判据在归属阶段排除，不会误立板块。"""
    if _is_section_heading(s):
        return None
    for text in ((s.subtitle or "").strip(), (s.title or "").strip()):
        t = (text or "").replace("\x0b", " ").strip()
        if not t:
            continue
        m = _NUMPATH_RE.match(t)
        if m and any(m.groups()):
            return tuple(int(g) for g in m.groups() if g)
        m = re.match(r"^\s*[（(]\s*([一二三四五六七八九十]+)\s*[)）]\s*$", t)
        if m:
            n = _cn_to_int(m.group(1))
            if n:
                return (n,)
        m = re.match(r"^\s*第([0-9一二三四五六七八九十]+)[章节部分篇]?\s*$", t)
        if m:
            n = _cn_to_int(m.group(1))
            if n:
                return (n,)
    return None


def split_top_sections(
    induced: list[dict], slides: Sequence["SlideInfo"],
) -> list[dict] | None:
    """两级编号树（第一级=板块，第二级=章）：在结构切章结果上，检测编号章内
    「连续 ≥2 页同号且异于章号」的编号页段（含其后的无号挂靠页），将其剥离为
    与各编号章平级的顶层板块——即稿子存在第二套编号体系（如 test3 稿编号章 01-10
    为章、02 招生和 03 问答板块剥离后与章平级）。

    归属规则（按优先级，0 token）：
    - 显式多级格式编号（1.1 / 1-2）→ 格式即层级，直接定父；
    - 原序连续性：分标题编号序列在原稿中连续成段 → 同属一个板块
      （主板块编号 = 其章编号序列起点，如章 01-02 → 板块 01）；
    - 章内连续 ≥2 页同号异号段 → 平级脱出为新板块（编号 = 成员角标）；
    - 无编号页（开篇/装饰/纲领等）角色保持开放、不强行归类，
      仅按原序挂靠跟随其所在板块；孤立单页异号数字（如「12」）→ 正文内容。
    节点含 "_num"（板块编号，顶层排序依据：编号升序 = 板块 01→02→03）。
    返回顶层节点列表 [{"title", "pages", "children", "_num"}]（children 非
    None 的为主板块，其子节点是编号章）；无剥离段时返回 None（单层稿）。"""
    info = {s.index: s for s in slides}
    top: list[dict] = []
    main_children: list[dict] = []
    seg_nodes: list[dict] = []
    main_ch_nums: list[int | None] = []

    def _seg_pages(g_pages: list[int], head_num: int | None) -> list[list[int]]:
        """章内剥离段：连续 ≥2 个同号异章编号页 + 其挂靠无号页。"""
        segs: list[list[int]] = []
        cur: list[int] = []
        owner: str = "chapter"   # 无号页跟随前一编号页归属
        for p in g_pages:
            s = info[p]
            if _is_section_heading(s):
                continue
            num = _page_numbering(s)
            first = num[0] if num else None
            if head_num is not None and first == head_num:
                owner = "chapter"
                if cur:
                    segs.append(cur)
                    cur = []
            elif first is not None:
                cur_first = _page_numbering(info[cur[0]]) if cur else None
                if cur and first != (cur_first[0] if cur_first else None):
                    segs.append(cur)
                    cur = []
                cur.append(p)
                owner = "seg"
            else:
                if owner == "seg":
                    cur.append(p)
                else:
                    if cur:
                        segs.append(cur)
                        cur = []
        if cur:
            segs.append(cur)
        return [seg for seg in segs
                if sum(1 for p in seg if _page_numbering(info[p])) >= 2]

    for g in induced:
        pages = list(g["pages"])
        head = pages[0] if _is_section_heading(info[pages[0]]) else None
        head_num = _divider_section_num(info[head]) if head else None
        segs = _seg_pages(pages, head_num) if head else []
        if segs:
            seg_set = {p for seg in segs for p in seg}
            rest = [p for p in pages if p not in seg_set]
            if len(rest) >= 2:  # 剥离后章仍需分标题+内容，否则放弃剥离
                for seg in segs:
                    t = (info[seg[0]].title or "").replace("\x0b", " ").strip()
                    num0 = _page_numbering(info[seg[0]])
                    seg_nodes.append({"title": t[:20] or "未命名板块",
                                      "pages": seg, "children": None,
                                      "_num": num0[0] if num0 else None})
                pages = rest
        if head:
            main_children.append({"title": g["title"], "pages": pages})
            main_ch_nums.append(head_num)
        elif not seg_nodes:
            main_children.append({"title": g["title"], "pages": pages})
            main_ch_nums.append(None)
        else:
            main_children.append({"title": g["title"], "pages": pages})
            main_ch_nums.append(None)
    if not seg_nodes:
        return None
    all_main = sorted(p for c in main_children for p in c["pages"])
    main_pages = [p for c in main_children for p in c["pages"]]
    # 主板块编号 = 其章编号序列起点（原序连续段，如章 01-10 → 板块 01）
    main_nums = sorted(n for n in main_ch_nums if n is not None)
    top.append({"title": f"正文主体（{len(main_children)}个章节）",
                "pages": main_pages, "children": main_children,
                "_all_pages": all_main,
                "_num": main_nums[0] if main_nums else None})
    top.extend(seg_nodes)
    top.sort(key=lambda t: min(t["pages"]))
    logger.info("两级编号树：剥离 %d 个顶层板块 %s（0 token）",
                len(seg_nodes), [t["title"] for t in top[1:]])
    return top


def _is_pure_divider(s: "SlideInfo") -> bool:
    """纯分隔页：标题是节号且几乎无正文内容（body_len ≤ 10 且无图表）。
    只看标题——副标题上的数字是子节角标（如「招生政策问答/03」），不是分隔页。"""
    if not DIVIDER_TITLE_RE.match((s.title or "").strip()):
        return False
    if s.has_image or s.has_table or s.has_chart:
        return False
    return s.body_len <= 10


def trim_internal_dividers(
    order: Sequence[int], induced: list[dict] | None,
    slides: Sequence["SlideInfo"], *,
    keep_chapter_start: bool = True,
) -> tuple[list[int], list[str]]:
    """删除章节内部的冗余纯分隔页，为制作目录做准备。

    只删除标题为节号且无正文的纯分隔页（如"01""第二章"），
    内容页全部保留。章首页的分隔页保留（作为章节锚点）。
    连续两个分隔页时删除后者（避免连排）。
    """
    info = {s.index: s for s in slides}
    if induced:
        chap_of = {}
        for i, g in enumerate(induced):
            for p in g["pages"]:
                chap_of[p] = i
    else:
        chap_of = {}
    # 第一遍：章内冗余分隔页去重
    result: list[int] = []
    removed: list[str] = []
    seen_chapters: set[int] = set()
    for p in order:
        s = info.get(p)
        if s and _is_pure_divider(s):
            ci = chap_of.get(p, -1)
            if ci in seen_chapters and ci != -1:
                removed.append(f"页{p}《{s.title or s.subtitle or '?'}》"
                               f"（章内冗余分隔页）")
                continue
            if ci != -1:
                seen_chapters.add(ci)
            if not keep_chapter_start:
                removed.append(f"页{p}《{s.title or s.subtitle or '?'}》（分隔页）")
                continue
        result.append(p)
    # 第二遍：连续分隔页去重——保留前一个，删后一个
    final: list[int] = []
    prev_was_divider = False
    for p in result:
        s = info.get(p)
        is_div = s is not None and _is_pure_divider(s)
        if is_div and prev_was_divider:
            removed.append(f"页{p}《{s.title or s.subtitle or '?'}》"
                           f"（与前页连续分隔页）")
            continue
        final.append(p)
        prev_was_divider = is_div
    return final, removed


def _has_org_signal(s: "SlideInfo") -> bool:
    """标题/副标题或正文任意位置含机构/事件关键词（校名、公司、发布会等）。

    真实封面的校名/公司常独占一行且排在作者行之后，副标题只取正文第一行
    时会漏掉（如标题「第5章 隐私保护」、首行作者、次行才是学院），
    故 analyze 阶段对全文扫描的结果记录在 s.org_signal。
    """
    if s.org_signal:
        return True
    text = f"{s.title or ''}{s.subtitle or ''}"
    if any(w in text for w in COVER_ORG_KEYWORDS):
        return True
    # 日期模式也是封面的正信号（封面常含"2024年X月"等日期）
    return bool(COVER_DATE_RE.search(text))


def _pick_cover(slides: Sequence["SlideInfo"]) -> "SlideInfo | None":
    """内容识别定位封面（不依赖位置）。

    候选门控：短而无章节词的唯一标题、无项目符号、段落 ≤12、非数字分隔页。
    优先级：机构正信号 >（段落数, 字数, 首页先验, 页号）。零候选才以严格稀疏
    门控回退首页先验；乱序稿首页为正文则不标。
    """
    freq = Counter(_norm_title(s.title) for s in slides if s.special is None)
    candidates = [
        s for s in slides
        if _looks_like_cover(s)
        and freq.get(_norm_title(s.title), 0) == 1
        and s.bullet_count == 0
        and s.paragraph_count <= COVER_MAX_PARAS
        and not _is_divider(s)
    ]
    if candidates:
        return min(candidates, key=lambda s: (
            0 if _has_org_signal(s) else 1,
            s.paragraph_count, s.body_len, 0 if s.index == 1 else 1, s.index))
    first = slides[0]
    if (first.special is None and first.bullet_count == 0
            and first.paragraph_count <= COVER_FALLBACK_MAX_PARAS
            and first.body_len <= COVER_FALLBACK_BODY_MAX):
        return first
    return None


def _slide_stats(slide) -> tuple[int, int, int]:
    """正文特征 (字数, 段落数, 项目符号数)，不含标题形状。"""
    body_len = paragraph_count = bullet_count = 0
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            text = para.text.strip()
            if not text or shape == slide.shapes.title:
                continue
            body_len += len(text)
            paragraph_count += 1
            if para.level and para.level > 0:
                bullet_count += 1
    return body_len, paragraph_count, bullet_count


def _extract_subtitle_and_numbers(slide, title: str) -> tuple[str, list[int]]:
    """提取 (副标题=正文首个非空短行, 段首编号列表)；二者都是廉价区分信号。"""
    subtitle = ""
    numbers: list[int] = []
    for shape in slide.shapes:
        if not shape.has_text_frame or shape == slide.shapes.title:
            continue
        for para in shape.text_frame.paragraphs:
            text = para.text.strip()
            if not text or text == title:
                continue
            if not subtitle and len(text) <= TEXT_SNIPPET_LEN // 2:
                subtitle = text
            m = ORDINAL_LEAD.match(text)
            if m:
                v = int(m.group(1))
                if 1 <= v <= ORDINAL_MAX:
                    numbers.append(v)
    return subtitle, numbers


def _extract_ordinal(subtitle: str, numbers: list[int]) -> int | None:
    """同标题组内序号；页内列举（≥3 个不同编号）或无信号时返回 None。"""
    distinct = set(numbers)
    if len(distinct) >= ORDINAL_INLINE_MIN:
        return None
    m = ORDINAL_LEAD.match(subtitle or "")
    if m:
        return int(m.group(1))
    return numbers[0] if len(distinct) == 1 else None


def _stage_score(text: str) -> dict[int, int]:
    scores: dict[int, int] = {}
    for stage, words in STAGE_KEYWORDS.items():
        score = sum(text.count(w) * wt for w, wt in words)
        if score:
            scores[stage] = score
    return scores


def _classify_stage(title: str, subtitle: str, snippet: str) -> int | None:
    """汇报阶段 0..5；副标题优先（同标题组靠副标题区分），平局取更早阶段。"""
    scores = _stage_score(subtitle or "")
    if not scores:
        scores = _stage_score(f"{title} {snippet[:60]}")
    if not scores:
        return None
    best = max(scores.values())
    return min(k for k, v in scores.items() if v == best)


def analyze_pptx(path: str) -> AnalysisResult:
    """读取 PPT，识别每页主题、特殊页、封面与章节结构，返回结构化摘要。"""
    p = validate_pptx_path(path)
    logger.info("analyze start: %s", p)
    prs = Presentation(str(p))
    slides_info: list[SlideInfo] = []
    for i, slide in enumerate(prs.slides, start=1):
        title = _slide_title(slide)
        full_text = _slide_text(slide)
        snippet = full_text.replace("\n", " ")[:TEXT_SNIPPET_LEN]
        try:
            layout = slide.slide_layout.name
        except Exception:
            layout = ""
        body_len, para_cnt, bullet_cnt = _slide_stats(slide)
        subtitle, numbers = _extract_subtitle_and_numbers(slide, title)
        org_signal = any(w in full_text for w in COVER_ORG_KEYWORDS)
        slides_info.append(
            SlideInfo(
                index=i,
                title=title or "(无标题)",
                text_snippet=snippet,
                n_shapes=len(slide.shapes),
                has_image=_has_image(slide),
                has_table=_has_table(slide),
                has_chart=_has_chart(slide),
                layout=layout,
                special=_classify_special(title),
                body_len=body_len,
                paragraph_count=para_cnt,
                bullet_count=bullet_cnt,
                subtitle=subtitle,
                ordinal=_extract_ordinal(subtitle, numbers),
                stage=_classify_stage(title or "", subtitle, snippet),
                org_signal=org_signal,
            )
        )
    # 封面：内容识别优先（对乱序健壮），零候选才严格回退首页先验。
    if slides_info:
        cover = _pick_cover(slides_info)
        if cover is not None:
            cover.special = "封面"
    chapters = detect_chapters(slides_info)
    logger.info("analyze done: %d slides, %d chapters", len(slides_info), len(chapters))
    return AnalysisResult(
        file=str(p), slide_count=len(slides_info), slides=slides_info,
        chapters=chapters)


def detect_chapters(slides: Sequence[SlideInfo]) -> list[Chapter]:
    """章节边界：优先编号模式（第X章/1.1/Chapter N），否则按相同标题聚合，
    特殊页独立切章。"""
    if not slides:
        return []
    numbered = [any(pat.match(s.title) for pat in CHAPTER_PATTERNS) for s in slides]
    if any(numbered):
        return _chapters_by_number(slides, numbered)
    return _chapters_by_title_grouping(slides)


def _chapters_by_number(slides: Sequence[SlideInfo], numbered: Sequence[bool]) -> list[Chapter]:
    chapters: list[Chapter] = []
    current: Chapter | None = None
    for s, is_chap in zip(slides, numbered):
        if is_chap or s.special is not None:
            # 编号页或特殊页触发切分
            if current is not None:
                current.end = s.index - 1
                if current.end >= current.start:
                    chapters.append(current)
                current = None
            chapters.append(Chapter(title=s.title, start=s.index, end=s.index))
        else:
            if current is None:
                current = Chapter(title="正文", start=s.index, end=s.index)
            else:
                current.end = s.index
    if current is not None:
        chapters.append(current)
    return chapters or [Chapter(title="前言", start=1, end=slides[-1].index)]


def _chapters_by_title_grouping(slides: Sequence[SlideInfo]) -> list[Chapter]:
    chapters: list[Chapter] = []
    current_title: str | None = None
    current_start: int | None = None
    for s in slides:
        title = s.title.strip() or "(无标题)"
        if s.special is not None:  # 特殊页独立成章
            if current_title is not None:
                chapters.append(Chapter(current_title, current_start, s.index - 1))
                current_title = current_start = None
            chapters.append(Chapter(title, s.index, s.index))
            continue
        if title != current_title:  # 标题变化即新章
            if current_title is not None:
                chapters.append(Chapter(current_title, current_start, s.index - 1))
            current_title, current_start = title, s.index
    if current_title is not None:
        chapters.append(Chapter(current_title, current_start, slides[-1].index))
    return chapters or [Chapter(title="前言", start=1, end=slides[-1].index)]


def _norm_title(t: str) -> str:
    """标题归一化：去全部空白。"""
    return re.sub(r"\s+", "", t or "")


def refine_within_groups(
    order: Sequence[int], slides: Sequence[SlideInfo],
    groups_map: dict | None = None, *, preserve_narrative: bool = False,
) -> tuple[list[int], int]:
    """组内确定性证据精排，组的整体位置不动，特殊页不参与；异常整体脱出。

    1. 序号槽位回填（③）：带编号页在其已占槽位间按编号升序；编号重复跳过该组；
    2. 无编号连续块按阶段理顺（④兜底）：块内 ≥2 个不同阶段才动作，无标签置
       块尾，不跨越已归位编号页。
    分组来源：groups_map（章 id）> 相同标题。
    preserve_narrative=True（章内已由 LLM 按目的精排）时只做编号归位、跳过
    阶段 2，避免固定学术阶段序把目的差异化叙事（科普/答辩/总结…）抹平。
    返回 (新序, 修正数)。
    """
    try:
        n = len(slides)
        norm = normalize_order(order, n)
        info = {s.index: s for s in slides}

        def group_key(idx: int):
            s = info[idx]
            if s.special is not None:
                return None
            if groups_map is not None:
                gid = groups_map.get(idx)
                return ("induced", gid) if gid is not None else None
            return ("title", _norm_title(s.title))

        groups: dict = {}
        for pos, idx in enumerate(norm):
            k = group_key(idx)
            if k is not None:
                groups.setdefault(k, []).append(pos)

        result = list(norm)
        fixed = 0
        for positions in groups.values():
            if len(positions) < 2:
                continue
            # 阶段 1：序号槽位回填
            tagged = [(p, result[p], info[result[p]].ordinal) for p in positions
                      if info[result[p]].ordinal is not None]
            if len(tagged) >= 2:
                ords = [o for _, _, o in tagged]
                if len(set(ords)) == len(ords) and ords != sorted(ords):
                    sorted_pages = [idx for _, idx, _ in sorted(tagged, key=lambda x: x[2])]
                    for (p, _, _), idx in zip(tagged, sorted_pages):
                        result[p] = idx
                    fixed += 1
            # 阶段 2：无编号连续块按阶段理顺（不跨编号页）
            # LLM 已按目的排出叙事序时跳过，防止目的差异被固定阶段序覆盖
            if preserve_narrative:
                continue
            no_tag = [p for p in positions if info[result[p]].ordinal is None]
            i = 0
            while i < len(no_tag):
                j = i
                while j + 1 < len(no_tag) and no_tag[j + 1] == no_tag[j] + 1:
                    j += 1
                block = no_tag[i:j + 1]
                if len(block) >= 2:
                    pages = [result[p] for p in block]
                    known = [info[x].stage for x in pages if info[x].stage is not None]
                    if len(set(known)) >= 2:
                        ranked = [x for _, x in sorted(
                            enumerate(pages),
                            key=lambda t: (info[t[1]].stage is None,
                                           info[t[1]].stage if info[t[1]].stage is not None else 0,
                                           t[0]))]
                        if ranked != pages:
                            for p, x in zip(block, ranked):
                                result[p] = x
                            fixed += 1
                i = j + 1
        return normalize_order(result, n), fixed
    except Exception as e:
        logger.warning("组内证据精排异常，脱出并保留入参顺序: %s", e)
        return list(order), 0


def _compose_base_with_inner_orders(
    induced: list[dict], slides: Sequence[SlideInfo], purpose: str,
    fallback_order: Sequence[int], chapters: Sequence[Chapter] | None,
    llm_kw: dict, trace: dict, inner_llm: bool = True,
) -> list[int]:
    """大 PPT 逐章排序：每章统一走 LLM 按目的精排（两级编号树的归属已在
    split_top_sections 切章层定型，章内不再做单层绑定）；LLM 不可用回退粗排
    相对序。章节块本身由上游切定，这里只在块内调整讲述顺序，不跨块移动。"""
    n = len(slides)
    kind_of = {s.index: s.special for s in slides}
    fixed = {i for i in range(1, n + 1) if kind_of.get(i) in SPECIAL_KINDS}
    pos_fb = {idx: p for p, idx in enumerate(fallback_order)}
    base = [i for i in fallback_order if i in fixed]
    calls = ok = 0
    for g in induced:
        pages = [p for p in g["pages"] if p not in fixed]
        seq = None
        if inner_llm and 2 <= len(pages) <= WITHIN_CHAPTER_LLM_MAX:
            calls += 1
            seq = order_pages_within_chapter_by_llm(
                g["title"], pages, slides, purpose, chapters=chapters, **llm_kw)
        if seq is not None:
            ok += 1
            base.extend(seq)
        else:
            base.extend(sorted(pages, key=lambda p: pos_fb.get(p, 0)))
    trace["within_chapter_calls"] = calls
    trace["within_chapter_ok"] = ok
    return base


def resolve_order(
    analysis: "AnalysisResult", purpose: str, *,
    sort_level: str = "chapter", use_llm: bool = False,
    model: str = "", base_url: str = "", api_key: str = "",
    anchor_special: bool = False, smart: bool = False,
) -> tuple[list[int], dict]:
    """排序唯一权威：粗排 → 章节检定（证据升级链）→ 组内精排，返回 (1-based 序, trace)。

    smart 证据链：⓪ 原 PPT 分标题结构切章（0 token，章块不拆散）→ ① 规则
    标题归组（0 token）→ ② LLM 单次章节聚类；章序由槽位投票决定，章内由
    LLM/证据精排调整讲述顺序；③ 特殊页锚定兜底。前级不可用才升级。
    """
    slides = analysis.slides
    llm_kw = {"model": model, "base_url": base_url, "api_key": api_key}
    trace: dict = {
        "base": "llm" if use_llm else "keyword",
        "chapter_source": "none",
        "induce_mode": "none",
        "induced_chapter_count": 0, "induced_chapters": [],
        "within_chapter_calls": 0, "within_chapter_ok": 0,
        "groups_refined": 0, "anchor_applied": False,
        "top_sections": [],
        "heading_bound": [], "heading_dropped": [],
    }
    # 粗排：大 PPT smart 路径跳过全页模型粗排（长 prompt 不可靠，章序/章内
    # 另有小粒度调用负责），用零成本关键词序作位置兜底。
    skip_full_llm = smart and len(slides) > INDUCE_SINGLE_MAX
    if use_llm and not skip_full_llm:
        order = plan_reorder_by_llm(
            slides, purpose, chapters=analysis.chapters,
            sort_level=sort_level, **llm_kw)
    else:
        order = plan_reorder_by_purpose(slides, purpose, sort_level=sort_level)
        if skip_full_llm:
            trace["base"] = "keyword(skipped_full_llm)"

    if not smart:
        if anchor_special:
            order = anchor_special_pages(order, slides)
            trace["anchor_applied"] = True
        order, hm, hd = bind_section_headings(order, slides)
        trace["heading_bound"], trace["heading_dropped"] = hm, hd
        return order, trace

    # smart 证据升级链
    induced: list[dict] | None = None
    large = len(slides) > INDUCE_SINGLE_MAX
    # ⓪ 结构分标题切章（0 token，章节成块、章内保原序）——最高优先
    induced = induce_chapters_by_structure(slides)
    if induced is not None:
        trace["induce_mode"] = "structural_sections"
    else:
        induced = induce_chapters_by_rules(slides)
    if induced is not None and trace["induce_mode"] == "none":
        trace["induce_mode"] = "rule_titles"
    if induced is None and use_llm:
        # ② LLM 单次章节聚类
        trace["induce_mode"] = "single"
        induced = induce_chapters_by_llm(
            slides, purpose, chapters=analysis.chapters, **llm_kw)

    structural = trace["induce_mode"] == "structural_sections"
    if induced is not None:
        # 章序按目的重排（LLM 投票或确定性兜底）；结构章同样只排章序
        votes_board = 3 if large else 1      # 板块候选少（≤3），3 票已稳
        votes_chapter = 5 if large else 1    # 章候选多（可达10+），5 票压波动
        pos_base = {idx: p for p, idx in enumerate(order)}
        # 两级编号树：编号章内连续同号异章段剥离为顶层板块；板块间先做目的
        # 投票（LLM 优先，失败回退编号升序 板块 01→02/03），板块内章间按目的
        # 投票，单层稿走原路径
        top = split_top_sections(induced, slides) if structural else None
        if top is not None and len(top) < 2:
            top = None
        if top is not None:
            # 板块间目的投票：板块以「分标题展示名 + 自身页关键词」参与，
            # 与编号章投票同链路（order_chapters_by_llm）。失败/未用 LLM 时
            # 回退编号升序（板块 01 主体 → 板块 02/03），编号缺失按原序位置跟随。
            board_groups: list[dict] = []
            for t in top:
                head_pages = t.get("_all_pages", t["pages"])
                head = slides[head_pages[0] - 1] if head_pages else None
                title = t["title"]
                if head is not None and (_is_section_heading(head)
                                         or _is_reverse_numbered_heading(head)):
                    title = _heading_display_name(head)
                board_groups.append({"title": title, "pages": t["pages"]})
            board_order = (order_chapters_by_llm(
                               board_groups, slides, purpose,
                               votes=votes_board, **llm_kw)
                           if use_llm else None)
            if board_order is not None:
                top = [top[i] for i in board_order]
                logger.info("板块间目的投票：%s → %s",
                            [g["title"] for g in board_groups],
                            [board_groups[i]["title"] for i in board_order])
            else:
                top.sort(key=lambda t: (t.get("_num") is None,
                                        t.get("_num") if t.get("_num") is not None
                                        else min(pos_base.get(p, len(slides))
                                                 for p in t.get("_all_pages",
                                                                t["pages"]))))
            trace["top_sections"] = [
                {"title": t["title"], "pages": t.get("_all_pages", t["pages"])}
                for t in top]
            flat: list[dict] = []
            for t in top:
                if t.get("children") is not None:
                    sub = list(t["children"])
                    # 板块开篇（首个编号章前的引言组，含板块分标题页）恒定
                    # 板块首位、不参与目的投票——编号章只在彼此之间按目的排序，
                    # 保证目录上「板块 → 其章」的父子相邻结构
                    lead: dict | None = None
                    if len(sub) >= 2 and sub[0]["title"] == "开篇引言":
                        lead = sub.pop(0)
                    ch_order = (order_chapters_by_llm(
                                    sub, slides, purpose, votes=votes_chapter,
                                    **llm_kw)
                                if use_llm else None)
                    if ch_order is not None:
                        sub = [sub[i] for i in ch_order]
                    else:
                        sub.sort(key=lambda g: min(
                            pos_base.get(p, len(slides)) for p in g["pages"]))
                    if lead is not None:
                        sub.insert(0, lead)
                    flat.extend(sub)
                else:
                    flat.append({"title": t["title"], "pages": t["pages"]})
            induced = flat
        else:
            ch_order = (order_chapters_by_llm(
                            induced, slides, purpose, votes=votes_chapter,
                            **llm_kw)
                        if use_llm else None)
            if ch_order is not None:
                induced = [induced[i] for i in ch_order]
            else:
                # 排章不可用：按粗排中各章最早出现位置定章序（确定性兜底）
                induced.sort(key=lambda g: min(pos_base.get(p, len(slides))
                                               for p in g["pages"]))
        if large:
            composed = _compose_base_with_inner_orders(
                induced, slides, purpose, order, analysis.chapters,
                llm_kw, trace, inner_llm=use_llm)
            order = assemble_by_chapters(composed, slides, induced)
        else:
            order = assemble_by_chapters(order, slides, induced)
        if trace["chapter_source"] == "none":
            trace["chapter_source"] = ("structural_sections" if structural
                                       else "rule_titles"
                                       if trace["induce_mode"] == "rule_titles"
                                       else "llm_induced")
        trace.update(
            induced_chapter_count=len(induced),
            induced_chapters=[
                {"title": g["title"], "pages": g["pages"]} for g in induced],
        )
    else:  # 检定全部不可用 → 规则锚定兜底（仍产出可用顺序）
        order = anchor_special_pages(order, slides)
        trace.update(chapter_source="anchor_fallback", anchor_applied=True)

    groups_map = None
    if induced is not None:
        groups_map = {p: i for i, g in enumerate(induced) for p in g["pages"]}
    # 章内/全序已由 LLM 按目的排出叙事序时，只保留编号归位，跳过阶段重排
    narrative_by_llm = bool(
        use_llm and (trace.get("within_chapter_ok", 0) > 0 or not large))
    order, fixed = refine_within_groups(
        order, slides, groups_map, preserve_narrative=narrative_by_llm)
    trace["groups_refined"] = fixed
    if structural:
        order = lead_section_headings(order, slides, groups_map)
    order, hm, hd = bind_section_headings(order, slides, groups_map)
    trace["heading_bound"], trace["heading_dropped"] = hm, hd
    # 特殊页确定性锚定：封面/目录恒居首，总结/参考文献/致谢恒居尾
    order = anchor_special_pages(order, slides)
    trace["anchor_applied"] = True
    return order, trace


# --------------------------------------------------------------------------- #
# 关键词启发式粗排
# --------------------------------------------------------------------------- #
PURPOSE_SEEDS: dict[str, list[str]] = {
    "投资": ["商业", "价值", "市场", "规模", "收入", "融资", "盈利", "增长", "客户", "成本"],
    "汇报": ["总结", "成果", "指标", "结论", "建议", "进度", "计划", "目标", "里程碑", "复盘"],
    "技术": ["技术", "架构", "系统", "算法", "实现", "设计", "开发", "代码", "模块", "性能"],
    "教学": ["概念", "原理", "基础", "教程", "知识点", "示例", "练习", "讲解", "步骤", "公式"],
    "产品": ["需求", "用户", "功能", "体验", "交互", "原型", "迭代", "场景", "产品", "版本"],
}


def plan_reorder_by_purpose(
    slides: Sequence[SlideInfo], purpose: str, sort_level: str = "chapter"
) -> list[int]:
    """目的种子词启发式打分排序，返回 1-based 序；未知目的保持原序。

    每页 hits = TextRank 种子词权重和*10 + 种子词词频；chapter 模式按章
    （最高 prio, 平均 hits）排、章内原序；slide 模式按单页排。特殊页 prio：
    封面 3/目录 2/总结 1/普通 0/致谢 -1。
    """
    if not slides:
        return []
    sort_level = sort_level.lower()
    if sort_level not in ("chapter", "slide"):
        raise ValueError(f"sort_level 必须是 'chapter' 或 'slide'，收到: {sort_level}")
    purpose_lower = purpose.lower()
    matched_seeds: list[str] = []
    for key, seeds in PURPOSE_SEEDS.items():
        if key in purpose_lower:
            matched_seeds.extend(seeds)
    if not matched_seeds:
        return [s.index for s in slides]

    slide_hits: dict[int, float] = {}
    slide_prio: dict[int, int] = {}
    for s in slides:
        if _HAS_JIEBA:
            text = s.title + " " + s.text_snippet
            kw_weight = dict(jieba.analyse.textrank(text, topK=20, withWeight=True))
            seed_weight = sum(kw_weight.get(seed, 0.0) for seed in matched_seeds)
            freq = sum(1 for w in jieba.lcut(text) if w in matched_seeds)
            hits = seed_weight * 10 + freq
        else:
            text = (s.title + " " + s.text_snippet).lower()
            hits = sum(1 for kw in matched_seeds if kw in text)
        slide_hits[s.index] = hits
        slide_prio[s.index] = (
            3 if s.special == "封面" else 2 if s.special == "目录"
            else -1 if s.special == "致谢" else 1 if s.special == "总结" else 0)

    if sort_level == "slide":
        ordered = sorted(slides,
                         key=lambda s: (slide_prio[s.index], slide_hits[s.index], -s.index),
                         reverse=True)
        return [s.index for s in ordered]

    chapters = detect_chapters(slides)

    def chapter_score(ch: Chapter) -> tuple[int, float, int]:
        idxs = list(range(ch.start, ch.end + 1))
        return (max(slide_prio[i] for i in idxs),
                sum(slide_hits[i] for i in idxs) / len(idxs), -ch.start)

    result: list[int] = []
    for ch in sorted(chapters, key=chapter_score, reverse=True):
        result.extend(range(ch.start, ch.end + 1))
    return result


# --------------------------------------------------------------------------- #
# 低 Token 摘要与大模型全页排序（Python 拆分 + 大模型决策）
# --------------------------------------------------------------------------- #
def build_slide_briefs(
    slides: Sequence[SlideInfo], chapters: Sequence[Chapter] | None = None,
    top_k: int = 8,
) -> list[dict]:
    """每页压成低 token 结构化摘要（统计特征+关键词，不传正文）。"""
    idx_to_chapter: dict[int, str] = {}
    if chapters:
        for ch in chapters:
            for idx in range(ch.start, ch.end + 1):
                idx_to_chapter[idx] = ch.title

    briefs: list[dict] = []
    for s in slides:
        if _HAS_JIEBA:
            kws = jieba.analyse.textrank(
                s.title + " " + s.text_snippet, topK=top_k, withWeight=False)
        else:
            kws = []
        briefs.append({
            "index": s.index,
            "title": s.title,
            "special": s.special,
            "subtitle": s.subtitle,
            "stage": STAGE_NAMES.get(s.stage, ""),
            "ordinal": s.ordinal,
            "body_len": s.body_len,
            "paragraphs": s.paragraph_count,
            "bullets": s.bullet_count,
            "has_image": s.has_image,
            "has_table": s.has_table,
            "has_chart": s.has_chart,
            "layout": s.layout,
            "keywords": kws,
            "chapter": idx_to_chapter.get(s.index, ""),
        })
    return briefs


def _build_llm_prompt(briefs: list[dict], purpose: str) -> str:
    """低 token 全页排序 prompt：只给统计特征+关键词+章节归属，不给正文。"""
    lines = [
        f"你是 PPT 重组助手。下列页面可能处于乱序状态，请根据汇报目的「{purpose}」和内容逻辑重新排列，不要照抄当前编号。",
        "排序规则：",
        "1. 封面放最前，目录放第二，致谢/结束页放最后，参考文献靠近末尾；",
        "2. 其余页面按“背景介绍→内容/成果→问题论证→案例实践→过程分析→结论建议”的汇报逻辑排列；",
        "3. 同一章节、同一主题的页面尽量保持连续；标题相同的多页，按其副标题、阶段标签（背景→结论）、组内序号（若标注）排先后；",
        "4. 内容丰富（字数多、有图表）的页面通常更重要；",
        "5. 只有每行【开头的编号】是页码；括号里的字数/段数、“序号:N”、页数等数字绝对不能当作页码；",
        f"6. 只返回 JSON：{{\"order\": [页码列表]}}，必须恰好使用 1..{len(briefs)} 各一次、不重复不遗漏。",
        "",
        "页面列表（开头编号=页码 | 标题 [特殊页] | 副标题/阶段/序号 | 字数/段/项 | 图/表 | 关键词 | 章节）：",
    ]
    for b in briefs:
        tag = f"[{b['special']}]" if b["special"] else ""
        hints = []
        if b.get("subtitle"):
            hints.append(b["subtitle"])
        if b.get("stage"):
            hints.append(f"阶段:{b['stage']}")
        if b.get("ordinal") is not None:
            hints.append(f"序号:{b['ordinal']}")
        hint_str = " / ".join(hints) if hints else "-"
        stats = f"{b['body_len']}字/{b['paragraphs']}段/{b['bullets']}项"
        media_parts = []
        if b["has_image"]:
            media_parts.append("图")
        if b["has_table"]:
            media_parts.append("表")
        if b["has_chart"]:
            media_parts.append("图表")
        media = "、".join(media_parts) if media_parts else "无"
        kw_str = "、".join(b["keywords"]) if b["keywords"] else "无"
        ch = b.get("chapter", "")
        ch_str = f"章:{ch}" if ch else "章:无"
        lines.append(f"{b['index']}. {b['title']} {tag} | {hint_str} | {stats} | {media} | {kw_str} | {ch_str}")
    return "\n".join(lines)


def _llm_chat_json(
    prompt: str, *, model: str = "", base_url: str = "", api_key: str = "",
    temperature: float = 0.1, timeout: float = 90.0,
) -> dict | None:
    """调用 OpenAI 兼容接口（Ollama 等）解析 JSON；任何失败（含超时）返回 None。"""
    try:
        from openai import OpenAI
        client = OpenAI(api_key=api_key or os.getenv("OPENAI_API_KEY", ""),
                        base_url=base_url or os.getenv("OPENAI_BASE_URL") or None,
                        timeout=timeout)
    except ImportError:
        logger.warning("openai 库未安装")
        return None
    try:
        resp = client.chat.completions.create(
            model=model or os.getenv("OPENAI_MODEL", ""),
            messages=[{"role": "user", "content": prompt}],
            temperature=temperature,
            response_format={"type": "json_object"},
        )
        return json.loads(resp.choices[0].message.content.strip())
    except Exception as e:
        logger.warning("大模型调用失败: %s", e)
        return None


def plan_reorder_by_llm(
    slides: Sequence[SlideInfo], purpose: str, *,
    chapters: Sequence[Chapter] | None = None,
    model: str = "", base_url: str = "", api_key: str = "",
    sort_level: str = "chapter",
) -> list[int]:
    """低 token 大模型全页排序；非法/不可用回退关键词序。"""
    briefs = build_slide_briefs(slides, chapters=chapters)
    n = len(slides)
    data = _llm_chat_json(
        _build_llm_prompt(briefs, purpose),
        model=model, base_url=base_url, api_key=api_key)
    if data is not None:
        order = data.get("order", [])
        if sorted(order) == list(range(1, n + 1)):  # 必须是 1..N 排列
            logger.info("llm reorder success: %s", order)
            return order
        logger.warning("大模型返回的顺序非法，回退到关键词排序: %s", order)
    return plan_reorder_by_purpose(slides, purpose, sort_level=sort_level)


# --------------------------------------------------------------------------- #
# 章节检定（LLM 单次聚类）→ 排章投票 → 章内小粒度精排
# --------------------------------------------------------------------------- #
INDUCED_MIN_CHAPTERS = 2
INDUCED_MAX_CHAPTERS = 12
INDUCE_SINGLE_MAX = 30        # > 此值跳过全页模型粗排，改用章/章内小粒度调用
WITHIN_CHAPTER_LLM_MAX = 24   # 章内 LLM 精排允许的最大章页数


def _brief_line(b: dict) -> str:
    tag = f"[{b['special']}]" if b.get("special") else ""
    hints = " / ".join(x for x in (
        b.get("subtitle", ""),
        f"阶段:{b['stage']}" if b.get("stage") else "",
        f"序号:{b['ordinal']}" if b.get("ordinal") is not None else "",
    ) if x) or "-"
    kw = "、".join(b.get("keywords", [])) or "无"
    return f"{b['index']}. {b['title']} {tag} | {hints} | {kw}"


def _build_induce_prompt(briefs: list[dict], purpose: str) -> str:
    """第一遍：把（可能乱序的）页面聚成逻辑章节。"""
    lines = [
        f"你是 PPT 结构分析专家。这份 PPT 没有可解析的目录页、页面可能已乱序。"
        f"请依据每页的标题/副标题/关键词，把全部页面聚成 3..{INDUCED_MAX_CHAPTERS} 个逻辑章节，"
        f"使其适合用于「{purpose}」的汇报。",
        "要求：",
        "1. 主题相近、副标题同属一个主题的页归入同一章；封面、目录、参考文献、致谢各自单独成组；",
        "2. 每个页面恰好属于一个章节，不重复、不遗漏；",
        "3. 章节标题用 4-10 个字概括（如“研究背景”“系统实现”“测试验证”）；",
        "4. 只有行首编号是页码，括号里的字数/序号/页数等数字不是页码；",
        f"5. 只返回 JSON：{{\"chapters\":[{{\"title\":\"章节名\",\"pages\":[页码...]}},...]}}，"
        f"pages 合起来必须恰好是 1..{len(briefs)} 各一次。",
        "",
        "页面列表：",
    ]
    lines.extend(_brief_line(b) for b in briefs)
    return "\n".join(lines)


def induce_chapters_by_llm(
    slides: Sequence[SlideInfo], purpose: str, *,
    chapters: Sequence[Chapter] | None = None,
    model: str = "", base_url: str = "", api_key: str = "",
) -> list[dict] | None:
    """单次聚类出逻辑章节 [{"title","pages"}]；章数/覆盖校验失败返回 None。"""
    n = len(slides)
    if n < 2:
        return None
    briefs = build_slide_briefs(slides, chapters=chapters)
    data = _llm_chat_json(
        _build_induce_prompt(briefs, purpose),
        model=model, base_url=base_url, api_key=api_key)
    if not data:
        return None
    raw = data.get("chapters")
    if not isinstance(raw, list) or not (INDUCED_MIN_CHAPTERS <= len(raw) <= INDUCED_MAX_CHAPTERS):
        logger.warning("章节发现：返回章节数非法: %r", raw if raw is None else len(raw))
        return None
    groups: list[dict] = []
    all_pages: list[int] = []
    for item in raw:
        if not isinstance(item, dict):
            return None
        pages = []
        for p in item.get("pages", []):
            try:
                v = int(p)
            except (TypeError, ValueError):
                return None
            if 1 <= v <= n and v not in pages:
                pages.append(v)
        if not pages:
            continue
        title = str(item.get("title", "")).strip()[:20] or "未命名章节"
        groups.append({"title": title, "pages": pages})
        all_pages.extend(pages)
    if (len(groups) < INDUCED_MIN_CHAPTERS
            or sorted(all_pages) != list(range(1, n + 1))):
        logger.warning("章节发现：页面覆盖校验失败（%d 组/%d 页），脱出", len(groups), len(all_pages))
        return None
    logger.info("llm induced %d chapters: %s",
                len(groups), [g["title"] for g in groups])
    return groups


# 叙事槽位：模型只做“章→槽”语义归类，槽序由目的模板强制（风格硬约束）
SLOT_MEANING: dict[str, str] = {
    "cover": "封面/开场", "toc": "目录",
    "background": "研究背景与意义", "related": "国内外现状/相关基础",
    "requirement": "需求/问题/目标", "solution": "总体方案/总体设计",
    "keytech": "关键技术/原理", "implement": "系统实现/方法/产品",
    "test": "测试与验证", "case": "应用案例/市场",
    "result": "成果/价值/创新点", "summary": "总结与展望",
    "reference": "参考文献", "thanks": "致谢/结尾",
}

# 易混槽位判别口径——近义工科章小模型常连环错槽，随 prompt 下发
SLOT_HINTS: dict[str, str] = {
    "background": "项目由来、为什么做、研究意义",
    "related": "别人做过什么、国内外现状综述、理论基础",
    "requirement": "要解决什么问题、需求清单、目标指标（还没谈怎么解决）",
    "solution": "总体思路、架构设计、模块划分、方案选型（设计图纸，尚未落地）",
    "keytech": "难点攻关、核心算法、技术原理（难在哪、原理是什么）",
    "implement": "系统怎么开发出来的、代码/模块实现、平台产品落地、部署（实物已做出）",
    "test": "怎么验证、测试用例、实验数据、结果对比",
    "case": "具体使用场景、典型案例、客户/市场应用（别人拿去怎么用）",
    "result": "创新点、成果清单、价值收益、论文专利（最终得到什么）",
}
SLOT_DISTINGUISH = [
    "需求 vs 方案：讲“要什么/解决什么问题”归 requirement；讲“怎么设计、架构如何”归 solution。",
    "方案 vs 关键技术：整体架构/模块划分/选型归 solution；某个难点的原理、算法攻关归 keytech。",
    "关键技术 vs 系统实现：讲“原理/难点”归 keytech；讲“开发落地过程、代码实现、平台部署”归 implement。",
    "标题原词优先：含“方案/总体设计/架构”多为 solution，含“关键技术/原理/算法”多为 keytech，"
    "含“实现/开发/平台/部署”多为 implement；正文细节只作辅助，不要被个别句子带偏。",
    "案例 vs 成果：具体使用场景/客户应用归 case；指标、创新点、价值、论文专利归 result。",
]

PURPOSE_SLOT_TEMPLATES: dict[str, list[str]] = {
    "技术汇报": ["cover", "toc", "background", "related", "requirement",
                "solution", "keytech", "implement", "test", "case",
                "result", "summary", "reference", "thanks"],
    "投资路演": ["cover", "toc", "result", "case", "implement", "solution",
                "keytech", "requirement", "related", "background", "test",
                "summary", "reference", "thanks"],
    "教学讲解": ["cover", "toc", "background", "related", "requirement",
                "keytech", "solution", "implement", "case", "test",
                "result", "summary", "reference", "thanks"],
    # 科普：背景意义先行，通俗原理+身边应用，弱化实验/参考
    "科普宣讲": ["cover", "toc", "background", "related", "requirement",
                "keytech", "solution", "case", "result", "implement",
                "summary", "test", "reference", "thanks"],
    # 学术答辩：问题→方法→实验→结果→结论，参考文献靠后但保留
    "学术答辩": ["cover", "toc", "background", "related", "requirement",
                "keytech", "solution", "implement", "test", "result",
                "case", "summary", "reference", "thanks"],
    # 工作总结/述职：成果先行，再讲做了什么，最后问题与计划
    "工作总结": ["cover", "toc", "result", "implement", "case", "background",
                "requirement", "solution", "keytech", "test", "summary",
                "related", "reference", "thanks"],
    # 项目汇报：背景需求→方案技术→落地→测试→成果
    "项目汇报": ["cover", "toc", "background", "requirement", "solution",
                "keytech", "implement", "test", "result", "case",
                "summary", "related", "reference", "thanks"],
    # 产品发布：痛点→产品→核心功能→演示→价值→展望
    "产品发布": ["cover", "toc", "background", "requirement", "solution",
                "keytech", "implement", "case", "result", "summary",
                "test", "related", "reference", "thanks"],
}
_DEFAULT_TEMPLATE_KEY = "技术汇报"


def _select_slot_template(purpose: str) -> list[str]:
    """按目的关键词选槽位模板；未知目的回退技术叙事模板。
    顺序先具体后宽泛，避免“产品发布”被“产品”、“工作总结”被“汇报”抢占。"""
    p = purpose or ""
    if any(w in p for w in ("投资", "商业", "路演", "融资", "创业", "市场", "销售")):
        return PURPOSE_SLOT_TEMPLATES["投资路演"]
    if any(w in p for w in ("科普", "普及", "大众科学")):
        return PURPOSE_SLOT_TEMPLATES["科普宣讲"]
    if any(w in p for w in ("答辩", "毕业论文", "学位论文", "开题")):
        return PURPOSE_SLOT_TEMPLATES["学术答辩"]
    if any(w in p for w in ("年终", "年度", "工作总结", "述职", "复盘", "半年总结")):
        return PURPOSE_SLOT_TEMPLATES["工作总结"]
    if any(w in p for w in ("项目", "课题", "结题", "立项")):
        return PURPOSE_SLOT_TEMPLATES["项目汇报"]
    if any(w in p for w in ("发布会", "新品", "产品发布", "产品介绍", "新品发布")):
        return PURPOSE_SLOT_TEMPLATES["产品发布"]
    if any(w in p for w in ("教学", "课程", "培训", "授课", "课堂", "讲座", "研修")):
        return PURPOSE_SLOT_TEMPLATES["教学讲解"]
    return PURPOSE_SLOT_TEMPLATES[_DEFAULT_TEMPLATE_KEY]


# 章标题 → 槽位的确定性推断（0 token）：漏章按标题原词插回模板位置而非沉尾。
# 顺序即优先级，先具体后宽泛。
_SLOT_TITLE_RULES: list[tuple[str, tuple[str, ...]]] = [
    ("cover", ("封面",)),
    ("toc", ("目录", "内容提要", "目次")),
    ("thanks", ("致谢", "谢谢")),
    ("reference", ("参考文献",)),
    ("summary", ("总结", "展望", "结论", "结语")),
    ("background", ("背景", "意义", "动机", "由来")),
    ("related", ("国内外", "现状", "相关工作", "研究综述", "理论基础")),
    ("requirement", ("需求", "目标", "问题定义")),
    ("keytech", ("关键技术", "原理", "算法", "难点")),
    ("solution", ("方案", "总体设计", "架构", "选型")),
    ("implement", ("实现", "开发", "平台", "部署", "落地")),
    ("test", ("测试", "验证", "实验", "评估")),
    ("case", ("案例", "应用场景", "典型应用")),
    ("result", ("成果", "价值", "创新", "收益", "专利")),
]
_SPECIAL_TO_SLOT = {"封面": "cover", "目录": "toc", "总结": "summary",
                    "参考文献": "reference", "致谢": "thanks"}


def _guess_slot_by_title(title: str) -> str | None:
    """按章标题原词猜槽位；无匹配 None。"""
    t = title or ""
    for slot, words in _SLOT_TITLE_RULES:
        if any(w in t for w in words):
            return slot
    return None


def _chapter_guess_ranks(
    chapter_infos: list[dict], template: Sequence[str],
    special_kinds: Sequence[list[str]] | None = None,
) -> dict[int, int]:
    """每章规则槽秩：标题原词优先，其次特殊页多数派（需过半）；猜不出不给秩。"""
    ranks: dict[int, int] = {}
    for i, ch in enumerate(chapter_infos):
        slot = _guess_slot_by_title(ch.get("title", ""))
        if slot is None and special_kinds:
            kinds = [k for k in special_kinds[i] if k in _SPECIAL_TO_SLOT]
            if kinds:
                kind = max(set(kinds), key=kinds.count)
                if kinds.count(kind) * 2 >= len(kinds):
                    slot = _SPECIAL_TO_SLOT[kind]
        if slot is not None and slot in template:
            ranks[i] = list(template).index(slot)
    return ranks


def _complete_order_by_template(
    order: Sequence[int], llm_ranks: dict[int, int],
    guess_rank: dict[int, int], k: int, template_len: int,
) -> list[int]:
    """漏章按规则槽秩插回模板位置（同槽多章按章号稳定排序）；猜不出追加末尾。"""
    big = template_len
    ranked: list[tuple[int, int | None]] = [(c, llm_ranks.get(c)) for c in order]
    missing = sorted((c for c in range(k) if c not in set(order)),
                     key=lambda c: (guess_rank.get(c, big), c))
    for m in missing:
        g = guess_rank.get(m, big)
        pos = next((i for i, (_, r) in enumerate(ranked)
                    if (r if r is not None else big) > g), len(ranked))
        ranked.insert(pos, (m, g if g < big else None))
    return [c for c, _ in ranked]


def _build_chapter_order_prompt(chapter_infos: list[dict], purpose: str) -> str:
    """第二遍：模型只做“章→槽”语义分类，槽序由目的模板强制。"""
    template = _select_slot_template(purpose)
    chain = " → ".join(SLOT_MEANING[s] for s in template)
    n = len(chapter_infos)
    lines = [
        f"下列章节来自一份用于「{purpose}」的 PPT。请把每个章节归入最贴近的"
        "叙事槽位，系统会严格按槽位模板的顺序排列（不要自己重排槽位）。",
        f"「{purpose}」的槽位模板（从先到后）：{chain}。",
        "槽位判别口径：",
    ]
    for s in template:
        if s in SLOT_HINTS:
            lines.append(f"  - {s}（{SLOT_MEANING[s]}）：{SLOT_HINTS[s]}")
    lines += [
        "易混槽位辨析（重点，勿串行）：",
        *[f"  - {d}" for d in SLOT_DISTINGUISH],
        "规则：每个章恰好归入一个槽；空槽跳过；一个槽有多章时按 slots 数组"
        "中的先后作为槽内顺序；封面/目录/参考文献/致谢归入对应专用槽；"
        "优先依据章标题原词判定槽位。"
        "chapter 字段必须填每行【编号N】中的 N（从 0 开始的列表序号），"
        "禁止填标题文字里自带的数字（如标题「01 总体情况」的 01）。",
        f"【完整性硬要求】共 {n} 个章（编号 0 到 {n - 1}），slots 数组必须"
        f"恰好包含 {n} 项，每个编号 0..{n - 1} 都要出现且仅出现一次，"
        "严禁遗漏任何章或只返回部分章；不确定归入哪个槽时也要给出一个最接近的槽。",
        "只返回 JSON：{\"reasoning\":\"一句话归类依据\","
        "\"slots\":[{\"chapter\":章节编号,\"slot\":槽位key}]}，"
        f"槽位 key 只能取：{', '.join(template)}。",
        "",
        "章节列表（【编号N】. 标题（页数） 主导阶段 代表关键词）：",
    ]
    for i, ch in enumerate(chapter_infos):
        stage = f"[{ch['stage']}]" if ch.get("stage") else ""
        lines.append(f"【编号{i}】. {ch['title']}（{len(ch['pages'])}页）{stage}"
                     f"关键词：{'、'.join(ch['keywords'])}")
    return "\n".join(lines)


def _order_chapters_once(
    chapter_infos: list[dict], k: int, purpose: str,
    llm_kw: dict, temperature: float, guess_rank: dict[int, int] | None = None,
    special_allowed: dict[str, set[int]] | None = None,
) -> list[int] | None:
    """单次“章→槽”归类 + 规则槽序组装；漏章确定性补位，有效章不足一半判失败。

    普通章（不含对应特殊页）被塞进 cover/toc/reference/thanks 专用槽时
    判无效——目的域不匹配时模型常乱填专用槽（如招生稿对学术答辩模板）。
    special_allowed: {槽位key: 含对应类型特殊页的章索引集合}。"""
    data = _llm_chat_json(
        _build_chapter_order_prompt(chapter_infos, purpose),
        temperature=temperature, **llm_kw)
    if not data:
        return None
    template = _select_slot_template(purpose)
    slot_rank = {s: i for i, s in enumerate(template)}
    special_slots = {"cover", "toc", "reference", "thanks"}
    special_allowed = special_allowed or {}
    # 标题自带编号 → 列表索引（模型常把标题里的「01」当成章节编号回传）
    title_num_to_idx: dict[str, int] = {}
    for i, ch in enumerate(chapter_infos):
        m = re.match(r"^\s*0*(\d{1,2})\b", str(ch.get("title", "")))
        if m and m.group(1) not in title_num_to_idx:
            title_num_to_idx[m.group(1)] = i

    def parse_chapter(raw) -> int | None:
        # JSON 整数 = 模型按指令回传列表序号，直接用
        if isinstance(raw, int):
            return raw
        token = str(raw).strip()
        m2 = re.match(r"^(0+)(\d{1,2})$", token)
        # 带前导零的字符串（「01」）几乎必为标题自带编号，按标题反查列表索引
        if m2 and m2.group(2) in title_num_to_idx:
            return title_num_to_idx[m2.group(2)]
        try:
            v = int(token)
            return v if 0 <= v < k else None
        except (TypeError, ValueError):
            return None

    raw_items = data.get("slots", []) or []
    placed: dict[int, tuple[int, int]] = {}
    rejected_special: list[tuple[int, str, int]] = []
    for n, item in enumerate(raw_items):
        if not isinstance(item, dict):
            continue
        ch = parse_chapter(item.get("chapter"))
        if ch is None or not (0 <= ch < k) or ch in placed:
            continue
        slot = str(item.get("slot", "")).strip()
        # 不含对应类型特殊页的普通章不允许占用专用槽（防目的域错配乱归类）
        if slot in special_slots and ch not in special_allowed.get(slot, set()):
            rejected_special.append((ch, slot, n))
            continue
        rank = slot_rank.get(slot, len(template))
        placed[ch] = (rank, n)
    # 严格门控致空（模型把全部章塞进专用槽，常见于内容域与目的模板完全不匹配）：
    # 宽松放行——专用槽有固定秩（致谢/参考在末尾），槽内按回传顺序，等同保持原序，
    # 避免白白浪费一轮投票
    if not placed and rejected_special:
        for ch, slot, n in rejected_special:
            if ch in placed:
                continue
            placed[ch] = (slot_rank.get(slot, len(template)), n)
    order = sorted(placed, key=lambda c: placed[c])
    if not order:
        # 连一个有效槽位归类都没有（返回损坏/全被门控且无法宽松放行）才丢弃本轮；
        # 只要有 ≥1 个有效章就采纳，漏掉的章由 _complete_order_by_template
        # 按规则槽秩确定性补回（猜不出则原序追加），避免模型欠填白白浪费投票。
        return None
    return _complete_order_by_template(
        order, {c: r for c, (r, _) in placed.items()},
        guess_rank or {}, k, len(template))


def _borda_merge_orders(orders: list[list[int]], k: int) -> list[int]:
    """Borda 计数聚合多轮排列：出现给位置分，缺席惩罚 k 分；分低靠前。"""
    scores = [0.0] * k
    for o in orders:
        seen = set(o)
        for pos, ch in enumerate(o):
            scores[ch] += pos
        for ch in range(k):
            if ch not in seen:
                scores[ch] += k
    return sorted(range(k), key=lambda c: (scores[c], c))


def order_chapters_by_llm(
    induced: list[dict], slides: Sequence[SlideInfo], purpose: str, *,
    model: str = "", base_url: str = "", api_key: str = "",
    votes: int = 1,
) -> list[int] | None:
    """按目的排章，返回 0-based 章索引排列；votes>1 做 Borda 自一致投票。"""
    k = len(induced)
    if k < 2:
        return None
    info = {s.index: s for s in slides}
    chapter_infos: list[dict] = []
    for ch in induced:
        kw_bag: dict[str, int] = {}
        stage_bag: dict[int, int] = {}
        for p in ch["pages"]:
            s = info.get(p)
            if s is None:
                continue
            if _HAS_JIEBA:
                for w in jieba.analyse.textrank(
                        s.title + " " + s.text_snippet, topK=6, withWeight=False):
                    kw_bag[w] = kw_bag.get(w, 0) + 1
            if s.stage is not None:
                stage_bag[s.stage] = stage_bag.get(s.stage, 0) + 1
        top_kw = [w for w, _ in sorted(kw_bag.items(), key=lambda x: -x[1])[:8]]
        stage = STAGE_NAMES[max(stage_bag, key=stage_bag.get)] if stage_bag else ""
        chapter_infos.append({"title": ch["title"], "pages": ch["pages"],
                              "keywords": top_kw, "stage": stage})
    # 漏章补位规则槽秩（标题原词/特殊页多数派），各轮共用
    special_kinds = [
        [info[p].special for p in ch["pages"] if p in info and info[p].special]
        for ch in induced]
    guess_rank = _chapter_guess_ranks(
        chapter_infos, _select_slot_template(purpose), special_kinds)
    # 专用槽 → 含对应类型特殊页的章索引（普通章占专用槽判无效）
    slot_special = {"cover": "封面", "toc": "目录",
                    "reference": "参考文献", "thanks": "致谢"}
    special_allowed: dict[str, set[int]] = {
        slot: {i for i, kinds in enumerate(special_kinds) if kind in kinds}
        for slot, kind in slot_special.items()}
    llm_kw = {"model": model, "base_url": base_url, "api_key": api_key}
    orders: list[list[int]] = []
    n_votes = max(1, votes)
    for r in range(n_votes):
        o = _order_chapters_once(chapter_infos, k, purpose, llm_kw,
                                 temperature=0.1 if r == 0 else 0.15,
                                 guess_rank=guess_rank,
                                 special_allowed=special_allowed)
        if o is not None:
            orders.append(o)
    if not orders:
        logger.warning("章节排序 %d 轮均失败，脱出", n_votes)
        return None
    if len(orders) == 1:
        order = orders[0]
    else:
        order = _borda_merge_orders(orders, k)
        logger.info("章节排序自一致投票：%d/%d 轮有效，各轮 %s",
                    len(orders), n_votes, orders)
    logger.info("llm chapter order for %r: %s", purpose, order)
    return order


# --------------------------------------------------------------------------- #
# 0-token 结构/标题检定
# --------------------------------------------------------------------------- #
RULE_MIN_GROUPS = 2
RULE_MAX_GROUPS = 12
RULE_DOMINANT_RATIO = 0.60   # 单组占正文比例超此值 → 判标题无区分度


def _title_groups_raw(slides: Sequence[SlideInfo]) -> list[dict]:
    """按归一化页标题对正文页全局归组（不受乱序影响），特殊页不参与。"""
    kind_of = {s.index: s.special for s in slides}
    buckets: dict[str, list[int]] = {}
    display: dict[str, str] = {}
    for s in slides:
        if kind_of.get(s.index) in SPECIAL_KINDS:
            continue
        key = _norm_title(s.title) or f"__p{s.index}"  # 空标题各自成组（显碎）
        buckets.setdefault(key, []).append(s.index)
        display.setdefault(key, (s.title or "").strip()[:20])
    return [{"title": display[k] or "未命名章节", "pages": sorted(ps)}
            for k, ps in buckets.items()]


def induce_chapters_by_structure(slides: Sequence["SlideInfo"]) -> list[dict] | None:
    """0-token 结构切章（最高优先，章节成块不拆散）：沿原始页序以分标题页
    （第N节/章、稀疏数字分隔页）为界，每个分标题拥有到下一个分标题前的
    全部正文页；首个分标题前的正文归「开篇引言」。

    采用条件：≥2 个分标题、结构块 2..12 个、每个分标题至少带 1 页内容
    （防空分标题成孤块）。不满足返回 None 升级标题归组/LLM。
    """
    kind_of = {s.index: s.special for s in slides}
    heads = [s for s in slides if _is_section_heading(s)]
    if len(heads) < 2:
        return None
    bounds = [s.index for s in heads]
    groups: list[dict] = []
    pre = [s.index for s in slides
           if s.index < bounds[0] and kind_of.get(s.index) is None]
    if pre:
        groups.append({"title": "开篇引言", "pages": pre})
    for k, h in enumerate(heads):
        end = bounds[k + 1] - 1 if k + 1 < len(heads) else len(slides)
        pages = [p for p in range(h.index, end + 1)
                 if kind_of.get(p) is None]
        # 分标题 + 至少 1 页内容，避免空分标题成孤块
        if len(pages) < 2:
            logger.info("结构切章：分标题 P%d《%s》无内容页，放弃结构检定",
                        h.index, (h.title or "").replace("\x0b", " ")[:16])
            return None
        name = _heading_display_name(h) or f"第{k + 1}节"
        groups.append({"title": name, "pages": pages})
    if not (INDUCED_MIN_CHAPTERS <= len(groups) <= INDUCED_MAX_CHAPTERS):
        logger.info("结构切章 %d 块（超出 %d..%d），转后续检定",
                    len(groups), INDUCED_MIN_CHAPTERS, INDUCED_MAX_CHAPTERS)
        return None
    covered = sorted(p for g in groups for p in g["pages"])
    body = sorted(s.index for s in slides if kind_of.get(s.index) is None)
    if covered != body:  # 构造上应天然覆盖，防御性校验
        logger.info("结构切章覆盖不全，转后续检定")
        return None
    logger.info("章节检定命中结构分标题（0 token，章节成块）：%s",
                [f"{g['title']}({len(g['pages'])}页)" for g in groups])
    return groups


def induce_chapters_by_rules(slides: Sequence[SlideInfo]) -> list[dict] | None:
    """0-token 标题归组：组数 2..12 且无绝对主导组时采用，否则 None 升级 LLM。"""
    groups = _title_groups_raw(slides)
    body_n = sum(len(g["pages"]) for g in groups)
    if body_n == 0:
        return None
    if not (RULE_MIN_GROUPS <= len(groups) <= RULE_MAX_GROUPS):
        logger.info("标题归组 %d 组（超出 %d..%d），转 LLM 聚类",
                    len(groups), RULE_MIN_GROUPS, RULE_MAX_GROUPS)
        return None
    biggest = max(len(g["pages"]) for g in groups)
    if biggest / body_n > RULE_DOMINANT_RATIO:
        logger.info("标题归组主导组占 %.0f%%，判标题无区分度，转 LLM 聚类",
                    100.0 * biggest / body_n)
        return None
    logger.info("章节检定命中标题归组（0 token）：%s",
                [g["title"] for g in groups])
    return groups


# 各汇报目的的章内叙事准则（注入章内精排 prompt）。
# 每个目的给出该场景听众期待的叙事线，模型据此重排章内页面（②版，目的驱动）。
_DEFAULT_WITHIN_GUIDE = (
    "按「背景动机→相关基础→方案设计→实现方法→过程结果→结论建议」的技术"
    "叙事线组织：背景与定义在前，实现与过程居中，结论与总结收在章末。")
WITHIN_CHAPTER_GUIDES: dict[str, str] = {
    "科普宣讲": (
        "听众没有专业背景，需要先被现象吸引再理解原理：把身边现象、真实案例、"
        "直观效果页提到章首（概念/定义之前），再展开原理与技术；公式推导、"
        "实验参数、参考文献放在本章末尾。"),
    "学术答辩": (
        "按论文论证链（问题→方法→实验→结果→结论）组织：让核心方法、创新点、"
        "关键结果尽早出现，不被铺垫性内容压住；与其他技术横向对比的承上过渡页"
        "放在本章主体内容之后。"),
    "投资人路演": (
        "投资人先关心价值与可信度：把数据、客户案例、效果截图等价值实证页提到"
        "原理推导与背景之前；纯理论推导、参考文献放在本章后部。"),
    "课堂教学": (
        "按知识点递进组织：概念→讲解→例题→练习，例题与课堂练习紧跟其对应"
        "知识点，小结收在章末。"),
    "培训讲座": (
        "学员希望尽快上手：把操作步骤、流程页提到理论背景之前；分步页之间"
        "严格遵守编号顺序。"),
    "产品发布": (
        "先抓住观众注意力：把最有冲击力的功能演示、核心卖点前置，技术实现"
        "细节靠后；价格与上市信息放在结尾。"),
    "产品评审": (
        "评审需要快速聚焦待决策问题：把风险与待决策问题前置，再按需求→方案"
        "→进度展开。"),
    "技术汇报": _DEFAULT_WITHIN_GUIDE,
    "项目汇报": (
        "按项目闭环（目标→方案→过程→成果→计划）组织：背景与目标在章首，"
        "成果与验收数据放到实施过程之后的显眼位置。"),
    "工作总结": (
        "述职先亮成绩：把核心成果与关键指标提到本章前部，再叙述工作过程，"
        "反思与计划靠后。"),
    "招生综合宣讲": (
        "先用亮点吸引考生：把荣誉、数据、校园图片等亮点页前置，历史沿革"
        "靠后；报考信息与联系方式放在章末。"),
    "综合汇报": _DEFAULT_WITHIN_GUIDE,
}


def _within_chapter_guide(purpose: str) -> str:
    """取目的专属章内准则；未配置按关键词回退，再回退通用技术叙事。"""
    p = purpose or ""
    if p in WITHIN_CHAPTER_GUIDES:
        return WITHIN_CHAPTER_GUIDES[p]
    if any(w in p for w in ("科普", "普及")):
        return WITHIN_CHAPTER_GUIDES["科普宣讲"]
    if any(w in p for w in ("答辩", "论文", "开题")):
        return WITHIN_CHAPTER_GUIDES["学术答辩"]
    if any(w in p for w in ("融资", "投资", "路演", "商业")):
        return WITHIN_CHAPTER_GUIDES["投资人路演"]
    if any(w in p for w in ("培训", "讲座")):
        return WITHIN_CHAPTER_GUIDES["培训讲座"]
    return _DEFAULT_WITHIN_GUIDE


def _build_within_chapter_prompt(
    briefs: list[dict], title: str, pages: Sequence[int], purpose: str,
) -> str:
    """章内精排 prompt（一章 ≤24 页，小粒度模型可靠）。"""
    pset = sorted(pages)
    guide = _within_chapter_guide(purpose)
    lines = [
        f"下列 {len(pset)} 页同属 PPT 章节《{title}》。",
        f"本章将用于「{purpose or '通用汇报'}」场景。请只依据页面真实内容，"
        "按该场景的叙事逻辑重新排列这些页面。",
        f"排序准则：{guide}",
        "页内编号（1./2./3.）构成的递进步骤不得打散。",
        '只返回 JSON：{"reasoning":"排序理由","order":[页号...]}，'
        f"order 必须恰好包含 {pset} 各一次，不得遗漏。",
        "",
        "页面列表：",
    ]
    wanted = set(pset)
    lines.extend(_brief_line(b) for b in briefs if b["index"] in wanted)
    return "\n".join(lines)


def _fill_missing_by_original_order(
    order: list[int], pages: Sequence[int],
) -> list[int]:
    """章内漏页按原稿邻接位置插回（而非统一扔章末）。

    对每个漏页：在原稿序列中找最近的、已在结果里的前邻 → 插到其后；
    没有前邻（漏页在章首）则找最近后邻插到其前；都没有则追加末尾。
    漏页按原稿顺序依次处理，连续多页漏排时自然成链归位。"""
    result = list(order)
    present = set(result)
    for m in pages:  # pages 即章内原稿序
        if m in present:
            continue
        orig_idx = pages.index(m)
        anchor: int | None = None
        after = True
        for q in reversed(pages[:orig_idx]):
            if q in present:
                anchor, after = q, True
                break
        if anchor is None:
            for q in pages[orig_idx + 1:]:
                if q in present:
                    anchor, after = q, False
                    break
        if anchor is None:
            result.append(m)
        else:
            at = result.index(anchor) + (1 if after else 0)
            result.insert(at, m)
        present.add(m)
    return result


def order_pages_within_chapter_by_llm(
    title: str, pages: Sequence[int], slides: Sequence[SlideInfo], purpose: str, *,
    chapters: Sequence[Chapter] | None = None,
    model: str = "", base_url: str = "", api_key: str = "",
) -> list[int] | None:
    """章内自然讲述序；超页/不可用返回 None。漏页按原稿邻接位置补回，
    有效页不足一半判整章失败（保原序）。"""
    pages = list(pages)
    if len(pages) < 2 or len(pages) > WITHIN_CHAPTER_LLM_MAX:
        return None
    briefs = build_slide_briefs(slides, chapters=chapters)
    data = _llm_chat_json(
        _build_within_chapter_prompt(briefs, title, pages, purpose),
        model=model, base_url=base_url, api_key=api_key)
    if not isinstance(data, dict):
        return None
    wanted = set(pages)
    order: list[int] = []
    for x in data.get("order", []):
        try:
            v = int(x)
        except (TypeError, ValueError):
            continue
        if v in wanted and v not in order:
            order.append(v)
    if len(order) < max(2, len(pages) // 2):
        logger.warning("章内《%s》排序有效页不足（%d/%d），脱出",
                       title, len(order), len(pages))
        return None
    missing = [p for p in pages if p not in set(order)]
    if missing:
        logger.info("章内《%s》补漏 %d 页（按原稿邻接插回）：%s",
                    title, len(missing), missing)
        order = _fill_missing_by_original_order(order, pages)
    return order


def assemble_by_chapters(
    order: Sequence[int], slides: Sequence[SlideInfo], induced: list[dict],
) -> list[int]:
    """按已排章节组装全序：封面/目录置首、总结/参考/致谢置尾（即使被聚进章也
    抽出）；章内保持入参相对序；未覆盖页追加正文末尾。"""
    n = len(slides)
    norm = normalize_order(order, n)
    kind_of = {s.index: s.special for s in slides}
    head = [i for i in norm if kind_of.get(i) in ("封面", "目录")]
    tail = [i for i in norm if kind_of.get(i) in ("总结", "参考文献", "致谢")]
    fixed = set(head) | set(tail)
    pos_in = {idx: p for p, idx in enumerate(norm)}
    result = list(head)
    covered: set[int] = set()
    for ch in induced:
        pages = sorted((p for p in ch["pages"] if p not in fixed),
                       key=lambda x: pos_in.get(x, 0))
        result.extend(pages)
        covered.update(pages)
    result.extend(i for i in norm if i not in fixed and i not in covered)
    result.extend(tail)
    return normalize_order(result, n)


def normalize_order(order: Iterable[int], n: int) -> list[int]:
    """校验显式顺序：必须是 1..n 的排列。"""
    order_list = [int(x) for x in order]
    if sorted(order_list) != list(range(1, n + 1)):
        raise ValueError(f"顺序必须是 1..{n} 的排列，收到: {order_list}")
    return order_list


def anchor_special_pages(
    order: Sequence[int], slides: Sequence[SlideInfo],
) -> list[int]:
    """特殊页确定性锚定（模型只排正文语义）：
    封面…目录…【正文保持入参相对序】…总结…参考文献…致谢。"""
    n = len(slides)
    norm = normalize_order(order, n)
    kind_of = {s.index: s.special for s in slides}

    def take(kind: str) -> list[int]:
        return [i for i in norm if kind_of.get(i) == kind]

    head = take("封面") + take("目录")
    tail = take("总结") + take("参考文献") + take("致谢")
    fixed = set(head) | set(tail)
    body = [i for i in norm if i not in fixed]
    return normalize_order(head + body + tail, n)


def lead_section_headings(
    order: Sequence[int], slides: Sequence["SlideInfo"],
    groups_map: dict | None = None,
) -> list[int]:
    """结构切章专用：分标题必须位于其归属内容最前。
    只做块内前移、不跨块；章内顺序已由 LLM/证据精排调好，标题领头即可。
    groups_map（页→章 id，两级编号树剥离后的归属）优先；缺省回退原稿区间。"""
    info = {s.index: s for s in slides}
    seq = list(order)
    heads = sorted(p for p in info if _is_section_heading(info[p]))
    if len(heads) < 2:
        return seq
    last_idx = max(info) + 1
    owned = {}
    if groups_map:
        by_gid: dict = {}
        for p, gid in groups_map.items():
            by_gid.setdefault(gid, []).append(p)
        for h in heads:
            owned[h] = [p for p in by_gid.get(groups_map.get(h), [])
                        if p != h and not _is_section_heading(info[p])]
    else:
        for k, h in enumerate(heads):
            end = heads[k + 1] if k + 1 < len(heads) else last_idx
            owned[h] = [p for p in range(h + 1, end)
                        if p in info and not _is_section_heading(info[p])]
    for h in heads:  # 按原序依次前移，避免相互错位
        related = [p for p in owned[h] if p in seq]
        if not related:
            continue
        anchor = min(seq.index(p) for p in related)
        if seq.index(h) > anchor:
            seq.remove(h)
            seq.insert(anchor, h)
    return seq


def bind_section_headings(
    order: Sequence[int], slides: Sequence["SlideInfo"],
    groups_map: dict | None = None,
) -> tuple[list[int], list[str], list[str]]:
    """结构硬规则：分标题页必须绑定自己的内容，绝不允许两个分标题紧挨着。

    归属（不依赖语义、对乱序健壮）：优先 groups_map（两级编号树剥离后的
    页→章归属）；缺省按原始页序（分标题 h 拥有到下一个分标题前的全部正文页）。
    对最终顺序迭代修复每个「h 后紧跟分标题」或「h 悬空末尾」的违规：
    - 归属页仍在顺序中 → 把 h 移到最早归属页之前（强制绑定相关内容）；
    - 归属页已不在（被清理）且 h 是纯空分隔页（纯编号、无正文无图表）→ 删除；
    - h 自带正文（如「第X节+本节提纲」）且无内容可绑 → 视作内容页保留。
    幂等：无违规时原样返回。返回 (新顺序, 绑定说明, 删除说明)。
    """
    info = {s.index: s for s in slides}
    seq = [p for p in order if p in info]
    heading_idxs = sorted(p for p in info if _is_section_heading(info[p]))
    if not heading_idxs or not seq:
        return seq, [], []
    # 归属：groups_map 优先（板块剥离后原稿区间会混入板块页），缺省原稿区间
    owned: dict[int, list[int]] = {}
    if groups_map:
        by_gid: dict = {}
        for p, gid in groups_map.items():
            by_gid.setdefault(gid, []).append(p)
        for h in heading_idxs:
            owned[h] = [p for p in by_gid.get(groups_map.get(h), [])
                        if p != h and not _is_section_heading(info[p])]
    else:
        last_idx = max(info) + 1  # 页号未必连续，终点取最大页号
        for k, h in enumerate(heading_idxs):
            end = heading_idxs[k + 1] if k + 1 < len(heading_idxs) else last_idx
            owned[h] = [p for p in range(h + 1, end)
                        if p in info and not _is_section_heading(info[p])]
    present = set(seq)
    # 自带正文、且无归属内容可绑的分标题：按内容页对待，不参与违规判定、不删
    standalone = {
        h for h in heading_idxs
        if not any(p in present for p in owned[h]) and not _is_pure_divider(info[h])
    }
    effective = set(heading_idxs) - standalone
    moves: list[str] = []
    removed: list[str] = []

    def _name(p: int) -> str:
        s = info[p]
        return (s.title or s.subtitle or "?").replace("\x0b", " ").replace("\n", " ")[:20]

    for _ in range(len(seq) * 2 + 4):
        bad = -1
        for i in range(len(seq) - 1):
            if seq[i] in effective and seq[i + 1] in effective:
                bad = i
                break
        if bad < 0 and seq[-1] in effective:
            bad = len(seq) - 1
        if bad < 0:
            break
        h = seq[bad]
        related = [p for p in owned.get(h, []) if p in seq]
        if related:
            anchor = min(related, key=lambda p: seq.index(p))  # 当前最早归属页
            pos = seq.index(anchor)
            seq.pop(bad)
            if pos > bad:
                pos -= 1
            seq.insert(pos, h)
            moves.append(f"页{h}《{_name(h)}》绑定到其内容页 P{anchor} 之前")
        else:
            # 有效分标题且无内容可绑：必为纯空分隔页，删除
            seq.pop(bad)
            effective.discard(h)
            removed.append(f"页{h}《{_name(h)}》（无归属内容，删除空分标题）")
    return seq, moves, removed


# --------------------------------------------------------------------------- #
# 写操作：重排与目录页
# --------------------------------------------------------------------------- #
def _rebuild_sld_id_lst(prs: Presentation, order: Sequence[int]) -> None:
    """按 1-based 序重建 sldIdLst（python-pptx 无公开 move API）。

    order 为子集（删页）时：
    1. pop 被删页的关系 → 部件不可达，保存时不再写入（无孤儿部件）；
    2. 保留页部件名重编号为 slide1..slideK（新顺序），避免部件名空洞，
       保证后续 add_slide（插入目录页）按 len+1 取名时不与现有部件冲突。
    """
    sld_id_lst = prs.slides._sldIdLst  # type: ignore[attr-defined]
    elements = list(sld_id_lst)
    dropped = {int(x) for x in range(1, len(elements) + 1)} - {int(x) for x in order}
    for elem in elements:
        sld_id_lst.remove(elem)
    for orig_idx in order:
        sld_id_lst.append(elements[orig_idx - 1])
    if dropped:
        for idx in dropped:
            rId = elements[idx - 1].get(qn("r:id"))
            if rId and rId in prs.part.rels:
                prs.part.rels.pop(rId)
        for new_num, orig_idx in enumerate(order, start=1):
            rId = elements[orig_idx - 1].get(qn("r:id"))
            part = prs.part.rels[rId].target_part
            part.partname = PackURI(f"/ppt/slides/slide{new_num}.xml")


def reorder_slides(
    src_path: str, order: Sequence[int], dst_path: str, *, force: bool = False
) -> str:
    """按显式顺序生成新 PPT，返回输出路径。

    order 可以是 1..n 的完整排列，也可以是子集（删除页时用）。
    """
    src = validate_pptx_path(src_path)
    dst = check_output_path(dst_path, src, force=force)
    prs = Presentation(str(src))
    n = len(prs.slides)
    order_list = [int(x) for x in order]
    valid = all(1 <= x <= n for x in order_list)
    if len(order_list) == n and sorted(order_list) == list(range(1, n + 1)):
        norm_order = order_list
    elif valid and len(set(order_list)) == len(order_list):
        norm_order = order_list  # 子集：只输出指定页
        logger.info("reorder: subset %d/%d pages", len(norm_order), n)
    else:
        raise ValueError(f"顺序包含无效或重复页码: {order_list}")
    logger.info("reorder: %s -> %s, order=%s", src, dst, norm_order)
    _rebuild_sld_id_lst(prs, norm_order)
    prs.save(str(dst))
    logger.info("reorder saved: %s", dst)
    return str(dst)


def reorganize_by_purpose(
    src_path: str, purpose: str, dst_path: str, *, force: bool = False,
    sort_level: str = "chapter", use_llm: bool = False,
    model: str = "", base_url: str = "", api_key: str = "",
    anchor_special: bool = False, smart: bool = False,
) -> tuple[list[int], str, dict]:
    """按汇报目的重组，返回 (新顺序, 输出路径, trace)。

    trace 包含各环节的决策信息，可用于生成变动说明。
    章内冗余/连续的纯分隔页始终自动清理（结构兜底），内容页与章首分隔页保留；
    结构切章路径下该步骤通常为空操作。
    """
    src = validate_pptx_path(src_path)
    dst = check_output_path(dst_path, src, force=force)
    analysis = analyze_pptx(str(src))
    order, trace = resolve_order(
        analysis, purpose, sort_level=sort_level, use_llm=use_llm,
        model=model, base_url=base_url, api_key=api_key,
        anchor_special=anchor_special, smart=smart)
    # 结构兜底：自动清理章内冗余/连续纯分隔页（无需用户开关）
    order, removed = trim_internal_dividers(
        order, trace.get("induced_chapters"), analysis.slides)
    if removed:
        logger.info("自动清理 %d 个冗余分隔页: %s", len(removed), removed)
    # 清理可能造成新的分标题相邻/悬空，再做一次结构硬规则绑定（幂等）
    order, hm, hd = bind_section_headings(order, analysis.slides)
    if hm:
        logger.info("分标题绑定 %d 处: %s", len(hm), hm)
    if hd:
        logger.info("删除空分标题 %d 页: %s", len(hd), hd)
    logger.info("reorganize purpose=%r sort_level=%r use_llm=%r smart=%r trace=%r order=%s",
                purpose, sort_level, use_llm, smart, trace, order)
    prs = Presentation(str(src))
    _rebuild_sld_id_lst(prs, order)
    prs.save(str(dst))
    logger.info("reorganize saved: %s", dst)
    return order, str(dst), trace


def _clear_slide_shapes(slide) -> None:
    """删除页面上的所有形状（含从版式继承的占位符），得到真正的空白页。"""
    sp_tree = slide.shapes._spTree  # type: ignore[attr-defined]
    # 收集所有形状元素后再移除（避免边遍历边修改）
    to_remove = [
        child for child in list(sp_tree)
        if child.tag in (
            "{http://schemas.openxmlformats.org/presentationml/2006/main}sp",
            "{http://schemas.openxmlformats.org/presentationml/2006/main}pic",
            "{http://schemas.openxmlformats.org/presentationml/2006/main}grpSp",
            "{http://schemas.openxmlformats.org/presentationml/2006/main}cxnSp",
            "{http://schemas.openxmlformats.org/presentationml/2006/main}graphicFrame",
        )
    ]
    for elem in to_remove:
        sp_tree.remove(elem)


def _toc_sections(slides: Sequence[SlideInfo]) -> list[tuple[str, int, int]]:
    """构建目录条目：只列分标题（章/板块），不逐页罗列。

    返回 (标题, 页码, 层级) 三元组：层级 0=顶层板块、1=板块内章。
    仅当常规（编号标题）与反向（名称标题+编号正文）两种形态并存——即两级
    编号稿（如板块 01-03 + 编号章 01-10）——才区分层级；单一形态稿全部
    记 0（不分级，渲染保持原样式）。

    策略：
    - 有分标题时，以分标题为界切段，条目名取分标题自身
      （「02 升学与就业支持」），页码指向分标题页；首个分标题
      前的正文作为「开篇」一条；带角标的编号内容页不另立条目；
    - 无分标题时，按连续相同标题聚合（同标题跨页只列一次）；
    - 封面/目录/参考文献/致谢不列入目录，并触发分段。
    """
    skip_specials = {"封面", "目录", "参考文献", "致谢"}
    # 分标题识别：常规形态（编号标题）+ 反向形态（名称标题+编号正文首行，
    # 如「招生政策问答 03」剥离板块页），保证任何格式的分标题都进目录
    def _is_toc_heading(s: "SlideInfo") -> bool:
        return _is_section_heading(s) or _is_reverse_numbered_heading(s)

    reverse_pages = {s.index for s in slides if _is_reverse_numbered_heading(s)}
    has_regular = any(_is_section_heading(s) for s in slides)
    two_level = bool(reverse_pages) and has_regular
    has_heading = bool(reverse_pages) or has_regular

    sections: list[tuple[str, int, int]] = []
    cur_title: str | None = None
    cur_start: int | None = None
    cur_level = 0

    def flush():
        nonlocal cur_title, cur_start
        if cur_title is not None:
            sections.append((cur_title, cur_start, cur_level))
        cur_title = None
        cur_start = None

    if has_heading:
        for s in slides:
            if s.special in skip_specials:
                flush()
                continue
            if _is_toc_heading(s):
                flush()
                cur_title = _heading_display_name(s)
                cur_start = s.index
                cur_level = (0 if s.index in reverse_pages
                             else 1 if two_level else 0)
                continue
            if cur_title is None:  # 首个分标题前的开篇正文
                cur_title = s.title
                cur_start = s.index
                cur_level = 0
    else:
        for s in slides:
            if s.special in skip_specials:
                flush()
                continue
            if cur_title is None:
                cur_title = s.title
                cur_start = s.index
                cur_level = 0
            elif s.title != cur_title:
                flush()
                cur_title = s.title
                cur_start = s.index
                cur_level = 0
    flush()
    return sections


def insert_toc(
    src_path: str, dst_path: str, *, title: str = "目录", force: bool = False
) -> str:
    """插入一页目录（基于段落汇总），返回输出路径。

    目录页插在封面之后（无封面时插在最前）；页码按插入后的新位置计。
    目录项按段落汇总（分隔页切段 / 同标题聚合），不逐页罗列。
    两级编号稿做层级差分渲染：顶层板块（L0）加粗大字，板块内章（L1）
    缩进小字；单级稿保持原样式（18pt 平铺）。
    """
    src = validate_pptx_path(src_path)
    dst = check_output_path(dst_path, src, force=force)
    analysis = analyze_pptx(str(src))
    sections = _toc_sections(analysis.slides)
    has_levels = any(lv for _, _, lv in sections)
    prs = Presentation(str(src))
    # 选择占位符最少的版式（优先真正空白），然后清空该页所有继承形状
    layouts = list(prs.slide_layouts)
    best_layout = min(layouts, key=lambda ly: len(list(ly.placeholders))) \
        if layouts else prs.slide_layouts[6]
    toc_slide = prs.slides.add_slide(best_layout)
    _clear_slide_shapes(toc_slide)
    title_box = toc_slide.shapes.add_textbox(
        Inches(0.5), Inches(0.3), Inches(9), Inches(0.8))
    tf = title_box.text_frame
    tf.text = title
    tf.paragraphs[0].runs[0].font.size = Pt(32)
    body_box = toc_slide.shapes.add_textbox(
        Inches(0.8), Inches(1.4), Inches(8.4), Inches(5.5))
    btf = body_box.text_frame
    btf.word_wrap = True
    # 插入 1 页目录后，所有列出的段落起始页码 +1
    for i, (sec_title, start, level) in enumerate(sections):
        line = btf.paragraphs[0] if i == 0 else btf.add_paragraph()
        text = f"{sec_title}  ……  P{start + 1}"
        if has_levels and level == 1:
            text = "　　" + text  # 板块内章缩进
        line.text = text
        try:
            run = line.runs[0]
            if has_levels and level == 0:
                run.font.size = Pt(20)
                run.font.bold = True
            elif has_levels:
                run.font.size = Pt(15)
            else:
                run.font.size = Pt(18)
        except Exception:
            pass
    # 新页默认在末尾，移到封面之后（无封面则最前）
    sld_id_lst = prs.slides._sldIdLst  # type: ignore[attr-defined]
    toc_elem = list(sld_id_lst)[-1]
    sld_id_lst.remove(toc_elem)
    cover_first = bool(analysis.slides) and analysis.slides[0].special == "封面"
    sld_id_lst.insert(1 if cover_first else 0, toc_elem)
    prs.save(str(dst))
    logger.info("toc inserted: %s (%d sections, after_cover=%s)",
                dst, len(sections), cover_first)
    return str(dst)


# --------------------------------------------------------------------------- #
# 自然语言变动解释（写盘后向用户说明做了什么、为什么）
# --------------------------------------------------------------------------- #
_INDUCE_MODE_DESC = {
    "structural_sections": "分标题结构切章（0 token，按「第N节/编号」自动分块）",
    "rule_titles": "标题归组（0 token，按标题语义聚合）",
    "single": "全页 LLM 排序",
    "none": "未触发章节检定",
}

_CHAPTER_SOURCE_DESC = {
    "structural_sections": "结构分标题",
    "rule_titles": "规则标题归组",
    "llm_induced": "LLM 聚类",
    "anchor_fallback": "规则锚定兜底（检定全部不可用时）",
    "none": "未分章",
}


def explain_reorganization(
    trace: dict, order: Sequence[int],
    slides: Sequence["SlideInfo"],
) -> str:
    """根据 trace 生成自然语言变动说明，写盘后展示给用户。

    说明重组各环节做了什么、为什么这样做，帮助用户理解结果。
    """
    lines: list[str] = []
    n = len(slides)

    # 1. 章节检定
    induce = trace.get("induce_mode", "none")
    chap_src = trace.get("chapter_source", "none")
    chap_count = trace.get("induced_chapter_count", 0)

    if induce != "none" and chap_count > 0:
        desc = _INDUCE_MODE_DESC.get(induce, induce)
        lines.append(f"章节检定：{desc}，共 {chap_count} 章")
        # 列出各章标题
        chapters = trace.get("induced_chapters", [])
        if chapters:
            chap_names = [g["title"][:12] for g in chapters]
            lines.append(f"  各章：{' / '.join(chap_names)}")
    elif chap_src == "anchor_fallback":
        lines.append("章节检定：无可用的结构信号，采用规则锚定兜底")
    else:
        lines.append("章节检定：未分章（小稿或无结构信号）")

    # 2. LLM 投票
    base_mode = trace.get("base", "keyword")
    if "skipped_full_llm" in base_mode:
        lines.append(f"粗排：页数 {n} > {INDUCE_SINGLE_MAX}，跳过全页 LLM，用关键词序兜底")
    elif base_mode == "llm":
        lines.append("粗排：全页 LLM 排序")

    within_calls = trace.get("within_chapter_calls", 0)
    within_ok = trace.get("within_chapter_ok", 0)
    if within_calls > 0:
        lines.append(f"章内精排：{within_ok}/{within_calls} 章由 LLM 调整页内顺序")

    # 3. 两级编号剥离
    top_sections = trace.get("top_sections", [])
    if top_sections:
        names = [s.get("title", "")[:10] for s in top_sections]
        lines.append(f"两级编号剥离：{len(top_sections)} 个顶层板块（{' / '.join(names)}）")

    # 4. 编号归位与分标题绑定
    refined = trace.get("groups_refined", 0)
    if refined:
        lines.append(f"编号归位：修正 {len(refined)} 处序号错位")

    bound = trace.get("heading_bound", [])
    dropped = trace.get("heading_dropped", [])
    if bound:
        lines.append(f"分标题绑定：修正 {len(bound)} 处")
    if dropped:
        lines.append(f"空分标题清理：删除 {len(dropped)} 页")

    # 5. 锚定
    anchor = trace.get("anchor_applied", False)
    if anchor:
        lines.append("特殊页锚定：封面/目录→首位，总结/参考文献/致谢→末尾")

    # 6. 主要变动
    orig = list(range(1, n + 1))
    if list(order) == orig:
        lines.append("变动结论：最终顺序与原稿一致（LLM 投票后认为原序最优）")
    else:
        moved = sum(1 for i, o in enumerate(order) if o != orig[i]) if len(order) == n else 0
        if moved > 0:
            lines.append(f"变动结论：{moved} 页位置发生变化")
        # 找出最大的几处变动
        changes = []
        for i, o in enumerate(order):
            if o != orig[i] and len(changes) < 3:
                changes.append(f"  原第{o}页→新第{i+1}位")
        if changes:
            lines.append("主要变动：")
            lines.extend(changes)

    return "\n".join(lines)


# --------------------------------------------------------------------------- #
# 预览（不写文件）
# --------------------------------------------------------------------------- #
def preview_reorder(src_path: str, order: Sequence[int]) -> dict:
    src = validate_pptx_path(src_path)
    analysis = analyze_pptx(str(src))
    norm = normalize_order(order, len(analysis.slides))
    title_by_idx = {s.index: s.title for s in analysis.slides}
    return {
        "file": str(src),
        "order": norm,
        "preview": [
            {"new_position": pos + 1, "original_index": orig, "title": title_by_idx[orig]}
            for pos, orig in enumerate(norm)
        ],
    }


def preview_reorganize(
    src_path: str, purpose: str, sort_level: str = "chapter",
    use_llm: bool = False, model: str = "", base_url: str = "", api_key: str = "",
    anchor_special: bool = False, smart: bool = False,
) -> dict:
    src = validate_pptx_path(src_path)
    analysis = analyze_pptx(str(src))
    order, trace = resolve_order(
        analysis, purpose, sort_level=sort_level, use_llm=use_llm,
        model=model, base_url=base_url, api_key=api_key,
        anchor_special=anchor_special, smart=smart)
    # 与 reorganize_by_purpose 写盘路径保持一致：自动清理冗余分隔页，
    # 再做一次分标题绑定（幂等），保证预览顺序与最终文件完全相同。
    order, _removed = trim_internal_dividers(
        order, trace.get("induced_chapters"), analysis.slides)
    order, _moves, _drops = bind_section_headings(order, analysis.slides)
    title_by_idx = {s.index: s.title for s in analysis.slides}
    return {
        "file": str(src),
        "purpose": purpose,
        "sort_level": sort_level,
        "anchor_special": anchor_special or bool(smart),
        "trace": trace,
        "order": order,
        "preview": [
            {"new_position": pos + 1, "original_index": orig, "title": title_by_idx[orig]}
            for pos, orig in enumerate(order)
        ],
    }


# --------------------------------------------------------------------------- #
# 日志配置
# --------------------------------------------------------------------------- #
def setup_logging(log_dir: str | None = None, level: int = logging.INFO) -> None:
    if log_dir is None:
        log_dir = str(Path(__file__).resolve().parent / "logs")
    Path(log_dir).mkdir(parents=True, exist_ok=True)
    root = logging.getLogger("ppt_reorganize")
    root.setLevel(level)
    if not root.handlers:  # 避免重复 handler
        fh = logging.FileHandler(Path(log_dir) / "ppt_reorganize.log", encoding="utf-8")
        fh.setFormatter(logging.Formatter(
            "%(asctime)s [%(levelname)s] %(name)s: %(message)s"))
        root.addHandler(fh)
        sh = logging.StreamHandler()
        sh.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))
        root.addHandler(sh)


def to_json(obj) -> str:
    """序列化 dataclass/字典为 JSON。"""
    if hasattr(obj, "to_dict"):
        obj = obj.to_dict()
    return json.dumps(obj, ensure_ascii=False, indent=2)
