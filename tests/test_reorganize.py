"""pptx_reorganize 单元测试。

覆盖正常输入、边界情况与异常情况，共 11 个用例。

运行：
    python -m unittest tests.test_reorganize
或：
    python tests\test_reorganize.py
"""
import os
import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

# 让测试在直接运行时也能导入项目根的模块
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pptx_reorganize as core  # noqa: E402


def _copy_fixture(name: str, dest_dir: str) -> str:
    """把仓库内置合成用例复制到临时目录后返回路径。

    生成脚本已移除，sample/test1 等用例以静态文件随仓库提交；复制是为了让
    每个测试在隔离的临时目录中工作（含写回输入路径的异常用例）。
    """
    src = Path(__file__).resolve().parent.parent / "ppts" / name
    dst = os.path.join(dest_dir, name)
    shutil.copyfile(src, dst)
    return dst


class TestAnalyze(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmpdir = tempfile.TemporaryDirectory()
        cls.sample = _copy_fixture("sample.pptx", cls.tmpdir.name)

    @classmethod
    def tearDownClass(cls):
        cls.tmpdir.cleanup()

    def test_analyze_returns_structure(self):
        """正常：分析样例 PPT 返回 7 页 + 4 章节。"""
        r = core.analyze_pptx(self.sample)
        self.assertEqual(r.slide_count, 7)
        self.assertEqual(len(r.slides), 7)
        self.assertGreaterEqual(len(r.chapters), 3)
        # 首页识别为封面
        self.assertEqual(r.slides[0].special, "封面")
        # 至少有一个目录页
        self.assertTrue(any(s.special == "目录" for s in r.slides))

    def test_special_summary_patterns(self):
        """总结尾页：内容总结/本章小结/结论锚尾；章内正文「…技术总结」不锚尾。"""
        c = core._classify_special
        self.assertEqual(c("内容总结"), "总结")
        self.assertEqual(c("本章小结"), "总结")
        self.assertEqual(c("全文总结与展望"), "总结")
        self.assertEqual(c("结论"), "总结")
        self.assertIsNone(c("数据隐私保护技术总结"))
        self.assertIsNone(c("结构化设计方法"))

    def test_analyze_synthetic_small_deck(self):
        """脱敏小稿 test1（无分标题）端到端分析：封面/致谢识别、无异常。"""
        path = _copy_fixture("test1.pptx", type(self).tmpdir.name)
        r = core.analyze_pptx(path)
        self.assertGreaterEqual(r.slide_count, 12)
        self.assertTrue(any(s.special == "封面" for s in r.slides))
        # 无分标题小稿：结构切章不命中
        self.assertIsNone(core.induce_chapters_by_structure(r.slides))


class TestReorder(unittest.TestCase):
    def setUp(self):
        self.tmpdir = tempfile.TemporaryDirectory()
        self.sample = _copy_fixture("sample.pptx", self.tmpdir.name)
        self.out = os.path.join(self.tmpdir.name, "out.pptx")

    def tearDown(self):
        self.tmpdir.cleanup()

    def test_reorder_normal(self):
        """正常：重排 [3,1,2,5,4,6,7] 后输出文件 7 页。"""
        core.reorder_slides(self.sample, [3, 1, 2, 5, 4, 6, 7], self.out)
        from pptx import Presentation
        self.assertEqual(len(list(Presentation(self.out).slides)), 7)

    def test_reorder_identity(self):
        """边界：恒等顺序 1..N 应成功（不改顺序）。"""
        core.reorder_slides(self.sample, [1, 2, 3, 4, 5, 6, 7], self.out)
        from pptx import Presentation
        self.assertEqual(len(list(Presentation(self.out).slides)), 7)

    def test_reorder_subset(self):
        """子集：只输出指定页（删除页模式）。"""
        core.reorder_slides(self.sample, [1, 3, 5, 7], self.out)
        from pptx import Presentation
        self.assertEqual(len(list(Presentation(self.out).slides)), 4)

    def test_reorder_invalid_permutation(self):
        """异常：重复页码或越界页码应报 ValueError。"""
        with self.assertRaises(ValueError):
            core.reorder_slides(self.sample, [1, 1, 2, 3, 4, 5, 6], self.out)  # 重复
        with self.assertRaises(ValueError):
            core.reorder_slides(self.sample, [0, 2, 3, 4, 5, 6, 7], self.out)  # 越界

    def test_reorder_overwrite_source_rejected(self):
        """安全：输出路径与源相同时拒绝。"""
        with self.assertRaises(ValueError):
            core.reorder_slides(self.sample, [1, 2, 3, 4, 5, 6, 7], self.sample)

    def test_reorder_overwrite_existing_rejected(self):
        """安全：默认拒绝覆盖已存在文件。"""
        Path(self.out).write_bytes(b"old")
        with self.assertRaises(FileExistsError):
            core.reorder_slides(self.sample, [1, 2, 3, 4, 5, 6, 7], self.out)
        # --force 可覆盖
        core.reorder_slides(self.sample, [1, 2, 3, 4, 5, 6, 7], self.out, force=True)


class TestReorganizeByPurpose(unittest.TestCase):
    def setUp(self):
        self.tmpdir = tempfile.TemporaryDirectory()
        self.sample = _copy_fixture("sample.pptx", self.tmpdir.name)
        self.out = os.path.join(self.tmpdir.name, "out.pptx")

    def tearDown(self):
        self.tmpdir.cleanup()

    def test_reorganize_investor_purpose(self):
        """正常：按投资人目的重组，封面在前、致谢在后。"""
        order, _, _ = core.reorganize_by_purpose(
            self.sample, "投资人汇报", self.out
        )
        self.assertEqual(len(order), 7)
        # 封面应在最前
        analysis = core.analyze_pptx(self.sample)
        titles = [s.title for s in analysis.slides]
        first_orig = order[0]
        self.assertEqual(titles[first_orig - 1], "封面")
        # 致谢应在最后
        last_orig = order[-1]
        self.assertEqual(titles[last_orig - 1], "致谢")

    def test_reorganize_unknown_purpose_unchanged(self):
        """边界：未知目的应保持原顺序。"""
        order, _, _ = core.reorganize_by_purpose(
            self.sample, "完全无法匹配的关键词xyz", self.out
        )
        self.assertEqual(order, [1, 2, 3, 4, 5, 6, 7])

class TestPathValidation(unittest.TestCase):
    def test_non_pptx_rejected(self):
        """异常：非 .pptx 扩展名拒绝。"""
        with self.assertRaises(ValueError):
            core.validate_pptx_path("foo.txt", must_exist=False)
        with self.assertRaises(ValueError):
            core.validate_pptx_path("noext", must_exist=False)

    def test_dotdot_in_path_rejected(self):
        """安全：路径中含 '..' 拒绝。"""
        with self.assertRaises(ValueError):
            core.validate_pptx_path("../evil.pptx", must_exist=False)

    def test_missing_file_rejected(self):
        """异常：must_exist 时不存在文件报错。"""
        with self.assertRaises(FileNotFoundError):
            core.validate_pptx_path("not_exists.pptx", must_exist=True)


class TestTOC(unittest.TestCase):
    def setUp(self):
        self.tmpdir = tempfile.TemporaryDirectory()
        self.sample = _copy_fixture("sample.pptx", self.tmpdir.name)
        self.out = os.path.join(self.tmpdir.name, "out.pptx")

    def tearDown(self):
        self.tmpdir.cleanup()

    def test_toc_inserts_extra_slide(self):
        """可选功能：toc 后页数 +1。"""
        core.insert_toc(self.sample, self.out)
        from pptx import Presentation
        n = len(list(Presentation(self.out).slides))
        self.assertEqual(n, 8)  # 原 7 + 目录 1


class TestTOCHeadingFormats(unittest.TestCase):
    """任何格式的分标题都应进目录：常规（编号标题）与反向（名称标题+编号正文）。"""

    def _s(self, idx, title, subtitle="", body_len=None, paras=None, special=None):
        """构造带正文特征的 SlideInfo（body 默认恰为 标题+副标题 长度和）。"""
        return core.SlideInfo(
            index=idx, title=title, text_snippet="", n_shapes=2,
            has_image=False, has_table=False, layout="Blank", special=special,
            body_len=len(title) + len(subtitle) if body_len is None else body_len,
            paragraph_count=(2 if subtitle else 0) if paras is None else paras,
            subtitle=subtitle)

    def test_reverse_numbered_heading_enters_toc(self):
        """反向形态分标题（名称标题+编号正文首行）进目录，多段内容页不误判。"""
        slides = [
            self._s(1, "封面", special="封面"),
            self._s(2, "05", "卓越的人才培养"),                    # 常规形态分标题
            self._s(3, "育人特色", "05", body_len=132, paras=9),   # 编号内容页
            self._s(4, "常见问题解答", "03", body_len=15, paras=3),    # 反向内容页（多段）
            self._s(5, "招生政策问答", "02"),                   # 反向形态分标题
            self._s(6, "总体情况", "02", body_len=18, paras=3),    # 反向内容页
        ]
        sections = core._toc_sections(slides)
        names = [t for t, _, _ in sections]
        self.assertIn("05 卓越的人才培养", names)
        self.assertIn("02 招生政策问答", names)   # 反向形态进目录
        self.assertNotIn("常见问题解答", names)           # 多段内容页不是分标题
        self.assertEqual(len(sections), 2)            # 内容页不切分目录段
        # 两形态并存 → 分级：常规编号=板块内章(L1)，反向=顶层板块(L0)
        self.assertEqual(sections[0], ("05 卓越的人才培养", 2, 1))
        self.assertEqual(sections[1], ("02 招生政策问答", 5, 0))

    def test_single_form_deck_not_leveled(self):
        """单一形态稿（如「第N节」）不分级，全部条目层级 0。"""
        slides = [
            self._s(1, "第1节 技术初探"),
            self._s(2, "线性表", body_len=100, paras=8),
            self._s(3, "第2节 差分隐私"),
            self._s(4, "噪声", body_len=80, paras=6),
        ]
        sections = core._toc_sections(slides)
        self.assertEqual([t for t, _, _ in sections],
                         ["第1节 技术初探", "第2节 差分隐私"])
        self.assertTrue(all(lv == 0 for _, _, lv in sections))

    def test_reverse_heading_display_name(self):
        """反向形态展示名：编号在前、名称在后。"""
        s = self._s(1, "招生政策问答", "02")
        self.assertEqual(core._heading_display_name(s), "02 招生政策问答")


class TestTopSectionLeadPinning(unittest.TestCase):
    """两级编号稿：主板块开篇组（含板块分标题）恒定板块首位，不参与目的投票。"""

    @staticmethod
    def _deck():
        """10 页：封面 / 白01开篇(反向形态) / 黄01章(3页) / 黄02章(含章尾白03段2页)。"""
        slides = [_slide(1, "封面", "封面"),
                  core.SlideInfo(index=2, title="选择示例大学的十大理由",
                                 text_snippet="", n_shapes=2, has_image=False,
                                 has_table=False, layout="Blank", special=None,
                                 body_len=13, paragraph_count=2, subtitle="01"),
                  TestSectionAware._s(3, "高校选择因素", "01"),
                  TestSectionAware._heading(4, "01", "特色人才培养"),
                  TestSectionAware._s(5, "培养模式内容", "01"),
                  TestSectionAware._s(6, "地位延伸", "01"),
                  TestSectionAware._heading(7, "02", "升学与就业支持"),
                  TestSectionAware._s(8, "科研平台", "02"),
                  TestSectionAware._s(9, "常见问题解答", "03"),
                  TestSectionAware._s(10, "问答内容", "03")]
        return slides

    def test_lead_opening_pinned_before_voted_chapters(self):
        """投票把黄章任意重排，开篇引言仍恒在主板块首位、白板块殿后。"""
        slides = self._deck()
        analysis = core.AnalysisResult(
            file="x.pptx", slide_count=10, slides=slides, chapters=[])
        base = list(range(1, 11))
        # 第1次调用=板块投票（返回 None → 回退编号升序 白01→白03）；
        # 第2次调用=主板块黄章投票（反序 → 黄02 先、黄01 后）
        with mock.patch.object(core, "plan_reorder_by_llm", return_value=base), \
             mock.patch.object(core, "order_chapters_by_llm",
                               side_effect=[None, [1, 0]]):
            order, trace = core.resolve_order(
                analysis, "技术汇报", use_llm=True, smart=True)
        titles = [g["title"] for g in trace["induced_chapters"]]
        self.assertEqual(titles[0], "开篇引言")          # 开篇恒在板块首位
        self.assertEqual(titles[1], "02 升学与就业支持")  # 投票序生效
        self.assertEqual(titles[2], "01 特色人才培养")
        self.assertEqual(titles[3], "常见问题解答")           # 白板块殿后
        pos = {p: i for i, p in enumerate(order)}
        self.assertLess(pos[2], min(pos[p] for p in (4, 5, 6, 7, 8)))  # 开篇 < 全部黄章
        self.assertGreater(pos[9], max(pos[p] for p in (4, 5, 6, 7, 8)))  # 白板块殿后

    def test_board_vote_reorders_top_sections(self):
        """板块间目的投票：LLM 可把白板块排到主板块之前（票序生效）。"""
        slides = self._deck()
        analysis = core.AnalysisResult(
            file="x.pptx", slide_count=10, slides=slides, chapters=[])
        base = list(range(1, 11))
        # 第1次调用=板块投票（[1,0] → 白03 板块排到主板块前）；
        # 第2次调用=主板块黄章投票（[1,0] → 黄02 先）
        with mock.patch.object(core, "plan_reorder_by_llm", return_value=base), \
             mock.patch.object(core, "order_chapters_by_llm",
                               side_effect=[[1, 0], [1, 0]]):
            order, trace = core.resolve_order(
                analysis, "技术汇报", use_llm=True, smart=True)
        titles = [g["title"] for g in trace["induced_chapters"]]
        self.assertEqual(titles[0], "常见问题解答")           # 板块票序生效：白板块先
        self.assertEqual(titles[1], "开篇引言")           # 开篇仍恒在主板块首位
        self.assertEqual(titles[2], "02 升学与就业支持")
        pos = {p: i for i, p in enumerate(order)}
        self.assertLess(pos[9], pos[2])   # 白板块整体在主板块开篇之前
        self.assertLess(pos[2], min(pos[p] for p in (4, 5, 6, 7, 8)))  # 开篇绑定不破


def _slide(idx, title, special=None):
    """构造最小 SlideInfo，供锚定/封面识别测试使用。"""
    return core.SlideInfo(
        index=idx, title=title, text_snippet="", n_shapes=1,
        has_image=False, has_table=False, layout="Blank", special=special,
    )


class TestAnchorSpecialPages(unittest.TestCase):
    def test_anchor_moves_special_to_head_and_tail(self):
        """正常：封面/目录锚定到开头、致谢锚定到末尾，正文保持模型给的相对顺序。"""
        slides = [
            _slide(1, "XX大学", "封面"),
            _slide(2, "目录", "目录"),
            _slide(3, "正文A"),
            _slide(4, "谢谢", "致谢"),
            _slide(5, "正文B"),
        ]
        # 模型给出的乱序：致谢、正文A、封面、正文B、目录
        anchored = core.anchor_special_pages([4, 3, 1, 5, 2], slides)
        self.assertEqual(anchored[0], 1)   # 封面在首
        self.assertEqual(anchored[1], 2)   # 目录第二
        self.assertEqual(anchored[-1], 4)  # 致谢在尾
        self.assertEqual(anchored[2:4], [3, 5])  # 正文相对顺序保持
        self.assertEqual(sorted(anchored), [1, 2, 3, 4, 5])  # 仍是合法排列

    def test_anchor_full_tail_order(self):
        """正常：结尾按 总结→参考文献→致谢 锚定。"""
        slides = [
            _slide(1, "封面", "封面"),
            _slide(2, "正文A"),
            _slide(3, "总结", "总结"),
            _slide(4, "参考文献", "参考文献"),
            _slide(5, "致谢", "致谢"),
        ]
        anchored = core.anchor_special_pages([5, 2, 1, 4, 3], slides)
        self.assertEqual(anchored, [1, 2, 3, 4, 5])

    def test_anchor_rejects_invalid_order(self):
        """异常：非法排列传入锚定函数应报 ValueError。"""
        slides = [_slide(i, f"第{i}页") for i in range(1, 4)]
        with self.assertRaises(ValueError):
            core.anchor_special_pages([1, 1, 2], slides)


def _heading(idx, title, body=0):
    """分标题页/正文页构造（body>0 表示自带正文的「第X节+提纲」页）。"""
    return core.SlideInfo(
        index=idx, title=title, text_snippet="", n_shapes=1,
        has_image=False, has_table=False, layout="Blank", special=None,
        body_len=body, paragraph_count=1 if body else 0, bullet_count=0)


class TestBindSectionHeadings(unittest.TestCase):
    def test_heading_relocated_before_its_content(self):
        """两个分标题紧挨：前一个的归属内容在别处 → 移到归属内容前强制绑定。"""
        slides = [
            _heading(5, "第1节 隐私保护", body=30),   # 拥有 P6
            _heading(6, "内容1"),
            _heading(7, "第2节 线性表", body=30),     # 拥有 P8
            _heading(8, "内容2"),
        ]
        order, moves, removed = core.bind_section_headings([5, 7, 6, 8], slides)
        self.assertEqual(order, [5, 6, 7, 8])
        self.assertEqual(len(moves), 2)
        self.assertEqual(removed, [])

    def test_empty_divider_without_content_is_removed(self):
        """纯空分隔页相邻且无归属内容 → 删除；自带正文的分标题保留。"""
        slides = [
            _heading(1, "01"),                # 纯空分隔页，无归属内容
            _heading(2, "02"),                # 纯空分隔页，拥有 P3
            _heading(3, "内容页"),
            _heading(4, "第3节 安全多方计算", body=20),  # 自带正文，保留
        ]
        order, moves, removed = core.bind_section_headings([1, 2, 3, 4], slides)
        self.assertEqual(order, [2, 3, 4])
        self.assertEqual(len(removed), 1)
        self.assertIn("1", removed[0])

    def test_numbered_content_page_not_treated_as_heading(self):
        """纯编号标题但正文丰富（94~264字/多段/有图）的是内容页，不是分标题；
        只有稀疏章首页（~15字/3段）才算分标题。"""
        chapter_start = _heading(15, "03")                 # 默认 0 字 → 分标题
        content = _heading(16, "03", body=164)            # 丰富正文 → 内容页
        content.paragraph_count = 29
        self.assertTrue(core._is_section_heading(chapter_start))
        self.assertFalse(core._is_section_heading(content))
        content_pic = _heading(23, "05", body=94)
        content_pic.has_image = True
        content_pic.paragraph_count = 23
        self.assertFalse(core._is_section_heading(content_pic))
        # 稀疏的真实章首页不被判为相邻
        order, moves, removed = core.bind_section_headings(
            [15, 16], [chapter_start, content])
        self.assertEqual(order, [15, 16])
        self.assertEqual((moves, removed), ([], []))

    def test_named_section_heading_recognized_with_body(self):
        """「第N节 名称」即使带提纲正文也仍是分标题（不受稀疏门控限制）。"""
        h = _heading(5, "第1节 隐私保护技术初探", body=37)
        h.paragraph_count = 4
        self.assertTrue(core._is_section_heading(h))

    def test_subtitle_badge_not_treated_as_heading(self):
        """正文页副标题带「01」角标不算分标题，绑定不改变顺序。"""
        s = core.SlideInfo(
            index=1, title="国家历批次重点建设大学", text_snippet="",
            n_shapes=1, has_image=False, has_table=False, layout="Blank",
            body_len=50, paragraph_count=2, subtitle="01")
        self.assertFalse(core._is_section_heading(s))
        order, moves, removed = core.bind_section_headings(
            [1, 2, 3], [s, _heading(2, "02"), _heading(3, "正文")])
        self.assertEqual(order, [1, 2, 3])
        self.assertEqual((moves, removed), ([], []))

    def test_bind_is_idempotent(self):
        """对已合规顺序重复绑定，结果不变。"""
        slides = [
            _heading(5, "第1节 背景", body=30),
            _heading(6, "内容1"),
            _heading(7, "第2节 技术", body=30),
            _heading(8, "内容2"),
        ]
        once, _, _ = core.bind_section_headings([5, 6, 7, 8], slides)
        twice, m2, r2 = core.bind_section_headings(once, slides)
        self.assertEqual(once, [5, 6, 7, 8])
        self.assertEqual(twice, once)
        self.assertEqual((m2, r2), ([], []))


class TestCoverDetection(unittest.TestCase):
    def test_looks_like_cover_by_content(self):
        """封面内容识别：短的机构名为封面，章节性标题不是。"""
        self.assertTrue(core._looks_like_cover(_slide(17, "示例大学")))
        self.assertFalse(core._looks_like_cover(_slide(1, "精典案例分析与实践")))
        self.assertFalse(core._looks_like_cover(_slide(2, "主题研究与主要成果")))
        # 已被识别为其他特殊页的不再当封面候选
        self.assertFalse(core._looks_like_cover(_slide(15, "道德经", "致谢")))

    @staticmethod
    def _rich(idx, title, paras=3, body=40, bullets=0, special=None):
        return core.SlideInfo(
            index=idx, title=title, text_snippet="", n_shapes=1,
            has_image=False, has_table=False, layout="Blank", special=special,
            body_len=body, paragraph_count=paras, bullet_count=bullets)

    def test_pick_cover_finds_real_cover_in_shuffled_deck(self):
        """乱序稿：首页是重复章标题正文页，真封面在任意位置仍被内容识别找到。"""
        slides = [
            self._rich(1, "关键技术研究"),                 # 首页是正文（3 段）
            self._rich(2, "关键技术研究"),                 # 标题重复
            self._rich(3, "系统实现"),
            self._rich(4, "测试与验证"),
            self._rich(7, "示例大学", paras=1, body=21),  # 唯一短标题稀疏页
        ]
        cover = core._pick_cover(slides)
        self.assertIsNotNone(cover)
        self.assertEqual(cover.index, 7)

    def test_pick_cover_dense_first_page_no_fallback(self):
        """无内容候选且首页是多段正文页：不回退误判，返回 None。"""
        slides = [
            self._rich(1, "项目研究汇报", paras=3, body=50),
            self._rich(2, "项目研究汇报", paras=3, body=50),
            self._rich(3, "项目研究汇报", paras=3, body=50),
        ]
        self.assertIsNone(core._pick_cover(slides))

    def test_pick_cover_fallback_sparse_first_page(self):
        """无内容候选但首页极稀疏（1 行字）：允许首页先验回退。"""
        slides = [
            self._rich(1, "课程实验报告", paras=1, body=8),  # 含负词，过不了内容候选
            self._rich(2, "项目研究汇报", paras=3, body=50),
        ]
        cover = core._pick_cover(slides)
        self.assertIsNotNone(cover)
        self.assertEqual(cover.index, 1)

    def test_pick_cover_allows_multi_line_real_cover(self):
        """真实封面有 8 段短行（校名/题目/作者/导师/日期）仍应被选中。"""
        slides = [
            self._rich(1, "主题研究与主要成果", paras=7, body=70, bullets=3),
            self._rich(9, "示例大学", paras=8, body=95),  # 乱序后跑到第 9 位
        ]
        self.assertEqual(core._pick_cover(slides).index, 9)

    def test_pick_cover_rejects_dense_divider_page(self):
        """无项目符号但段落 >12 的密集页不参与封面候选。"""
        slides = [self._rich(3, "测试与验证", paras=13, body=200)]
        self.assertIsNone(core._pick_cover(slides))

    def test_pick_cover_org_signal_beats_sparsity(self):
        """多候选：含机构关键词者优先（即使正文稍多）；同档按稀疏度，再按首页。"""
        slides = [
            self._rich(1, "某某大学", paras=2, body=30),
            self._rich(5, "答辩论文", paras=1, body=10),
        ]
        self.assertEqual(core._pick_cover(slides).index, 1)
        # 同档：首页先验
        slides = [
            self._rich(1, "某某大学", paras=1, body=10),
            self._rich(5, "某某学院", paras=1, body=10),
        ]
        self.assertEqual(core._pick_cover(slides).index, 1)

    @staticmethod
    def _divider(idx, title, sub):
        s = core.SlideInfo(
            index=idx, title=title, text_snippet="", n_shapes=1,
            has_image=False, has_table=False, layout="Blank", special=None,
            body_len=10, paragraph_count=2, bullet_count=0)
        s.subtitle = sub
        return s

    def test_pick_cover_promo_deck_date_cover_vs_numeric_dividers(self):
        """宣讲稿：真封面标题是日期、校名在副标题；数字编号分隔页不得中选。"""
        cover = self._divider(1, "2024年X月", "示例大学")
        sec1 = self._divider(2, "选择示例大学的十大理由", "01")
        sec2 = self._divider(6, "01", "特色人才培养")
        sec3 = self._divider(48, "招生政策问答", "02")
        picked = core._pick_cover([cover, sec1, sec2, sec3])
        self.assertEqual(picked.index, 1)


class TestSectionAware(unittest.TestCase):
    """两级编号树：页编号解析 + 黄章内连续同号异章段剥离为顶层板块。"""

    @staticmethod
    def _s(idx, title="", subtitle="", body_len=80, paras=6, special=None):
        """构造带角标的内容页（默认有正文，不会被判成分标题）。"""
        return core.SlideInfo(
            index=idx, title=title, text_snippet="", n_shapes=1,
            has_image=False, has_table=False, layout="Blank", special=special,
            body_len=body_len, paragraph_count=paras, subtitle=subtitle)

    @classmethod
    def _heading(cls, idx, title, subtitle=""):
        return cls._s(idx, title=title, subtitle=subtitle, body_len=5, paras=2)

    def test_page_numbering_formats(self):
        """01→(1,)、1.1→(1,1)、1-1→(1,1)、1.1.2→(1,1,2)、第三章/（一）→(3,)/(1,)。"""
        cases = [
            ("01", "", (1,)), ("12", "", (12,)),
            ("1.1", "", (1, 1)), ("2-1", "", (2, 1)),
            ("1.1.2", "", (1, 1, 2)), ("", "第三章", (3,)),
            ("", "（一）", (1,)), ("定义与背景", "", None),
        ]
        for i, (title, sub, want) in enumerate(cases, start=1):
            s = self._s(i, title=title, subtitle=sub)
            self.assertEqual(core._page_numbering(s), want, msg=f"case {title!r}")

    def test_heading_page_numbering_none(self):
        """分标题页自身编号是章锚，_page_numbering 返回 None。"""
        s = self._heading(1, "10", "升学与就业支持")
        self.assertIsNone(core._page_numbering(s))

    def test_split_detaches_consecutive_foreign_badges(self):
        """黄10章内白02(3页)+挂靠、白03(2页)连续段 → 剥离为两个顶层板块。"""
        slides = [
            self._heading(1, "10", "升学与就业支持"),
            self._s(2, "校区布局", "10"),
            self._s(3, "招生政策问答", "02"),
            self._s(4, "招生流程", "02"),
            self._s(5, "招生补充"),                 # 无号挂靠白02
            self._s(6, "常见问题解答", "03"),
            self._s(7, "问答二", "03"),
        ]
        induced = [{"title": "10 升学与就业支持", "pages": [1, 2, 3, 4, 5, 6, 7]}]
        top = core.split_top_sections(induced, slides)
        self.assertIsNotNone(top)
        self.assertEqual(len(top), 3)
        main = top[0]
        self.assertIsNotNone(main["children"])
        self.assertEqual(main["children"][0]["pages"], [1, 2])
        self.assertIsNone(top[1]["children"])
        self.assertEqual(top[1]["pages"], [3, 4, 5])
        self.assertEqual(top[1]["title"], "招生政策问答")
        self.assertEqual(top[2]["pages"], [6, 7])

    def test_split_ignores_isolated_badge(self):
        """孤立单页异号数字（Africa 页正文「12」）不立板块 → 返回 None。"""
        slides = [
            self._heading(1, "09", "校园文化生活"),
            self._s(2, "Africa", "12"),
            self._s(3, "国际化成果", "09"),
        ]
        induced = [{"title": "09 校园文化生活", "pages": [1, 2, 3]}]
        self.assertIsNone(core.split_top_sections(induced, slides))

    def test_split_same_badge_stays_in_chapter(self):
        """同号角标（==章号）是章成员，不剥离。"""
        slides = [
            self._heading(1, "02", "升学与就业支持"),
            self._s(2, "科研平台", "02"),
            self._s(3, "科研团队", "02"),
        ]
        induced = [{"title": "02 升学与就业支持", "pages": [1, 2, 3]}]
        self.assertIsNone(core.split_top_sections(induced, slides))

    def test_split_multilevel_numbering(self):
        """点分编号 1.1/1.2 在章号 2 的章内 → 首段 1 异章，连续 2 页即剥离。"""
        slides = [
            self._heading(1, "2", "第二章"),
            self._s(2, "概述", "2"),
            self._s(3, "要点A", "1.1"),
            self._s(4, "要点B", "1.2"),
        ]
        induced = [{"title": "2 第二章", "pages": [1, 2, 3, 4]}]
        top = core.split_top_sections(induced, slides)
        self.assertIsNotNone(top)
        self.assertEqual(top[1]["pages"], [3, 4])

    def test_split_abandoned_when_chapter_would_empty(self):
        """剥离后章只剩分标题 → 放弃剥离，返回 None。"""
        slides = [
            self._heading(1, "02", "升学与就业支持"),
            self._s(2, "成果A", "03"),
            self._s(3, "成果B", "03"),
        ]
        induced = [{"title": "02 升学与就业支持", "pages": [1, 2, 3]}]
        self.assertIsNone(core.split_top_sections(induced, slides))

    def test_no_heading_returns_none(self):
        """无分标题的开篇组不参与剥离。"""
        slides = [self._s(1, "标题1"), self._s(2, "标题2", "02"), self._s(3, "标题3", "02")]
        induced = [{"title": "开篇引言", "pages": [1, 2, 3]}]
        self.assertIsNone(core.split_top_sections(induced, slides))

    def test_divider_section_num_chinese(self):
        """第三章 → 3, 第5节 → 5。"""
        s3 = core.SlideInfo(index=1, title="第三章", text_snippet="",
                            n_shapes=0, has_image=False, has_table=False,
                            layout="", subtitle="")
        self.assertEqual(core._divider_section_num(s3), 3)
        s5 = core.SlideInfo(index=2, title="第5节", text_snippet="",
                           n_shapes=0, has_image=False, has_table=False,
                           layout="", subtitle="")
        self.assertEqual(core._divider_section_num(s5), 5)
        s10 = core.SlideInfo(index=3, title="10", text_snippet="",
                            n_shapes=0, has_image=False, has_table=False,
                            layout="", subtitle="")
        self.assertEqual(core._divider_section_num(s10), 10)


class TestTrimDividers(unittest.TestCase):
    """删除章内冗余分隔页：章首页分隔页保留，章内其余删除。"""

    @staticmethod
    def _build_divider_slides():
        """10 页 PPT，分隔页 01（页1）和 02（页6）。"""
        plan = [
            "01", "概述", "背景", "目标", "方案",
            "02", "实现", "测试", "总结", "展望",
        ]
        return [_slide(i, title, None) for i, title in enumerate(plan, 1)]

    def test_trim_keeps_chapter_start_removes_internal(self):
        """章1含分隔页01(章首保留)、章2含分隔页02(章首保留)，
        跨章时章2的分隔页02因在新章首次出现→保留。"""
        slides = self._build_divider_slides()
        induced = [
            {"title": "前半", "pages": [1, 2, 3, 4, 5]},
            {"title": "后半", "pages": [6, 7, 8, 9, 10]},
        ]
        order = list(range(1, 11))
        result, removed = core.trim_internal_dividers(order, induced, slides)
        # 章首页1(01)和页6(02)都保留，无内部冗余
        self.assertEqual(result, order)
        self.assertEqual(len(removed), 0)

    def test_trim_removes_internal_divider(self):
        """章含两个分隔页→第二个删除。"""
        slides = self._build_divider_slides()
        induced = [{"title": "整篇", "pages": list(range(1, 11))}]
        order = list(range(1, 11))
        result, removed = core.trim_internal_dividers(order, induced, slides)
        # 页1(01)是章首→保留，页6(02)是章内冗余→删除
        self.assertEqual(result, [1, 2, 3, 4, 5, 7, 8, 9, 10])
        self.assertEqual(len(removed), 1)
        self.assertIn("02", removed[0])

    def test_trim_all_dividers_when_not_keeping_start(self):
        """keep_chapter_start=False → 所有分隔页都删除。"""
        slides = self._build_divider_slides()
        induced = [{"title": "整篇", "pages": list(range(1, 11))}]
        order = list(range(1, 11))
        result, removed = core.trim_internal_dividers(
            order, induced, slides, keep_chapter_start=False)
        self.assertEqual(result, [2, 3, 4, 5, 7, 8, 9, 10])
        self.assertEqual(len(removed), 2)

    def test_trim_no_dividers_unchanged(self):
        """无分隔页 → 原序不变。"""
        slides = [_slide(i, f"标题{i}", None) for i in range(1, 6)]
        induced = [{"title": "整篇", "pages": list(range(1, 6))}]
        result, removed = core.trim_internal_dividers(
            list(range(1, 6)), induced, slides)
        self.assertEqual(result, list(range(1, 6)))
        self.assertEqual(len(removed), 0)

    def test_trim_no_induced_positional(self):
        """无 induced → 按位置去重（遇到分隔页就标记，后续同位置删）。
        induced=None 时 chap_of 为空，所有分隔页视为不同章→都保留。"""
        slides = self._build_divider_slides()
        result, removed = core.trim_internal_dividers(
            list(range(1, 11)), None, slides)
        self.assertEqual(result, list(range(1, 11)))
        self.assertEqual(len(removed), 0)

    def test_trim_consecutive_dividers(self):
        """连续两个纯分隔页→删后一个；跨章连续也删。"""
        plan = ["01", "02", "概述", "03", "实现", "04", "05", "测试"]
        slides = [_slide(i, t, None) for i, t in enumerate(plan, 1)]
        induced = [
            {"title": "A", "pages": [1, 2, 3]},
            {"title": "B", "pages": [4, 5, 6]},
            {"title": "C", "pages": [7, 8]},
        ]
        result, removed = core.trim_internal_dividers(
            list(range(1, 9)), induced, slides)
        # 页1(01)章A首→保留，页2(02)章A内冗余→删，页3→保留，
        # 页4(03)章B首→保留，页5→保留，页6(04)章B内冗余→删，
        # 页7(05)章C首→保留，页8→保留
        # 第二遍：无连续分隔页（页1和页4之间有页3隔开）
        self.assertIn(1, result)
        self.assertNotIn(2, result)
        self.assertIn(4, result)
        self.assertNotIn(6, result)
        self.assertIn(7, result)
        # 确保结果中无连续分隔页
        div_pages = [p for p in result
                     if core._is_pure_divider(slides[p-1])]
        for i in range(1, len(div_pages)):
            prev_p = div_pages[i-1]
            cur_p = div_pages[i]
            gap = result.index(cur_p) - result.index(prev_p)
            self.assertGreater(gap, 1, f"连续分隔页: {prev_p},{cur_p}")

    def test_trim_consecutive_cross_chapter(self):
        """跨章连续分隔页：两章首页都是分隔页且相邻→删后者。"""
        plan = ["01", "内容A", "02", "内容B"]
        slides = [_slide(i, t, None) for i, t in enumerate(plan, 1)]
        induced = [
            {"title": "A", "pages": [1, 2]},
            {"title": "B", "pages": [3, 4]},
        ]
        # 故意把两章首页分隔页放一起
        order = [1, 3, 2, 4]
        result, removed = core.trim_internal_dividers(order, induced, slides)
        # 页1(01)章A首→保留，页3(02)章B首→保留(第一遍)，
        # 但第二遍：页1和页3连续→删页3
        self.assertIn(1, result)
        self.assertNotIn(3, result)  # 连续→删
        self.assertIn(2, result)
        self.assertIn(4, result)
        # 确保无连续分隔页
        div_pages = [p for p in result
                     if core._is_pure_divider(slides[p-1])]
        for i in range(1, len(div_pages)):
            self.assertGreater(
                result.index(div_pages[i]) - result.index(div_pages[i-1]), 1)


def _build_real_like_slides():
    """构造与真实 PPT 结构一致的 18 页（标题 + special 标记）。"""
    plan = [
        ("示例大学", "封面"), ("内容提要", "目录"),
        ("主题名称及内涵解析", None),
        ("主题研究与主要成果", None), ("主题研究与主要成果", None),
        ("主题研究与主要成果", None), ("主题研究与主要成果", None),
        ("主题研究与主要成果", None),
        ("研究问题选定及论证", None), ("研究问题选定及论证", None),
        ("研究问题选定及论证", None),
        ("精典案例分析与实践", None), ("精典案例分析与实践", None),
        ("精典案例分析与实践", None),
        ("研究过程与问题分析", None),
        ("研究结论与主要建议", "总结"), ("参考文献", "参考文献"),
        ("道德经", "致谢"),
    ]
    return [_slide(i, title, special) for i, (title, special) in enumerate(plan, 1)]


class TestOrdinalAndStage(unittest.TestCase):
    def test_ordinal_from_subtitle(self):
        """正常：副标题带编号且正文无多编号 → 取副标题编号。"""
        self.assertEqual(core._extract_ordinal("2.漏洞同源性判别模块", [2]), 2)

    def test_ordinal_single_number_anywhere(self):
        """正常：副标题无编号但全页仅一个段首编号 → 采用该编号。"""
        self.assertEqual(core._extract_ordinal("关键功能模块实现", [1]), 1)

    def test_ordinal_inline_list_excluded(self):
        """脱出：同一页列了 1、2、3（页内列举）→ 不取序号。"""
        self.assertIsNone(core._extract_ordinal("系统架构设计", [1, 2, 3]))

    def test_ordinal_no_numbers(self):
        """脱出：无任何编号 → None。"""
        self.assertIsNone(core._extract_ordinal("核心矛盾", []))

    def test_stage_subtitle_priority_conclusion(self):
        """副标题优先：组标题含“案例/实践”，副标题“测试结论”仍判为结论阶段(5)。"""
        self.assertEqual(core._classify_stage("精典案例分析与实践", "测试结论", ""), 5)

    def test_stage_design_beats_test(self):
        """权重：副标题“测试案例设计”中“设计”(权重2)压过“测试”(1) → 设计阶段(2)。"""
        self.assertEqual(core._classify_stage("精典案例分析与实践", "测试案例设计", ""), 2)

    def test_stage_process_result(self):
        self.assertEqual(core._classify_stage("精典案例分析与实践", "测试过程与结果", ""), 4)

    def test_stage_no_signal(self):
        """脱出：无任何阶段词 → None（不硬贴标签）。"""
        self.assertIsNone(core._classify_stage("研究问题选定及论证", "关于Structure2vec", ""))


class TestEnforceOrdinal(unittest.TestCase):
    def setUp(self):
        self.slides = _build_real_like_slides()
        # 真实 PPT：原4=序号1、原5=序号2、原6=序号4；原7/8 是页内列举，无序号
        self.slides[3].ordinal, self.slides[4].ordinal, self.slides[5].ordinal = 1, 2, 4
        self.base_order = list(range(1, 19))

    def test_slot_backfill_fixes_numbered_pages_only(self):
        """核心：模型给的组内顺序 5,4,6 被槽位回填修正为 4,5,6，无编号的 7,8 不动。"""
        order = [1, 2, 3, 5, 4, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18]
        out, fixed = core.refine_within_groups(order, self.slides)
        self.assertEqual(out[:8], [1, 2, 3, 4, 5, 6, 7, 8])
        self.assertEqual(fixed, 1)

    def test_conflicting_ordinals_skipped(self):
        """脱出：同组出现重复编号（证据冲突）→ 不修正。"""
        self.slides[3].ordinal = 1
        self.slides[4].ordinal = 1  # 与原4冲突
        order = [1, 2, 3, 5, 4, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18]
        out, fixed = core.refine_within_groups(order, self.slides)
        self.assertEqual(out, order)
        self.assertEqual(fixed, 0)

    def test_already_sorted_skipped(self):
        """脱出：已按编号有序 → 不动，fixed=0。"""
        out, fixed = core.refine_within_groups(self.base_order, self.slides)
        self.assertEqual(out, self.base_order)
        self.assertEqual(fixed, 0)

    def test_stage_block_after_numbered_pages(self):
        """④兜底：序号归位后，尾部无编号连续块 8(阶段4),7(阶段2) 理顺为 7,8。"""
        self.slides[6].stage, self.slides[7].stage = 2, 4  # 原7架构设计、原8界面展示
        order = [1, 2, 3, 5, 4, 6, 8, 7, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18]
        out, fixed = core.refine_within_groups(order, self.slides)
        self.assertEqual(out[:8], [1, 2, 3, 4, 5, 6, 7, 8])
        self.assertEqual(fixed, 2)  # 序号修正 1 处 + 阶段块 1 处

    def test_stage_block_with_none_tail(self):
        """④兜底：全无编号组 10(阶段1),11(无标签),9(阶段0) → 9,10,11（无标签置尾）。"""
        self.slides[8].stage, self.slides[9].stage, self.slides[10].stage = 0, 1, None
        order = [1, 2, 3, 4, 5, 6, 7, 8, 10, 11, 9, 12, 13, 14, 15, 16, 17, 18]
        out, fixed = core.refine_within_groups(order, self.slides)
        self.assertEqual(out[8:11], [9, 10, 11])
        self.assertEqual(fixed, 1)

    def test_stage_non_contiguous_not_crossed(self):
        """保守：无编号页被编号页隔开（不连续）时不跨块移动。"""
        self.slides[6].stage, self.slides[7].stage = 2, 4  # 原7、原8
        # 组内排列：8,4,7,5,6 —— 无编号页 8、7 不连续
        order = [1, 2, 3, 8, 4, 7, 5, 6, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18]
        out, fixed = core.refine_within_groups(order, self.slides)
        self.assertEqual(out[3:8], [8, 4, 7, 5, 6])  # 保持不动
        self.assertEqual(fixed, 0)

    def test_stage_insufficient_distinct_skipped(self):
        """脱出：块内只有 1 种已知阶段（其余无标签）→ 证据不足，不动。"""
        self.slides[8].stage, self.slides[9].stage, self.slides[10].stage = 0, 0, None
        order = [1, 2, 3, 4, 5, 6, 7, 8, 10, 11, 9, 12, 13, 14, 15, 16, 17, 18]
        out, fixed = core.refine_within_groups(order, self.slides)
        self.assertEqual(out[8:11], [10, 11, 9])
        self.assertEqual(fixed, 0)

    def test_preserve_narrative_keeps_llm_stage_order(self):
        """目的叙事保留：LLM 章内序 8(阶段4),7(阶段2) 不再被阶段精修掰回 7,8，
        但编号槽位回填仍然生效。"""
        self.slides[6].stage, self.slides[7].stage = 2, 4
        order = [1, 2, 3, 5, 4, 6, 8, 7, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18]
        out, fixed = core.refine_within_groups(
            order, self.slides, preserve_narrative=True)
        self.assertEqual(out[:8], [1, 2, 3, 4, 5, 6, 8, 7])  # 编号修正、阶段不动
        self.assertEqual(fixed, 1)


class TestWithinChapterGuides(unittest.TestCase):
    def test_every_purpose_has_distinct_guide(self):
        """12 个目的都有准则，且科普/答辩/路演等与默认技术叙事不同。"""
        purposes = ["投资人路演", "技术汇报", "课堂教学", "招生综合宣讲",
                    "产品评审", "综合汇报", "科普宣讲", "学术答辩", "工作总结",
                    "项目汇报", "产品发布", "培训讲座"]
        guides = {p: core._within_chapter_guide(p) for p in purposes}
        self.assertTrue(all(guides.values()))
        default = core._DEFAULT_WITHIN_GUIDE
        for p in ("科普宣讲", "学术答辩", "投资人路演", "产品发布", "工作总结"):
            self.assertNotEqual(guides[p], default)

    def test_free_text_purpose_fallback(self):
        """未枚举但带目的词的自由文本也能命中差异化准则。"""
        self.assertEqual(core._within_chapter_guide("面向大众做科普"),
                         core.WITHIN_CHAPTER_GUIDES["科普宣讲"])
        self.assertEqual(core._within_chapter_guide("毕业答辩用"),
                         core.WITHIN_CHAPTER_GUIDES["学术答辩"])
        self.assertEqual(core._within_chapter_guide("未知场景"),
                         core._DEFAULT_WITHIN_GUIDE)

    def test_prompt_contains_purpose_guideline(self):
        """章内 prompt 必须嵌入目的名、该目的专属准则与编号保护。"""
        slides = _build_real_like_slides()
        briefs = core.build_slide_briefs(slides)
        prompt = core._build_within_chapter_prompt(
            briefs, "第3节 差分隐私", [49, 50, 51], "科普宣讲")
        self.assertIn("科普宣讲", prompt)
        self.assertIn("身边现象", prompt)
        prompt2 = core._build_within_chapter_prompt(
            briefs, "第3节 差分隐私", [49, 50, 51], "学术答辩")
        self.assertIn("论文论证链", prompt2)
        self.assertIn("编号", prompt2)
        self.assertNotEqual(prompt, prompt2)


class TestAssembleByChapters(unittest.TestCase):
    def setUp(self):
        self.slides = _build_real_like_slides()

    def test_head_chapter_body_tail(self):
        """正常：特殊页置首尾，章节按给定顺序，章内保持入参相对序。"""
        induced = [
            {"title": "主题部分", "pages": [3, 4, 5, 6, 7, 8]},
            {"title": "问题与案例", "pages": [9, 10, 11, 12, 13, 14, 15]},
        ]
        # 入参：封面/目录在前；A 组逆序、B 组逆序；尾部正常
        order = [1, 2, 8, 7, 6, 5, 4, 3, 15, 14, 13, 12, 11, 10, 9, 16, 17, 18]
        out = core.assemble_by_chapters(order, self.slides, induced)
        self.assertEqual(out[:2], [1, 2])
        self.assertEqual(out[2:8], [8, 7, 6, 5, 4, 3])    # A 组章内保序
        self.assertEqual(out[8:15], [15, 14, 13, 12, 11, 10, 9])  # B 组
        self.assertEqual(out[-3:], [16, 17, 18])

    def test_special_pages_extracted_even_when_clustered(self):
        """即使模型把封面/目录聚进了章节，组装时仍抽到首尾。"""
        induced = [
            {"title": "含封面的组", "pages": [1, 3, 4, 5]},
            {"title": "其余", "pages": [2] + list(range(6, 19))},
        ]
        out = core.assemble_by_chapters(list(range(1, 19)), self.slides, induced)
        self.assertEqual(out[0], 1)
        self.assertEqual(out[1], 2)
        self.assertEqual(out[-1], 18)
        self.assertEqual(sorted(out), list(range(1, 19)))
        # 组1 正文（3,4,5）先于组2 正文
        self.assertEqual(out[2:5], [3, 4, 5])


class TestInducedChaptersValidation(unittest.TestCase):
    def setUp(self):
        self.slides = _build_real_like_slides()

    def test_valid_payload(self):
        payload = {"chapters": [
            {"title": "开篇", "pages": [1, 2]},
            {"title": "主体", "pages": list(range(3, 16))},
            {"title": "结尾", "pages": [16, 17, 18]},
        ]}
        with mock.patch.object(core, "_llm_chat_json", return_value=payload):
            out = core.induce_chapters_by_llm(self.slides, "技术汇报")
        self.assertEqual(len(out), 3)
        self.assertEqual(out[1]["title"], "主体")

    def test_missing_page_returns_none(self):
        """脱出：聚类未覆盖全部页面 → None。"""
        payload = {"chapters": [
            {"title": "开篇", "pages": [1, 2]},
            {"title": "主体", "pages": list(range(3, 17))},  # 缺 17/18
        ]}
        with mock.patch.object(core, "_llm_chat_json", return_value=payload):
            self.assertIsNone(core.induce_chapters_by_llm(self.slides, "技术汇报"))

    def test_too_few_chapters_returns_none(self):
        """脱出：只聚出 1 章 → None。"""
        payload = {"chapters": [{"title": "全部", "pages": list(range(1, 19))}]}
        with mock.patch.object(core, "_llm_chat_json", return_value=payload):
            self.assertIsNone(core.induce_chapters_by_llm(self.slides, "技术汇报"))

    def test_llm_unavailable_returns_none(self):
        """脱出：模型调用失败 → None（由上层继续锚定兜底）。"""
        with mock.patch.object(core, "_llm_chat_json", return_value=None):
            self.assertIsNone(core.induce_chapters_by_llm(self.slides, "技术汇报"))


class TestChapterOrdering(unittest.TestCase):
    def setUp(self):
        self.slides = _build_real_like_slides()

    def test_borda_merge(self):
        # 多数轮 [0,1,2] 压过单轮 [1,0,2]
        self.assertEqual(
            core._borda_merge_orders([[0, 1, 2], [0, 1, 2], [1, 0, 2]], 3),
            [0, 1, 2])
        # 某轮漏章（章 2 缺席得惩罚分 k=3）：序 0,1 与 0,2,1 聚合
        out = core._borda_merge_orders([[0, 1], [0, 2, 1]], 3)
        self.assertEqual(out[0], 0)
        self.assertEqual(sorted(out), [0, 1, 2])  # 全覆盖

    def test_chapter_order_tolerant_fill(self):
        """模型漏掉 1 页封面组：主序保留，封面章前置补位。"""
        slides = []
        for i, sp in enumerate([("项目汇报封面", "封面"), ("背景内容", None), ("方案内容", None)], start=1):
            slides.append(core.SlideInfo(
                index=i, title=sp[0], text_snippet="", n_shapes=1,
                has_image=False, has_table=False, layout="", special=sp[1]))
        induced = [
            {"title": "封面", "pages": [1]},
            {"title": "背景", "pages": [2]},
            {"title": "方案", "pages": [3]},
        ]
        with mock.patch.object(core, "_HAS_JIEBA", False), \
             mock.patch.object(core, "_llm_chat_json",
                               return_value={"slots": [
                                   {"chapter": 1, "slot": "background"},
                                   {"chapter": 2, "slot": "solution"}]}):  # 漏 0
            out = core.order_chapters_by_llm(
                induced, slides, "技术汇报")
        self.assertEqual(out, [0, 1, 2])
        # 仅 1 个有效章（章2）：新语义下也采纳，漏章 0/1 按规则槽秩确定性补回
        with mock.patch.object(core, "_HAS_JIEBA", False), \
             mock.patch.object(core, "_llm_chat_json",
                               return_value={"slots": [
                                   {"chapter": 2, "slot": "solution"}]}):
            out = core.order_chapters_by_llm(
                induced, slides, "技术汇报")
        self.assertEqual(out, [0, 1, 2])

    def test_zero_padded_string_chapter_resolves_by_title(self):
        """模型回传标题自带编号「01」时按标题反查列表索引，不发生 off-by-one。"""
        infos = [
            {"title": "01 选择理由", "pages": [1, 2], "keywords": [], "stage": ""},
            {"title": "02 招生情况", "pages": [3], "keywords": [], "stage": ""},
            {"title": "03 常见问题解答", "pages": [4], "keywords": [], "stage": ""},
        ]
        resp = {"slots": [
            {"chapter": "01", "slot": "implement"},
            {"chapter": "02", "slot": "background"},
            {"chapter": "03", "slot": "case"}]}
        with mock.patch.object(core, "_llm_chat_json", return_value=resp):
            out = core._order_chapters_once(
                infos, 3, "技术汇报", {}, 0.1, guess_rank={}, special_allowed={})
        # background(章1) < implement(章0) < case(章2)
        self.assertEqual(out, [1, 0, 2])

    def test_prompt_requires_complete_enumeration(self):
        """方案2：prompt 必须声明 slots 数量恰好等于章数、编号全覆盖。"""
        infos = [
            {"title": f"{i:02d} 板块{i}", "pages": [i + 1],
             "keywords": [], "stage": ""}
            for i in range(1, 4)]
        p = core._build_chapter_order_prompt(infos, "招生综合宣讲")
        self.assertIn("恰好包含 3 项", p)
        self.assertIn("编号 0 到 2", p)
        self.assertIn("严禁遗漏任何章", p)

    def test_special_slot_gate_and_relaxed_fallback(self):
        """普通章占专用槽：严格门控拒绝；全部被拒时宽松放行保持原序。"""
        infos = [
            {"title": "背景", "pages": [1], "keywords": [], "stage": ""},
            {"title": "方案", "pages": [2], "keywords": [], "stage": ""},
            {"title": "成果", "pages": [3], "keywords": [], "stage": ""},
        ]
        # 全部塞 thanks（无致谢页）→ 严格门控致空 → 宽松二次扫描 → 原序
        resp_all = {"slots": [
            {"chapter": 0, "slot": "thanks"},
            {"chapter": 1, "slot": "thanks"},
            {"chapter": 2, "slot": "thanks"}]}
        with mock.patch.object(core, "_llm_chat_json", return_value=resp_all):
            out = core._order_chapters_once(
                infos, 3, "技术汇报", {}, 0.1, guess_rank={}, special_allowed={})
        self.assertEqual(out, [0, 1, 2])
        # 仅 1 章归入正常槽、2 章乱塞 cover/toc：采纳有效章 0，乱塞的 1/2 不走
        # 专用槽，按原序补全（不再整轮丢弃）
        resp_mix = {"slots": [
            {"chapter": 0, "slot": "background"},
            {"chapter": 1, "slot": "cover"},
            {"chapter": 2, "slot": "toc"}]}
        with mock.patch.object(core, "_llm_chat_json", return_value=resp_mix):
            out = core._order_chapters_once(
                infos, 3, "技术汇报", {}, 0.1, guess_rank={}, special_allowed={})
        self.assertEqual(out, [0, 1, 2])
        # 零可解析归类（空 slots / 越界章号）才丢弃本轮 → None
        for resp_none in ({"slots": []},
                          {"slots": [{"chapter": 9, "slot": "background"}]}):
            with mock.patch.object(core, "_llm_chat_json", return_value=resp_none):
                out = core._order_chapters_once(
                    infos, 3, "技术汇报", {}, 0.1, guess_rank={},
                    special_allowed={})
            self.assertIsNone(out)
        # 含封面的章占 cover 槽合法（special_allowed 放行）
        resp_ok = {"slots": [
            {"chapter": 0, "slot": "cover"},
            {"chapter": 1, "slot": "background"},
            {"chapter": 2, "slot": "result"}]}
        allowed = {"cover": {0}, "toc": set(), "reference": set(), "thanks": set()}
        with mock.patch.object(core, "_llm_chat_json", return_value=resp_ok):
            out = core._order_chapters_once(
                infos, 3, "技术汇报", {}, 0.1, guess_rank={},
                special_allowed=allowed)
        self.assertEqual(out, [0, 1, 2])
        # 板块欠填：模型只回 1 个章（主板块 cover，合法）→ 采纳并补全其余章，
        # 不再整轮失败（复现招生宣讲板块投票 3 轮全失败场景）
        resp_underfill = {"slots": [{"chapter": "01", "slot": "cover"}]}
        board_infos = [
            {"title": "01 选择示例大学的十大理由", "pages": [1, 2],
             "keywords": [], "stage": ""},
            {"title": "02 招生政策问答", "pages": [3],
             "keywords": [], "stage": ""},
            {"title": "03 常见问题解答", "pages": [4],
             "keywords": [], "stage": ""},
        ]
        with mock.patch.object(core, "_llm_chat_json",
                               return_value=resp_underfill):
            out = core._order_chapters_once(
                board_infos, 3, "招生综合宣讲", {}, 0.1, guess_rank={},
                special_allowed=allowed)
        self.assertEqual(out, [0, 1, 2])

    def test_slot_template_forces_style(self):
        """同组章在不同目的下经槽位模板必然产出不同顺序（风格硬约束）。"""
        slides = []
        titles = ["研究背景", "成果价值", "系统实现", "应用案例"]
        for i, t in enumerate(titles, start=1):
            slides.append(core.SlideInfo(
                index=i, title=t, text_snippet="", n_shapes=1,
                has_image=False, has_table=False, layout=""))
        induced = [{"title": t, "pages": [i]} for i, t in enumerate(titles, start=1)]
        # 模型始终按输入序归类到各自正确槽位（模拟稳定语义分类）
        slots_resp = {"slots": [
            {"chapter": 0, "slot": "background"},
            {"chapter": 1, "slot": "result"},
            {"chapter": 2, "slot": "implement"},
            {"chapter": 3, "slot": "case"}]}
        with mock.patch.object(core, "_HAS_JIEBA", False), \
             mock.patch.object(core, "_llm_chat_json", return_value=slots_resp):
            tech = core.order_chapters_by_llm(induced, slides, "技术汇报")
            inv = core.order_chapters_by_llm(induced, slides, "投资路演")
            teach = core.order_chapters_by_llm(induced, slides, "教学讲解")
        self.assertEqual(tech, [0, 2, 3, 1])   # 背景→实现→案例→成果
        self.assertEqual(inv, [1, 3, 2, 0])    # 成果→案例→实现→背景
        self.assertEqual(teach, [0, 2, 3, 1])

    def test_within_chapter_order_valid_and_invalid(self):
        with mock.patch.object(core, "_llm_chat_json",
                               return_value={"order": [5, 3, 4, 6]}):
            out = core.order_pages_within_chapter_by_llm(
                "实现", [3, 4, 5, 6], self.slides, "技术汇报")
        self.assertEqual(out, [5, 3, 4, 6])
        # 漏 1 页：宽容补尾（按页号升序），不丢弃整次调用
        with mock.patch.object(core, "_llm_chat_json",
                               return_value={"order": [3, 4, 5]}):
            out = core.order_pages_within_chapter_by_llm(
                "实现", [3, 4, 5, 6], self.slides, "技术汇报")
        self.assertEqual(out, [3, 4, 5, 6])
        # 有效页不足一半 → None，交上层回退
        with mock.patch.object(core, "_llm_chat_json",
                               return_value={"order": [3]}):
            self.assertIsNone(core.order_pages_within_chapter_by_llm(
                "实现", [3, 4, 5, 6], self.slides, "技术汇报"))
        # 超阈值不调用直接 None
        with mock.patch.object(core, "_llm_chat_json") as m:
            self.assertIsNone(core.order_pages_within_chapter_by_llm(
                "大章", list(range(1, 30)), self.slides, "技术汇报"))
            m.assert_not_called()

    def test_fill_missing_by_original_neighbor(self):
        """漏页按原稿邻接插回（不是统一扔章末）。"""
        f = core._fill_missing_by_original_order
        # 中间漏页：原序 [48,49,50,51,52]，LLM 漏 49 → 插到前邻 48 之后
        self.assertEqual(
            f([48, 50, 51, 52], [48, 49, 50, 51, 52]),
            [48, 49, 50, 51, 52])
        # 漏页的前邻也被 LLM 重排到远处：仍紧随前邻（保持 LLM 叙事，只补漏）
        self.assertEqual(
            f([52, 51, 48, 50], [48, 49, 50, 51, 52]),
            [52, 51, 48, 49, 50])
        # 章首漏页：找最近后邻插到其前
        self.assertEqual(
            f([49, 50, 51], [48, 49, 50, 51]),
            [48, 49, 50, 51])
        # 连续多页漏排：成链归位
        self.assertEqual(
            f([48, 52], [48, 49, 50, 51, 52]),
            [48, 49, 50, 51, 52])
        # 全覆盖原样返回
        self.assertEqual(f([3, 4, 5], [3, 4, 5]), [3, 4, 5])

    def test_within_chapter_missing_inserted_at_neighbor(self):
        """端到端：章内精排漏掉概述页时回到分标题前邻之后，而非章末。"""
        with mock.patch.object(core, "_llm_chat_json",
                               return_value={"order": [48, 50, 51, 52]}):
            out = core.order_pages_within_chapter_by_llm(
                "招生政策问答", [48, 49, 50, 51, 52],
                self.slides, "招生综合宣讲")
        self.assertEqual(out, [48, 49, 50, 51, 52])


class TestStructuralInduction(unittest.TestCase):
    def test_named_sections_split_with_opening(self):
        """第N节分标题切章：开篇 + 各节，页区间完整、特殊页排除。"""
        slides = [_slide(1, "封面", "封面"), _heading(2, "背景焦点", body=40),
                  _heading(3, "引言问题", body=40),
                  _heading(4, "第1节 技术初探", body=20)]
        slides += [_heading(i, f"初探内容{i}", body=50) for i in range(5, 9)]
        slides += [_heading(9, "第2节 线性表", body=20)]
        slides += [_heading(i, f"匿名内容{i}", body=50) for i in range(10, 13)]
        slides += [_slide(13, "内容总结", "总结")]
        groups = core.induce_chapters_by_structure(slides)
        self.assertIsNotNone(groups)
        self.assertEqual([g["title"] for g in groups],
                         ["开篇引言", "第1节 技术初探", "第2节 线性表"])
        self.assertEqual([g["pages"] for g in groups],
                         [[2, 3], [4, 5, 6, 7, 8], [9, 10, 11, 12]])

    def test_numeric_dividers_split_numbered_content_not(self):
        """稀疏数字分隔页切章；标题带编号角标的内容页（164字）不切。"""
        slides = [_slide(1, "封面", "封面"), _heading(2, "导语", body=40),
                  _heading(3, "01"), _heading(4, "背景", body=50)]
        rich = _heading(5, "03", body=164)       # 思源式编号内容页
        rich.paragraph_count = 29
        slides += [rich, _heading(6, "02"), _heading(7, "成果", body=50)]
        groups = core.induce_chapters_by_structure(slides)
        self.assertIsNotNone(groups)
        self.assertEqual([g["pages"] for g in groups],
                         [[2], [3, 4, 5], [6, 7]])

    def test_single_or_zero_headings_returns_none(self):
        slides = [_slide(1, "封面", "封面"),
                  _heading(2, "第1节 唯一一节", body=20),
                  _heading(3, "内容", body=50)]
        self.assertIsNone(core.induce_chapters_by_structure(slides))
        plain = [_slide(1, "封面", "封面"), _heading(2, "正文", body=50)]
        self.assertIsNone(core.induce_chapters_by_structure(plain))

    def test_empty_heading_returns_none(self):
        """两个分标题紧邻、前一个无内容页 → 放弃结构检定，升级后续链路。"""
        slides = [_slide(1, "封面", "封面"),
                  _heading(2, "第1节 空标题", body=20),
                  _heading(3, "第2节 有内容", body=20),
                  _heading(4, "内容", body=50)]
        self.assertIsNone(core.induce_chapters_by_structure(slides))


class TestStructuralResolveOrder(unittest.TestCase):
    def _deck(self):
        """33 页：封面 / 2 引言 / 第1节(P4..16) / 第2节(P17..32) / 总结。"""
        slides = [_slide(1, "封面", "封面"),
                  _heading(2, "背景焦点", body=40), _heading(3, "引言问题", body=40),
                  _heading(4, "第1节 技术方案", body=20)]
        slides += [_heading(i, f"方案内容{i}", body=50) for i in range(5, 17)]
        slides.append(_heading(17, "第2节 商业成果", body=20))
        slides += [_heading(i, f"成果内容{i}", body=50) for i in range(18, 33)]
        slides.append(_slide(33, "内容总结", "总结"))
        return slides

    def test_blocks_contiguous_heading_leads_inner_may_reorder(self):
        """粗排把两章穿插到极致：章块完整连续、分标题领头；章内跟随粗排可变动。"""
        slides = self._deck()
        analysis = core.AnalysisResult(
            file="x.pptx", slide_count=33, slides=slides, chapters=[])
        # 粗排：第2节（倒序）→ 第1节（倒序）→ 引言，特殊页夹在任意位置
        fallback = [1] + list(range(32, 16, -1)) + list(range(16, 3, -1)) + [2, 3, 33]
        with mock.patch.object(core, "plan_reorder_by_purpose", return_value=fallback):
            order, trace = core.resolve_order(
                analysis, "无法匹配目的xyz", use_llm=False, smart=True)
        self.assertEqual(trace["induce_mode"], "structural_sections")
        self.assertEqual(trace["chapter_source"], "structural_sections")
        self.assertEqual(order[0], 1)
        self.assertEqual(order[-1], 33)
        # 章块连续（集合一致即可，章内允许被粗排/精排改动）
        self.assertEqual(sorted(order[1:17]), list(range(17, 33)))
        self.assertEqual(sorted(order[17:30]), list(range(4, 17)))
        self.assertEqual(sorted(order[30:32]), [2, 3])
        pos = {p: i for i, p in enumerate(order)}
        # 分标题始终在本章所有内容之前（块内最前）
        self.assertEqual(order[1], 17)
        self.assertEqual(order[17], 4)
        self.assertLess(pos[17], min(pos[p] for p in range(18, 33)))
        self.assertLess(pos[4], min(pos[p] for p in range(5, 17)))


class TestRuleTitleInduction(unittest.TestCase):
    def setUp(self):
        self.slides = _build_real_like_slides()

    def test_groups_by_title(self):
        groups = core.induce_chapters_by_rules(self.slides)
        self.assertIsNotNone(groups)
        by_title = {g["title"]: g["pages"] for g in groups}
        self.assertEqual(by_title["主题研究与主要成果"], [4, 5, 6, 7, 8])
        self.assertEqual(by_title["精典案例分析与实践"], [12, 13, 14])
        # 特殊页（封面/目录/总结/参考/致谢）不参与
        all_pages = [p for g in groups for p in g["pages"]]
        self.assertEqual(sorted(all_pages), list(range(3, 16)))

    def test_dominant_title_returns_none(self):
        """单一标题吞掉绝大多数正文 → 标题无区分度，升级 LLM 层。"""
        for s in self.slides[2:15]:
            s.title = "同一个标题"
        self.assertIsNone(core.induce_chapters_by_rules(self.slides))

    def test_too_many_title_groups_returns_none(self):
        """标题全不同（过碎）→ 升级 LLM 层。"""
        for i, s in enumerate(self.slides[2:15], start=3):
            s.title = f"完全不同的标题{i}"
        self.assertIsNone(core.induce_chapters_by_rules(self.slides))

    def test_shuffle_does_not_affect_groups(self):
        """打乱只改变位置，全局标题归组结果不变（组内页序排序后比较）。"""
        g1 = core.induce_chapters_by_rules(self.slides)
        shuffled = [self.slides[i] for i in [0, 17, 5, 2, 11, 1, 7, 14, 9, 3, 13,
                                             6, 10, 4, 8, 12, 15, 16]]
        g2 = core.induce_chapters_by_rules(shuffled)
        norm = lambda gs: sorted((g["title"], tuple(sorted(g["pages"]))) for g in gs)
        self.assertEqual(norm(g1), norm(g2))


class TestResolveOrderCascade(unittest.TestCase):
    def setUp(self):
        self.slides = _build_real_like_slides()
        self.analysis = core.AnalysisResult(
            file="x.pptx", slide_count=18, slides=self.slides, chapters=[])
        self.induced = [
            {"title": "主题部分", "pages": [3, 4, 5, 6, 7, 8]},
            {"title": "问题与案例", "pages": [9, 10, 11, 12, 13, 14, 15]},
        ]
        # 粗排：头部正常，A 组逆序、B 组逆序
        self.base = [1, 2, 8, 7, 6, 5, 4, 3, 15, 14, 13, 12, 11, 10, 9, 16, 17, 18]

    def test_toc_impossible_then_llm_induction(self):
        """核心：规则标题无区分度 → smart+llm 自动启用 LLM 章节发现并按模型序组装。"""
        # 让 0-token 标题归组失效（全部同名），强制升级到 LLM 层
        for s in self.slides[2:15]:
            s.title = "同一标题"
        with mock.patch.object(core, "plan_reorder_by_llm", return_value=self.base), \
             mock.patch.object(core, "induce_chapters_by_llm", return_value=self.induced), \
             mock.patch.object(core, "order_chapters_by_llm", return_value=[1, 0]):
            order, trace = core.resolve_order(
                self.analysis, "技术汇报", sort_level="slide",
                use_llm=True, smart=True)
        self.assertEqual(trace["chapter_source"], "llm_induced")
        self.assertEqual(trace["induce_mode"], "single")
        self.assertEqual(trace["induced_chapter_count"], 2)
        # 章节顺序被反转为 B、A
        self.assertEqual(order[:2], [1, 2])
        self.assertEqual(order[2:9], [15, 14, 13, 12, 11, 10, 9])
        self.assertEqual(order[9:15], [8, 7, 6, 5, 4, 3])
        self.assertEqual(order[-3:], [16, 17, 18])

    def test_toc_impossible_rule_titles_without_llm(self):
        """目录不可得且未启用 LLM：0-token 标题归组仍能恢复章节，特殊页归位。"""
        order, trace = core.resolve_order(
            self.analysis, "无法匹配的目的xyz", sort_level="slide",
            use_llm=False, smart=True)
        self.assertEqual(trace["chapter_source"], "rule_titles")
        self.assertEqual(trace["induce_mode"], "rule_titles")
        self.assertEqual(order[0], 1)
        self.assertEqual(order[-1], 18)
        self.assertEqual(sorted(order), list(range(1, 19)))

    def test_toc_impossible_all_fail_uses_anchor(self):
        """标题无区分度且未启用 LLM → 特殊页锚定兜底，链路不报错不放弃。"""
        for s in self.slides[2:15]:
            s.title = "同一标题"
        order, trace = core.resolve_order(
            self.analysis, "无法匹配的目的xyz", sort_level="slide",
            use_llm=False, smart=True)
        self.assertEqual(trace["chapter_source"], "anchor_fallback")
        self.assertTrue(trace["anchor_applied"])
        self.assertEqual(order[0], 1)
        self.assertEqual(order[-1], 18)
        self.assertEqual(sorted(order), list(range(1, 19)))

    def test_refine_uses_induced_groups_map(self):
        """LLM 聚类分组下，组内序号槽位回填同样生效。"""
        self.slides[3].ordinal, self.slides[4].ordinal, self.slides[5].ordinal = 1, 2, 4
        gmap = {p: 0 for p in range(3, 9)}
        order = [1, 2, 3, 5, 4, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18]
        out, fixed = core.refine_within_groups(order, self.slides, groups_map=gmap)
        self.assertEqual(out[:8], [1, 2, 3, 4, 5, 6, 7, 8])
        self.assertEqual(fixed, 1)


class TestChapterSlotPrompt(unittest.TestCase):
    def _infos(self):
        return [
            {"title": t, "pages": list(range(1, 3)), "keywords": [], "stage": ""}
            for t in ("总体方案设计", "关键技术研究", "系统实现")]

    def test_prompt_carries_slot_hints_and_distinctions(self):
        """近义槽位：判别口径与成对辨析必须随 prompt 下发，且投资模板成果前置。"""
        prompt = core._build_chapter_order_prompt(self._infos(), "投资路演")
        self.assertIn("槽位判别口径", prompt)
        self.assertIn("设计图纸", prompt)              # solution 口径
        self.assertIn("实物已做出", prompt)            # implement 口径
        self.assertIn("标题原词优先", prompt)
        self.assertIn("方案 vs 关键技术", prompt)
        # 投资路演模板：成果在方案/关键技术之前
        self.assertLess(prompt.index("成果/价值"), prompt.index("总体方案"))

    def test_tech_template_orders_solution_before_keytech_before_implement(self):
        """技术汇报模板硬序：方案 → 关键技术 → 实现（模型无权改动）。"""
        tpl = core.PURPOSE_SLOT_TEMPLATES["技术汇报"]
        self.assertLess(tpl.index("solution"), tpl.index("keytech"))
        self.assertLess(tpl.index("keytech"), tpl.index("implement"))

    def test_guess_slot_by_title_for_nine_chapters(self):
        """9 个真值章标题全部能按原词确定性猜槽。"""
        cases = {
            "研究背景与意义": "background", "国内外研究现状": "related",
            "需求分析": "requirement", "总体方案设计": "solution",
            "关键技术研究": "keytech", "系统实现": "implement",
            "测试与验证": "test", "典型应用案例": "case",
            "研究成果与价值": "result",
        }
        for title, slot in cases.items():
            self.assertEqual(core._guess_slot_by_title(title), slot, title)
        self.assertIsNone(core._guess_slot_by_title("莫名其妙的章名"))

    def test_omitted_chapter_inserted_at_template_position(self):
        """投资模板下漏掉的“需求分析”插在关键技术之后、现状之前，而非沉尾。"""
        tpl = core.PURPOSE_SLOT_TEMPLATES["投资路演"]
        infos = [{"title": t, "pages": [1], "keywords": [], "stage": ""}
                 for t in ("成果", "案例", "实现", "方案", "关键技术",
                           "现状", "背景", "测试", "需求分析")]
        guess = core._chapter_guess_ranks(infos, tpl)
        # 模型只排了前 8 章，漏掉 8=需求分析
        order = [0, 1, 2, 3, 4, 5, 6, 7]
        llm_ranks = {i: tpl.index(s) for i, s in enumerate(
            ["result", "case", "implement", "solution", "keytech",
             "related", "background", "test"])}
        out = core._complete_order_by_template(order, llm_ranks, guess, 9, len(tpl))
        self.assertEqual(sorted(out), list(range(9)))
        self.assertEqual(out.index(8), 5)       # 需求在关键技术(4)后、现状(5原)前
        self.assertEqual(out, [0, 1, 2, 3, 4, 8, 5, 6, 7])

    def test_unguessable_omitted_chapter_appended(self):
        """规则也猜不出槽位的漏章追加末尾，不破坏模型已排顺序。"""
        tpl = core.PURPOSE_SLOT_TEMPLATES["技术汇报"]
        out = core._complete_order_by_template(
            [0, 1], {0: 2, 1: 3}, {}, 3, len(tpl))
        self.assertEqual(out, [0, 1, 2])


class TestEmptyAndEdge(unittest.TestCase):
    def test_empty_presentation(self):
        """边界：空 PPT（0 页）的章节识别。"""
        from pptx import Presentation
        with tempfile.TemporaryDirectory() as d:
            p = os.path.join(d, "empty.pptx")
            Presentation().save(p)
            r = core.analyze_pptx(p)
            self.assertEqual(r.slide_count, 0)
            self.assertEqual(r.chapters, [])

    def test_normalize_order_wrong_length(self):
        """异常：normalize_order 长度不一致。"""
        with self.assertRaises(ValueError):
            core.normalize_order([1, 2, 3], n=7)


if __name__ == "__main__":
    unittest.main(verbosity=2)
