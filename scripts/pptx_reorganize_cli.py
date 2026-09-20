"""AI PPT 结构重组 - 命令行入口（Script/CLI）。

可独立运行，无需 nanobot 或大模型。被 nanobot 等 Agent Runtime 通过 SKILL.md 调用。

子命令：
  analyze     分析 PPT 结构，输出结构化 JSON（只读，低 Token）
  reorganize  按汇报目的启发式重排，输出新 PPT
  reorder     按显式页码顺序重排，输出新 PPT
  toc         在首页前插入目录页（可选功能）

安全：
  - 写操作默认 dry-run 预览，需 --yes 才真正写盘
  - 默认禁止覆盖源文件 / 已存在文件，需 --force 覆盖
  - 路径校验、扩展名校验、运行日志

用法示例见 README.md 与 SKILL.md。
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path

# 本脚本位于 scripts/ 子目录：把项目根目录加入 sys.path，
# 才能导入根下的核心库（pptx_reorganize）与配置模块（model_config）。
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

# 大模型配置：统一由 model_config 加载（.env 文件 → 内置 Ollama 默认值）
# 外部已设置的环境变量优先；nanobot 等 Agent 可通过环境变量注入任意模型。
import model_config
model_config.apply_model_env()

import pptx_reorganize as core


def _confirm(prompt: str, assume_yes: bool) -> bool:
    if assume_yes:
        print(f"[确认-跳过] {prompt}")
        return True
    try:
        ans = input(f"{prompt} [y/N]: ").strip().lower()
    except (EOFError, KeyboardInterrupt):
        return False
    return ans in ("y", "yes")


def cmd_analyze(args) -> int:
    result = core.analyze_pptx(args.pptx)
    print(core.to_json(result))
    return 0


def cmd_reorganize(args) -> int:
    sort_level = getattr(args, "sort_level", "chapter")
    use_llm = getattr(args, "use_llm", False)
    anchor = getattr(args, "anchor_special", False)
    smart = getattr(args, "smart", False)
    model = getattr(args, "model", "")
    base_url = getattr(args, "base_url", "")
    api_key = getattr(args, "api_key", "")
    mode = "LLM" if use_llm else "关键词"
    if smart:
        mode += "+智能级联(归组→聚类→锚定)"
    else:
        if anchor:
            mode += "+特殊页锚定"
    preview = core.preview_reorganize(
        args.pptx, args.purpose, sort_level=sort_level, use_llm=use_llm,
        model=model, base_url=base_url, api_key=api_key,
        anchor_special=anchor, smart=smart,
    )
    meta = preview.get("trace") or {}
    # 冗余分隔页已在预览/写盘路径内部自动清理；此处再做一次幂等绑定，
    # 作为结构安全网（无违规时原样返回）。
    analysis = core.analyze_pptx(args.pptx)
    bound_order, hb_moves, hb_drops = core.bind_section_headings(
        preview["order"], analysis.slides)
    if hb_moves or hb_drops:
        title_by_idx = {s.index: s.title for s in analysis.slides}
        preview = dict(preview)
        preview["order"] = bound_order
        preview["preview"] = [
            {"new_position": pos + 1, "original_index": orig, "title": title_by_idx[orig]}
            for pos, orig in enumerate(bound_order)
        ]
    if smart:
        source = meta.get("chapter_source", "none")
        fixed = meta.get("groups_refined", 0)
        if source == "rule_titles":
            titles = [c["title"] for c in meta.get("induced_chapters", [])]
            wc, wok = meta.get("within_chapter_calls", 0), meta.get("within_chapter_ok", 0)
            inner = f"，章内模型精排 {wok}/{wc} 章" if wc else ""
            print(f"章节检定：页标题全局归组（①，0 token，"
                  f"{meta.get('induced_chapter_count')} 章{inner}），规则精修 {fixed} 处。")
            print(f"  发现章节：{' | '.join(titles)}")
        elif source == "llm_induced":
            titles = [c["title"] for c in meta.get("induced_chapters", [])]
            mode_txt = "分批聚类→归一" if meta.get("induce_mode") == "batched" else "单次聚类"
            wc, wok = meta.get("within_chapter_calls", 0), meta.get("within_chapter_ok", 0)
            inner = f"，章内模型精排 {wok}/{wc} 章" if wc else ""
            print(f"章节检定：大模型{mode_txt}（②，"
                  f"{meta.get('induced_chapter_count')} 章{inner}），"
                  f"规则精修 {fixed} 处。\n  发现章节：{' | '.join(titles)}")
        elif source == "anchor_fallback":
            print("标题归组与大模型章节检定均不可用 → 特殊页锚定兜底（③，未放弃排序）。")
    if hb_moves:
        print(f"分标题绑定内容 {len(hb_moves)} 页：")
        for m in hb_moves:
            print(f"  - {m}")
    if hb_drops:
        print(f"删除空分标题 {len(hb_drops)} 页：")
        for d in hb_drops:
            print(f"  - {d}")
    print(f"=== 重组预览（sort_level={sort_level}, 排序方式={mode}，dry-run，未写盘）===")
    print(json.dumps(preview, ensure_ascii=False, indent=2))
    if args.dry_run:
        print("提示：--dry-run 已设置，未生成文件。去掉 --dry_run 并加 --yes 执行写盘。")
        return 0
    if not _confirm(f"确认写入到 {args.output}？", args.yes):
        print("已取消。")
        return 1
    order, out, trace = core.reorganize_by_purpose(
        args.pptx, args.purpose, args.output, force=args.force,
        sort_level=sort_level, use_llm=use_llm,
        model=model, base_url=base_url, api_key=api_key,
        anchor_special=anchor, smart=smart,
    )
    print(f"完成：{out}（顺序 {order}）")
    explanation = core.explain_reorganization(trace, order, core.analyze_pptx(args.pptx).slides)
    if explanation:
        print(f"\n--- 变动说明 ---\n{explanation}")
    return 0


def cmd_reorder(args) -> int:
    order = [int(x.strip()) for x in args.order.split(",") if x.strip()]
    preview = core.preview_reorder(args.pptx, order)
    print("=== 重排预览（dry-run，未写盘）===")
    print(json.dumps(preview, ensure_ascii=False, indent=2))
    if args.dry_run:
        print("提示：--dry-run 已设置，未生成文件。去掉 --dry_run 并加 --yes 执行写盘。")
        return 0
    if not _confirm(f"确认写入到 {args.output}？", args.yes):
        print("已取消。")
        return 1
    out = core.reorder_slides(args.pptx, order, args.output, force=args.force)
    print(f"完成：{out}")
    return 0


def cmd_toc(args) -> int:
    src = core.validate_pptx_path(args.pptx)
    analysis = core.analyze_pptx(args.pptx)
    print("=== 将插入的目录结构 ===")
    chapters = analysis.chapters or [
        core.Chapter(title=s.title, start=s.index, end=s.index) for s in analysis.slides
    ]
    for ch in chapters:
        print(f"  {ch.title}  ……  P{ch.start}")
    if args.dry_run:
        print("提示：--dry-run 已设置，未生成文件。去掉 --dry_run 并加 --yes 执行写盘。")
        return 0
    if not _confirm(f"确认写入到 {args.output}？", args.yes):
        print("已取消。")
        return 1
    out = core.insert_toc(args.pptx, args.output, title=args.title, force=args.force)
    print(f"完成：{out}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="pptx_reorganize_cli",
        description="AI PPT 结构重组（基于 python-pptx）",
    )
    sub = p.add_subparsers(dest="command", required=True)

    pa = sub.add_parser("analyze", help="分析 PPT 结构（只读）")
    pa.add_argument("pptx", help="输入 .pptx 路径")
    pa.set_defaults(func=cmd_analyze)

    pr = sub.add_parser("reorganize", help="按汇报目的重排")
    pr.add_argument("pptx", help="输入 .pptx 路径")
    pr.add_argument("--purpose", required=True, help="汇报目的，如 '投资人汇报'/'技术分享'")
    pr.add_argument("--output", "-o", required=True, help="输出 .pptx 路径")
    pr.add_argument("--sort-level", choices=["chapter", "slide"], default="chapter",
                    help="排序粒度: chapter=按章节排序(章节内原序, 默认); slide=单页独立排序")
    pr.add_argument("--use-llm", action="store_true",
                    help="用大模型排序（低 token：Python 拆分后只传标题+关键词），失败自动回退关键词")
    pr.add_argument("--anchor-special", action="store_true",
                    help="规则锚定特殊页：封面/目录置首，总结/参考文献/致谢置尾（模型只排正文）")
    pr.add_argument("--smart", action="store_true",
                    help="智能级联：①规则标题归组→②大模型聚类检定章节（配合--use-llm，"
                         "消耗更多token）→③特殊页锚定兜底；每级失败自动继续下一级")
    pr.add_argument("--model", default="", help="大模型名称，默认读环境变量 OPENAI_MODEL")
    pr.add_argument("--base-url", default="", help="API 地址，默认读环境变量 OPENAI_BASE_URL")
    pr.add_argument("--api-key", default="", help="API Key，默认读环境变量 OPENAI_API_KEY（勿写入代码）")
    pr.add_argument("--dry-run", action="store_true", help="只预览不写盘")
    pr.add_argument("--yes", action="store_true", help="跳过交互确认")
    pr.add_argument("--force", action="store_true", help="覆盖已存在的输出文件")
    pr.set_defaults(func=cmd_reorganize)

    prd = sub.add_parser("reorder", help="按显式页码顺序重排")
    prd.add_argument("pptx", help="输入 .pptx 路径")
    prd.add_argument("--order", required=True, help='页码顺序，如 "3,1,2,4"')
    prd.add_argument("--output", "-o", required=True, help="输出 .pptx 路径")
    prd.add_argument("--dry-run", action="store_true", help="只预览不写盘")
    prd.add_argument("--yes", action="store_true", help="跳过交互确认")
    prd.add_argument("--force", action="store_true", help="覆盖已存在的输出文件")
    prd.set_defaults(func=cmd_reorder)

    pt = sub.add_parser("toc", help="插入目录页（可选功能）")
    pt.add_argument("pptx", help="输入 .pptx 路径")
    pt.add_argument("--output", "-o", required=True, help="输出 .pptx 路径")
    pt.add_argument("--title", default="目录", help="目录页标题")
    pt.add_argument("--dry-run", action="store_true", help="只预览不写盘")
    pt.add_argument("--yes", action="store_true", help="跳过交互确认")
    pt.add_argument("--force", action="store_true", help="覆盖已存在的输出文件")
    pt.set_defaults(func=cmd_toc)
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    core.setup_logging()
    try:
        return args.func(args)
    except (ValueError, FileNotFoundError, FileExistsError) as e:
        print(f"错误：{e}", file=sys.stderr)
        return 2
    except Exception as e:  # noqa: BLE001
        print(f"未知错误：{e}", file=sys.stderr)
        return 3


if __name__ == "__main__":
    raise SystemExit(main())
