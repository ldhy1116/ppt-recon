# PPT 智能重组工具

按自然语言描述的汇报目的，自动识别 PPT 的章节结构并重组页面顺序，输出**仍可用
PowerPoint 编辑**的 .pptx。同一份 PPT，说一句「面向招生宣讲」或「我要上课用」，
即可得到不同叙事顺序的版本，核心判据是**听众视角的理解顺畅度**。

完整链路：**自然语言 → 轻量智能体（chat.py）→ Skill 定义 → Python CLI → 核心库 → 结果**

---

## 1. 项目简介

- 读取已有 .pptx，逐页提取标题/正文/编号/特殊页等结构信号（不依赖母版配色）
- 自动切章：分标题编号（「01」「第2节」「Chapter 3」）、两级编号树（板块-章）
- 支持 12 种汇报目的：科普宣讲、学术答辩、工作总结、项目汇报、产品发布、
  投资人路演、招生综合宣讲、培训讲座、课堂教学、技术汇报、产品评审、综合汇报
- 本地小模型（Ollama qwen2.5:3b，约 2GB 显存/内存即可跑）驱动语义排序；
  **未配置模型时自动回退 0-token 关键词模式**，始终可用
- 支持删旧目录/插新目录（两级编号稿生成**分级目录**：板块加粗、章缩进）

## 2. 用户场景

| 场景 | 说法示例 |
|------|---------|
| 招生老师同一套宣讲稿 | 「按招生综合宣讲重组，问答放前面」 |
| 研究生组会/学术答辩 | 「按学术答辩的顺序排，背景方法在前」 |
| 教师拿课件上课 | 「我要课堂教学用，概念引入在前」 |
| 产品/项目汇报 | 「面向投资人路演重组，结论先行」 |
| 只想看结构 | 「分析一下这份 PPT 有哪些章节」 |
| 手动微调 | 「把第 3 页移到最前面」「按 3,1,2,4 的顺序重排」 |

## 3. 安装方法

```powershell
# 1) Python 3.10+
pip install -r requirements.txt

# 2) 可选：本地大模型（推荐，语义排序效果更好）
#    安装 Ollama：https://ollama.com
ollama pull qwen2.5:3b
ollama serve

# 3) 测试用例已随仓库提供（ppts\test1/2/3.pptx；也可把自己的 .pptx 放进 ppts\）
```

大模型配置（OpenAI 兼容接口；[model_config.py](model_config.py) 统一加载，
内置本地 Ollama 默认值，可跳过此步）：

```powershell
# 方式 A：复制 .env.example 为 .env，修改其中的值即可（.env 不入库）
# 方式 B：设置系统环境变量（优先级高于 .env）
$env:OPENAI_BASE_URL = "http://localhost:11434/v1"
$env:OPENAI_API_KEY  = "ollama"     # 本地占位值，非真实密钥
$env:OPENAI_MODEL    = "qwen2.5:3b"
# 方式 C：nanobot 等 Agent Runtime 直接注入环境变量，项目无需感知
```

配置优先级：已设环境变量 > .env 文件 > 内置默认值（Ollama）。

## 4. 运行方法

**方式一：一键脚本**（依次跑 test1/test2/test3 三个场景）

```powershell
.\run_examples.ps1
```

**方式二：自然语言入口（推荐）**

```powershell
python chat.py "把 test3.pptx 按招生综合宣讲重组，删去旧目录增加新目录，存到 test3_intro"
```

- PPT 只写文件名，程序自动在 `ppts\` 下查找
- 产物默认写入 `data\`；`存到 xxx` 指定文件名（可省略 .pptx）
- `python chat.py --parse "把 test1.pptx 按技术汇报重组"`：只解析意图不执行（调试用）
- 写盘前先打印新顺序预览；交互式 REPL 中需输入 y 确认后才写盘

**方式三：CLI 直调**

```powershell
# 只分析结构（不写盘）
python pptx_reorganize_cli.py analyze "ppts/test3.pptx"

# 重组
python pptx_reorganize_cli.py reorganize "ppts/test3.pptx" `
  --purpose "招生综合宣讲" -o "data/test3_intro.pptx" `
  --smart --use-llm --trim --yes --force
```

| 参数 | 作用 |
|------|------|
| `--purpose` | 12 种汇报目的之一（必填） |
| `-o` | 输出路径 |
| `--smart` | 智能检定级联（结构切章→标题归组→LLM 聚类→锚定兜底） |
| `--use-llm` | 启用大模型（失败自动回退关键词模式） |
| `--trim` | 清理章内冗余分隔页 |
| `--yes --force` | 跳过确认 / 覆盖已存在产物 |
| `--dry-run` | 只预览不写盘 |

## 5. 工具/Skill 使用方式

[SKILL.md](SKILL.md) 是供 AI 智能体（如 Trae）调用的技能定义，声明了：

- 触发时机（用户要求为特定演讲场景重排/重组 PPT）
- 主入口命令、12 种目的的自然语言归一规则
- 排序管线、trace 字段解读、限制说明

在支持 Skill 的智能体中直接说自然语言需求即可自动调用；无智能体时用第 4 节的
CLI/chat.py 两种方式等价运行。

## 6. 开源依赖及许可证

本项目采用 [MIT License](LICENSE)。

| 依赖 | 许可证 | 用途 |
|------|--------|------|
| [python-pptx](https://github.com/scanny/python-pptx) | MIT | PPTX 读写（不依赖 PowerPoint） |
| [jieba](https://github.com/fxsjy/jieba) | MIT | 中文分词与 TextRank 关键词（0-token 粗排/兜底） |
| [openai-python](https://github.com/openai/openai-python) | MIT | OpenAI 兼容接口调用本地/远程大模型 |
| lxml / Pillow / XlsxWriter | BSD/MIT | python-pptx 间接依赖 |

默认模型 [Qwen2.5-3B](https://huggingface.co/Qwen/Qwen2.5-3B)（Apache-2.0），
通过 [Ollama](https://ollama.com) 本地运行；模型为**可选**组件，不安装不影响
0-token 模式运行。

## 7. 测试方法

```powershell
# 单元测试（107 项，1 项按环境跳过）：切章/两级树/槽位投票/补漏/目录/锚定等
python -m unittest tests.test_reorganize -v

# 端到端验证（结构分析，不耗 token）
python pptx_reorganize_cli.py analyze "ppts/test3.pptx"
```

测试数据为程序化生成的**虚构内容**，以静态文件随仓库提交，三个用例分别
覆盖三种典型稿件形态：

| 用例 | 页数 | 形态 | 覆盖链路 |
|------|------|------|---------|
| test1.pptx | 18 | 研究报告、无分标题 | 标题归组切章 + 全页 LLM 排序 |
| test2.pptx | 57 | 课件、单层「第N节」分标题 | 结构切章 + 章投票 + 章内精排 |
| test3.pptx | 60 | 招生宣讲、两级编号 | 两级树剥离 + 分层投票 + 分级目录 |

用户真实 PPT 放入 `ppts\` 即被 .gitignore 排除。

## 8. 已知问题

1. **跨运行小幅波动**：投票用非零温度采样，同指令重跑时近义章（如「师资」与
   「培养」）可能互换；已用「板块 3 票 / 章 5 票 Borda + 降温 0.15」收敛，
   板块序在实测中保持稳定
2. **3b 模型能力上限**：模糊指令（如「正式一点」）无法识别目的，请直接说
   目的词（「答辩」「招生」）；偶发漏排由确定性邻接补漏兜底，不丢页
3. **耗时**：大稿（>30 页）端到端约 5-8 分钟（多轮投票 + 逐章精排，本地 3b）
4. **语言**：关键词与目的体系为中文，非中文 PPT 效果下降
5. **纯图片/无文本 PPT**：无法提取标题，退化为封面/总结锚定 + 原序
6. **目录页**：基于文本框重建，原 PPT 复杂的目录母版样式不会保留

---

详细设计与实验数据见 [docs/REPORT.md](docs/REPORT.md)。
