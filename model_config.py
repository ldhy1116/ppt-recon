"""大模型环境配置加载器（单一数据源）。

本模块是项目中唯一负责设置模型环境变量的入口。chat.py、CLI 和核心库
都不自行硬编码模型名、端点或密钥，统一通过本模块加载配置。

配置来源（优先级从高到低）：
1. 已显式设置的环境变量（含外部 .env 文件、nanobot 注入、系统 export 等）
2. 项目根目录 .env 文件（用户私有配置，已被 .gitignore 排除）
3. 内置默认值（本地 Ollama qwen2.5:3b，零配置即可运行）

设计原则：
- 项目内部不写死任何模型/端点/密钥，只在此处集中管理
- setdefault 语义：已有环境变量不覆盖，外部注入优先
- 不配置任何模型时自动回退 0-token 关键词模式，核心功能不受影响
- nanobot 等 Agent Runtime 可通过环境变量注入任意模型，本项目无需感知
"""
from __future__ import annotations

import os
from pathlib import Path

# 内置默认值（本地 Ollama，占位值非真实密钥）
_DEFAULTS = {
    "OPENAI_BASE_URL": "http://localhost:11434/v1",
    "OPENAI_API_KEY": "ollama",
    "OPENAI_MODEL": "qwen2.5:3b",
}


def _load_dotenv_if_exists() -> None:
    """如果项目根目录有 .env 文件，读取并设置环境变量（不覆盖已有值）。

    纯手写解析，不引入 python-dotenv 依赖。仅支持 KEY=VALUE 行格式，
    忽略空行和 # 开头的注释行。
    """
    env_path = Path(__file__).resolve().parent / ".env"
    if not env_path.is_file():
        return
    for line in env_path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if "=" not in line:
            continue
        key, _, value = line.partition("=")
        key = key.strip()
        value = value.strip().strip("'\"")
        os.environ.setdefault(key, value)


def apply_model_env() -> None:
    """加载模型配置到环境变量（唯一对外接口）。

    在 chat.py 和 CLI 的入口处调用一次即可。核心库通过 os.getenv 读取。
    已显式设置的环境变量优先，.env 文件次之，内置默认值兜底。
    """
    _load_dotenv_if_exists()
    for key, default in _DEFAULTS.items():
        os.environ.setdefault(key, default)
