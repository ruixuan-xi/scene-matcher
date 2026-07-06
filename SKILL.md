---
name: scene-matcher
description: >
  语义向量场景匹配系统。将用户提交的 Excel 需求表与 Supabase pgvector 场景库进行语义相似度匹配，
  自动判定「已有场景 / 待确认 / 新场景」三档结果，并输出带判定结论的匹配结果 Excel。
  This skill should be used when the user asks to analyze requirements, match scenes,
  分析需求, 场景匹配, 匹配需求, process feedback, 反馈入库, generate template, or 生成模板.
agent_created: true
---

# Scene Matcher — 语义向量场景匹配

## Overview

This skill wraps the semantic scene matching pipeline that compares user-submitted requirements
(Excel) against a Supabase pgvector scene database using local sentence-transformers embeddings
and cosine similarity. The pipeline determines whether each requirement matches an existing scene,
needs human review, or represents a new scene.

## When to Use

- User says: 分析需求, 匹配场景, 跑一下匹配, 分析这个 Excel, process this requirement
- User provides an `.xlsx` path and wants matching results
- User wants to process feedback (is this scene ready to import?)
- User wants to generate the requirement collection template

## Prerequisites

Before running, verify these are available:

1. **Python venv**: `C:/Users/13251/.workbuddy/binaries/python/envs/scene_matcher/Scripts/python.exe`
2. **`.env` file** at `project/` root with:
   - `SUPABASE_URL` + `SUPABASE_SERVICE_KEY` (Supabase REST API)
   - `DEEPSEEK_API_KEY` (LLM normalization)
3. **Embedding model** at `project/.model_cache/all-MiniLM-L6-v2/` (auto-downloads on first run if missing)

See `references/setup.md` for detailed configuration.

## Core Workflows

### 1. Requirement Matching (main pipeline)

```
python <skill_dir>/project/semantic_match.py --input <input.xlsx>
```

For local development, the project is at:
`D:/Desktop/test_folder/项目/WorkBuddy_project/scene_matcher/.workbuddy/skills/scene-matcher/project/`

### 2. Feedback Processing

When the user has filled in the "是否入库" (Import?) column in the results Excel:

```
python <skill_dir>/project/semantic_match.py --input <result.xlsx> --apply-feedback
```

### 3. Generate Template

To generate a fresh requirement collection template:

```
python <skill_dir>/project/semantic_match.py --input /dev/null
```

## Judgment Rules (3-tier)

| Similarity | Label | Color | Action |
|-----------|-------|-------|--------|
| ≥ 0.85 | 已有场景 | Blue | Do not re-import |
| 0.75–0.85 | 待确认 | Yellow | Manual review needed |
| < 0.75 | 新场景 | Green | Recommend import |

## Key Design Decisions

- **Department filtering**: Only match within the same department (LLM-normalized), then cosine similarity
- **LLM normalization**: Both header mapping and department normalization use DeepSeek API (no hardcoded dictionaries)
- **Local matching**: All cosine similarity computed locally with numpy; Supabase only used for data storage via REST API
- **REST API only**: No psycopg2 direct connection (Supabase DB endpoint is IPv6-only, sandbox cannot route)

## Project Structure

```
scene-matcher/
├── SKILL.md              ← Skill 入口定义（本文件）
├── references/
│   └── setup.md          ← 环境配置说明
└── project/              ← 所有项目文件
    ├── semantic_match.py ← 主脚本
    ├── .env               ← 密钥（手动创建，从 .env.example 复制）
    ├── .env.example       ← 环境变量模板
    ├── requirements.txt
    ├── setup_pgvector.sql
    ├── .model_cache/     ← 嵌入模型缓存
    ├── 需求收集模板.xlsx  ← 需求收集模板
    ├── README.md          ← 项目说明
    ├── PRD.md             ← 需求文档
    ├── AGENT_PROMPT.md    ← Agent 提示词
    └── diagrams/           ← 架构图
```
