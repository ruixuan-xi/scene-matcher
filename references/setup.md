# Scene Matcher 环境配置

> 所有项目文件位于 skill 的 `project/` 子目录下。

## 目录结构

```
skill root/
└── project/              ← 所有项目文件在此
    ├── semantic_match.py
    ├── .env              ← 手动创建，填入密钥
    ├── .env.example
    ├── requirements.txt
    ├── setup_pgvector.sql
    ├── .model_cache/     ← 嵌入模型缓存
    ├── 需求收集模板.xlsx
    └── ...
```

## 所需环境变量 (`.env`)

在 `project/` 目录下创建 `.env` 文件：

```env
# Supabase REST API（必填）
SUPABASE_URL=https://hbiiwjqewgbeoalpwwrv.supabase.co
SUPABASE_SERVICE_KEY=sb_secret_xxx

# DeepSeek API（必填，用于 LLM 表头映射 + 部门归一化）
DEEPSEEK_API_KEY=sk-xxx
DEEPSEEK_API_URL=https://api.deepseek.com/v1/chat/completions
```

## Python 环境

```
venv: C:/Users/13251/.workbuddy/binaries/python/envs/scene_matcher/
python: C:/Users/13251/.workbuddy/binaries/python/envs/scene_matcher/Scripts/python.exe
```

## 依赖包

```
openpyxl>=3.1.5
sentence-transformers>=3.0
numpy>=2.0
requests>=2.31
python-dotenv>=1.0
huggingface_hub>=0.20
```

安装：
```
C:/Users/13251/.workbuddy/binaries/python/envs/scene_matcher/Scripts/pip install -r <skill_dir>/project/requirements.txt
```

## 嵌入模型

all-MiniLM-L6-v2 (384 维, ~80MB)，首次运行自动下载到 `project/.model_cache/` 目录。
使用 HF 镜像加速：`HF_ENDPOINT=https://hf-mirror.com`

## Supabase 数据库

- 表: `scenes` (含 `embedding` 列，类型 `vector(384)`)
- 索引: ivfflat (已创建，但实际匹配走本地 numpy 余弦相似度)
- 数据通道: REST API only (`service_role` key)
- DB 直连不可用（IPv6 only）

## 模板文件

`project/需求收集模板.xlsx` 是需求收集模板，程序首次运行时会自动生成。
7 个必填列：部门、需求提交人、工作流程名称、具体工作步骤说明、业务痛点、涉及到的软件系统和网页、企业名称。
