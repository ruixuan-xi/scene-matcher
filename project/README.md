# 场景收集与智能去重沉淀（Supabase pgvector 版）

基于 **Supabase PostgreSQL + pgvector + all-MiniLM-L6-v2** 的 RPA 场景去重系统。输入业务需求表，自动判断每个需求是否与历史场景库重复。

## 判定规则（三档）

| 相似度 | 判定 | 颜色 | 操作 |
|---|---|---|---|
| ≥ 0.85 | 已有场景 | 蓝 | 不重复入库 |
| 0.75 ~ 0.85 | 待确认 | 黄 | 人工判断 |
| < 0.75 | 新场景 | 绿 | 建议入库 |

步骤不详细（无分步编号或每步 < 8 字）→ 退回补充（红）

## 智能表头映射

上传的 Excel 列名若与标准 7 列不完全一致，脚本不会直接拒绝，而是按 LLM 智能映射：

- 通过 DeepSeek API 自动映射相似列名（如「工作内容」→「工作流程名称」、「业务场景描述」→「具体工作步骤说明」）
- 已映射的单元格会添加批注，说明来源列名
- 无法映射的列以 **黄色高亮** 标注「⚠ 请补充」
- 生成 `{原文件名}_待补充.xlsx`，用户补全后重新上传即可继续匹配

> ⚠️ 格式不匹配时不会执行语义匹配，必须先补全模板。

## 输出列

| 序号 | 列名 | 说明 |
|---|---|---|
| 1~7 | 原始输入列 | 原样保留（通过智能映射补齐的列会标注来源批注） |
| 8 | 流程是否详细 | ✅ 详细 / 🔴 不详细 |
| 9 | Top1场景 | 最接近的场景名称 |
| 10 | Top1场景流程 | 匹配场景的步骤说明 |
| 11 | 相似度 | 0~1 余弦相似度 |
| 12 | 自动判定 | 已有场景/待确认/新场景/退回补充 |
| 13 | 建议操作 | 具体建议 |
| 14 | 是否入库 | 人工反馈列（是/否） |

## 快速开始

```bash
# 1. 安装依赖
pip install -r requirements.txt

# 2. 配置环境变量
cp .env.example .env
# 编辑 .env，填入 SUPABASE_URL + SUPABASE_SERVICE_KEY + DEEPSEEK_API_KEY

# 3. 运行匹配
python semantic_match.py --input 需求收集模板.xlsx

# 4. 对匹配结果文件填写「是否入库」后，执行反馈入库
python semantic_match.py --input 需求收集模板_匹配结果.xlsx --apply-feedback
```

### 表头不匹配时

如果上传的 Excel 列名与标准模板不一致，脚本会生成 `xxx_待补充.xlsx`：

```bash
python semantic_match.py --input 用户上传表.xlsx
# 输出：用户上传表_待补充.xlsx
```

请用户补全黄色高亮区域后重新上传，再执行匹配。

首次运行会自动下载 `all-MiniLM-L6-v2` 模型（约 80MB）到 `.model_cache` 目录，之后纯本地 CPU 运算；向量计算在本地完成，结果和场景元数据存储在 Supabase PostgreSQL。

## 技术栈

- **向量模型**：`sentence-transformers/all-MiniLM-L6-v2`（384 维，本地 CPU）
- **向量数据库**：Supabase PostgreSQL + `pgvector` 扩展（仅存储，不执行向量检索）
- **业务存储**：Supabase PostgreSQL `scenes` 表
- **数据访问**：Supabase REST API（`supabase-py` 或 `requests` + Service Role Key）
- **相似度计算**：本地 numpy 余弦相似度
- **LLM 归一化**：DeepSeek API（表头映射 + 部门归一化）
- **Excel 处理**：`openpyxl`

## 架构说明

当前版本采用 **REST API 模式**：
- 不再使用 `psycopg2` 直连数据库（Supabase DB 端点仅 IPv6，沙箱无法路由）
- 所有数据读写通过 Supabase REST API（`SUPABASE_URL` + `SUPABASE_SERVICE_KEY`）
- 向量相似度计算在本地完成（拉取全量 embeddings 后 numpy 批量计算）

## 文件说明

| 文件 | 说明 |
|---|---|
| `semantic_match.py` | 核心代码（REST API 模式） |
| `setup_pgvector.sql` | Supabase 初始化 SQL（遗留参考，已无需执行） |
| `.env.example` | 环境变量模板 |
| `.env` | 环境变量配置（手动创建，不提交到 Git） |
| `AGENT_PROMPT.md` | 智能体提示词（面向 WorkBuddy 智能体） |
| `需求收集模板.xlsx` | 需求收集模板 |
| `PRD.md` | 产品需求文档 |
