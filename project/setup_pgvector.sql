-- ─────────────────────────────────────────────
-- ⚠️  此文件为遗留参考，已不再使用
--
-- 当前架构：REST API + 本地 numpy 余弦匹配
-- Supabase 仅作数据存储，不再执行向量检索
-- 如需参考，可保留本文件但无需执行
-- ─────────────────────────────────────────────

-- 在 Supabase SQL Editor 中执行一次即可

CREATE EXTENSION IF NOT EXISTS vector;

CREATE TABLE IF NOT EXISTS scenes (
    id BIGSERIAL PRIMARY KEY,
    scene_id TEXT UNIQUE NOT NULL,
    scene_title TEXT,
    process_steps TEXT,
    department TEXT,
    owner TEXT,
    platform TEXT,
    source TEXT,
    pain TEXT,
    embedding VECTOR(384)
);

CREATE INDEX IF NOT EXISTS idx_scenes_embedding ON scenes USING ivfflat (embedding vector_cosine_ops);
