-- ============================================================
-- 吉林高考志愿数据库 Schema  v2.0
-- SQLite（后续可平滑升级到 PostgreSQL）
-- ============================================================

-- 1. 省份 / 地区元数据
CREATE TABLE IF NOT EXISTS province (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    name        TEXT NOT NULL UNIQUE,        -- 省份名称，如"吉林"
    city_level  TEXT,                        -- 城市水平标签
    region      TEXT                         -- 大区（东北 / 华北 / 华东…）
);

-- 2. 院校表（完整字段）
CREATE TABLE IF NOT EXISTS school (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    name        TEXT NOT NULL,               -- 院校名称
    province_id INTEGER REFERENCES province(id),
    city        TEXT,                        -- 所在城市
    type        TEXT,                        -- 类型：综合 / 理工 / 师范…
    tags        TEXT,                        -- 院校标签（985/211/双一流…）
    pub_priv    TEXT DEFAULT '公办',          -- 公私性质
    ruanke      TEXT,                        -- 软科评级 A+ / A / B+ / B / 无
    city_level  TEXT,                        -- 城市水平标签
    nat_rank    INTEGER,                    -- 全国院校排名
    school_level TEXT,                       -- 院校水平（C9联盟/部委直属/卓越工程师…）
    admin_unit  TEXT,                        -- 隶属单位（教育部/工信部/省属…）
    bao_yan     TEXT,                        -- 保研率（如"58.6%"）
    transfer    TEXT,                        -- 转专业政策
    master_cnt  INTEGER,                    -- 全校硕士专业数
    doctor_cnt  INTEGER,                    -- 全校博士专业数
    ruanke_rank INTEGER,                    -- 软科全国排名（数值）
    charter_url TEXT,                        -- 招生章程链接
    UNIQUE(name, city)
);

-- 3. 院校专业组
CREATE TABLE IF NOT EXISTS major_group (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    gcode       TEXT NOT NULL,               -- 院校专业组代码（原"院校专业组代码"）
    school_id   INTEGER REFERENCES school(id),
    year        INTEGER NOT NULL,            -- 招生年份
    ke_lei      TEXT NOT NULL,               -- 科类：物理 / 历史
    batch       TEXT NOT NULL,               -- 批次：本科批 / 提前批A段…
    subj_req    TEXT,                        -- 选科要求
    gmin_score  REAL,                        -- 专业组最低分
    gmin_rank   INTEGER,                    -- 专业组最低位次
    disc_eval   TEXT,                        -- 学科评估（如"四轮：A；五轮：A+"）
    group_plan  INTEGER,                    -- 专业组计划人数
    admit_cnt   INTEGER,                    -- 专业组录取人数
    UNIQUE(gcode, year)
);

-- 4. 专业录取明细（核心大表，一行=一个专业在某年的录取数据）
CREATE TABLE IF NOT EXISTS major_score (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    major_group_id  INTEGER REFERENCES major_group(id),
    major_name      TEXT NOT NULL,            -- 专业名称（简称）
    major_full_name TEXT,                     -- 专业全称（含备注，如"计算机科学与技术(国家专项计划)"）
    year            INTEGER NOT NULL,         -- 数据年份（2023/2024/2025）
    min_score       REAL,                     -- 最低分
    min_rank        INTEGER,                  -- 最低位次
    max_score       REAL,                     -- 最高分
    max_rank        INTEGER,                  -- 最高位次
    avg_score       REAL,                     -- 平均分（如有）
    tuition         REAL,                     -- 学费
    plan_count      INTEGER,                  -- 招生计划人数
    admit_count     INTEGER,                  -- 录取人数
    study_years     INTEGER,                  -- 学制（年）
    remark          TEXT,                     -- 专业备注
    major_gate      TEXT,                     -- 门类（工学/理学/医学…）
    major_class     TEXT,                     -- 专业类（计算机类/电子信息类…）
    major_level     TEXT,                     -- 专业水平（国一/省一/…）
    master_point    TEXT,                     -- 本专业硕士点
    doctor_point    TEXT,                     -- 本专业博士点
    is_new          TEXT                      -- 是否新增专业
);

-- 5. 本科专业目录（教育部标准 13门类→92类别→770专业）
CREATE TABLE IF NOT EXISTS major_catalog (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    level       TEXT NOT NULL DEFAULT '本科', -- 本科 / 专科
    gate        TEXT NOT NULL,               -- 门类名称
    gate_code   TEXT,                        -- 门类代码
    category    TEXT NOT NULL,               -- 类别名称
    cat_code    TEXT,                        -- 类别代码
    major_name  TEXT NOT NULL,               -- 专业名称
    major_code  TEXT,                        -- 专业代码
    UNIQUE(level, major_name)
);

-- 6. 实验班映射
CREATE TABLE IF NOT EXISTS syban_mapping (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    school_name TEXT NOT NULL,
    class_name  TEXT NOT NULL,               -- 实验班名称
    full_name   TEXT,                        -- 实验班全称
    major_name  TEXT NOT NULL,               -- 分流专业名称
    UNIQUE(school_name, class_name, major_name)
);

-- 7. 用户（预留，当前 MVP 不启用）
CREATE TABLE IF NOT EXISTS user (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    open_id     TEXT UNIQUE,
    nickname    TEXT,
    created_at  TEXT DEFAULT (datetime('now'))
);

-- 8. 用户方案历史（预留）
CREATE TABLE IF NOT EXISTS user_plan (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id     INTEGER REFERENCES user(id),
    profile     TEXT,                        -- JSON：考生信息快照
    plan_json   TEXT,                        -- JSON：完整志愿方案
    created_at  TEXT DEFAULT (datetime('now'))
);

-- 9. 聊天记录（预留）
CREATE TABLE IF NOT EXISTS chat_history (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id     INTEGER REFERENCES user(id),
    role        TEXT,                        -- user / assistant
    content     TEXT,
    created_at  TEXT DEFAULT (datetime('now'))
);

-- ============================================================
-- 索引
-- ============================================================
CREATE INDEX IF NOT EXISTS idx_mg_year_ke      ON major_group(year, ke_lei);
CREATE INDEX IF NOT EXISTS idx_mg_batch        ON major_group(batch);
CREATE INDEX IF NOT EXISTS idx_ms_group_year   ON major_score(major_group_id, year);
CREATE INDEX IF NOT EXISTS idx_ms_major_name   ON major_score(major_name);
CREATE INDEX IF NOT EXISTS idx_ms_full_name    ON major_score(major_full_name);
CREATE INDEX IF NOT EXISTS idx_school_name     ON school(name);
CREATE INDEX IF NOT EXISTS idx_mc_gate         ON major_catalog(gate);
CREATE INDEX IF NOT EXISTS idx_mc_category     ON major_catalog(category);
CREATE INDEX IF NOT EXISTS idx_syban_school    ON syban_mapping(school_name);
