"""
数据访问层 v3 — SQLite 连接池 + LRU 缓存 + 结构化查询 + 方案持久化
对 planner.py 提供与原 load_raw_df() 输出格式完全兼容的 DataFrame，
同时提供专业目录查询、院校搜索、数据统计、方案存储等增强功能。
"""
import os, sys, json, sqlite3, threading
from collections import defaultdict
from datetime import datetime
from functools import lru_cache
import pandas as pd

if getattr(sys, 'frozen', False):
    _BASE = os.path.dirname(sys.executable)
else:
    _BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

DB_PATH = os.path.join(_BASE, 'data', 'gaokao.db')
# 在 git worktree 中运行时，data/ 不存在，自动回退到主仓库路径
if not os.path.exists(DB_PATH):
    _main_repo = os.path.normpath(os.path.join(_BASE, '..', '..', '..'))
    _fallback  = os.path.join(_main_repo, 'data', 'gaokao.db')
    if os.path.exists(_fallback):
        DB_PATH = _fallback

_df_cache = None


# ═══════════════════════════════════════════════════════════
# 连接池管理（线程安全，SQLite 读写分离）
# ═══════════════════════════════════════════════════════════

class _ConnPool:
    """
    轻量级 SQLite 连接池
    - 读连接：线程本地存储，每线程复用一个只读连接（SQLite WAL 模式支持并发读）
    - 写连接：全局单例 + 锁（SQLite 单写者模型）
    """
    def __init__(self, db_path: str):
        self._db_path = db_path
        self._local = threading.local()    # 线程本地读连接
        self._write_conn = None
        self._write_lock = threading.Lock()

    def get_read(self) -> sqlite3.Connection:
        """获取当前线程的只读连接（复用）"""
        conn = getattr(self._local, 'conn', None)
        if conn is None:
            conn = sqlite3.connect(f"file:{self._db_path}?mode=ro", uri=True)
            conn.row_factory = sqlite3.Row
            conn.execute("PRAGMA query_only = ON")
            self._local.conn = conn
        return conn

    def get_write(self) -> sqlite3.Connection:
        """获取写连接（全局单例，带锁）"""
        if self._write_conn is None:
            with self._write_lock:
                if self._write_conn is None:
                    self._write_conn = sqlite3.connect(self._db_path, check_same_thread=False)
                    self._write_conn.row_factory = sqlite3.Row
                    self._write_conn.execute("PRAGMA journal_mode=WAL")
                    self._write_conn.execute("PRAGMA synchronous=NORMAL")
        return self._write_conn

    @property
    def write_lock(self):
        return self._write_lock

    def close_all(self):
        """关闭所有连接（应用退出时调用）"""
        conn = getattr(self._local, 'conn', None)
        if conn:
            try: conn.close()
            except: pass
            self._local.conn = None
        if self._write_conn:
            try: self._write_conn.close()
            except: pass
            self._write_conn = None


# 全局连接池实例
_pool = _ConnPool(DB_PATH) if os.path.exists(DB_PATH) else None

# 自动迁移：为已有数据库补充新列（在函数定义后调用，见模块末尾 _migrate_user_plan()）


def _get_conn(readonly=True):
    """获取数据库连接（通过连接池）"""
    global _pool
    if _pool is None:
        if not os.path.exists(DB_PATH):
            raise FileNotFoundError(f"数据库不存在: {DB_PATH}")
        _pool = _ConnPool(DB_PATH)
    return _pool.get_read() if readonly else _pool.get_write()


def db_exists() -> bool:
    """检查数据库是否存在"""
    return os.path.exists(DB_PATH)


def _migrate_user_plan():
    """为已有数据库的 user_plan 表添加 student_province 列（幂等）"""
    if not os.path.exists(DB_PATH):
        return
    conn = _get_conn(readonly=False)
    with _pool.write_lock:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(user_plan)")
        cols = {row[1] for row in cur.fetchall()}
        if 'student_province' not in cols:
            cur.execute("ALTER TABLE user_plan ADD COLUMN student_province TEXT DEFAULT '吉林'")
            conn.commit()


def _migrate_major_group_sheng_yuan():
    """
    为已有数据库的 major_group 表添加 sheng_yuan_di 列并重建 UNIQUE 约束（幂等）。
    SQLite 不支持直接修改 UNIQUE 约束，需要重建表。
    """
    if not os.path.exists(DB_PATH):
        return
    conn = _get_conn(readonly=False)
    with _pool.write_lock:
        cur = conn.cursor()
        # 检查列是否已存在
        cur.execute("PRAGMA table_info(major_group)")
        cols = {row[1] for row in cur.fetchall()}
        if 'sheng_yuan_di' in cols:
            return  # 已迁移，跳过

        cur.executescript("""
            PRAGMA foreign_keys = OFF;

            CREATE TABLE IF NOT EXISTS major_group_new (
                id             INTEGER PRIMARY KEY AUTOINCREMENT,
                gcode          TEXT NOT NULL,
                school_id      INTEGER REFERENCES school(id),
                year           INTEGER NOT NULL,
                ke_lei         TEXT NOT NULL,
                batch          TEXT NOT NULL,
                sheng_yuan_di  TEXT NOT NULL DEFAULT '吉林',
                subj_req       TEXT,
                gmin_score     REAL,
                gmin_rank      INTEGER,
                disc_eval      TEXT,
                group_plan     INTEGER,
                admit_cnt      INTEGER,
                UNIQUE(gcode, year, sheng_yuan_di)
            );

            INSERT INTO major_group_new
                (id, gcode, school_id, year, ke_lei, batch, sheng_yuan_di,
                 subj_req, gmin_score, gmin_rank, disc_eval, group_plan, admit_cnt)
            SELECT id, gcode, school_id, year, ke_lei, batch, '吉林',
                   subj_req, gmin_score, gmin_rank, disc_eval, group_plan, admit_cnt
            FROM major_group;

            DROP TABLE major_group;
            ALTER TABLE major_group_new RENAME TO major_group;

            CREATE INDEX IF NOT EXISTS idx_mg_year_ke    ON major_group(year, ke_lei);
            CREATE INDEX IF NOT EXISTS idx_mg_batch      ON major_group(batch);
            CREATE INDEX IF NOT EXISTS idx_mg_sheng_yuan ON major_group(sheng_yuan_di);

            PRAGMA foreign_keys = ON;
        """)
        conn.commit()
        print("[db] major_group 迁移完成：已添加 sheng_yuan_di 列")


# ═══════════════════════════════════════════════════════════
# LRU 查询缓存
# ═══════════════════════════════════════════════════════════
# 对静态数据（专业目录、院校列表等）使用 lru_cache 避免重复查询
# 缓存随进程生命周期存在，数据更新后需重启或调用 clear_cache()

def clear_cache():
    """清除所有 LRU 缓存（数据更新后调用）"""
    global _df_cache
    _df_cache = None
    _cached_major_catalog.cache_clear()
    _cached_major_tree.cache_clear()
    _cached_search_majors.cache_clear()
    _cached_syban_map.cache_clear()
    _cached_stats.cache_clear()


# ═══════════════════════════════════════════════════════════
# 1. 核心数据加载（兼容 planner.py）
# ═══════════════════════════════════════════════════════════

def load_raw_df() -> pd.DataFrame:
    """
    从 SQLite 加载主数据，输出完整 DataFrame（含所有生源地）。
    包含原 Excel 所有关键列 + 新增的院校/专业附加信息列。
    planner.py 在 Python 层按 student_province 过滤。
    """
    global _df_cache
    if _df_cache is not None:
        return _df_cache

    if not os.path.exists(DB_PATH):
        raise FileNotFoundError(
            f"数据库不存在: {DB_PATH}\n"
            f"请先运行 python -m engine.migrate_to_db 完成迁移"
        )

    conn = _get_conn(readonly=True)
    sql = """
    SELECT
        mg.year           AS "年份",
        mg.sheng_yuan_di  AS "生源地",
        mg.ke_lei         AS "科类",
        mg.batch          AS "批次",
        s.pub_priv        AS "公私性质",
        s.name         AS "院校名称",
        ms25.major_name AS "专业名称",
        ms25.major_full_name AS "专业全称",
        ms25.remark    AS "专业备注",
        mg.gcode       AS "院校专业组代码",
        p.name         AS "所在省",
        s.city         AS "城市",
        s.type         AS "类型",
        s.tags         AS "院校标签",
        s.ruanke       AS "软科评级",
        s.city_level   AS "城市水平标签",
        mg.subj_req    AS "选科要求",
        mg.gmin_score  AS "专业组最低分",
        mg.gmin_rank   AS "专业组最低位次",
        ms25.tuition   AS "学费",
        ms25.min_score AS "最低分",
        ms24.min_score AS "最低分_1",
        ms23.min_score AS "最低分_2",
        ms25.min_rank  AS "最低位次",
        ms24.min_rank  AS "最低分位次",
        ms25.max_score AS "最高分",
        ms25.max_rank  AS "最高位次",
        s.nat_rank     AS "院校排名",
        mg.disc_eval   AS "学科评估",
        ms25.plan_count AS "计划人数",
        ms25.admit_count AS "录取人数",
        ms25.study_years AS "学制",
        ms25.major_gate  AS "门类",
        ms25.major_class AS "专业类",
        ms25.major_level AS "专业水平",
        ms25.master_point AS "本专业硕士点",
        ms25.doctor_point AS "本专业博士点",
        ms25.is_new      AS "是否新增",
        s.school_level   AS "院校水平",
        s.admin_unit     AS "隶属单位",
        s.bao_yan        AS "保研率",
        s.transfer       AS "转专业情况",
        s.master_cnt     AS "全校硕士专业数",
        s.doctor_cnt     AS "全校博士专业数",
        s.ruanke_rank    AS "软科排名",
        s.charter_url    AS "招生章程",
        mg.group_plan    AS "专业组计划人数"
    FROM major_score ms25
    JOIN major_group mg ON ms25.major_group_id = mg.id
    JOIN school s       ON mg.school_id = s.id
    LEFT JOIN province p ON s.province_id = p.id
    LEFT JOIN (
        SELECT major_group_id, major_name, MIN(min_score) AS min_score,
               MIN(min_rank) AS min_rank
        FROM major_score WHERE year = 2024
        GROUP BY major_group_id, major_name
    ) ms24
        ON ms24.major_group_id = mg.id
        AND ms24.major_name = ms25.major_name
    LEFT JOIN (
        SELECT major_group_id, major_name, MIN(min_score) AS min_score,
               MIN(min_rank) AS min_rank
        FROM major_score WHERE year = 2023
        GROUP BY major_group_id, major_name
    ) ms23
        ON ms23.major_group_id = mg.id
        AND ms23.major_name = ms25.major_name
    WHERE ms25.year = 2025
    """
    df = pd.read_sql_query(sql, conn)
    _df_cache = df
    return df


@lru_cache(maxsize=1)
def _cached_syban_map() -> dict:
    """LRU 缓存：实验班映射"""
    conn = _get_conn()
    cur = conn.cursor()
    cur.execute("SELECT school_name, class_name, major_name FROM syban_mapping")
    tmp = defaultdict(set)
    for school, cls, major in cur.fetchall():
        tmp[(school, cls)].add(major)
    return {k: frozenset(v) for k, v in tmp.items()}


def load_syban_map() -> dict:
    """
    从 SQLite 加载实验班映射（LRU 缓存），返回格式与 sybandb.py 完全相同：
    {(院校名称, 实验班名称): frozenset[分流专业名]}
    """
    if not os.path.exists(DB_PATH):
        return {}
    return _cached_syban_map()


# ═══════════════════════════════════════════════════════════
# 2. 专业目录查询（major_catalog）
# ═══════════════════════════════════════════════════════════

@lru_cache(maxsize=4)
def _cached_major_catalog(level: str) -> tuple:
    """LRU 缓存：专业目录（返回 tuple 以支持 lru_cache）"""
    conn = _get_conn()
    cur = conn.cursor()
    cur.execute("""SELECT gate, gate_code, category, cat_code, major_name, major_code
                   FROM major_catalog WHERE level=? ORDER BY gate_code, cat_code, major_code""",
                (level,))
    return tuple(dict(r) for r in cur.fetchall())


def get_major_catalog(level='本科') -> list[dict]:
    """
    返回专业目录完整树形列表（LRU 缓存）
    [{'gate':'工学','gate_code':'08','category':'计算机类','cat_code':'0809','major_name':'计算机科学与技术','major_code':'080901'}, ...]
    """
    return list(_cached_major_catalog(level))


@lru_cache(maxsize=4)
def _cached_major_tree(level: str) -> tuple:
    """LRU 缓存：专业目录树"""
    catalog = _cached_major_catalog(level)
    tree = []
    gate_idx = {}
    cat_idx = {}

    for row in catalog:
        gate = row['gate']
        cat = row['category']

        if gate not in gate_idx:
            gate_node = {'name': gate, 'code': row['gate_code'], 'children': []}
            tree.append(gate_node)
            gate_idx[gate] = gate_node

        gate_key = (gate, cat)
        if gate_key not in cat_idx:
            cat_node = {'name': cat, 'code': row['cat_code'], 'children': []}
            gate_idx[gate]['children'].append(cat_node)
            cat_idx[gate_key] = cat_node

        cat_idx[gate_key]['children'].append({
            'name': row['major_name'],
            'code': row['major_code'],
        })

    return tuple(tree)


def get_major_tree(level='本科') -> list[dict]:
    """
    返回专业目录树形结构（LRU 缓存，前端级联选择器可用）
    [{'name':'工学','code':'08','children':[...]}, ...]
    """
    return list(_cached_major_tree(level))


@lru_cache(maxsize=64)
def _cached_search_majors(keyword: str, level: str) -> tuple:
    """LRU 缓存：专业搜索结果"""
    conn = _get_conn()
    cur = conn.cursor()
    kw = f'%{keyword}%'
    cur.execute("""SELECT gate, category, major_name, major_code
                   FROM major_catalog
                   WHERE level=? AND (major_name LIKE ? OR category LIKE ? OR gate LIKE ?)
                   ORDER BY gate, category, major_name""",
                (level, kw, kw, kw))
    return tuple(dict(r) for r in cur.fetchall())


def search_majors(keyword: str, level='本科') -> list[dict]:
    """按关键词搜索专业（模糊匹配，LRU 缓存最近64个查询）"""
    return list(_cached_search_majors(keyword, level))


# ═══════════════════════════════════════════════════════════
# 3. 院校查询
# ═══════════════════════════════════════════════════════════

def search_schools(keyword: str = '', province: str = '', tags: str = '',
                   limit: int = 50) -> list[dict]:
    """
    搜索院校，支持名称/省份/标签过滤（连接池复用）
    返回 [{'name','city','province','type','tags','ruanke','city_level'}, ...]
    """
    conn = _get_conn()
    cur = conn.cursor()
    conditions = []
    params = []
    if keyword:
        conditions.append("s.name LIKE ?")
        params.append(f'%{keyword}%')
    if province:
        conditions.append("p.name = ?")
        params.append(province)
    if tags:
        conditions.append("s.tags LIKE ?")
        params.append(f'%{tags}%')

    # 注意：conditions 列表中每项都使用 ? 占位符，params 严格对应，无 SQL 注入风险
    where = (' WHERE ' + ' AND '.join(conditions)) if conditions else ''
    sql = f"""SELECT s.name, s.city, p.name as province, s.type, s.tags,
                     s.ruanke, s.city_level, s.pub_priv
              FROM school s
              LEFT JOIN province p ON s.province_id = p.id
              {where}
              ORDER BY s.tags DESC, s.name
              LIMIT ?"""
    cur.execute(sql, params + [min(limit, 200)])
    return [dict(r) for r in cur.fetchall()]


def get_school_majors(school_name: str, year: int = 2025, ke_lei: str = '物理') -> list[dict]:
    """
    查询某院校某年度某科类的所有专业组及专业（连接池复用）
    返回 [{'gcode','batch','subj_req','gmin_score','majors':[{'name','min_score','min_rank','tuition'}]}, ...]
    """
    conn = _get_conn()
    cur = conn.cursor()
    cur.execute("""
        SELECT mg.gcode, mg.batch, mg.subj_req, mg.gmin_score,
               ms.major_name, ms.min_score, ms.min_rank, ms.tuition
        FROM major_group mg
        JOIN school s ON mg.school_id = s.id
        JOIN major_score ms ON ms.major_group_id = mg.id AND ms.year = ?
        WHERE s.name = ? AND mg.year = ? AND mg.ke_lei = ?
        ORDER BY mg.gcode, ms.min_score DESC
    """, (year, school_name, year, ke_lei))

    groups = {}
    for r in cur.fetchall():
        gc = r['gcode']
        if gc not in groups:
            groups[gc] = {
                'gcode': gc, 'batch': r['batch'],
                'subj_req': r['subj_req'], 'gmin_score': r['gmin_score'],
                'majors': []
            }
        groups[gc]['majors'].append({
            'name': r['major_name'],
            'min_score': r['min_score'],
            'min_rank': r['min_rank'],
            'tuition': r['tuition'],
        })
    return list(groups.values())


# ═══════════════════════════════════════════════════════════
# 4. 数据统计
# ═══════════════════════════════════════════════════════════

@lru_cache(maxsize=1)
def _cached_stats() -> dict:
    """LRU 缓存：数据库统计"""
    conn = _get_conn()
    cur = conn.cursor()
    stats = {}
    _ALLOWED_TABLES = {'province', 'school', 'major_group', 'major_score',
                        'major_catalog', 'syban_mapping'}
    for table in _ALLOWED_TABLES:
        # table 来自硬编码白名单，安全使用 f-string
        cur.execute(f"SELECT COUNT(*) FROM [{table}]")
        stats[table] = cur.fetchone()[0]

    # 分科类/批次统计
    cur.execute("""SELECT ke_lei, batch, COUNT(DISTINCT gcode) as n_groups
                   FROM major_group WHERE year=2025
                   GROUP BY ke_lei, batch ORDER BY ke_lei, batch""")
    stats['groups_by_ke_batch'] = [dict(r) for r in cur.fetchall()]

    # 分数区间分布
    cur.execute("""SELECT
        CASE
            WHEN min_score >= 650 THEN '650+'
            WHEN min_score >= 600 THEN '600-649'
            WHEN min_score >= 550 THEN '550-599'
            WHEN min_score >= 500 THEN '500-549'
            WHEN min_score >= 450 THEN '450-499'
            ELSE '<450'
        END as score_range,
        COUNT(*) as cnt
        FROM major_score WHERE year=2025
        GROUP BY score_range ORDER BY score_range DESC""")
    stats['score_distribution'] = [dict(r) for r in cur.fetchall()]

    return stats


def get_stats() -> dict:
    """返回数据库基本统计（LRU 缓存）"""
    return _cached_stats()


# ═══════════════════════════════════════════════════════════
# 5. 方案持久化（user_plan 表）
# ═══════════════════════════════════════════════════════════

def save_plan(profile: dict, plan_json: dict, user_id: int = None) -> int:
    """
    保存用户方案到数据库，返回 plan_id（写连接 + 全局锁）
    """
    student_province = profile.get('student_province', '吉林')
    conn = _get_conn(readonly=False)
    with _pool.write_lock:
        cur = conn.cursor()
        cur.execute("""INSERT INTO user_plan (user_id, student_province, profile, plan_json, created_at)
                       VALUES (?, ?, ?, ?, ?)""",
                    (user_id, student_province,
                     json.dumps(profile, ensure_ascii=False),
                     json.dumps(plan_json, ensure_ascii=False),
                     datetime.now().isoformat()))
        conn.commit()
        return cur.lastrowid


def load_plans(user_id: int = None, limit: int = 10) -> list[dict]:
    """
    读取用户历史方案（最近N条，连接池复用）
    返回 [{'id','profile':dict,'plan_json':dict,'created_at':str}, ...]
    """
    conn = _get_conn()
    cur = conn.cursor()
    if user_id is not None:
        cur.execute("""SELECT id, profile, plan_json, created_at
                       FROM user_plan WHERE user_id=?
                       ORDER BY created_at DESC LIMIT ?""", (user_id, limit))
    else:
        cur.execute("""SELECT id, profile, plan_json, created_at
                       FROM user_plan
                       ORDER BY created_at DESC LIMIT ?""", (limit,))
    results = []
    for r in cur.fetchall():
        try:
            profile = json.loads(r['profile']) if r['profile'] else {}
        except (json.JSONDecodeError, TypeError):
            profile = {}
        try:
            plan_json = json.loads(r['plan_json']) if r['plan_json'] else {}
        except (json.JSONDecodeError, TypeError):
            plan_json = {}
        results.append({
            'id': r['id'],
            'profile': profile,
            'plan_json': plan_json,
            'created_at': r['created_at'],
        })
    return results


def delete_plan(plan_id: int) -> bool:
    """删除指定方案（写连接 + 全局锁）"""
    conn = _get_conn(readonly=False)
    with _pool.write_lock:
        cur = conn.cursor()
        cur.execute("DELETE FROM user_plan WHERE id=?", (plan_id,))
        conn.commit()
        return cur.rowcount > 0


# ═══════════════════════════════════════════════════════════
# 6. 聊天记录（预留）
# ═══════════════════════════════════════════════════════════

def save_chat(role: str, content: str, user_id: int = None) -> int:
    """保存一条聊天记录（写连接 + 全局锁）"""
    conn = _get_conn(readonly=False)
    with _pool.write_lock:
        cur = conn.cursor()
        cur.execute("""INSERT INTO chat_history (user_id, role, content, created_at)
                       VALUES (?, ?, ?, ?)""",
                    (user_id, role, content, datetime.now().isoformat()))
        conn.commit()
        return cur.lastrowid


def load_chat(user_id: int = None, limit: int = 50) -> list[dict]:
    """读取聊天记录（连接池复用）"""
    conn = _get_conn()
    cur = conn.cursor()
    if user_id is not None:
        cur.execute("""SELECT id, role, content, created_at
                       FROM chat_history WHERE user_id=?
                       ORDER BY created_at DESC LIMIT ?""", (user_id, limit))
    else:
        cur.execute("""SELECT id, role, content, created_at
                       FROM chat_history
                       ORDER BY created_at DESC LIMIT ?""", (limit,))
    results = [dict(r) for r in cur.fetchall()]
    results.reverse()  # 时间正序
    return results


def _migrate_major_direct():
    """创建 major_direct 表（专业+学校直填模式，幂等）"""
    if not os.path.exists(DB_PATH):
        return
    conn = _get_conn(readonly=False)
    with _pool.write_lock:
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS major_direct (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                school_name  TEXT NOT NULL,
                school_id    INTEGER REFERENCES school(id),
                major_name   TEXT NOT NULL,
                major_full   TEXT,
                year         INTEGER NOT NULL,
                ke_lei       TEXT,
                batch        TEXT,
                sheng_yuan   TEXT NOT NULL,
                min_score    REAL,
                min_rank     INTEGER,
                max_score    REAL,
                avg_score    REAL,
                plan_count   INTEGER,
                admit_count  INTEGER,
                tuition      REAL,
                study_years  INTEGER,
                subj_req     TEXT,
                remark       TEXT,
                major_class  TEXT,
                pub_priv     TEXT DEFAULT '公办',
                UNIQUE(school_name, major_name, year, sheng_yuan, ke_lei, batch)
            );
            CREATE INDEX IF NOT EXISTS idx_md_year_prov ON major_direct(year, sheng_yuan);
            CREATE INDEX IF NOT EXISTS idx_md_ke_lei    ON major_direct(ke_lei);
            CREATE INDEX IF NOT EXISTS idx_md_school    ON major_direct(school_name);
        """)
        conn.commit()


def load_direct_df(province: str) -> 'pd.DataFrame':
    """加载专业+学校直填模式数据（辽宁/重庆/山东/浙江/河北）"""
    conn = _get_conn()
    sql = """
    SELECT
        md.year        AS "年份",
        md.sheng_yuan  AS "生源地",
        md.ke_lei      AS "科类",
        md.batch       AS "批次",
        md.school_name AS "院校名称",
        md.major_name  AS "专业名称",
        md.major_full  AS "专业全称",
        md.min_score   AS "最低分",
        md.min_rank    AS "最低位次",
        md.max_score   AS "最高分",
        md.avg_score   AS "平均分",
        md.plan_count  AS "计划人数",
        md.admit_count AS "录取人数",
        md.tuition     AS "学费",
        md.study_years AS "学制",
        md.subj_req    AS "选科要求",
        md.remark      AS "专业备注",
        md.major_class AS "专业类",
        md.pub_priv    AS "公私性质",
        s.city         AS "城市",
        s.type         AS "院校类型",
        s.tags         AS "院校标签",
        s.city_level   AS "城市等级",
        s.school_level AS "院校层级",
        s.ruanke_rank  AS "软科排名",
        s.nat_rank     AS "全国排名",
        s.bao_yan      AS "保研率",
        s.master_cnt   AS "硕士点数量",
        s.doctor_cnt   AS "博士点数量"
    FROM major_direct md
    LEFT JOIN school s ON md.school_id = s.id
    WHERE md.sheng_yuan = ?
    """
    df = pd.read_sql_query(sql, conn, params=(province,))
    return df


# 模块加载时自动执行迁移（幂等，已有列时 no-op）
try:
    _migrate_user_plan()
except Exception:
    pass
try:
    _migrate_major_direct()
except Exception:
    pass
