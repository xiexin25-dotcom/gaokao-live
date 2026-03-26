"""
专业+学校直填模式数据迁移脚本
将辽宁/重庆/山东/浙江/河北的 Excel 导入 major_direct 表

这5省无院校专业组代码，每行 = 1所学校的1个专业的1年录取数据。

用法：
    python -m engine.migrate_direct          # 从项目根目录
    python engine/migrate_direct.py          # 直接运行
"""
import os, sys, sqlite3
from collections import defaultdict

if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

if getattr(sys, 'frozen', False):
    _BASE = os.path.dirname(sys.executable)
else:
    _BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

DB_PATH     = os.path.join(_BASE, 'data', 'gaokao.db')
FOLDER_2026 = os.path.join(_BASE, '2026all')

# ── 省份文件配置 ──────────────────────────────────────────────
# 键 = 省份名，值 = {file, is_xls, ke_lei_col, pub_priv_col, school_col,
#                    major_col, score_cfg, school_info_col}
#
# score_cfg: list of (year, score_col_or_idx, rank_col_or_idx)
#   - 若为 str: 直接用 col[name] 取绝对列位置
#   - 若为 int: 用 min_score_pos[int] / min_rank_pos[int] 取相对位置
_PROV_FILES = {
    '河北': {
        'file': '09-河北.xls', 'is_xls': True,
        'ke_lei_col': '科目', 'pub_priv_col': '办学性质',
        'school_col': '院校', 'major_col': '专业',
        'school_info_col': '院校省份',   # Format B 命名式
        # 2025=min_score_pos[0], 2024=pos[1], 2023=pos[2]
        'score_cfg': [(2025, 0, 0), (2024, 1, 1), (2023, 2, 2)],
        'use_投档': False,
    },
    '辽宁': {
        'file': '16-辽宁.xlsx', 'is_xls': False,
        'ke_lei_col': '科目', 'pub_priv_col': '办学性质',
        'school_col': '院校', 'major_col': '专业',
        'school_info_col': '院校省份',
        # 2025=投档最低分, 2024=min_score_pos[0], 2023=pos[1]
        'score_cfg': [(2025, '投档最低分', '最低位次'), (2024, 0, 0), (2023, 1, 1)],
        'use_投档': True,
    },
    '重庆': {
        'file': '19-重庆.xlsx', 'is_xls': False,
        'ke_lei_col': '科类', 'pub_priv_col': '公私性质',
        'school_col': '院校名称', 'major_col': '专业名称',
        'school_info_col': '所在省',   # Format A 偏移式（与吉林相同）
        # 2025=pos[0], 2024=pos[2]（pos[1]是重复的2025数据）, 2023=pos[3]
        'score_cfg': [(2025, 0, 0), (2024, 2, 2), (2023, 3, 3)],
        'use_投档': False,
    },
    '山东': {
        'file': '20-山东.xls', 'is_xls': True,
        'ke_lei_col': '科目', 'pub_priv_col': '办学性质',
        'school_col': '院校', 'major_col': '专业',
        'school_info_col': '院校省份',
        # 2025=投档最低分, 2024=min_score_pos[0], 2023=pos[1]
        'score_cfg': [(2025, '投档最低分', '最低位次'), (2024, 0, 0), (2023, 1, 1)],
        'use_投档': True,
    },
    '浙江': {
        'file': '27-浙江.xls', 'is_xls': True,
        'ke_lei_col': '科目', 'pub_priv_col': '办学性质',
        'school_col': '院校', 'major_col': '专业',
        'school_info_col': '院校省份',
        # 2025=min_score_pos[0], 2024=pos[1], 2023=pos[2]
        'score_cfg': [(2025, 0, 0), (2024, 1, 1), (2023, 2, 2)],
        'use_投档': False,
    },
}

# 山东/浙江 ke_lei 映射：科目='综合' → ke_lei='综合'（不分物理/历史）
_KE_LEI_NO_SPLIT = {'山东', '浙江'}


# ── 工具函数 ──────────────────────────────────────────────────

def _sf(v):
    if v is None: return None
    try:
        f = float(v)
        return f if f == f else None
    except (ValueError, TypeError):
        return None

def _si(v):
    f = _sf(v)
    return int(f) if f is not None else None

def _ss(v):
    if v is None: return None
    s = str(v).strip()
    return s if s and s not in ('None', 'nan', '') else None

def _row_val(row, idx, default=None):
    if idx is None or idx < 0 or idx >= len(row):
        return default
    return row[idx]


# ── 列映射构建（同 migrate_2026all 相同逻辑）────────────────────

def _build_col_map(headers):
    positions = defaultdict(list)
    col = {}
    for i, h in enumerate(headers):
        if not h: continue
        positions[h].append(i)
        if h not in col:
            col[h] = i
    school_start = col.get('所在省', col.get('院校省份', len(headers)))
    gmin_positions = [p for p in positions.get('专业组最低分', []) if p < school_start]
    first_gmin = min(gmin_positions) if gmin_positions else 0
    min_score_pos = sorted([p for p in positions.get('最低分', [])
                             if p > first_gmin and p < school_start])
    rank_names = ['最低位次', '最低分位次']
    min_rank_pos = sorted([p for n in rank_names
                            for p in positions.get(n, [])
                            if p > first_gmin and p < school_start])
    max_score_pos = sorted([p for p in positions.get('最高分', [])
                             if p > first_gmin and p < school_start])
    return {
        'col': col,
        'positions': dict(positions),
        'school_start': school_start,
        'min_score_pos': min_score_pos,
        'min_rank_pos': min_rank_pos,
        'max_score_pos': max_score_pos,
    }


def _find_header_row(rows_preview):
    """找含 '年份' 和 '院校'/'院校名称' 的表头行（前4行）"""
    for i, row in enumerate(rows_preview):
        names = [str(v).strip() if v else '' for v in row]
        has_school = '院校' in names or '院校名称' in names
        has_year   = '年份' in names
        has_major  = '专业' in names or '专业名称' in names
        if has_school and has_year and has_major:
            return i, names
    return None, None


# ── 单文件导入 ────────────────────────────────────────────────

def import_province(conn, province, cfg):
    fpath = os.path.join(FOLDER_2026, cfg['file'])
    if not os.path.exists(fpath):
        print(f"  [!] 文件不存在: {cfg['file']}")
        return None

    # 读取文件
    if cfg['is_xls']:
        try:
            import xlrd
            wb = xlrd.open_workbook(fpath)
            ws = wb.sheets()[0]
            all_rows = [[ws.cell_value(r, c) for c in range(ws.ncols)]
                        for r in range(ws.nrows)]
        except Exception as e:
            print(f"  [!] 读取失败: {e}")
            return None
    else:
        try:
            import openpyxl
            wb = openpyxl.load_workbook(fpath, read_only=True, data_only=True)
            ws = wb.active
            all_rows = [list(r) for r in ws.iter_rows(values_only=True)]
            wb.close()
        except Exception as e:
            print(f"  [!] 读取失败: {e}")
            return None

    hdr_idx, headers = _find_header_row(all_rows[:4])
    if hdr_idx is None:
        print(f"  [!] 未找到表头")
        return None

    cm = _build_col_map(headers)
    col = cm['col']
    score_cfg = cfg['score_cfg']

    # 预加载缓存
    cur = conn.cursor()
    cur.execute("SELECT id, name FROM province")
    prov_cache = {r[1]: r[0] for r in cur.fetchall()}
    cur.execute("SELECT id, name, city FROM school")
    school_cache = {(r[1], r[2]): r[0] for r in cur.fetchall()}

    rows_processed = 0
    inserted = 0

    for raw_row in all_rows[hdr_idx + 1:]:
        if not any(v for v in raw_row):
            continue

        school_name = _ss(_row_val(raw_row, col.get(cfg['school_col'])))
        major_name  = _ss(_row_val(raw_row, col.get(cfg['major_col'])))
        if not school_name or not major_name:
            continue

        rows_processed += 1

        # 基本字段
        year_val   = _si(_row_val(raw_row, col.get('年份'))) or 2025
        sheng_yuan = (_ss(_row_val(raw_row, col.get('生源地')))
                      or _ss(_row_val(raw_row, col.get('省份')))
                      or province)
        ke_lei_raw = (_ss(_row_val(raw_row, col.get(cfg['ke_lei_col'])))
                      or ('综合' if province in _KE_LEI_NO_SPLIT else ''))
        ke_lei     = ke_lei_raw if ke_lei_raw else '综合'
        batch      = _ss(_row_val(raw_row, col.get('批次'))) or ''
        pub_priv   = (_ss(_row_val(raw_row, col.get(cfg['pub_priv_col'])))
                      or '公办')
        major_full = (_ss(_row_val(raw_row, col.get('专业全称')))
                      or _ss(_row_val(raw_row, col.get('专业备注'))))
        remark     = _ss(_row_val(raw_row, col.get('专业备注')))
        subj_req   = _ss(_row_val(raw_row, col.get('选科要求'))) or '不限'
        plan_cnt   = _si(_row_val(raw_row, col.get('计划人数')))
        study_yrs  = _si(_row_val(raw_row, col.get('学制')))
        tuition    = _sf(_row_val(raw_row, col.get('学费')))
        major_cls  = (_ss(_row_val(raw_row, col.get('专业类')))
                      or _ss(_row_val(raw_row, col.get('专业类别'))))

        # 学校信息
        if '院校省份' in col:
            sch_prov   = _ss(_row_val(raw_row, col.get('院校省份'))) or ''
            city       = _ss(_row_val(raw_row, col.get('院校城市'))) or ''
            city_level = _ss(_row_val(raw_row, col.get('城市等级'))) or ''
            tags       = _ss(_row_val(raw_row, col.get('院校标签'))) or ''
            school_lv  = _ss(_row_val(raw_row, col.get('院校层级')))
            admin_unit = _ss(_row_val(raw_row, col.get('隶属部门')))
            sch_type   = _ss(_row_val(raw_row, col.get('院校类型'))) or ''
            nat_rank   = _si(_row_val(raw_row, col.get('院校排名')))
            bao_yan    = _ss(_row_val(raw_row, col.get('保研率')))
            transfer   = _ss(_row_val(raw_row, col.get('转专业情况')))
            master_cnt = _si(_row_val(raw_row, col.get('硕士点数量')))
            doctor_cnt = _si(_row_val(raw_row, col.get('博士点数量')))
            charter    = _ss(_row_val(raw_row, col.get('招生简章')))
            ruanke     = _ss(_row_val(raw_row, col.get('软科评级')))
            ruanke_rk  = _si(_row_val(raw_row, col.get('软科排名')))
        else:
            ss = cm['school_start']
            sch_prov   = _ss(_row_val(raw_row, ss))     or ''
            city       = _ss(_row_val(raw_row, ss + 1)) or ''
            city_level = _ss(_row_val(raw_row, ss + 2)) or ''
            tags       = _ss(_row_val(raw_row, ss + 3)) or ''
            school_lv  = _ss(_row_val(raw_row, ss + 4))
            admin_unit = _ss(_row_val(raw_row, ss + 6))
            sch_type   = _ss(_row_val(raw_row, ss + 7)) or ''
            nat_rank   = _si(_row_val(raw_row, ss + 11))
            bao_yan    = _ss(_row_val(raw_row, ss + 10))
            transfer   = _ss(_row_val(raw_row, ss + 12))
            master_cnt = _si(_row_val(raw_row, ss + 13))
            doctor_cnt = _si(_row_val(raw_row, ss + 15))
            charter    = _ss(_row_val(raw_row, ss + 17))
            ruanke     = _ss(_row_val(raw_row, ss + 18))
            ruanke_rk  = _si(_row_val(raw_row, ss + 19))

        # 1. province
        if sch_prov and sch_prov not in prov_cache:
            cur.execute("INSERT OR IGNORE INTO province(name) VALUES(?)", (sch_prov,))
            cur.execute("SELECT id FROM province WHERE name=?", (sch_prov,))
            r = cur.fetchone()
            prov_cache[sch_prov] = r[0] if r else None
        prov_id = prov_cache.get(sch_prov)

        # 2. school
        sch_key = (school_name, city)
        if sch_key not in school_cache:
            cur.execute("""INSERT OR IGNORE INTO school
                           (name, province_id, city, type, tags, pub_priv,
                            city_level, nat_rank, school_level, admin_unit, bao_yan,
                            transfer, master_cnt, doctor_cnt, ruanke_rank, charter_url, ruanke)
                           VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                        (school_name, prov_id, city, sch_type, tags, pub_priv,
                         city_level, nat_rank, school_lv, admin_unit, bao_yan,
                         transfer, master_cnt, doctor_cnt, ruanke_rk, charter, ruanke))
            cur.execute("SELECT id FROM school WHERE name=? AND city=?", (school_name, city))
            r = cur.fetchone()
            school_cache[sch_key] = r[0] if r else None
        sch_id = school_cache.get(sch_key)

        # 3. 多年分数提取
        def _get_score(spec, score_positions, rank_positions):
            """spec 可以是 (str, str) 代表直接用列名，或 int 代表用位次索引"""
            if isinstance(spec, str):
                idx = col.get(spec)
                return _sf(_row_val(raw_row, idx)) if idx is not None else None
            else:
                return _sf(_row_val(raw_row, score_positions[spec])) \
                       if spec < len(score_positions) else None

        def _get_rank(spec, rank_positions):
            if isinstance(spec, str):
                idx = col.get(spec)
                return _si(_row_val(raw_row, idx)) if idx is not None else None
            else:
                return _si(_row_val(raw_row, rank_positions[spec])) \
                       if spec < len(rank_positions) else None

        # 4. 插入各年数据
        for (yr, s_spec, r_spec) in score_cfg:
            s = _get_score(s_spec, cm['min_score_pos'], cm['min_rank_pos'])
            r = _get_rank(r_spec, cm['min_rank_pos'])
            if s is None:
                continue
            cur.execute("""INSERT OR IGNORE INTO major_direct
                           (school_name, school_id, major_name, major_full, year,
                            ke_lei, batch, sheng_yuan, min_score, min_rank,
                            plan_count, admit_count, tuition, study_years,
                            subj_req, remark, major_class, pub_priv)
                           VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                        (school_name, sch_id, major_name, major_full, yr,
                         ke_lei, batch, sheng_yuan, s, r,
                         plan_cnt if yr == year_val else None,
                         None, tuition if yr == year_val else None,
                         study_yrs if yr == year_val else None,
                         subj_req, remark, major_cls, pub_priv))
            if yr == year_val:
                inserted += 1

    conn.commit()
    return rows_processed, inserted


# ── 主迁移逻辑 ────────────────────────────────────────────────

def run():
    if not os.path.exists(DB_PATH):
        print(f"[!] 数据库不存在: {DB_PATH}")
        sys.exit(1)

    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")

    # 确保 major_direct 表存在
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

    total_rows = 0
    total_ins  = 0
    ok = 0
    skip = 0

    for province, cfg in _PROV_FILES.items():
        print(f"  导入: {cfg['file']} ({province}) ...", end=' ', flush=True)
        result = import_province(conn, province, cfg)
        if result is None:
            print("跳过")
            skip += 1
        else:
            rows, ins = result
            print(f"rows={rows:,}, inserted={ins:,}")
            total_rows += rows
            total_ins  += ins
            ok += 1

    conn.close()
    print(f"\n{'='*50}")
    print(f"  完成: {ok} 个省份成功，{skip} 个跳过")
    print(f"  总计: {total_rows:,} 行处理，{total_ins:,} 条 major_direct 记录")
    print(f"{'='*50}\n")


if __name__ == '__main__':
    run()
