"""
2026all 多省数据迁移脚本
将 data/../2026all/ 下的各省 Excel 文件导入 gaokao.db

支持格式：
  Format A（主格式）：含 院校专业组代码、生源地 列，64-76 列不等
  - .xlsx 用 openpyxl，.xls 用 xlrd

用法：
    python -m engine.migrate_2026all          # 从项目根目录
    python engine/migrate_2026all.py          # 直接运行
"""
import os, sys, sqlite3
from collections import defaultdict

if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

if getattr(sys, 'frozen', False):
    _BASE = os.path.dirname(sys.executable)
else:
    _BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

DB_PATH      = os.path.join(_BASE, 'data', 'gaokao.db')
FOLDER_2026  = os.path.join(_BASE, '2026all')
BACKUP_FOLDER = os.path.join(FOLDER_2026, '_backup')


# ── 工具函数 ──────────────────────────────────────────────

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


# ── 表头解析 ──────────────────────────────────────────────

def _find_header_row(rows_preview):
    """
    在前4行中找表头行，返回 (hdr_row_index, headers_list)。
    支持两种格式：
      Format A: 含 '生源地' + '院校专业组代码'
      Format B: 含 '省份'  + '专业组代码'（省份列即生源地）
    """
    for i, row in enumerate(rows_preview):
        names = [str(v).strip() if v else '' for v in row]
        has_syd   = '生源地' in names or '省份' in names
        has_gcode = '院校专业组代码' in names or '专业组代码' in names
        if has_syd and has_gcode:
            return i, names
    return None, None


def _build_col_map(headers):
    """
    构建列映射字典，处理重复列名（最低分/最低位次等出现多次）。
    返回:
      col       - {name: first_index}（唯一列名的快速查找）
      positions - {name: [index, ...]}（所有出现位置）
      school_start - '所在省' 的位置（学校信息区起始）
    """
    positions = defaultdict(list)
    col = {}
    for i, h in enumerate(headers):
        if not h:
            continue
        positions[h].append(i)
        if h not in col:
            col[h] = i

    # Format A 用 '所在省'，Format B 用 '院校省份'
    school_start = col.get('所在省', col.get('院校省份', len(headers)))

    # 找专业组最低分在分数区内的所有位置（用于确定年份块）
    gmin_positions  = [p for p in positions.get('专业组最低分', []) if p < school_start]
    # 专业级别各年最低分位置（排除专业组级别之前的孤立最低分）
    first_gmin = min(gmin_positions) if gmin_positions else 0
    min_score_pos = sorted([p for p in positions.get('最低分', []) if p > first_gmin and p < school_start])
    # 最低位次：兼容 '最低位次' / '最低分位次'
    rank_col_names = ['最低位次', '最低分位次']
    min_rank_pos  = sorted([p for name in rank_col_names
                             for p in positions.get(name, [])
                             if p > first_gmin and p < school_start])
    max_score_pos = sorted([p for p in positions.get('最高分', []) if p > first_gmin and p < school_start])

    return {
        'col':           col,
        'positions':     dict(positions),
        'school_start':  school_start,
        'gmin_pos':      gmin_positions,       # [2025, 2024?, ...]
        'min_score_pos': min_score_pos,        # [2025, 2024, 2023, ...]
        'min_rank_pos':  min_rank_pos,
        'max_score_pos': max_score_pos,
    }


def _row_val(row, idx, default=None):
    if idx is None or idx >= len(row):
        return default
    return row[idx]


# ── 单文件导入 ────────────────────────────────────────────

def import_file(conn, filepath, is_xls=False):
    """
    导入单个省份文件，返回 (rows_processed, scores_inserted) 或 None（格式不兼容）
    """
    fname = os.path.basename(filepath)

    # ── 读取工作表 ──
    if is_xls:
        try:
            import xlrd
        except ImportError:
            print(f"  [!] xlrd 未安装，跳过 {fname}")
            return None
        try:
            wb  = xlrd.open_workbook(filepath)
            ws  = wb.sheets()[0]
            all_rows = [
                [ws.cell_value(r, c) for c in range(ws.ncols)]
                for r in range(ws.nrows)
            ]
        except Exception as e:
            print(f"  [!] 读取失败 {fname}: {e}")
            return None
    else:
        try:
            import openpyxl
            wb  = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
            ws  = wb.active
            all_rows = [list(r) for r in ws.iter_rows(values_only=True)]
            wb.close()
        except Exception as e:
            print(f"  [!] 读取失败 {fname}: {e}")
            return None

    # ── 找表头行 ──
    hdr_idx, headers = _find_header_row(all_rows[:4])
    if hdr_idx is None:
        print(f"  [!] 未找到有效表头（需含生源地+院校专业组代码），跳过 {fname}")
        return None

    cm = _build_col_map(headers)
    col = cm['col']

    # ── 缓存 ──
    prov_cache   = {}   # school_province_name -> province.id
    school_cache = {}   # (school_name, city)  -> school.id
    group_cache  = {}   # (gcode, year, sheng_yuan_di) -> major_group.id

    # 预加载已有的 province / school / group（避免重复查询）
    cur = conn.cursor()
    cur.execute("SELECT id, name FROM province")
    for row in cur.fetchall():
        prov_cache[row[0 if isinstance(row, (list,tuple)) else 'id'] if isinstance(row,(list,tuple)) else row['name']] = None
    # 重新以 name 为 key
    cur.execute("SELECT id, name FROM province")
    prov_cache = {r[1]: r[0] for r in cur.fetchall()}
    cur.execute("SELECT id, name, city FROM school")
    school_cache = {(r[1], r[2]): r[0] for r in cur.fetchall()}
    cur.execute("SELECT id, gcode, year, sheng_yuan_di FROM major_group")
    group_cache = {(r[1], r[2], r[3]): r[0] for r in cur.fetchall()}

    rows_processed = 0
    scores_inserted = 0

    for raw_row in all_rows[hdr_idx + 1:]:
        if not any(v for v in raw_row):
            continue  # 空行
        # 检查关键字段
        # Format A: 院校专业组代码；Format B: 院校代码 + '-' + 专业组代码
        if '院校专业组代码' in col:
            gcode = _ss(_row_val(raw_row, col['院校专业组代码']))
        else:
            sch_c = _ss(_row_val(raw_row, col.get('院校代码')))
            grp_c = _ss(_row_val(raw_row, col.get('专业组代码')))
            gcode = f"{sch_c}-{grp_c}" if sch_c and grp_c else None

        school_name = _ss(_row_val(raw_row, col.get('院校名称', col.get('院校'))))
        if not gcode or not school_name:
            continue

        rows_processed += 1

        # 基本字段（兼容 Format A / Format B 列名差异）
        year_val     = _si(_row_val(raw_row, col.get('年份'))) or 2025
        sheng_yuan   = (_ss(_row_val(raw_row, col.get('生源地')))
                        or _ss(_row_val(raw_row, col.get('省份')))
                        or '吉林')
        ke_lei       = (_ss(_row_val(raw_row, col.get('科类')))
                        or _ss(_row_val(raw_row, col.get('科目')))
                        or '')
        batch        = _ss(_row_val(raw_row, col.get('批次'))) or ''
        pub_priv     = (_ss(_row_val(raw_row, col.get('公私性质')))
                        or _ss(_row_val(raw_row, col.get('办学性质')))
                        or '公办')
        major_name   = (_ss(_row_val(raw_row, col.get('专业名称')))
                        or _ss(_row_val(raw_row, col.get('专业')))
                        or '')
        major_full   = _ss(_row_val(raw_row, col.get('专业全称')))
        remark       = _ss(_row_val(raw_row, col.get('专业备注')))
        subj_req     = _ss(_row_val(raw_row, col.get('选科要求'))) or ''
        plan_count   = _si(_row_val(raw_row, col.get('计划人数')))
        study_years  = _si(_row_val(raw_row, col.get('学制')))
        tuition      = _sf(_row_val(raw_row, col.get('学费')))
        major_gate   = _ss(_row_val(raw_row, col.get('门类')))
        major_class  = _ss(_row_val(raw_row, col.get('专业类')))
        major_level  = _ss(_row_val(raw_row, col.get('专业水平')))
        master_pt    = _ss(_row_val(raw_row, col.get('本专业硕士点')))
        doctor_pt    = _ss(_row_val(raw_row, col.get('本专业博士点')))
        is_new       = _ss(_row_val(raw_row, col.get('是否新增')))
        disc_eval    = _ss(_row_val(raw_row, col.get('学科评估')))
        group_plan   = _si(_row_val(raw_row, col.get('专业组计划人数')))

        # 专业组最低分（2025）
        gmin_score = _sf(_row_val(raw_row, cm['gmin_pos'][0])) if cm['gmin_pos'] else None
        gmin_rank_positions = cm['positions'].get('专业组最低位次',
                              cm['positions'].get('专业组\\n最低位次', []))
        gmin_rank_pos = [p for p in gmin_rank_positions if p < cm['school_start']]
        gmin_rank  = _si(_row_val(raw_row, gmin_rank_pos[0])) if gmin_rank_pos else None
        admit_cnt_g = _si(_row_val(raw_row, col.get('专业组录取人数')))

        # 多年分数：0=2025, 1=2024, 2=2023
        def _yr_score(n):
            s = _sf(_row_val(raw_row, cm['min_score_pos'][n])) if n < len(cm['min_score_pos']) else None
            r = _si(_row_val(raw_row, cm['min_rank_pos'][n]))  if n < len(cm['min_rank_pos'])  else None
            return s, r
        s25, r25 = _yr_score(0)
        s24, r24 = _yr_score(1)
        s23, r23 = _yr_score(2)
        max25 = _sf(_row_val(raw_row, cm['max_score_pos'][0])) if cm['max_score_pos'] else None
        adm25 = _si(_row_val(raw_row, col.get('录取人数')))

        # 学校信息（Format A: offset-based from '所在省'；Format B: named columns）
        if '院校省份' in col:
            # Format B
            province_name = _ss(_row_val(raw_row, col.get('院校省份'))) or ''
            city          = _ss(_row_val(raw_row, col.get('院校城市'))) or ''
            city_level    = _ss(_row_val(raw_row, col.get('城市等级'))) or ''
            tags          = _ss(_row_val(raw_row, col.get('院校标签'))) or ''
            school_lv     = _ss(_row_val(raw_row, col.get('院校层级')))
            admin_unit    = _ss(_row_val(raw_row, col.get('隶属部门')))
            sch_type      = _ss(_row_val(raw_row, col.get('院校类型'))) or ''
            nat_rank      = _si(_row_val(raw_row, col.get('院校排名')))
            bao_yan       = _ss(_row_val(raw_row, col.get('保研率')))
            transfer      = _ss(_row_val(raw_row, col.get('转专业情况')))
            master_cnt    = _si(_row_val(raw_row, col.get('硕士点数量')))
            doctor_cnt    = _si(_row_val(raw_row, col.get('博士点数量')))
            charter_url   = _ss(_row_val(raw_row, col.get('招生简章')))
            ruanke        = _ss(_row_val(raw_row, col.get('软科评级')))
            ruanke_rank   = _si(_row_val(raw_row, col.get('软科排名')))
        else:
            # Format A: offset-based from '所在省'
            school_start = cm['school_start']
            province_name = _ss(_row_val(raw_row, school_start))          or ''
            city          = _ss(_row_val(raw_row, school_start + 1))      or ''
            city_level    = _ss(_row_val(raw_row, school_start + 2))      or ''
            tags          = _ss(_row_val(raw_row, school_start + 3))      or ''
            school_lv     = _ss(_row_val(raw_row, school_start + 4))
            admin_unit    = _ss(_row_val(raw_row, school_start + 6))
            sch_type      = _ss(_row_val(raw_row, school_start + 7))      or ''
            nat_rank      = _si(_row_val(raw_row, school_start + 11))
            bao_yan       = _ss(_row_val(raw_row, school_start + 10))
            transfer      = _ss(_row_val(raw_row, school_start + 12))
            master_cnt    = _si(_row_val(raw_row, school_start + 13))
            doctor_cnt    = _si(_row_val(raw_row, school_start + 15))
            charter_url   = _ss(_row_val(raw_row, school_start + 17))
            ruanke        = _ss(_row_val(raw_row, school_start + 18))
            ruanke_rank   = _si(_row_val(raw_row, school_start + 19))

        # ── 1. province ──
        if province_name and province_name not in prov_cache:
            cur.execute("INSERT OR IGNORE INTO province(name, city_level) VALUES(?,?)",
                        (province_name, city_level))
            cur.execute("SELECT id FROM province WHERE name=?", (province_name,))
            r = cur.fetchone()
            prov_cache[province_name] = r[0] if r else None
        prov_id = prov_cache.get(province_name)

        # ── 2. school ──
        sch_key = (school_name, city)
        if sch_key not in school_cache:
            cur.execute("""INSERT OR IGNORE INTO school
                           (name, province_id, city, type, tags, pub_priv, ruanke,
                            city_level, nat_rank, school_level, admin_unit, bao_yan,
                            transfer, master_cnt, doctor_cnt, ruanke_rank, charter_url)
                           VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                        (school_name, prov_id, city, sch_type, tags, pub_priv, ruanke,
                         city_level, nat_rank, school_lv, admin_unit, bao_yan,
                         transfer, master_cnt, doctor_cnt, ruanke_rank, charter_url))
            cur.execute("SELECT id FROM school WHERE name=? AND city=?", (school_name, city))
            r = cur.fetchone()
            school_cache[sch_key] = r[0] if r else None
        sch_id = school_cache.get(sch_key)
        if sch_id is None:
            continue

        # ── 3. major_group ──
        grp_key = (gcode, year_val, sheng_yuan)
        if grp_key not in group_cache:
            cur.execute("""INSERT OR IGNORE INTO major_group
                           (gcode, school_id, year, ke_lei, batch, sheng_yuan_di,
                            subj_req, gmin_score, gmin_rank, disc_eval, group_plan, admit_cnt)
                           VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
                        (gcode, sch_id, year_val, ke_lei, batch, sheng_yuan,
                         subj_req, gmin_score, gmin_rank, disc_eval, group_plan, admit_cnt_g))
            cur.execute("SELECT id FROM major_group WHERE gcode=? AND year=? AND sheng_yuan_di=?",
                        (gcode, year_val, sheng_yuan))
            r = cur.fetchone()
            group_cache[grp_key] = r[0] if r else None
        grp_id = group_cache.get(grp_key)
        if grp_id is None or not major_name:
            continue

        # ── 4. major_score ──
        def _insert_score(year, min_s, min_r, max_s=None, adm=None):
            if min_s is None:
                return
            cur.execute("""INSERT OR IGNORE INTO major_score
                           (major_group_id, major_name, major_full_name, year,
                            min_score, min_rank, max_score, tuition, plan_count,
                            admit_count, study_years, remark,
                            major_gate, major_class, major_level,
                            master_point, doctor_point, is_new)
                           VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                        (grp_id, major_name, major_full, year,
                         min_s, min_r, max_s, tuition, plan_count,
                         adm, study_years, remark,
                         major_gate, major_class, major_level,
                         master_pt, doctor_pt, is_new))

        _insert_score(2025, s25, r25, max25, adm25)
        _insert_score(2024, s24, r24)
        _insert_score(2023, s23, r23)
        scores_inserted += 1

    conn.commit()
    return rows_processed, scores_inserted


# ── 主迁移逻辑 ────────────────────────────────────────────

def ensure_sheng_yuan_column(conn):
    """确保 major_group 有 sheng_yuan_di 列（表重建迁移）"""
    cur = conn.cursor()
    cur.execute("PRAGMA table_info(major_group)")
    cols = {row[1] for row in cur.fetchall()}
    if 'sheng_yuan_di' in cols:
        return  # 已存在

    print("[...] major_group 表迁移：添加 sheng_yuan_di 列 ...")
    conn.executescript("""
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
    print("[OK] major_group 迁移完成")


def collect_files():
    """
    收集 2026all/ 下所有可导入的 Format A 文件。
    优先用主目录文件，.xls 也支持。
    """
    files = []
    if not os.path.isdir(FOLDER_2026):
        print(f"[!] 2026all 目录不存在: {FOLDER_2026}")
        return files

    for fname in sorted(os.listdir(FOLDER_2026)):
        if fname.startswith('_'):
            continue
        if fname.lower().endswith('.xlsx'):
            files.append((os.path.join(FOLDER_2026, fname), False))
        elif fname.lower().endswith('.xls'):
            files.append((os.path.join(FOLDER_2026, fname), True))
    return files


def run():
    if not os.path.exists(DB_PATH):
        print(f"[!] 数据库不存在: {DB_PATH}")
        print("    请先运行 python -m engine.migrate_to_db 初始化数据库")
        sys.exit(1)

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")

    # Step 1: 确保 sheng_yuan_di 列存在
    ensure_sheng_yuan_column(conn)

    # Step 2: 收集文件
    files = collect_files()
    print(f"\n[...] 找到 {len(files)} 个文件，开始导入...\n")

    total_rows = 0
    total_scores = 0
    ok_count = 0
    skip_count = 0

    for fpath, is_xls in files:
        fname = os.path.basename(fpath)
        print(f"  导入: {fname} ...", end=' ', flush=True)
        result = import_file(conn, fpath, is_xls=is_xls)
        if result is None:
            print("跳过")
            skip_count += 1
        else:
            rows, scores = result
            print(f"rows={rows:,}, scores={scores:,}")
            total_rows   += rows
            total_scores += scores
            ok_count += 1

    conn.close()
    print(f"\n{'='*50}")
    print(f"  导入完成: {ok_count} 个文件成功，{skip_count} 个跳过")
    print(f"  总计: {total_rows:,} 行处理，{total_scores:,} 条 major_score 记录")
    print(f"{'='*50}\n")


if __name__ == '__main__':
    run()
