"""
高考志愿规划系统 · 本地版
python app.py → http://localhost:5000
"""
import os, sys, json, time, threading
import pandas as pd
from flask import Flask, render_template, request, jsonify, send_file, make_response

def _safe_int(val, default=0):
    """安全的 int 转换，非数字时返回默认值"""
    try:
        return int(val)
    except (TypeError, ValueError):
        return default

sys.path.insert(0, os.path.dirname(__file__))
from engine.planner import (build_plan, build_plan_direct, mc_simulate,
                             optimize_plan, export_excel, export_excel_direct,
                             KW_LIB, EXCLUDE_PRESETS, load_raw_df,
                             _DIRECT_PROV_CFG)
from engine.system_prompt import SYSTEM_PROMPT, AI_REVIEW_PROMPT, ZXF_REVIEW_PROMPT
from engine import db as gaokao_db

# ── PyInstaller 打包路径兼容 ──────────────────────────────
if getattr(sys, 'frozen', False):
    # 打包后：模板在 _MEIPASS，数据/输出在 exe 同级目录
    _BUNDLE = sys._MEIPASS
    BASE    = os.path.dirname(sys.executable)
    app = Flask(__name__, template_folder=os.path.join(_BUNDLE, 'templates'),
                static_folder=os.path.join(_BUNDLE, 'static'))
else:
    BASE = os.path.dirname(os.path.abspath(__file__))
    app = Flask(__name__)
OUT_DIR = os.path.join(BASE, 'outputs')
os.makedirs(OUT_DIR, exist_ok=True)
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # 10MB 请求体上限

# ── 当前会话方案（内存，按 session_id 隔离 + LRU 淘汰） ──
import uuid as _uuid
_SESSIONS = {}        # {session_id: {plan, mc, profile, ...}}
_SESSION_TS = {}      # {session_id: last_access_timestamp} — 淘汰用
_MAX_SESSIONS = 200   # 最多保留 200 个 session
SETTINGS = {}         # 用户设置（API Key 等）
_SESSION_LOCK = threading.RLock()
_HISTORY_LOCK = threading.Lock()
_CHAT_LAST_TIME = {}  # {session_id: timestamp} — 按用户限流
PLAN_VERSION = 'v3.5' # 每次规则变更时递增，旧SESSION自动失效

def _get_sid():
    """从请求 cookie 或 header 获取 session_id，不存在则生成（同一请求内缓存）"""
    from flask import g
    if hasattr(g, '_gaokao_sid'):
        return g._gaokao_sid
    sid = request.cookies.get('gaokao_sid') or request.headers.get('X-Session-Id')
    if not sid:
        sid = str(_uuid.uuid4())[:12]
    g._gaokao_sid = sid
    return sid

def _get_session(sid=None):
    """获取当前用户的 SESSION dict（线程安全 + LRU 淘汰）"""
    if sid is None:
        sid = _get_sid()
    with _SESSION_LOCK:
        if sid not in _SESSIONS:
            _SESSIONS[sid] = {}
            # LRU 淘汰：超过上限时删除最久未活跃的 session
            if len(_SESSIONS) > _MAX_SESSIONS:
                oldest = sorted(_SESSION_TS, key=_SESSION_TS.get)
                for old_sid in oldest[:len(_SESSIONS) - _MAX_SESSIONS]:
                    _SESSIONS.pop(old_sid, None)
                    _SESSION_TS.pop(old_sid, None)
                    _CHAT_LAST_TIME.pop(old_sid, None)
        _SESSION_TS[sid] = time.time()
        return _SESSIONS[sid]


class _SessionProxy(dict):
    """代理对象：在请求上下文中自动路由到正确的 per-user SESSION"""
    def _target(self):
        try:
            return _get_session()
        except RuntimeError:
            # 非请求上下文（启动阶段），返回空 dict
            return {}
    def __getitem__(self, k):   return self._target()[k]
    def __setitem__(self, k, v):self._target()[k] = v
    def __delitem__(self, k):   del self._target()[k]
    def __contains__(self, k):  return k in self._target()
    def __iter__(self):         return iter(self._target())
    def __len__(self):          return len(self._target())
    def get(self, k, d=None):   return self._target().get(k, d)
    def pop(self, k, *a):       return self._target().pop(k, *a)
    def setdefault(self, k, d=None): return self._target().setdefault(k, d)

SESSION = _SessionProxy()   # 全局变量保持不变，底层按 session_id 隔离

# ── 全局模板变量：API_BASE 支持子路径部署 ──────────────────
API_BASE = os.environ.get('API_BASE', '').rstrip('/')

@app.context_processor
def _inject_api_base():
    return {'API_BASE': API_BASE}

@app.after_request
def _set_session_cookie(response):
    """首次访问时设置 gaokao_sid cookie"""
    if not request.cookies.get('gaokao_sid'):
        sid = _get_sid()
        response.set_cookie('gaokao_sid', sid, max_age=86400*30, httponly=True, samesite='Lax')
    return response

# ── 历史方案持久化（双通道：SQLite 优先，JSON 文件回退）────
import json as _json

def _history_files():
    """返回按时间降序排列的历史方案文件路径列表"""
    import glob
    files = glob.glob(os.path.join(OUT_DIR, 'plan_*.json'))
    files.sort(reverse=True)
    return files

def _load_history():
    """从数据库或磁盘加载最近10条历史方案"""
    # 优先从 SQLite 加载
    if gaokao_db.db_exists():
        try:
            db_plans = gaokao_db.load_plans(limit=10)
            history = []
            for p in db_plans:
                pj = p.get('plan_json', {})
                prof = p.get('profile', {})
                history.append({
                    'ts':       p.get('created_at', '')[:16].replace('T', ' '),
                    'score':    prof.get('score', 0),
                    'ke':       prof.get('ke_lei', '物理'),
                    'tgt':      '|'.join(prof.get('target_kw', [])[:3]),
                    'n_vols':   pj.get('n_vols', 0),
                    'rush_rate': pj.get('rush_rate', 0),
                    'vols_out': pj.get('vols_out', []),
                    'mc':       pj.get('mc', {}),
                    'profile':  prof,
                    'stats':    pj.get('stats', {}),
                    'plan_version': pj.get('plan_version', PLAN_VERSION),
                    '_db_id':   p['id'],
                })
            if history:
                return history
        except Exception:
            pass
    # 回退：从 JSON 文件加载
    history = []
    for fpath in _history_files()[:10]:
        try:
            with open(fpath, 'r', encoding='utf-8') as f:
                history.append(_json.load(f))
        except Exception:
            pass
    return history

def _save_history_entry(entry):
    """将单条方案保存到 SQLite + JSON 文件双写"""
    import datetime
    # 1. 写入 SQLite
    if gaokao_db.db_exists():
        try:
            profile = entry.get('profile', {})
            plan_json = {
                'vols_out':     entry.get('vols_out', []),
                'mc':           entry.get('mc', {}),
                'stats':        entry.get('stats', {}),
                'n_vols':       entry.get('n_vols', 0),
                'rush_rate':    entry.get('rush_rate', 0),
                'plan_version': entry.get('plan_version', PLAN_VERSION),
            }
            gaokao_db.save_plan(profile, plan_json)
        except Exception as e:
            import logging
            logging.warning(f"方案保存到SQLite失败: {e}")
    # 2. 同时写 JSON 文件（兼容旧流程）
    fname = 'plan_' + datetime.datetime.now().strftime('%Y%m%d_%H%M%S') + '.json'
    fpath = os.path.join(OUT_DIR, fname)
    try:
        with open(fpath, 'w', encoding='utf-8') as f:
            _json.dump(entry, f, ensure_ascii=False, indent=2)
        files = _history_files()
        for old in files[10:]:
            try: os.remove(old)
            except Exception: pass
    except Exception as e:
        import logging
        logging.warning(f"方案保存到JSON文件失败: {e}")

PLAN_HISTORY = _load_history()   # 启动时从磁盘恢复

def validate_plan(plan_vols, score):
    """验证方案合法性，返回 (ok:bool, warnings:list)"""
    warnings = []

    # ── 平行志愿核心规则提示（信息级，不影响 plan_ok 判定） ──────
    INFO_RULE = (
        "📌【吉林平行志愿铁规】一次投档，不补充投档。"
        "若投档后因体检/单科成绩不达标被退档，本轮所有后续志愿作废，"
        "仅能参加征集志愿或下一批次。请务必确保体检、单科满足每所院校要求。"
    )

    if len(plan_vols) < 10:
        warnings.append(f"⚠️ 当前分数候选院校不足，仅生成 {len(plan_vols)} 个志愿，建议降低筛选条件")

    # 高体检/政审退档风险专业关键词
    HIGH_RISK_KWS = ['军事', '军队', '国防', '公安', '警察', '司法', '军医', '海军', '空军',
                     '武警', '飞行', '航海', '轮机', '船舶驾驶', '消防', '特警']

    # 直填模式：简化校验（无 gmin25/sc6/top6 字段）
    is_direct = any(v.get('major') for v in plan_vols[:1])
    for v in plan_vols:
        tp = v.get('tp')
        if is_direct:
            # 直填模式：检查专业名称中的高风险关键词
            major_name = v.get('major', '')
            if tp == '冲' and any(kw in major_name for kw in HIGH_RISK_KWS):
                v.setdefault('warn_tuidan', True)
                v.setdefault('warn_msg_tuidan',
                    f"⚠️ 含高退档风险专业：{major_name}。"
                    f"此类专业有体检/政审/视力等额外要求，请核实身体条件后再填。")
            continue

        gmin = float(v.get('gmin25') or 0)
        sc6 = v.get('sc6')

        # 冲志愿：退档风险专项提示
        if tp == '冲':
            risky = [m['name'] for m in v.get('top6', [])
                     if any(kw in m['name'] for kw in HIGH_RISK_KWS)]
            if risky:
                v.setdefault('warn_tuidan', True)
                v.setdefault('warn_msg_tuidan',
                    f"⚠️ 含高退档风险专业：{'、'.join(risky)}。"
                    f"此类专业有体检/政审/视力等额外要求，"
                    f"不达标将被退档且本轮志愿全废，请核实身体条件后再填。")

        if tp == '稳':
            if gmin > score:
                warnings.append(f"⚠️ 稳志愿「{v['school']}」gmin={gmin:.0f} > 考生分{score}（应为≤），可能是旧数据")
            if sc6 and sc6 > score:
                warnings.append(f"⚠️ 稳志愿「{v['school']}」sc6={sc6:.0f} > 考生分{score}，命中率极低")
        if tp == '保':
            if gmin > score:
                warnings.append(f"❌ 保志愿「{v['school']}」gmin={gmin:.0f} > 考生分{score}，不安全！")

    return (len(warnings) == 0, [INFO_RULE] + warnings)

# ── 首页 ─────────────────────────────────────────────────
def _no_cache(resp):
    """给页面响应附加禁止缓存的头，确保每次重启后浏览器取最新内容"""
    resp.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    resp.headers['Pragma'] = 'no-cache'
    resp.headers['Expires'] = '0'
    return resp

@app.route('/')
def index():
    return _no_cache(make_response(render_template('index.html')))

@app.route('/api/history')
def api_history():
    """返回历史方案列表（不含完整 vols 数据，只含摘要）"""
    summary = [{
        'idx':       i,
        'ts':        h['ts'],
        'score':     h['score'],
        'ke':        h['ke'],
        'tgt':       h['tgt'],
        'n_vols':    h['n_vols'],
        'rush_rate': h['rush_rate'],
    } for i, h in enumerate(PLAN_HISTORY)]
    return jsonify({'history': summary})

@app.route('/api/history/<int:idx>', methods=['POST'])
def api_history_restore(idx):
    """恢复某条历史方案为当前方案"""
    if idx < 0 or idx >= len(PLAN_HISTORY):
        return jsonify({'error': '历史记录不存在'}), 404
    h = PLAN_HISTORY[idx]
    if h.get('plan_version') and h.get('plan_version') != PLAN_VERSION:
        return jsonify({
            'error': f"历史方案版本({h.get('plan_version')})与当前引擎({PLAN_VERSION})不兼容，请重新生成方案"
        }), 409
    # 从 vols_out 重建 plan_vols 供 simulate/optimize 使用
    restored_vols = []
    for v in h.get('vols_out', []):
        rv = dict(v)
        # simulate/optimize 需要 top6 字段（从 intent6 还原）
        rv['top6'] = rv.get('intent6', [])
        restored_vols.append(rv)
    with _SESSION_LOCK:
        SESSION['profile']      = h['profile']
        SESSION['mc']           = h['mc']
        SESSION['plan_version'] = h.get('plan_version', PLAN_VERSION)
        SESSION['plan'] = {
            'plan_vols':   restored_vols,
            'stats':       h.get('stats', {}),
            'all_results': {},
            'profile':     h['profile'],
            'mode':        h.get('mode', 'group'),
        }
    return jsonify({
        'ok':      True,
        'vols':    h['vols_out'],
        'mc':      h['mc'],
        'profile': h['profile'],
        'stats':   h['stats'],
        'mode':    h.get('mode', 'group'),
    })

# ── 规划页 ───────────────────────────────────────────────
@app.route('/plan')
def plan_page():
    import datetime as _dt
    # 打包后模板在 _MEIPASS，开发时在 BASE/templates
    if getattr(sys, 'frozen', False):
        tpl_path = os.path.join(sys._MEIPASS, 'templates', 'plan.html')
    else:
        tpl_path = os.path.join(BASE, 'templates', 'plan.html')
    try:
        mtime = _dt.datetime.fromtimestamp(os.path.getmtime(tpl_path))
        ver_str = mtime.strftime('%Y-%m-%d %H:%M')
    except Exception:
        ver_str = _dt.datetime.now().strftime('%Y-%m-%d')
    return _no_cache(make_response(render_template('plan.html', ver_str=ver_str)))

# ── MC页 ─────────────────────────────────────────────────
@app.route('/mc')
def mc_page():
    return _no_cache(make_response(render_template('mc.html')))

@app.route('/ai-jobs')
def ai_jobs_page():
    return _no_cache(make_response(render_template('ai_jobs.html')))

# ══ API ══════════════════════════════════════════════════

@app.route('/api/keywords')
def api_keywords():
    """返回专业关键词库和排除预设"""
    return jsonify({'kw_lib': KW_LIB, 'exclude_presets': EXCLUDE_PRESETS})

@app.route('/api/batches')
def api_batches():
    """返回指定省份的可用批次列表"""
    from engine.planner import get_province_batches
    province = request.args.get('province', '吉林')
    batches = get_province_batches(province)
    return jsonify({'province': province, 'batches': batches})

@app.route('/api/batch_status')
def api_batch_status():
    """返回当前会话各批次的填报状态"""
    with _SESSION_LOCK:
        bp = SESSION.get('batch_plans', {})
        filled = {}
        for k, v in bp.items():
            vols = v.get('plan', {}).get('plan_vols', [])
            filled[k] = len(vols)
    return jsonify({'filled': filled, 'active': SESSION.get('active_batch', '')})

@app.route('/api/generate', methods=['POST'])
def api_generate():
    """
    根据考生信息生成志愿方案
    Body JSON:
      score, ke_lei, target_kw, exclude_kw,
      exclude_northeast, min_city_rank, school_pref
    """
    body = request.json or {}
    try:
        score = int(body.get('score', 0))
        if not (300 <= score <= 750):
            return jsonify({'error': '分数范围应在300~750之间'}), 400

        ke_lei = body.get('ke_lei', '物理')
        student_prov = body.get('student_province', '吉林')
        # 3+3 省份（山东/浙江）不分物理历史，科类为"综合"
        if student_prov in _DIRECT_PROV_CFG and not _DIRECT_PROV_CFG[student_prov].get('ke_split'):
            ke_lei = '综合'
        elif ke_lei not in ('物理', '历史'):
            return jsonify({'error': f"科类必须为'物理'或'历史'，收到: {ke_lei}"}), 400

        fee_max_raw = body.get('fee_max', None)
        batch_key = body.get('batch', None)          # 批次 key（来自前端标签）
        profile = {
            'score':            score,
            'ke_lei':           ke_lei,
            'student_province': body.get('student_province', '吉林'),
            'target_kw':        body.get('target_kw', []),
            'exclude_kw':       body.get('exclude_kw', []),
            'strict_exclude':   body.get('strict_exclude', False),
            'exclude_northeast':body.get('exclude_northeast', False),
            'pref_provinces':   body.get('pref_provinces', []),
            'exclude_provinces':body.get('exclude_provinces', []),
            'include_types':    body.get('include_types', []),
            'exclude_types':    body.get('exclude_types', []),
            'min_city_rank':    int(body.get('min_city_rank', 4)),
            'school_pref':      body.get('school_pref', 'school'),
            'slope':            float(body.get('slope', 150.0)),
            'fee_max':          int(fee_max_raw) if fee_max_raw else None,
            'select_subjects':  body.get('select_subjects', []),
            'batch':            batch_key,             # 批次参数传递给引擎
        }

        include_zhuanxiang = body.get('include_zhuanxiang', False)

        t0 = time.time()
        student_province = profile.get('student_province', '吉林')
        is_direct = student_province in _DIRECT_PROV_CFG
        if is_direct:
            result = build_plan_direct(profile)
        else:
            result = build_plan(profile)
        elapsed = round(time.time()-t0, 2)

        plan_vols = result['plan_vols']

        # 直填模式不走专项计划过滤逻辑（无 gcode 字段）
        if not include_zhuanxiang and not is_direct:
            zx_gcodes = {v['gcode'] for v in plan_vols if v.get('has_zhuanxiang')}
            if zx_gcodes:
                plan_vols = [v for v in plan_vols if not v.get('has_zhuanxiang')]
                # 从 all_results 候选池中补充非专项计划的志愿组
                need = 40 - len(plan_vols)
                if need > 0:
                    used_gcodes = {v['gcode'] for v in plan_vols}
                    used_schools = {v['school'] for v in plan_vols}
                    all_results = result.get('all_results', {})
                    # 按 sc6 降序从全量候选中补充
                    fill_cands = sorted(
                        [r for r in all_results.values()
                         if r['gcode'] not in used_gcodes
                         and r['gcode'] not in zx_gcodes
                         and r['school'] not in used_schools
                         and not r.get('has_zhuanxiang')],
                        key=lambda r: -(r.get('sc6') or 0)
                    )
                    for r in fill_cands[:need]:
                        sc6 = r.get('sc6')
                        gmin = float(r.get('gmin25') or 0)
                        if gmin > score:
                            r['tp'] = '冲'
                        elif sc6 is not None and score - sc6 >= 10:
                            r['tp'] = '保'
                        else:
                            r['tp'] = '稳'
                        plan_vols.append(r)
                # 按冲稳保排序，重新编号
                tp_ord = {'冲': 0, '稳': 1, '保': 2}
                plan_vols.sort(key=lambda v: (tp_ord.get(v.get('tp','稳'), 1), -(v.get('sc6') or 0)))
                for i, v in enumerate(plan_vols):
                    v['vol_idx'] = i + 1
                result['plan_vols'] = plan_vols

        stats     = result['stats']

        # 直填模式 MC 仿真暂不支持（无专业组概念），返回空占位
        if is_direct:
            mc = {'rates': [], 'total_rate': 0, 'rush_rate': 0,
                  'exp_vol': None, 'exp_q': None, 'N': 0}
        else:
            # 快速MC（N=5000，基准）—— 使用 build_plan 算好的实际位次
            mc = mc_simulate(plan_vols, N=5000, seed=42, bias_lo=0, bias_hi=0, noise_pct=3.5,
                             student_rank=stats['student_rank'], student_score=score)

        with _SESSION_LOCK:
            SESSION['plan']    = result
            SESSION['mc']      = mc
            SESSION['profile'] = profile
            # 按批次存储（支持多批次独立规划）
            _bk = batch_key or '本科批'
            SESSION.setdefault('batch_plans', {})
            SESSION['batch_plans'][_bk] = {
                'plan': result, 'mc': mc, 'profile': profile,
            }
            SESSION['active_batch'] = _bk

        def _n(x):
            """Convert NaN/inf to None for JSON safety."""
            import math
            if x is None: return None
            try:
                if math.isnan(x) or math.isinf(x): return None
            except TypeError:
                pass
            return x

        # 序列化输出（去掉大型原始字段）
        vols_out = []
        for v in plan_vols:
            if is_direct:
                vols_out.append({
                    'vol_idx':    v['vol_idx'],
                    'tp':         v['tp'],
                    'school':     v['school'],
                    'major':      v.get('major', ''),
                    'city':       v.get('city', ''),
                    'ke_lei':     v.get('ke_lei', ''),
                    'batch':      v.get('batch', ''),
                    's25':        _n(v.get('s25')),
                    'diff':       _n(v.get('diff')),
                    'tags':       v.get('tags', '') or '',
                    'city_level': v.get('city_level', '') or '',
                    'school_lv':  v.get('school_lv', '') or '',
                    'subj_req':   v.get('subj_req', '') or '',
                    'tuition':    _n(v.get('tuition')),
                    'ruanke_rank': _n(v.get('ruanke_rank')),
                    'plan_count': v.get('plan_count'),
                })
            else:
                intent6 = v.get('top6', v.get('intent',[]))[:6]
                vols_out.append({
                    'vol_idx':   v['vol_idx'],
                    'tp':        v['tp'],
                    'school':    v['school'],
                    'city':      v['city'],
                    'province':  v.get('province', ''),
                    'lv_label':  v.get('lv_label','?'),
                    'cr_label':  v.get('cr_label','?'),
                    'school_lv': v['school_lv'],
                    'city_rank': v['city_rank'],
                    'gcode':     v['gcode'],
                    'gmin25':    v.get('gmin25'),
                    'gmin24':    v.get('gmin24'),
                    'gmin23':    v.get('gmin23'),
                    'sc6':       v.get('sc6'),
                    'safe':      v.get('safe', False),
                    'n_target':  v.get('n_target', 0),
                    'n_cold':    v.get('n_cold', 0),
                    'gmin_rank': v.get('gmin_rank'),
                    'intent6':   [{'name':m['name'],'s25':m.get('s25'),
                                   'diff':m.get('diff'),'kind':m.get('kind','other'),
                                   'fee': m.get('fee'), 'r25': m.get('r25'), 'r24': m.get('r24'),
                                   'syban_majors':m.get('syban_majors',[]),
                                   'syban_all':   m.get('syban_all',[]),
                                   'zhuanxiang':  m.get('zhuanxiang',''),
                                   'full_name':   m.get('full_name','')} for m in intent6],
                    'has_zhuanxiang':    v.get('has_zhuanxiang', False),
                    'zhuanxiang_types':  v.get('zhuanxiang_types', []),
                    'all_majors_count': len(v.get('majors',[])),
                    'dedup_count':      v.get('dedup_count', 0),
                    'diaoji':           v.get('diaoji', True),
                    'warn_few_majors':  v.get('warn_few_majors', False),
                    'warn_critical':    v.get('warn_critical', False),
                    'warn_msg':         v.get('warn_msg', ''),
                    'warn_excl_major':  v.get('warn_excl_major', False),
                    'warn_msg_excl':    v.get('warn_msg_excl', ''),
                    'warn_cold_anchor':  v.get('warn_cold_anchor', False),
                    'warn_msg_cold':     v.get('warn_msg_cold', ''),
                    'warn_tuidan':       v.get('warn_tuidan', False),
                    'warn_msg_tuidan':   v.get('warn_msg_tuidan', ''),
                })

        plan_ok, plan_warnings = validate_plan(plan_vols, score)
        with _SESSION_LOCK:
            SESSION['plan_version'] = PLAN_VERSION

        # 保存到历史列表（最多10条，新的在前）
        import datetime
        plan_mode = result.get('mode', 'group')
        hist_entry = {
            'ts':       datetime.datetime.now().strftime('%m-%d %H:%M'),
            'score':    score,
            'ke':       profile.get('ke_lei','物理'),
            'tgt':      '|'.join(profile.get('target_kw',[])[:3]),
            'n_vols':   len(plan_vols),
            'rush_rate':round(mc.get('rush_rate',0)*100,1),
            'vols_out': vols_out,
            'mc':       mc,
            'profile':  profile,
            'stats':    stats,
            'mode':     plan_mode,
        }
        with _HISTORY_LOCK:
            PLAN_HISTORY.insert(0, hist_entry)
            if len(PLAN_HISTORY) > 10:
                PLAN_HISTORY.pop()
        _save_history_entry(hist_entry)

        return jsonify({
            'ok':         True,
            'elapsed':    elapsed,
            'stats':      stats,
            'vols':       vols_out,
            'mc':         mc,
            'profile':    profile,
            'plan_ok':    plan_ok,
            'warnings':   plan_warnings,
            'mode':       result.get('mode', 'group'),
            'round_note': result.get('round_note', ''),
        })

    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/simulate', methods=['POST'])
def api_simulate():
    """对当前方案做MC模拟（自定义参数）"""
    if 'plan' not in SESSION:
        return jsonify({'error': '请先生成志愿方案'}), 400
    if SESSION.get('plan', {}).get('mode') == 'direct':
        return jsonify({'error': '直填模式暂不支持MC仿真'}), 400

    body = request.json or {}
    N         = min(int(body.get('N', 10000)), 300000)
    seed      = int(body.get('seed', 42))
    # 偏移区间：正 = 悲观（分数线上升），负 = 乐观（分数线下降）
    bias_lo   = float(body.get('bias_lo', 0))
    bias_hi   = float(body.get('bias_hi', 0))
    noise_pct = float(body.get('noise_pct', body.get('noise', 3.5)))

    # 在锁内做快照，防止并发修改
    import copy as _copy
    with _SESSION_LOCK:
        plan_vols    = _copy.deepcopy(SESSION['plan']['plan_vols'])
        profile      = _copy.deepcopy(SESSION['profile'])
        student_rank = SESSION['plan']['stats']['student_rank']

    score = profile['score']

    mc = mc_simulate(plan_vols, N=N, seed=seed,
                     bias_lo=bias_lo, bias_hi=bias_hi, noise_pct=noise_pct,
                     student_rank=student_rank, student_score=score)
    with _SESSION_LOCK:
        SESSION['mc'] = mc

    # 附加各志愿最高命中专业信息
    vols_rates = []
    for i, v in enumerate(plan_vols):
        rate  = mc['rates'][i] if i < len(mc['rates']) else 0
        top_m = mc['top_majors'][i] if i < len(mc['top_majors']) else None
        mhd   = mc['major_hits'][i] if i < len(mc['major_hits']) else {}
        intent6 = v.get('intent', v.get('top6',[]))[:6]
        mj_rates = []
        for m in intent6:
            cnt  = mhd.get(m['name'],0)
            mj_rates.append({'name':m['name'],'s25':m.get('s25'),
                             'diff':m.get('diff'),'rate':round(cnt/N,4),'count':cnt})
        from engine.planner import LV_LABEL, CR_LABEL
        vols_rates.append({
            'vol_idx':  v['vol_idx'],
            'tp':       v['tp'],
            'school':   v['school'],
            'city':     v['city'],
            'lv_label': v.get('lv_label', LV_LABEL.get(v['school_lv'],'?')),
            'cr_label': v.get('cr_label', CR_LABEL.get(v['city_rank'],'?')),
            'school_lv':v['school_lv'],
            'city_rank':v['city_rank'],
            'gcode':    v.get('gcode',''),
            'gmin25':   v.get('gmin25'),
            'sc6':      v.get('sc6'),
            'rate': rate,
            'top_major': top_m, 'major_rates': mj_rates,
        })

    return jsonify({**mc, 'vols_rates': vols_rates, 'student_rank_score': score})



@app.route('/api/optimize', methods=['POST'])
def api_optimize():
    """
    多轮MC迭代优化：替换无效志愿，逼近对话版效果。
    Body: { max_rounds: 10, mc_n: 8000, noise_pct: 15 }
    """
    if 'plan' not in SESSION:
        return jsonify({'error': '请先生成志愿方案'}), 400
    with _SESSION_LOCK:
        _plan_mode = SESSION['plan'].get('mode', 'group')
    if _plan_mode == 'direct':
        return jsonify({'error': '直填模式不支持迭代优化'}), 400

    body       = request.json or {}
    max_rounds = int(body.get('max_rounds', 10))
    mc_n       = min(int(body.get('mc_n', 8000)), 50000)
    noise_pct  = float(body.get('noise_pct', 3.5))

    try:
        import copy as _copy
        with _SESSION_LOCK:
            plan_snapshot = _copy.deepcopy(SESSION['plan'])

        t0     = time.time()
        opt    = optimize_plan(
            plan_snapshot,
            max_rounds=max_rounds,
            mc_n=mc_n,
            noise_pct=noise_pct,
            seed=42,
        )
        elapsed = round(time.time() - t0, 2)

        # 用优化后的方案更新 SESSION
        with _SESSION_LOCK:
            SESSION['plan']['plan_vols'] = opt['plan_vols']
            SESSION['mc']                = opt['mc']

        # 序列化历史摘要（每轮 exp_q + 变更列表）
        history_summary = []
        for h in opt['history']:
            history_summary.append({
                'round':   h['round'],
                'label':   h['label'],
                'exp_q':   h['mc']['exp_q'],
                'rush_rate': h['mc']['rush_rate'],
                'total_rate': h['mc']['total_rate'],
                'changes': h['changes'],
            })

        # 优化后志愿序列化（同 generate 格式）
        vols_out = []
        for v in opt['plan_vols']:
            intent6 = v.get('top6', v.get('intent', []))[:6]
            vols_out.append({
                'vol_idx':   v['vol_idx'],
                'tp':        v['tp'],
                'school':    v['school'],
                'city':      v['city'],
                'province':  v.get('province', ''),
                'lv_label':  v.get('lv_label', '?'),
                'cr_label':  v.get('cr_label', '?'),
                'school_lv': v['school_lv'],
                'city_rank': v['city_rank'],
                'gcode':     v['gcode'],
                'gmin25':    v.get('gmin25'),
                'gmin24':    v.get('gmin24'),
                'gmin23':    v.get('gmin23'),
                'sc6':       v.get('sc6'),
                'safe':      v.get('safe', False),
                'gmin_rank': v.get('gmin_rank'),
                'intent6':   [{'name': m['name'], 's25': m.get('s25'),
                               'diff': m.get('diff'), 'kind': m.get('kind','other'),
                               'fee': m.get('fee'), 'r25': m.get('r25'), 'r24': m.get('r24'),
                               'syban_majors': m.get('syban_majors', []),
                               'syban_all':    m.get('syban_all', [])} for m in intent6],
                'diaoji':           v.get('diaoji', True),
                'dedup_count':      v.get('dedup_count', 0),
                'warn_few_majors':  v.get('warn_few_majors', False),
                'warn_critical':    v.get('warn_critical', False),
                'warn_msg':         v.get('warn_msg', ''),
                'warn_excl_major':  v.get('warn_excl_major', False),
                'warn_msg_excl':    v.get('warn_msg_excl', ''),
                'warn_cold_anchor':  v.get('warn_cold_anchor', False),
                'warn_msg_cold':     v.get('warn_msg_cold', ''),
                'warn_tuidan':       v.get('warn_tuidan', False),
                'warn_msg_tuidan':   v.get('warn_msg_tuidan', ''),
            })

        return jsonify({
            'ok':          True,
            'elapsed':     elapsed,
            'best_round':  opt['best_round'],
            'rounds_run':  len(opt['history']) - 1,
            'vols':        vols_out,
            'mc':          opt['mc'],
            'history':     history_summary,
        })

    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500



@app.route('/api/optimize_constrained', methods=['POST'])
def api_optimize_constrained():
    """
    带用户约束的MC优化：
    Body: {
      locked_codes:     ["院校专业组代码", ...],   // 用户锁定的志愿，不替换
      excluded_schools: ["院校名称", ...],          // 用户排除的院校，不进入候选
      max_rounds: 10, mc_n: 8000, noise_pct: 15
    }
    返回：优化后方案 + diff（哪些志愿被替换了）
    """
    if 'plan' not in SESSION:
        return jsonify({'error': '请先生成志愿方案'}), 400
    if SESSION.get('plan', {}).get('mode') == 'direct':
        return jsonify({'error': '直填模式不支持约束优化'}), 400

    body             = request.json or {}
    locked_codes     = set(body.get('locked_codes', []))
    excluded_schools = set(body.get('excluded_schools', []))
    max_rounds       = int(body.get('max_rounds', 10))
    mc_n             = min(int(body.get('mc_n', 8000)), 50000)
    noise_pct        = float(body.get('noise_pct', 3.5))

    try:
        import copy as _copy
        with _SESSION_LOCK:
            plan_snapshot = _copy.deepcopy(SESSION['plan'])
            score = SESSION['profile']['score']

        t0 = time.time()

        # 记录优化前的方案，用于 diff 对比
        before_vols = plan_snapshot['plan_vols']
        before_map  = {v['gcode']: v for v in before_vols}

        opt = optimize_plan(
            plan_snapshot,
            max_rounds=max_rounds,
            mc_n=mc_n,
            noise_pct=noise_pct,
            seed=42,
            locked_codes=locked_codes,
            excluded_schools=excluded_schools,
        )
        elapsed = round(time.time() - t0, 2)

        # 更新 SESSION
        with _SESSION_LOCK:
            SESSION['plan']['plan_vols'] = opt['plan_vols']
            SESSION['mc']                = opt['mc']

        # 计算 diff：对比 before/after，找出被替换的位置
        diff = []
        for i, (bv, av) in enumerate(zip(before_vols, opt['plan_vols'])):
            if bv['gcode'] != av['gcode']:
                diff.append({
                    'vol_idx':     i + 1,
                    'tp':          av['tp'],
                    'old_school':  bv['school'],
                    'old_gcode':   bv['gcode'],
                    'old_gmin':    bv.get('gmin25'),
                    'old_lv':      bv.get('lv_label','?'),
                    'new_school':  av['school'],
                    'new_gcode':   av['gcode'],
                    'new_gmin':    av.get('gmin25'),
                    'new_lv':      av.get('lv_label','?'),
                })

        # 序列化优化后志愿（含 is_new / is_locked 标记）
        before_gcodes = {v['gcode'] for v in before_vols}
        vols_out = []
        for v in opt['plan_vols']:
            intent6 = v.get('top6', v.get('intent', []))[:6]
            vols_out.append({
                'vol_idx':   v['vol_idx'],
                'tp':        v['tp'],
                'school':    v['school'],
                'city':      v['city'],
                'province':  v.get('province', ''),
                'lv_label':  v.get('lv_label', '?'),
                'cr_label':  v.get('cr_label', '?'),
                'school_lv': v['school_lv'],
                'city_rank': v['city_rank'],
                'gcode':     v['gcode'],
                'gmin25':    v.get('gmin25'),
                'gmin24':    v.get('gmin24'),
                'gmin23':    v.get('gmin23'),
                'sc6':       v.get('sc6'),
                'safe':      v.get('safe', False),
                'gmin_rank': v.get('gmin_rank'),
                'is_new':    v['gcode'] not in before_gcodes,       # 新加入的志愿
                'is_locked': v['gcode'] in locked_codes,            # 用户锁定的
                'intent6':   [{'name': m['name'], 's25': m.get('s25'),
                               'diff': m.get('diff'), 'kind': m.get('kind','other'),
                               'fee': m.get('fee'), 'r25': m.get('r25'), 'r24': m.get('r24'),
                               'syban_majors': m.get('syban_majors', []),
                               'syban_all':    m.get('syban_all', [])} for m in intent6],
                'diaoji':           v.get('diaoji', True),
                'all_majors_count': v.get('all_majors_count', 0),
                'n_target':         v.get('n_target', 0),
                'n_cold':           v.get('n_cold', 0),
                'dedup_count':      v.get('dedup_count', 0),
                'warn_few_majors':  v.get('warn_few_majors', False),
                'warn_critical':    v.get('warn_critical', False),
                'warn_msg':         v.get('warn_msg', ''),
                'warn_excl_major':  v.get('warn_excl_major', False),
                'warn_msg_excl':    v.get('warn_msg_excl', ''),
                'warn_cold_anchor': v.get('warn_cold_anchor', False),
                'warn_msg_cold':    v.get('warn_msg_cold', ''),
                'warn_tuidan':      v.get('warn_tuidan', False),
                'warn_msg_tuidan':  v.get('warn_msg_tuidan', ''),
            })

        from engine.planner import LV_LABEL, CR_LABEL
        return jsonify({
            'ok':          True,
            'elapsed':     elapsed,
            'best_round':  opt['best_round'],
            'rounds_run':  len(opt['history']) - 1,
            'vols':        vols_out,
            'mc':          opt['mc'],
            'diff':        diff,
            'locked_count':    len(locked_codes),
            'excluded_count':  len(excluded_schools),
        })

    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500



# ── AI咨询页 ────────────────────────────────────────────────
@app.route('/chat')
def chat_page():
    return _no_cache(make_response(render_template('chat.html')))


@app.route('/api/settings', methods=['GET', 'POST'])
def api_settings():
    """获取/保存用户设置（Gemini API Key）"""
    global SETTINGS
    if request.method == 'POST':
        body = request.json or {}
        if 'api_key' in body:
            SETTINGS['api_key'] = body['api_key'].strip()
        return jsonify({'ok': True})
    key = SETTINGS.get('api_key', '')
    return jsonify({
        'has_key':     bool(key),
        'key_preview': ('AIza...' + key[-6:]) if key else '',
    })


@app.route('/api/chat', methods=['POST'])
def api_chat():
    """
    AI咨询：Gemini API（gemini-2.0-flash）
    Body: { message: str, history: [{role, content}, ...] }
    """
    sid = _get_sid()
    now = time.time()
    if now - _CHAT_LAST_TIME.get(sid, 0) < 3.0:
        return jsonify({'error': '请求过于频繁，请稍后再试（每3秒限1次）'}), 429
    _CHAT_LAST_TIME[sid] = now

    api_key = SETTINGS.get('api_key', '')
    if not api_key:
        return jsonify({'error': '请先在设置页填写 Gemini API Key'}), 400

    body    = request.json or {}
    message = body.get('message', '').strip()
    history = body.get('history', [])[:50]  # 限制历史长度，防止内存/token耗尽

    if not message or len(message) > 5000:
        return jsonify({'error': '消息不能为空且不超过5000字'}), 400

    # 构建方案摘要注入 system prompt
    plan_json  = '（尚未生成方案）'
    mc_summary = '（尚未运行 MC 仿真）'
    if 'plan' in SESSION:
        from copy import deepcopy
        with _SESSION_LOCK:
            vols = deepcopy(SESSION['plan']['plan_vols'])
        plan_items = []
        for v in vols:
            t6 = [m['name'] for m in v.get('top6', [])[:6]]
            plan_items.append({
                'no': v['vol_idx'], 'tp': v['tp'],
                'school': v['school'], 'city': v['city'],
                'lv': v.get('lv_label','?'),
                'gmin': v.get('gmin25'), 'sc6': v.get('sc6'), 'top6': t6,
            })
        plan_json = json.dumps(plan_items, ensure_ascii=False, indent=2)
        mc_data = SESSION.get('mc')
        if mc_data:
            rates = mc_data.get('rates', [])
            tops  = mc_data.get('top_majors', [])
            mc_summary = '\n'.join(
                f"#{v['vol_idx']} {v['school']}[{v['tp']}]: "
                f"{(rates[i] if i<len(rates) else 0):.1%} "
                f"最可能专业={(tops[i]['name'] if i<len(tops) and tops[i] else '无')}"
                for i, v in enumerate(vols)
            )

    system_text = SYSTEM_PROMPT.format(plan_json=plan_json, mc_summary=mc_summary)

    # Gemini contents 格式：历史 + 本次消息
    # system 通过 system_instruction 字段传入
    contents = []
    for h in history:
        role = 'user' if h['role'] == 'user' else 'model'
        contents.append({'role': role, 'parts': [{'text': h['content']}]})
    contents.append({'role': 'user', 'parts': [{'text': message}]})

    import urllib.request, urllib.error
    payload = json.dumps({
        'system_instruction': {'parts': [{'text': system_text}]},
        'contents': contents,
        'generationConfig': {'maxOutputTokens': 1024, 'temperature': 0.7},
    }).encode('utf-8')

    url = f'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={api_key}'
    req = urllib.request.Request(url, data=payload,
                                 headers={'Content-Type': 'application/json'},
                                 method='POST')
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            result = json.loads(resp.read().decode('utf-8'))
        # 解析 Gemini 响应
        reply = result['candidates'][0]['content']['parts'][0]['text']
        return jsonify({'ok': True, 'reply': reply})
    except urllib.error.HTTPError as e:
        err_body = e.read().decode('utf-8')
        try:
            err_msg = json.loads(err_body).get('error', {}).get('message', err_body)
        except Exception:
            err_msg = err_body
        return jsonify({'error': f'Gemini API错误 {e.code}: {err_msg}'}), 502
    except Exception as e:
        return jsonify({'error': str(e)}), 500


_AI_REVIEW_DIR = os.path.join(BASE, 'data', 'ai_review')
os.makedirs(_AI_REVIEW_DIR, exist_ok=True)
_AI_REVIEW_STATE = {}   # {sid: {status, review, applied, error, ...}}
_AI_REVIEW_LOCK = threading.Lock()


def _build_plan_summary(plan_data, profile, mc_data):
    """构建方案摘要供 AI 审核用——紧凑单行格式，减少 token 消耗"""
    plan_vols = plan_data.get('plan_vols', [])
    mc_rates = mc_data.get('rates', []) if mc_data else []
    lines = []
    for i, v in enumerate(plan_vols):
        if plan_data.get('mode') == 'direct':
            lines.append(f"#{v['vol_idx']}{v['tp']} {v['school']} {v.get('major','')} s25={v.get('s25')}")
        else:
            t6 = v.get('top6', v.get('intent', []))[:6]
            majors = '/'.join(f"{m['name']}{m.get('s25','')}" for m in t6)
            rate = mc_rates[i] if i < len(mc_rates) else 0
            gmin = v.get('gmin25', '?')
            sc6 = v.get('sc6', '?')
            lines.append(f"#{v['vol_idx']}{v['tp']} {v['school']}({v.get('lv_label','?')}) "
                         f"gmin={gmin} sc6={sc6} MC={rate:.0%} top6:[{majors}]")
    plan_text = '\n'.join(lines)
    return plan_text, ''  # mc 已内嵌到每行


def _apply_corrections_to_session(actions, sid):
    """将修正指令应用到 SESSION（线程安全）"""
    applied = []
    with _SESSION_LOCK:
        sess = _get_session(sid)
        if 'plan' not in sess:
            return ['错误：SESSION 中无方案']
        plan_vols = sess['plan']['plan_vols']

        for act in actions:
            if act.get('type') == 'remove_major':
                idx = act.get('vol_idx')
                mname = act.get('major_name', '')
                for v in plan_vols:
                    if v['vol_idx'] == idx:
                        orig_len = len(v.get('top6', []))
                        v['top6'] = [m for m in v.get('top6', []) if m.get('name') != mname]
                        v['intent'] = [m for m in v.get('intent', []) if m.get('name') != mname]
                        if len(v.get('top6', [])) < orig_len:
                            applied.append(f"#{idx} 移除专业「{mname}」")
                        break

        for act in actions:
            if act.get('type') == 'reclassify':
                idx = act.get('vol_idx')
                new_tp = act.get('new_tp', '稳')
                for v in plan_vols:
                    if v['vol_idx'] == idx:
                        old_tp = v.get('tp', '?')
                        v['tp'] = new_tp
                        applied.append(f"#{idx} {v['school']} {old_tp}->{new_tp}")
                        break

        remove_idxs = {act['vol_idx'] for act in actions if act.get('type') == 'remove_vol'}
        if remove_idxs:
            for v in plan_vols:
                if v['vol_idx'] in remove_idxs:
                    applied.append(f"移除 #{v['vol_idx']} {v['school']}")
            plan_vols = [v for v in plan_vols if v['vol_idx'] not in remove_idxs]
            tp_ord = {'冲': 0, '稳': 1, '保': 2}
            plan_vols.sort(key=lambda v: (tp_ord.get(v.get('tp', '稳'), 1), -(v.get('sc6') or 0)))
            for i, v in enumerate(plan_vols):
                v['vol_idx'] = i + 1
            sess['plan']['plan_vols'] = plan_vols

    return applied


def _serialize_plan(sid):
    """序列化当前 SESSION 方案为前端格式"""
    with _SESSION_LOCK:
        sess = _get_session(sid)
        from copy import deepcopy
        plan_vols = deepcopy(sess['plan']['plan_vols'])
        profile = deepcopy(sess.get('profile', {}))
        stats = sess['plan'].get('stats', {})
        mc = deepcopy(sess.get('mc', {}))
        mode = sess['plan'].get('mode', 'group')
    is_direct = mode == 'direct'

    def _n(x):
        import math
        if x is None: return None
        try:
            if math.isnan(x) or math.isinf(x): return None
        except TypeError: pass
        return x

    vols_out = []
    for v in plan_vols:
        if is_direct:
            vols_out.append({
                'vol_idx': v['vol_idx'], 'tp': v['tp'],
                'school': v['school'], 'major': v.get('major', ''),
                'city': v.get('city', ''), 's25': _n(v.get('s25')),
                'diff': _n(v.get('diff')), 'tags': v.get('tags', '') or '',
                'city_level': v.get('city_level', '') or '',
                'school_lv': v.get('school_lv', '') or '',
            })
        else:
            intent6 = v.get('top6', v.get('intent', []))[:6]
            vols_out.append({
                'vol_idx': v['vol_idx'], 'tp': v['tp'],
                'school': v['school'], 'city': v['city'],
                'province': v.get('province', ''),
                'lv_label': v.get('lv_label', '?'),
                'cr_label': v.get('cr_label', '?'),
                'school_lv': v.get('school_lv', 6),
                'city_rank': v.get('city_rank', 4),
                'gcode': v.get('gcode', ''),
                'gmin25': v.get('gmin25'), 'gmin24': v.get('gmin24'),
                'gmin23': v.get('gmin23'), 'sc6': v.get('sc6'),
                'safe': v.get('safe', False),
                'n_target': v.get('n_target', 0), 'n_cold': v.get('n_cold', 0),
                'gmin_rank': v.get('gmin_rank'),
                'intent6': [{'name': m['name'], 's25': m.get('s25'),
                             'diff': m.get('diff'), 'kind': m.get('kind', 'other'),
                             'fee': m.get('fee'), 'r25': m.get('r25'), 'r24': m.get('r24'),
                             'syban_majors': m.get('syban_majors', []),
                             'syban_all': m.get('syban_all', []),
                             'zhuanxiang': m.get('zhuanxiang', ''),
                             'full_name': m.get('full_name', '')} for m in intent6],
                'has_zhuanxiang': v.get('has_zhuanxiang', False),
                'zhuanxiang_types': v.get('zhuanxiang_types', []),
                'all_majors_count': len(v.get('majors', [])),
                'dedup_count': v.get('dedup_count', 0),
                'diaoji': v.get('diaoji', True),
                'warn_few_majors': v.get('warn_few_majors', False),
                'warn_critical': v.get('warn_critical', False),
                'warn_msg': v.get('warn_msg', ''),
                'warn_excl_major': v.get('warn_excl_major', False),
                'warn_msg_excl': v.get('warn_msg_excl', ''),
                'warn_cold_anchor': v.get('warn_cold_anchor', False),
                'warn_msg_cold': v.get('warn_msg_cold', ''),
                'warn_tuidan': v.get('warn_tuidan', False),
                'warn_msg_tuidan': v.get('warn_msg_tuidan', ''),
            })
    return vols_out, stats, mc, profile, mode, len(plan_vols)


def _ai_review_worker(sid, prompt_text):
    """后台线程：调用 claude CLI 完成审核+修正"""
    import subprocess
    try:
        print(f"[AI修正] 开始审核 sid={sid}, prompt={len(prompt_text)}字", flush=True)
        with _AI_REVIEW_LOCK:
            _AI_REVIEW_STATE[sid] = {'status': 'running'}

        # 将 prompt 写入文件备份
        prompt_file = os.path.join(_AI_REVIEW_DIR, '_prompt.txt')
        with open(prompt_file, 'w', encoding='utf-8') as f:
            f.write(prompt_text)

        proc = subprocess.run(
            ['claude', '-p', '--output-format', 'text',
             '--model', 'haiku',
             '--tools', '',                          # 禁用内置工具，纯文本对话
             '--mcp-config', '{"mcpServers":{}}',    # 禁用 MCP servers
             '--strict-mcp-config',
             '--dangerously-skip-permissions',        # 跳过权限提示，避免阻塞
            ],
            input=prompt_text, capture_output=True, text=True, timeout=120,
        )
        print(f"[AI修正] claude 完成 code={proc.returncode} stdout={len(proc.stdout)}字 stderr={len(proc.stderr)}字", flush=True)
        if proc.returncode != 0:
            with _AI_REVIEW_LOCK:
                _AI_REVIEW_STATE[sid] = {'status': 'error',
                    'error': f'Claude Code 退出码 {proc.returncode}: {proc.stderr[:500]}'}
            return

        raw = proc.stdout.strip()
        # 去除可能的 markdown 代码块
        if raw.startswith('```'):
            lines = raw.split('\n')
            if lines[0].startswith('```'):
                lines = lines[1:]
            if lines and lines[-1].strip() == '```':
                lines = lines[:-1]
            raw = '\n'.join(lines)

        review = json.loads(raw)

        # 自动应用修正
        corrections = review.get('corrections', [])
        applied = []
        if corrections:
            applied = _apply_corrections_to_session(corrections, sid)
            # 序列化修正后的方案
            vols_out, stats, mc, profile, mode, count = _serialize_plan(sid)
        else:
            vols_out, stats, mc, profile, mode, count = None, None, None, None, None, None

        with _AI_REVIEW_LOCK:
            _AI_REVIEW_STATE[sid] = {
                'status': 'done',
                'review': {
                    'summary': review.get('summary', ''),
                    'score': review.get('score'),
                    'issues': review.get('issues', []),
                    'suggestions': review.get('suggestions', []),
                },
                'applied': applied,
                'correction_note': review.get('correction_note', ''),
                'vols': vols_out,
                'stats': stats,
                'mc': mc,
                'profile': profile,
                'mode': mode,
                'count': count,
            }
    except json.JSONDecodeError as e:
        print(f"[AI修正] JSON解析失败: {e}, raw前200字: {raw[:200] if raw else 'empty'}", flush=True)
        with _AI_REVIEW_LOCK:
            _AI_REVIEW_STATE[sid] = {
                'status': 'done',
                'review': {'summary': raw[:3000] if raw else str(e), 'score': None,
                           'issues': [], 'suggestions': [], 'raw': True},
                'applied': [],
            }
    except subprocess.TimeoutExpired:
        print(f"[AI修正] 超时 sid={sid}", flush=True)
        with _AI_REVIEW_LOCK:
            _AI_REVIEW_STATE[sid] = {'status': 'error', 'error': 'Claude Code 审核超时（120秒）'}
    except Exception as e:
        print(f"[AI修正] 异常 sid={sid}: {e}", flush=True)
        import traceback; traceback.print_exc()
        with _AI_REVIEW_LOCK:
            _AI_REVIEW_STATE[sid] = {'status': 'error', 'error': str(e)}


@app.route('/api/ai_review', methods=['POST'])
def api_ai_review():
    """AI修正（全自动）：后台调用 claude CLI 审核 + 自动应用修正"""
    sid = _get_sid()
    if 'plan' not in SESSION:
        return jsonify({'error': '请先生成志愿方案'}), 400

    # 防重复提交
    with _AI_REVIEW_LOCK:
        cur = _AI_REVIEW_STATE.get(sid, {})
        if cur.get('status') == 'running':
            return jsonify({'ok': True, 'message': '审核进行中…'})

    from copy import deepcopy
    with _SESSION_LOCK:
        plan_data = deepcopy(SESSION['plan'])
        profile   = deepcopy(SESSION.get('profile', {}))
        mc_data   = deepcopy(SESSION.get('mc'))

    score = profile.get('score', 0)
    plan_json, mc_summary = _build_plan_summary(plan_data, profile, mc_data)

    body = request.get_json(silent=True) or {}
    user_prompt = body.get('user_prompt', '').strip()
    review_mode = body.get('mode', 'standard')  # standard | zhangxuefeng

    base_prompt = ZXF_REVIEW_PROMPT if review_mode == 'zhangxuefeng' else AI_REVIEW_PROMPT
    prompt_text = base_prompt.format(
        score=score,
        ke_lei=profile.get('ke_lei', '物理'),
        target_kw='、'.join(profile.get('target_kw', [])) or '未指定',
        exclude_kw='、'.join(profile.get('exclude_kw', [])) or '无',
        plan_json=plan_json,
    )
    if user_prompt:
        prompt_text += f'\n\n用户额外要求：{user_prompt}'

    t = threading.Thread(target=_ai_review_worker, args=(sid, prompt_text), daemon=True)
    t.start()

    return jsonify({'ok': True, 'message': '已启动 Claude Code 审核…'})


@app.route('/api/ai_review_status')
def api_ai_review_status():
    """轮询：检查审核进度"""
    sid = _get_sid()
    with _AI_REVIEW_LOCK:
        state = _AI_REVIEW_STATE.get(sid, {})
    status = state.get('status', 'idle')
    if status == 'done':
        return jsonify(state)
    if status == 'error':
        return jsonify({'status': 'error', 'error': state.get('error', '未知错误')})
    if status == 'running':
        return jsonify({'status': 'pending'})
    return jsonify({'status': 'idle'})


@app.route('/api/export_excel', methods=['POST'])
def api_export_excel():
    """导出Excel（内存流，避免 Windows 文件缓冲导致截断）"""
    if 'plan' not in SESSION:
        return jsonify({'error': '请先生成志愿方案'}), 400
    import io, copy as _copy
    with _SESSION_LOCK:
        plan_snap    = _copy.deepcopy(SESSION['plan'])
        profile_snap = _copy.deepcopy(SESSION['profile'])
        mc_snap      = _copy.deepcopy(SESSION.get('mc', {}))
    score    = profile_snap['score']
    if plan_snap.get('mode') == 'direct':
        prov = profile_snap.get('student_province', '直填')
        fname = f"zhiyuan_{prov}_{score}fen.xlsx"
        buf = io.BytesIO()
        export_excel_direct(plan_snap, buf)
    else:
        fname    = f"zhiyuan_{score}fen.xlsx"
        buf = io.BytesIO()
        export_excel(plan_snap, mc_snap, buf)
    # 同步写磁盘备份（供调试用）
    out_path = os.path.join(OUT_DIR, fname)
    with open(out_path, 'wb') as f:
        f.write(buf.getvalue())
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/api/status')
def api_status():
    """返回系统状态"""
    has_plan = 'plan' in SESSION
    # 版本校验：若SESSION里的方案来自旧版本，标记为需要重新生成
    version_ok = SESSION.get('plan_version') == PLAN_VERSION if has_plan else True
    return jsonify({
        'has_plan':   has_plan,
        'score':      SESSION.get('profile',{}).get('score') if has_plan else None,
        'plan_count': len(SESSION.get('plan',{}).get('plan_vols',[])),
        'version_ok': version_ok,
        'plan_version': PLAN_VERSION,
    })


REPORTS_DIR = os.path.join(BASE, 'reports')

@app.route('/api/tiqian')
def api_tiqian():
    """查询提前批可报志愿列表（与本科批40个名额完全独立）"""
    try:
        from engine.planner import build_tiqian
        score    = request.args.get('score', type=int)
        ke_lei   = request.args.get('ke_lei', '物理')
        batch    = request.args.get('batch', 'all')   # A / B / all
        score_lo = request.args.get('score_lo', type=int, default=None)
        score_hi = request.args.get('score_hi', type=int, default=None)
        if score is None:
            return jsonify({'error': '缺少score参数'}), 400
        select_subjects = request.args.get('select_subjects', '')
        profile = {
            'score': score, 'ke_lei': ke_lei, 'batch': batch,
            'score_lo': score_lo if score_lo is not None else score - 80,
            'score_hi': score_hi if score_hi is not None else score + 20,
            'select_subjects': [s.strip() for s in select_subjects.split(',') if s.strip()] if select_subjects else [],
        }
        result = build_tiqian(profile)
        return jsonify(result)
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/reports')
def api_reports():
    """返回 reports/ 目录下所有 HTML 报告的专业名称列表（去掉 _物理/_历史 后缀）"""
    if not os.path.isdir(REPORTS_DIR):
        return jsonify({'reports': []})
    names = set()
    for f in os.listdir(REPORTS_DIR):
        if f.endswith('.html'):
            stem = f[:-5]  # 去掉 .html
            # 去掉 _物理 / _历史 科类后缀，保留专业名
            for suffix in ('_物理', '_历史'):
                if stem.endswith(suffix):
                    stem = stem[:-len(suffix)]
                    break
            names.add(stem)
    return jsonify({'reports': sorted(names)})

@app.route('/reports/<path:filename>')
def serve_report(filename):
    """直接提供 reports/ 目录下的 HTML 报告文件（防路径遍历）"""
    from flask import send_from_directory
    return send_from_directory(REPORTS_DIR, filename, mimetype='text/html')


@app.route('/api/major_schools')
def api_major_schools():
    """返回某专业在2025年各院校的录取最低分数据，用于专业选择器hover浮窗"""
    import re as _re
    major = request.args.get('name', '').strip()
    ke    = request.args.get('ke', '物理').strip()
    if not major:
        return jsonify({'error': '缺少name参数'}), 400
    try:
        from engine.sybandb import load_syban_map
        df = load_raw_df()
        base = df[
            (df['年份'] == 2025) &
            (df['科类'] == ke) &
            (df['批次'] == '本科批') &
            (df['公私性质'] == '公办')
        ]

        # ── 直接开设该专业的院校 ──
        sub_direct = base[base['专业名称'] == major].copy()
        sub_direct['via_syban'] = ''          # 直接招生，无实验班标记

        # ── 通过实验班覆盖该专业的院校 ──
        syban_map = load_syban_map()          # {(院校, 实验班名): set(分流专业)}
        syban_rows = []
        direct_schools = set(sub_direct['院校名称'].tolist())
        for (school, syban_name), covered in syban_map.items():
            if major not in covered:
                continue
            # 在 base 里找到该实验班行
            matched = base[
                (base['院校名称'] == school) &
                (base['专业名称'] == syban_name)
            ].copy()
            if matched.empty:
                continue
            matched['via_syban'] = syban_name
            syban_rows.append(matched)

        if syban_rows:
            sub_syban = pd.concat(syban_rows, ignore_index=True)
            # 若该校已有直接专业行，实验班行不重复加入
            sub_syban = sub_syban[~sub_syban['院校名称'].isin(direct_schools)]
            # 同校可能有多个实验班专业组，保留分数最低（最易进入）的那条
            sub_syban['_s25_tmp'] = pd.to_numeric(sub_syban['最低分'], errors='coerce')
            sub_syban = (sub_syban
                         .sort_values('_s25_tmp', ascending=True)
                         .drop_duplicates(subset=['院校名称'], keep='first')
                         .drop(columns=['_s25_tmp']))
        else:
            sub_syban = pd.DataFrame(columns=sub_direct.columns)

        sub = pd.concat([sub_direct, sub_syban], ignore_index=True)
        sub['s25'] = pd.to_numeric(sub['最低分'], errors='coerce')
        sub = sub.dropna(subset=['s25']).sort_values('s25', ascending=False)
        sub = sub.drop_duplicates(subset=['院校名称'], keep='first')

        def lv(tag):
            tag = str(tag)
            if '985' in tag: return '985'
            if '211' in tag: return '211'
            if '省重点' in tag: return '省重点'
            return ''

        def parse_disc(val):
            """从'四轮：A；五轮：A+'中提取五轮结果，无五轮则取四轮"""
            if not val or (isinstance(val, float) and pd.isna(val)): return None
            s = str(val)
            m = _re.search(r'五轮[：:]\s*([A-Za-z+\-]+)', s)
            if m: return m.group(1).strip()
            m = _re.search(r'四轮[：:]\s*([A-Za-z+\-]+)', s)
            if m: return m.group(1).strip()
            return None

        # 计算省内排名（按院校全国排名升序）
        jl_df = df[(df['年份'] == 2025) & (df['所在省'] == '吉林')]
        jl_ranks = {}
        if len(jl_df) and '院校排名' in jl_df.columns:
            jl_sch = jl_df[['院校名称', '院校排名']].copy()
            jl_sch['nat'] = pd.to_numeric(jl_sch['院校排名'], errors='coerce')
            jl_sch = jl_sch.dropna(subset=['nat']).drop_duplicates('院校名称')
            jl_sch = jl_sch.sort_values('nat').reset_index(drop=True)
            for i, row in jl_sch.iterrows():
                jl_ranks[row['院校名称']] = i + 1

        schools = []
        for _, r in sub.iterrows():
            nat  = pd.to_numeric(r.get('院校排名') if '院校排名' in sub.columns else None, errors='coerce')
            rank = pd.to_numeric(r.get('最低位次'), errors='coerce')
            name = str(r['院校名称'])
            schools.append({
                'name':      name,
                'city':      str(r['城市']),
                'province':  str(r['所在省']),
                'score':     int(r['s25']),
                'rank':      int(rank) if pd.notna(rank) else None,
                'lv':        lv(r['院校标签']),
                'disc_eval': parse_disc(r.get('学科评估')) if '学科评估' in sub.columns else None,
                'nat_rank':  int(nat) if pd.notna(nat) else None,
                'prov_rank': jl_ranks.get(name),
                'via_syban': str(r.get('via_syban', '') or ''),
            })

        return jsonify({'major': major, 'ke': ke, 'count': len(schools), 'schools': schools})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/school_majors')
def api_school_majors():
    """返回某院校专业组2025年全部专业录取分数+位次，用于志愿卡片hover浮窗"""
    gcode  = request.args.get('gcode', '').strip()
    ke     = request.args.get('ke', '物理').strip()
    if not gcode:
        return jsonify({'error': '缺少gcode参数'}), 400
    try:
        df = load_raw_df()
        sub = df[
            (df['年份'] == 2025) &
            (df['科类'] == ke) &
            (df['批次'] == '本科批') &
            (df['院校专业组代码'] == gcode)
        ].copy()
        sub['s25'] = pd.to_numeric(sub['最低分'], errors='coerce')
        sub['r25'] = pd.to_numeric(sub['最低位次'], errors='coerce')
        sub = sub.dropna(subset=['s25']).sort_values('s25', ascending=False)
        majors = [{
            'name':  str(r['专业名称']),
            'score': int(r['s25']),
            'rank':  int(r['r25']) if pd.notna(r['r25']) else None,
            'plan':  int(r['计划人数']) if ('计划人数' in sub.columns and pd.notna(r.get('计划人数'))) else None,
        } for _, r in sub.iterrows()]
        school = str(sub.iloc[0]['院校名称']) if len(sub) else gcode
        return jsonify({'school': school, 'ke': ke, 'count': len(majors), 'majors': majors})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/save_plan', methods=['POST'])
def api_save_plan():
    """保存用户手动编辑后的方案（增删志愿/调整顺序/修改专业）"""
    if 'plan' not in SESSION:
        return jsonify({'error': '无当前方案'}), 400
    body  = request.json or {}
    edits = body.get('vols', [])

    from engine.planner import LV_LABEL, CR_LABEL
    with _SESSION_LOCK:
        plan_vols = SESSION['plan']['plan_vols']
        vol_map   = {str(v['gcode']): v for v in plan_vols}

        new_plan_vols = []
        for i, e in enumerate(edits):
            gcode = e.get('gcode', '')
            if gcode in vol_map:
                v = dict(vol_map[gcode])
            else:
                # 新插入的志愿组（来自 search_groups）
                slv = int(e.get('school_lv', 6))
                crk = int(e.get('city_rank', 4))
                v = {
                    'gcode':     gcode,
                    'school':    e.get('school', ''),
                    'city':      e.get('city', ''),
                    'school_lv': slv,
                    'city_rank': crk,
                    'lv_label':  LV_LABEL.get(slv, '其他'),
                    'cr_label':  CR_LABEL.get(crk, '其他'),
                    'gmin25':    e.get('gmin25'),
                    'gmin24':    e.get('gmin24'),
                    'gmin23':    e.get('gmin23'),
                    'sc6':       e.get('sc6'),
                    'gmin_rank': None,
                    'n_target':  0, 'n_cold': 0,
                    'safe':      False,
                    'majors':    [{'name': m['name'], 's25': m.get('s25'),
                                   'diff': m.get('diff', 0), 'kind': m.get('kind', 'other')}
                                  for m in e.get('intent6', [])],
                    'diaoji':    True,
                    'dedup_count': 0,
                    'warn_few_majors': False, 'warn_critical': False,
                    'warn_msg': '', 'warn_cold_anchor': False, 'warn_msg_cold': '',
                }
            v = dict(v)
            v['vol_idx'] = i + 1
            v['tp']  = e.get('tp', v.get('tp', '稳'))
            v['top6'] = [{'name': m['name'], 's25': m.get('s25'),
                          'diff': m.get('diff'), 'kind': m.get('kind', 'other')}
                         for m in e.get('intent6', [])]
            new_plan_vols.append(v)

        SESSION['plan']['plan_vols'] = new_plan_vols
    return jsonify({'ok': True, 'count': len(new_plan_vols)})


@app.route('/api/replenish', methods=['POST'])
def api_replenish():
    """
    补充志愿：用户删除部分志愿组后，自动补充到目标数量并重新排序。
    请求体: { target_count: 40 }  (可选，默认补充到40个)
    """
    if 'plan' not in SESSION:
        return jsonify({'error': '无当前方案'}), 400

    body = request.json or {}
    target_count = body.get('target_count', 40)

    with _SESSION_LOCK:
        snap_plan = __import__('copy').deepcopy(SESSION.get('plan', {}))
        snap_profile = __import__('copy').deepcopy(SESSION.get('profile', {}))

    current_vols = snap_plan.get('plan_vols', [])
    current_count = len(current_vols)

    if current_count >= target_count:
        return jsonify({'ok': True, 'msg': '志愿数量已满，无需补充',
                        'count': current_count, 'added': 0})

    need = target_count - current_count
    existing_gcodes = {str(v['gcode']) for v in current_vols}

    # 用户主动删除的志愿组，补充时也要排除（避免补回已删除的）
    exclude_gcodes = set(str(g) for g in body.get('exclude_gcodes', []))

    # 用当前 profile 重新生成完整方案
    try:
        t0 = time.time()
        full_result = build_plan(snap_profile)
        full_vols = full_result.get('plan_vols', [])

        # 从完整方案中筛选出：不在当前列表中 且 不在用户删除列表中
        candidates = [v for v in full_vols
                      if str(v['gcode']) not in existing_gcodes
                      and str(v['gcode']) not in exclude_gcodes]

        # 按原始排序取前 need 个
        added = candidates[:need]

        # 合并：当前志愿 + 新增志愿
        merged = list(current_vols) + added

        # 梯度重分类：对齐 build_plan 的冲/稳/保定义
        # 冲: gmin25 > score（不会被提档，安全冲击高校）
        # 稳: gmin25 ≤ score 且 sc6 在 [score-45, score-2]（能提档，sc6 保底）
        # 保: score - sc6 ≥ 10（大幅超过保底线）
        score = snap_profile.get('score', 500)
        for v in merged:
            sc6 = v.get('sc6')
            gmin = v.get('gmin25')
            gmin_f = float(gmin) if gmin is not None else 0

            if gmin_f > score:
                # 组最低分高于考生分 → 冲志愿
                v['tp'] = '冲'
            elif sc6 is not None:
                diff = score - sc6
                if diff < 0:
                    v['tp'] = '冲'     # sc6 > score，有调剂风险
                elif diff < 10:
                    v['tp'] = '稳'     # 稳区：sc6 接近考生分
                else:
                    v['tp'] = '保'     # 保区：大幅超过 sc6
            else:
                # sc6 为 None（专业少于6个），按 gmin 判断
                v['tp'] = '稳' if gmin_f <= score else '冲'

        # 按冲稳保分组，组内按sc6降序（与 build_plan 一致）
        tp_order = {'冲': 0, '稳': 1, '保': 2}
        merged.sort(key=lambda v: (tp_order.get(v.get('tp', '稳'), 1),
                                   -(v.get('sc6') or 0)))

        # 重新编号
        for i, v in enumerate(merged):
            v['vol_idx'] = i + 1

        # 序列化输出（与 api_generate 格式完全对齐，含 warn 字段）
        vols_out = []
        for v in merged:
            intent6 = v.get('top6', v.get('intent6', []))[:6]
            vols_out.append({
                'vol_idx':   v['vol_idx'],
                'tp':        v['tp'],
                'school':    v['school'],
                'city':      v.get('city', ''),
                'province':  v.get('province', ''),
                'lv_label':  v.get('lv_label','?'),
                'cr_label':  v.get('cr_label','?'),
                'school_lv': v.get('school_lv', 6),
                'city_rank': v.get('city_rank', 4),
                'gcode':     v['gcode'],
                'gmin25':    v.get('gmin25'),
                'gmin24':    v.get('gmin24'),
                'gmin23':    v.get('gmin23'),
                'sc6':       v.get('sc6'),
                'safe':      v.get('safe', False),
                'n_target':  v.get('n_target', 0),
                'n_cold':    v.get('n_cold', 0),
                'gmin_rank': v.get('gmin_rank'),
                'intent6':   [{'name':m['name'],'s25':m.get('s25'),
                               'diff':m.get('diff'),'kind':m.get('kind','other'),
                               'fee': m.get('fee'), 'r25': m.get('r25'), 'r24': m.get('r24'),
                               'syban_majors':m.get('syban_majors',[]),
                               'syban_all':   m.get('syban_all',[]),
                               'zhuanxiang':  m.get('zhuanxiang',''),
                               'full_name':   m.get('full_name','')} for m in intent6],
                'has_zhuanxiang':    v.get('has_zhuanxiang', False),
                'zhuanxiang_types':  v.get('zhuanxiang_types', []),
                'all_majors_count': len(v.get('majors',[])),
                'dedup_count':      v.get('dedup_count', 0),
                'diaoji':           v.get('diaoji', True),
                'warn_few_majors':  v.get('warn_few_majors', False),
                'warn_critical':    v.get('warn_critical', False),
                'warn_msg':         v.get('warn_msg', ''),
                'warn_excl_major':  v.get('warn_excl_major', False),
                'warn_msg_excl':    v.get('warn_msg_excl', ''),
                'warn_cold_anchor': v.get('warn_cold_anchor', False),
                'warn_msg_cold':    v.get('warn_msg_cold', ''),
                'warn_tuidan':      v.get('warn_tuidan', False),
                'warn_msg_tuidan':  v.get('warn_msg_tuidan', ''),
            })

        elapsed = round(time.time() - t0, 2)

        with _SESSION_LOCK:
            SESSION['plan']['plan_vols'] = merged

        return jsonify({
            'ok': True,
            'count': len(merged),
            'added': len(added),
            'elapsed': elapsed,
            'vols': vols_out,
        })
    except Exception as e:
        return jsonify({'error': f'补充失败: {str(e)}'}), 500


@app.route('/api/search_groups')
def api_search_groups():
    """搜索院校专业组，用于手动插入志愿"""
    q     = request.args.get('q', '').strip()
    ke    = request.args.get('ke', '物理').strip()
    score = _safe_int(request.args.get('score', 0))
    if not q or len(q) < 2:
        return jsonify({'error': '请输入至少2个字'}), 400
    try:
        from engine.planner import school_level, city_rank as cr_fn, LV_LABEL, CR_LABEL
        df  = load_raw_df()
        sub = df[
            (df['年份'] == 2025) &
            (df['科类'] == ke) &
            (df['批次'] == '本科批') &
            (df['院校名称'].str.contains(q, na=False, regex=False))
        ].copy()
        sub['s25'] = pd.to_numeric(sub['最低分'], errors='coerce')
        sub = sub.dropna(subset=['s25', '院校专业组代码'])

        groups = []
        for gcode, grp in sub.groupby('院校专业组代码'):
            school  = str(grp['院校名称'].iloc[0])
            city    = str(grp['城市'].iloc[0])   if '城市'        in grp.columns else '?'
            lv_tag  = str(grp['院校标签'].iloc[0]) if '院校标签'   in grp.columns else ''
            cr_tag  = str(grp['城市水平标签'].iloc[0]) if '城市水平标签' in grp.columns else ''
            slv     = school_level(lv_tag)
            crk     = cr_fn(cr_tag)
            gmin25  = float(grp['s25'].min())
            sorted_s = sorted(grp['s25'].dropna().tolist())
            sc6     = float(sorted_s[min(5, len(sorted_s)-1)]) if sorted_s else gmin25
            majors  = [{'name': str(r['专业名称']), 's25': int(r['s25'])}
                       for _, r in grp.sort_values('s25', ascending=False).iterrows()]
            groups.append({
                'gcode':     str(gcode),
                'school':    school,
                'city':      city,
                'school_lv': slv,
                'city_rank': crk,
                'lv_label':  LV_LABEL.get(slv, '其他'),
                'cr_label':  CR_LABEL.get(crk, '其他'),
                'gmin25':    gmin25,
                'sc6':       sc6,
                'majors':    majors,
            })
        groups.sort(key=lambda g: -g['gmin25'])
        return jsonify({'ok': True, 'count': len(groups), 'groups': groups[:40]})
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/search_direct')
def api_search_direct():
    """搜索直填模式院校专业，用于手动插入志愿"""
    q     = request.args.get('q', '').strip()
    ke    = request.args.get('ke', '物理').strip()
    prov  = request.args.get('prov', '').strip()
    score = _safe_int(request.args.get('score', 0))
    if not q or len(q) < 2:
        return jsonify({'error': '请输入至少2个字'}), 400
    if not prov:
        return jsonify({'error': '缺少省份参数'}), 400
    try:
        from engine.db import load_direct_df
        df = load_direct_df(prov)
        _BATCH_OK = {'本科批', '一段线'}
        sub = df[
            (df['年份'] == 2025) &
            (df['批次'].isin(_BATCH_OK)) &
            (df['院校名称'].str.contains(q, na=False, regex=False))
        ].copy()
        # 科类过滤（山东/浙江不分科）
        cfg = _DIRECT_PROV_CFG.get(prov, {})
        if cfg.get('ke_split') and '科类' in sub.columns:
            sub = sub[sub['科类'] == ke]
        sub['s25'] = pd.to_numeric(sub['最低分'], errors='coerce')
        sub = sub.dropna(subset=['s25'])
        # 只取公办
        if '公私性质' in sub.columns:
            sub = sub[sub['公私性质'] == '公办']
        results = []
        for _, r in sub.sort_values('s25', ascending=False).head(60).iterrows():
            s25 = float(r['s25'])
            results.append({
                'school':    str(r.get('院校名称', '')),
                'major':     str(r.get('专业名称', '')),
                'city':      str(r.get('城市', '')),
                'ke_lei':    str(r.get('科类', '')),
                'batch':     str(r.get('批次', '')),
                'subj_req':  str(r.get('选科要求', '不限')),
                'school_lv': str(r.get('院校层级', '')),
                'tags':      str(r.get('院校标签', '')),
                's25':       s25,
                'diff':      round(s25 - score, 1),
                'tuition':   r.get('学费'),
                'plan_count': r.get('计划人数'),
            })
        return jsonify({'ok': True, 'count': len(results), 'results': results})
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500


# ── 地图页 ─────────────────────────────────────────────────
@app.route('/map')
def map_page():
    return _no_cache(make_response(render_template('map.html')))


@app.route('/api/map_data')
def api_map_data():
    """返回全国院校城市分布数据，用于地图可视化"""
    ke = request.args.get('ke', '物理').strip()
    try:
        from engine.planner import school_level as _school_lv
        df = load_raw_df()

        mask = (df['年份'] == 2025) & (df['科类'] == ke) & (df['批次'] == '本科批')
        if '公私性质' in df.columns:
            mask &= (df['公私性质'] == '公办')
        sub = df[mask].copy()

        # 每所学校取最低分（跨所有专业组）
        sub['_s25'] = pd.to_numeric(sub['最低分'], errors='coerce')
        score_by_sch = sub.groupby('院校名称')['_s25'].min().reset_index()
        score_by_sch.columns = ['院校名称', 'min_score']

        # 每所学校的唯一信息
        keep_cols = [c for c in ['院校名称', '所在省', '城市', '院校标签'] if c in sub.columns]
        sch = sub[keep_cols].drop_duplicates('院校名称').copy()
        sch['school_lv'] = sch['院校标签'].apply(
            lambda x: _school_lv(x) if pd.notna(x) else 6
        )
        sch = sch.merge(score_by_sch, on='院校名称', how='left')

        # 按城市分组
        city_dict = {}
        for _, row in sch.iterrows():
            city = str(row.get('城市', '') or '')
            if not city or city == 'nan':
                city = str(row.get('所在省', '') or '')
            if not city or city == 'nan':
                continue
            prov = str(row.get('所在省', '') or '')
            lv   = int(row['school_lv'])
            name = str(row['院校名称'])
            sc   = row['min_score']
            sc_i = int(sc) if pd.notna(sc) else None

            if city not in city_dict:
                city_dict[city] = {
                    'city': city, 'province': prov,
                    'schools': [],
                    'lv_counts': {str(i): 0 for i in range(1, 7)},
                    'top_lv': 6,
                }
            city_dict[city]['schools'].append({'name': name, 'lv': lv, 'score': sc_i})
            city_dict[city]['lv_counts'][str(lv)] += 1

        cities = []
        for data in city_dict.values():
            for lv in range(1, 7):
                if data['lv_counts'].get(str(lv), 0) > 0:
                    data['top_lv'] = lv
                    break
            data['count'] = len(data['schools'])
            data['schools'].sort(key=lambda s: (s['lv'], -(s['score'] or 0)))
            data['schools'] = data['schools'][:30]  # 每城市最多30所
            cities.append(data)

        # 当前方案城市信息
        plan_cities = {}
        has_plan = 'plan' in SESSION
        if has_plan:
            from copy import deepcopy
            with _SESSION_LOCK:
                _plan_vols_snap = deepcopy(SESSION['plan']['plan_vols'])
            for v in _plan_vols_snap:
                city = v.get('city', '')
                if city:
                    if city not in plan_cities:
                        plan_cities[city] = []
                    plan_cities[city].append({
                        'school': v['school'],
                        'lv':     v['school_lv'],
                        'tp':     v['tp'],
                        'gmin25': v.get('gmin25'),
                    })

        return jsonify({
            'ok':           True,
            'cities':       cities,
            'plan_cities':  plan_cities,
            'has_plan':     has_plan,
            'ke':           ke,
            'total_schools': sum(d['count'] for d in cities),
        })
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500


# ══ 数据库增强 API ═════════════════════════════════════════

@app.route('/api/db/stats')
def api_db_stats():
    """返回数据库统计信息"""
    if not gaokao_db.db_exists():
        return jsonify({'error': '数据库不存在，请先运行迁移'}), 404
    return jsonify(gaokao_db.get_stats())

@app.route('/api/db/major_tree')
def api_db_major_tree():
    """返回专业目录树形结构（门类→类别→专业）"""
    level = request.args.get('level', '本科')
    return jsonify(gaokao_db.get_major_tree(level))

@app.route('/api/db/search_majors')
def api_db_search_majors():
    """按关键词搜索专业目录"""
    kw = request.args.get('q', '')
    level = request.args.get('level', '本科')
    if not kw:
        return jsonify([])
    return jsonify(gaokao_db.search_majors(kw, level))

@app.route('/api/db/search_schools')
def api_db_search_schools():
    """搜索院校"""
    kw = request.args.get('q', '')
    prov = request.args.get('province', '')
    tags = request.args.get('tags', '')
    limit = _safe_int(request.args.get('limit', 50), 50)
    return jsonify(gaokao_db.search_schools(kw, prov, tags, limit))

@app.route('/api/db/school_detail')
def api_db_school_detail():
    """查询院校的专业组和专业详情"""
    name = request.args.get('name', '')
    year = _safe_int(request.args.get('year', 2025), 2025)
    ke = request.args.get('ke_lei', '物理')
    if not name:
        return jsonify({'error': '请提供院校名称'}), 400
    return jsonify(gaokao_db.get_school_majors(name, year, ke))

@app.route('/api/db/save_plan', methods=['POST'])
def api_db_save_plan():
    """保存方案到数据库"""
    body = request.json or {}
    profile = body.get('profile', SESSION.get('profile', {}))
    plan_json = body.get('plan_json', {})
    plan_id = gaokao_db.save_plan(profile, plan_json)
    return jsonify({'ok': True, 'plan_id': plan_id})

@app.route('/api/db/plans')
def api_db_plans():
    """读取历史方案列表"""
    limit = _safe_int(request.args.get('limit', 10), 10)
    return jsonify(gaokao_db.load_plans(limit=limit))

@app.route('/api/db/plans/<int:plan_id>', methods=['DELETE'])
def api_db_delete_plan(plan_id):
    """删除指定方案"""
    ok = gaokao_db.delete_plan(plan_id)
    return jsonify({'ok': ok})


if __name__ == '__main__':
    import os as _os, sys as _sys
    # Windows GBK 终端 emoji 兼容
    if hasattr(_sys.stdout, 'reconfigure'):
        try: _sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        except Exception: pass
    # 打包环境下自动打开浏览器
    if getattr(_sys, 'frozen', False):
        import threading as _th2, webbrowser as _wb, time as _t2
        def _open(): _t2.sleep(2); _wb.open('http://localhost:5000')
        _th2.Thread(target=_open, daemon=True).start()
    print("\n" + "="*52)
    print("  🎓 高考志愿规划系统 · 本地版 v3.5")
    print("="*52)
    # 检测数据源
    _db = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), 'data', 'gaokao.db')
    _cache = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), 'data', 'df_cache.pkl')
    if _os.path.exists(_db):
        print("  预加载数据中（读取SQLite数据库）...", end='', flush=True)
    elif _os.path.exists(_cache):
        print("  预加载数据中（读取pickle缓存）...", end='', flush=True)
    else:
        print("  首次启动，正在读取Excel数据（约30秒，请耐心等待）...", end='', flush=True)
    try:
        import threading as _th
        _done = [False]
        def _dot():
            import time as _t
            while not _done[0]:
                print('.', end='', flush=True)
                _t.sleep(2)
        _th.Thread(target=_dot, daemon=True).start()
        df = load_raw_df()
        _done[0] = True
        print(f"\n  ✅ 数据加载成功: {len(df)} 行")
    except Exception as e:
        _done[0] = True
        print(f"\n  ❌ 数据加载失败: {e}")
        print("  请确认 data/2026_jilin_gaokao_data.xlsx 文件存在")
        _sys.exit(1)
    print(f"  🌐 地址: http://localhost:5000")
    print("  Ctrl+C 停止\n")
    app.run(debug=False, port=5000, host='127.0.0.1')
