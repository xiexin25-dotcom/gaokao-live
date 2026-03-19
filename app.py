"""
吉林省高考志愿规划系统 · 本地版
python app.py → http://localhost:5000
"""
import os, sys, json, time, threading
import pandas as pd
from flask import Flask, render_template, request, jsonify, send_file, make_response

sys.path.insert(0, os.path.dirname(__file__))
from engine.planner import (build_plan, mc_simulate, optimize_plan, export_excel,
                             KW_LIB, EXCLUDE_PRESETS, load_raw_df)
from engine.system_prompt import SYSTEM_PROMPT

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

# ── 当前会话方案（内存） ──────────────────────────────────
SESSION = {}
SETTINGS = {}         # 用户设置（API Key 等）
_SESSION_LOCK = threading.Lock()
_HISTORY_LOCK = threading.Lock()
_CHAT_LAST_TIME = 0.0   # 上次 chat 请求时间戳，用于速率限制
PLAN_VERSION = 'v3.1'   # 每次规则变更时递增，旧SESSION自动失效

# ── 历史方案持久化 ────────────────────────────────────────
import json as _json

def _history_files():
    """返回按时间降序排列的历史方案文件路径列表"""
    import glob
    files = glob.glob(os.path.join(OUT_DIR, 'plan_*.json'))
    files.sort(reverse=True)
    return files

def _load_history():
    """从磁盘加载最近10条历史方案"""
    history = []
    for fpath in _history_files()[:10]:
        try:
            with open(fpath, 'r', encoding='utf-8') as f:
                history.append(_json.load(f))
        except Exception:
            pass
    return history

def _save_history_entry(entry):
    """将单条方案写入磁盘"""
    import datetime
    fname = 'plan_' + datetime.datetime.now().strftime('%Y%m%d_%H%M%S') + '.json'
    fpath = os.path.join(OUT_DIR, fname)
    try:
        with open(fpath, 'w', encoding='utf-8') as f:
            _json.dump(entry, f, ensure_ascii=False, indent=2)
        # 超过10条时删除最旧的
        files = _history_files()
        for old in files[10:]:
            try: os.remove(old)
            except Exception: pass
    except Exception:
        pass

PLAN_HISTORY = _load_history()   # 启动时从磁盘恢复

def validate_plan(plan_vols, score):
    """验证方案合法性，返回 (ok:bool, warnings:list)"""
    warnings = []

    # ── 吉林省平行志愿核心规则警告（一次性全局） ──────────────────────────
    warnings.append(
        "📌【吉林平行志愿铁规】一次投档，不补充投档。"
        "若投档后因体检/单科成绩不达标被退档，本轮所有后续志愿作废，"
        "仅能参加征集志愿或下一批次。请务必确保体检、单科满足每所院校要求。"
    )

    if len(plan_vols) < 10:
        warnings.append(f"⚠️ 当前分数候选院校不足，仅生成 {len(plan_vols)} 个志愿，建议降低筛选条件")

    # 高体检/政审退档风险专业关键词
    HIGH_RISK_KWS = ['军事', '军队', '国防', '公安', '警察', '司法', '军医', '海军', '空军',
                     '武警', '飞行', '航海', '轮机', '船舶驾驶', '消防', '特警']

    for v in plan_vols:
        tp = v.get('tp'); gmin = float(v.get('gmin25') or 0)
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

    return (len(warnings) == 0, warnings)

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
    with _SESSION_LOCK:
        SESSION['profile']      = h['profile']
        SESSION['mc']           = h['mc']
        SESSION['plan_version'] = h.get('plan_version', PLAN_VERSION)
    return jsonify({
        'ok':      True,
        'vols':    h['vols_out'],
        'mc':      h['mc'],
        'profile': h['profile'],
        'stats':   h['stats'],
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

# ══ API ══════════════════════════════════════════════════

@app.route('/api/keywords')
def api_keywords():
    """返回专业关键词库和排除预设"""
    return jsonify({'kw_lib': KW_LIB, 'exclude_presets': EXCLUDE_PRESETS})

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
        if ke_lei not in ('物理', '历史'):
            return jsonify({'error': f"科类必须为'物理'或'历史'，收到: {ke_lei}"}), 400

        fee_max_raw = body.get('fee_max', None)
        profile = {
            'score':            score,
            'ke_lei':           ke_lei,
            'target_kw':        body.get('target_kw', []),
            'exclude_kw':       body.get('exclude_kw', []),
            'exclude_northeast':body.get('exclude_northeast', False),
            'pref_provinces':   body.get('pref_provinces', []),
            'exclude_provinces':body.get('exclude_provinces', []),
            'include_types':    body.get('include_types', []),
            'exclude_types':    body.get('exclude_types', []),
            'min_city_rank':    int(body.get('min_city_rank', 4)),
            'school_pref':      body.get('school_pref', 'school'),
            'slope':            float(body.get('slope', 150.0)),
            'fee_max':          int(fee_max_raw) if fee_max_raw else None,
        }

        t0 = time.time()
        result = build_plan(profile)
        elapsed = round(time.time()-t0, 2)

        plan_vols = result['plan_vols']
        stats     = result['stats']

        # 快速MC（N=5000，基准）—— 使用 build_plan 算好的实际位次
        mc = mc_simulate(plan_vols, N=5000, seed=42, bias_lo=0, bias_hi=0, noise_pct=8,
                         student_rank=stats['student_rank'], student_score=score)

        with _SESSION_LOCK:
            SESSION['plan']    = result
            SESSION['mc']      = mc
            SESSION['profile'] = profile

        # 序列化输出（去掉大型原始字段）
        vols_out = []
        for v in plan_vols:
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
                               'syban_majors':m.get('syban_majors',[]),
                               'syban_all':   m.get('syban_all',[])} for m in intent6],
                'all_majors_count': len(v.get('majors',[])),
                'dedup_count':      v.get('dedup_count', 0),
                'diaoji':           v.get('diaoji', True),
                'warn_few_majors':  v.get('warn_few_majors', False),
                'warn_critical':    v.get('warn_critical', False),
                'warn_msg':         v.get('warn_msg', ''),
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
        }
        with _HISTORY_LOCK:
            PLAN_HISTORY.insert(0, hist_entry)
            if len(PLAN_HISTORY) > 10:
                PLAN_HISTORY.pop()
        _save_history_entry(hist_entry)

        return jsonify({
            'ok':       True,
            'elapsed':  elapsed,
            'stats':    stats,
            'vols':     vols_out,
            'mc':       mc,
            'profile':  profile,
            'plan_ok':  plan_ok,
            'warnings': plan_warnings,
        })

    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500


@app.route('/api/simulate', methods=['POST'])
def api_simulate():
    """对当前方案做MC模拟（自定义参数）"""
    if 'plan' not in SESSION:
        return jsonify({'error': '请先生成志愿方案'}), 400

    body = request.json or {}
    N         = min(int(body.get('N', 10000)), 300000)
    seed      = int(body.get('seed', 42))
    # 偏移区间：正 = 悲观（分数线上升），负 = 乐观（分数线下降）
    bias_lo   = float(body.get('bias_lo', 0))
    bias_hi   = float(body.get('bias_hi', 0))
    noise_pct = float(body.get('noise', 8))

    plan_vols = SESSION['plan']['plan_vols']
    profile   = SESSION['profile']

    score = profile['score']
    student_rank = SESSION['plan']['stats']['student_rank']

    mc = mc_simulate(plan_vols, N=N, seed=seed,
                     bias_lo=bias_lo, bias_hi=bias_hi, noise_pct=noise_pct,
                     student_rank=student_rank, student_score=score)
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

    body       = request.json or {}
    max_rounds = int(body.get('max_rounds', 10))
    mc_n       = min(int(body.get('mc_n', 8000)), 50000)
    noise_pct  = float(body.get('noise_pct', 15.0))

    try:
        t0     = time.time()
        opt    = optimize_plan(
            SESSION['plan'],
            max_rounds=max_rounds,
            mc_n=mc_n,
            noise_pct=noise_pct,
            seed=42,
        )
        elapsed = round(time.time() - t0, 2)

        # 用优化后的方案更新 SESSION
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
                               'diff': m.get('diff'), 'kind': m.get('kind','other')} for m in intent6],
                'diaoji':           v.get('diaoji', True),
                'dedup_count':      v.get('dedup_count', 0),
                'warn_few_majors':  v.get('warn_few_majors', False),
                'warn_critical':    v.get('warn_critical', False),
                'warn_msg':         v.get('warn_msg', ''),
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
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500



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

    body             = request.json or {}
    locked_codes     = set(body.get('locked_codes', []))
    excluded_schools = set(body.get('excluded_schools', []))
    max_rounds       = int(body.get('max_rounds', 10))
    mc_n             = min(int(body.get('mc_n', 8000)), 50000)
    noise_pct        = float(body.get('noise_pct', 15.0))
    score            = SESSION['profile']['score']

    try:
        t0 = time.time()

        # 记录优化前的方案，用于 diff 对比
        before_vols = SESSION['plan']['plan_vols']
        before_map  = {v['gcode']: v for v in before_vols}

        opt = optimize_plan(
            SESSION['plan'],
            max_rounds=max_rounds,
            mc_n=mc_n,
            noise_pct=noise_pct,
            seed=42,
            locked_codes=locked_codes,
            excluded_schools=excluded_schools,
        )
        elapsed = round(time.time() - t0, 2)

        # 更新 SESSION
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
        after_gcodes = {v['gcode'] for v in before_vols}
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
                'is_new':    v['gcode'] not in after_gcodes,       # 新加入的志愿
                'is_locked': v['gcode'] in locked_codes,            # 用户锁定的
                'intent6':   [{'name': m['name'], 's25': m.get('s25'),
                               'diff': m.get('diff')} for m in intent6],
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
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500



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
    global _CHAT_LAST_TIME
    now = time.time()
    if now - _CHAT_LAST_TIME < 3.0:
        return jsonify({'error': '请求过于频繁，请稍后再试（每3秒限1次）'}), 429
    _CHAT_LAST_TIME = now

    api_key = SETTINGS.get('api_key', '')
    if not api_key:
        return jsonify({'error': '请先在设置页填写 Gemini API Key'}), 400

    body    = request.json or {}
    message = body.get('message', '').strip()
    history = body.get('history', [])

    if not message:
        return jsonify({'error': '消息不能为空'}), 400

    # 构建方案摘要注入 system prompt
    plan_json  = '（尚未生成方案）'
    mc_summary = '（尚未运行 MC 仿真）'
    if 'plan' in SESSION:
        vols = SESSION['plan']['plan_vols']
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


@app.route('/api/export_excel', methods=['POST'])
def api_export_excel():
    """导出Excel（内存流，避免 Windows 文件缓冲导致截断）"""
    if 'plan' not in SESSION:
        return jsonify({'error': '请先生成志愿方案'}), 400
    import io
    profile  = SESSION['profile']
    score    = profile['score']
    fname    = f"zhiyuan_{score}fen.xlsx"
    buf = io.BytesIO()
    export_excel(SESSION['plan'], SESSION.get('mc', {}), buf)
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


REPORTS_DIR = r'C:\Users\谢欣\major_report_system\reports'

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
        if not score:
            return jsonify({'error': '缺少score参数'}), 400
        profile = {
            'score': score, 'ke_lei': ke_lei, 'batch': batch,
            'score_lo': score_lo if score_lo is not None else score - 80,
            'score_hi': score_hi if score_hi is not None else score + 20,
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
    """直接提供 reports/ 目录下的 HTML 报告文件"""
    filepath = os.path.join(REPORTS_DIR, filename)
    if not os.path.isfile(filepath):
        return '报告文件不存在', 404
    return send_file(filepath, mimetype='text/html')


@app.route('/api/major_schools')
def api_major_schools():
    """返回某专业在吉林省2025年各院校的录取最低分数据，用于专业选择器hover浮窗"""
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

        # 计算吉林省内排名（按院校全国排名升序）
        jl_df = df[(df['年份'] == 2025) & (df['所在省'] == '吉林')]
        jl_ranks = {}
        if len(jl_df):
            jl_sch = jl_df[['院校名称', '院校排名']].copy()
            jl_sch['nat'] = pd.to_numeric(jl_sch['院校排名'], errors='coerce')
            jl_sch = jl_sch.dropna(subset=['nat']).drop_duplicates('院校名称')
            jl_sch = jl_sch.sort_values('nat').reset_index(drop=True)
            for i, row in jl_sch.iterrows():
                jl_ranks[row['院校名称']] = i + 1

        schools = []
        for _, r in sub.iterrows():
            nat  = pd.to_numeric(r.get('院校排名'), errors='coerce')
            rank = pd.to_numeric(r.get('最低位次'), errors='coerce')
            name = str(r['院校名称'])
            schools.append({
                'name':      name,
                'city':      str(r['城市']),
                'province':  str(r['所在省']),
                'score':     int(r['s25']),
                'rank':      int(rank) if (rank == rank) else None,
                'lv':        lv(r['院校标签']),
                'disc_eval': parse_disc(r.get('学科评估')),
                'nat_rank':  int(nat) if (nat == nat) else None,
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
            'rank':  int(r['r25']) if (r['r25'] == r['r25']) else None,
            'plan':  int(r['计划人数']) if pd.notna(r['计划人数']) else None,
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


@app.route('/api/search_groups')
def api_search_groups():
    """搜索院校专业组，用于手动插入志愿"""
    q     = request.args.get('q', '').strip()
    ke    = request.args.get('ke', '物理').strip()
    score = int(request.args.get('score', 0))
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
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500


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
            for v in SESSION['plan']['plan_vols']:
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
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500


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
    print("  🎓 吉林省高考志愿规划系统 · 本地版 v3.4")
    print("="*52)
    # 检测缓存
    _cache = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), 'data', 'df_cache.pkl')
    if _os.path.exists(_cache):
        print("  预加载数据中（读取缓存）...", end='', flush=True)
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
    app.run(debug=False, port=5000, host='0.0.0.0')
