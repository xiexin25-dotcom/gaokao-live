# 代码测试报告 · 第7轮

**测试时间**: 2026-03-24
**代码版本**: fix/hardcoded-paths 分支 · v3.5
**文件**: app.py (1579行) · engine/planner.py (1257行) · engine/db.py (502行)

---

## 一、历史问题修复确认（第6轮遗留 → 全部已修复 ✅）

| # | 优先级 | 问题 | 状态 | 修复位置 |
|---|--------|------|------|----------|
| 1 | P0 | `api_replenish` 梯度分类缺失 `sc6>score` 和 `sc6 is None` 分支 | ✅ 已修复 | app.py:1230-1248，三分支完整覆盖 |
| 2 | P1 | `api_export_excel` 读 SESSION 无锁 | ✅ 已修复 | app.py:901-905，`_SESSION_LOCK` + `deepcopy` |
| 3 | P1 | `export_excel` sheet1 `int(m['s25'])` 无 NaN 保护 | ✅ 已修复 | planner.py:1148，`_safe_s()` 函数含 `pd.notna()` |
| 4 | P1 | `api_replenish` vols_out 缺少 warn_* 字段 | ✅ 已修复 | app.py:1294-1302，完整输出所有 warn 字段 |
| 5 | P1 | `api_replenish` 梯度阈值与 `build_plan` 不一致 | ✅ 已修复 | app.py:1225-1248，注释明确对齐定义 |

---

## 二、本轮新发现问题

### Issue #1 · P1 — `api_chat` 读取 `SESSION['mc']` 无锁且无 deepcopy

**文件**: app.py:847-855
**类型**: 线程安全 / 数据竞争

`api_chat` 对 `plan_vols` 已正确使用锁+deepcopy（行835-836），但随后读取 `SESSION.get('mc')` 时**既无锁也无 deepcopy**：

```python
# 行 835-836：✅ 正确
with _SESSION_LOCK:
    vols = deepcopy(SESSION['plan']['plan_vols'])

# 行 847-855：❌ 无锁无 deepcopy
mc_data = SESSION.get('mc')          # 裸读
if mc_data:
    rates = mc_data.get('rates', []) # 引用原始对象
    tops  = mc_data.get('top_majors', [])
```

若用户在 AI 咨询同时点击"重新模拟"（`api_simulate` 写入 `SESSION['mc']`），`mc_data` 可能引用到半更新的 dict，导致 `rates` 和 `tops` 长度不一致，触发行 853-854 的 IndexError。

**修复建议**：将 mc 读取合并到已有的 `_SESSION_LOCK` 块内：
```python
with _SESSION_LOCK:
    vols = deepcopy(SESSION['plan']['plan_vols'])
    mc_data = deepcopy(SESSION.get('mc'))
```

---

### Issue #2 · P1 — `api_status` 读 SESSION 无锁保护

**文件**: app.py:919-931
**类型**: 线程安全 / TOCTOU

```python
def api_status():
    has_plan = 'plan' in SESSION                              # 裸读①
    version_ok = SESSION.get('plan_version') == PLAN_VERSION  # 裸读②
    return jsonify({
        'has_plan':   has_plan,
        'score':      SESSION.get('profile',{}).get('score'),  # 裸读③
        'plan_count': len(SESSION.get('plan',{}).get('plan_vols',[])),  # 裸读④
        ...
    })
```

4 次裸读之间无原子性保证。极端情况：`has_plan=True` 但 `plan_count=0`（plan 在读取间被清空），前端可能展示矛盾状态。

**修复建议**：单次锁内快照：
```python
with _SESSION_LOCK:
    snap = {
        'has_plan': 'plan' in SESSION,
        'score': SESSION.get('profile',{}).get('score'),
        ...
    }
return jsonify(snap)
```

---

### Issue #3 · P1 — `PLAN_HISTORY` 读写竞争

**文件**: app.py:237-283
**类型**: 线程安全 / TOCTOU

`api_generate` 修改 `PLAN_HISTORY` 时用 `_HISTORY_LOCK`（行470-473），但以下读操作**未加锁**：

1. `api_history`（行239-248）：遍历 `PLAN_HISTORY` 构造摘要列表
2. `api_history_restore`（行251-283）：`PLAN_HISTORY[idx]` 按索引取值

场景：用户 A 生成方案（insert+pop），用户 B 同时恢复历史方案 → 索引偏移，恢复到错误方案或 IndexError。

**修复建议**：`api_history` 和 `api_history_restore` 中读 `PLAN_HISTORY` 时也加 `_HISTORY_LOCK`。

---

### Issue #4 · P2 — `api_replenish` 未校验 `profile` 存在性

**文件**: app.py:1185-1193
**类型**: 健壮性

行 1185 只检查 `'plan' not in SESSION`，未检查 `'profile' not in SESSION`。若某种异常路径导致 `plan` 存在但 `profile` 缺失，行 1211 的 `build_plan(snap_profile)` 会因 `snap_profile={}` 导致 KeyError（`profile['score']`）。

**修复建议**：
```python
if 'plan' not in SESSION or 'profile' not in SESSION:
    return jsonify({'error': '无当前方案或考生信息'}), 400
```

---

### Issue #5 · P2 — `api_db_save_plan` 读 SESSION 无锁

**文件**: app.py:1519
**类型**: 线程安全

```python
profile = body.get('profile', SESSION.get('profile', {}))
```

`SESSION.get('profile', {})` 在请求上下文中通过 `_SessionProxy` 读取，而 `_SessionProxy._target()` → `_get_session()` 虽然自身加锁，但返回的 dict 引用后续不受保护。若另一个请求同时修改 `SESSION['profile']`，此处可能读到部分更新的数据。

**修复建议**：用锁+deepcopy 或要求 body 中必须包含 profile（不回退到 SESSION）。

---

### Issue #6 · P2 — `export_excel` 内部函数 `_safe_s` 参数名遮蔽外层循环变量

**文件**: planner.py:1148
**类型**: 代码质量 / 可维护性

```python
for i,v in enumerate(plan_vols):       # 外层循环变量 v
    ...
    def _safe_s(v):                     # 参数名也是 v → 遮蔽
        return int(v) if isinstance(v, (int,float)) and pd.notna(v) else '?'
```

当前逻辑正确（函数内的 `v` 是参数，不是循环变量），但命名冲突极易误导后续维护者。

**修复建议**：重命名参数为 `val`：
```python
def _safe_s(val): return int(val) if isinstance(val, (int,float)) and pd.notna(val) else '?'
```

---

### Issue #7 · P2 — `api_optimize` 序列化缺少 `n_target` / `n_cold` / `all_majors_count` 字段

**文件**: app.py:601-637
**类型**: 数据完整性

对比 `api_generate`（行407-450）和 `api_optimize_constrained`（行724-764）的 vols_out 序列化，`api_optimize` 的输出缺少：
- `n_target`（generate 行 427、constrained 行 752）
- `n_cold`（generate 行 428、constrained 行 753）
- `all_majors_count`（generate 行 438、constrained 行 751）

前端如果依赖这些字段（如显示"命中目标专业数"），优化后会显示 undefined。

**修复建议**：在 `api_optimize` 的 vols_out 序列化中补齐：
```python
'n_target':         v.get('n_target', 0),
'n_cold':           v.get('n_cold', 0),
'all_majors_count': len(v.get('majors', [])),
```

---

### Issue #8 · P3 — `_cached_stats` 使用 f-string 拼接 SQL 表名

**文件**: db.py:384
**类型**: 代码规范

```python
for table in ['province', 'school', ...]:
    cur.execute(f"SELECT COUNT(*) FROM {table}")
```

表名来自硬编码列表，无实际注入风险，但违反"不拼接 SQL"的最佳实践。未来若表名来源变化，可能引入注入漏洞。

---

### Issue #9 · P3 — `build_tiqian` 不过滤公私性质

**文件**: planner.py:678-682
**类型**: 业务逻辑

`build_plan` 过滤 `df['公私性质'] == '公办'`（行149），但 `build_tiqian` 未做此过滤。提前批结果可能包含民办院校，对于以公办为目标的考生可能造成困惑。

如果这是有意设计（提前批含军校/警校不分公私），建议在返回数据中加 `pub_priv` 字段以便前端区分。

---

## 三、累计修复跟踪（第1轮至第7轮）

| 轮次 | 修复数 | 关键修复项 |
|------|--------|-----------|
| 1→2 | 5 | select_subjects 传参、KeyError 防护、SESSION 多用户隔离、路径遍历修复 |
| 2→3 | 3 | 内存泄漏 LRU 淘汰、optimize 候选池重建、NaN 检查 |
| 3→4 | 3 | simulate/optimize SESSION 加锁+deepcopy、双缓存消除、import re 提升 |
| 4→5 | 2 | 前端 export 错误处理、api_chat/map_data 加锁 |
| 5→6 | 4 | api_save_plan 加锁、api_tiqian 选科、db.py SQL 扩列、RLock 升级 |
| 6→7 | 5 | api_replenish 梯度全修、export_excel 加锁、sheet1 NaN、warn 字段补全 |

---

## 四、总结

| 优先级 | 数量 | 概述 |
|--------|------|------|
| P1 | 3 | `api_chat` mc 竞争（#1）、`api_status` 无锁（#2）、`PLAN_HISTORY` 竞争（#3） |
| P2 | 4 | `api_replenish` profile 校验（#4）、`api_db_save_plan` 无锁（#5）、`_safe_s` 命名（#6）、`api_optimize` 字段缺失（#7） |
| P3 | 2 | SQL 拼接规范（#8）、提前批公私过滤（#9） |

**整体评价**：第6轮的全部5个遗留问题均已修复，代码质量持续提升。本轮新问题主要集中在线程安全的"最后一公里"——部分 SESSION/PLAN_HISTORY 读操作仍未被锁覆盖。核心算法（`build_plan`、`mc_simulate`、`optimize_plan`）逻辑稳健，无新的计算错误。
