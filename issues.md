# 代码审查与测试报告

审查日期：2026-03-15
测试方式：自动化运行时测试 + 静态代码分析

---

## 测试结果（15通过 / 3失败）

| 测试项 | 结果 |
|--------|------|
| engine.planner 导入 | ✅ |
| build_plan 各分数段（300/400/500/550/600/700，物理/历史）| ✅ |
| 铁律二：top6 分数降序 | ✅ |
| sc6 稳/保区全局降序 | ✅ |
| mc_simulate（N=2000）| ✅ |
| optimize_plan（3轮）| ✅ |
| export_excel | ✅ |
| exclude_northeast 过滤 | ✅ |
| BUG-1 ke_lei 无效值不报错 | ❌ |
| BUG-2 search_groups 正则注入导致500 | ❌ |
| BUG-3 低分候选不足无警告 | ❌ |

---

## Bug

### BUG-1 · `ke_lei` 传入无效值时静默返回空方案，无错误提示
**位置**：`engine/planner.py` 第 99 行 / `app.py` 第 174 行

`build_plan({'score': 550, 'ke_lei': '数学'})` 不抛异常，因科类过滤后数据为空，返回 0 个志愿，前端无任何错误提示。用户看到空白方案，不知原因。

**修复建议**：在 `build_plan` 入口添加校验：
```python
if ke_lei not in ('物理', '历史'):
    raise ValueError(f"科类必须为'物理'或'历史'，收到: {ke_lei!r}")
```

---

### BUG-2 · `search_groups` API 未禁用正则，特殊字符导致 500 错误
**位置**：`app.py` 第 838 行

```python
sub[sub['院校名称'].str.contains(q, na=False)]
```

`q` 为用户输入，`str.contains` 默认按正则处理。用户搜索 `[工程`（含未闭合方括号）时抛 `PatternError`，返回 500。

**修复建议**：加 `regex=False`：
```python
sub[sub['院校名称'].str.contains(q, na=False, regex=False)]
```

---

### BUG-3 · 分数极低时候选不足，无用户可见警告
**位置**：`app.py` `validate_plan()` 函数

300 分仅返回 4 个志愿，但 `validate_plan` 不检查志愿总数，API 的 `warnings` 字段为空，前端无提示。

**修复建议**：在 `validate_plan` 中添加：
```python
if len(plan_vols) < 10:
    warnings.append(f"⚠️ 当前条件下仅找到 {len(plan_vols)} 个候选志愿，建议放宽筛选条件")
```

---

## 改进建议

### IMPROVE-1 · README 标题版本号未更新（v2.0 → v3.4）
**位置**：`README.md` 第 1 行

`# 吉林省高考志愿规划系统 · 本地版 v2.0`

其余文件（`app.py`、`yijian_qidong.bat`）均已标注 v3.4，README 标题未同步。

---

### IMPROVE-2 · MC 仿真为纯 Python 循环，大 N 时性能差
**位置**：`engine/planner.py` `mc_simulate()`

N=10000 约需 5-10 秒，N=100000 约 50-100 秒，UI 会卡顿。建议用 numpy 向量化重写内层循环，可提速 10-50 倍。

---

### IMPROVE-3 · 缺少自动化测试套件
核心函数 `build_plan / mc_simulate / optimize_plan` 均无单元测试，建议新建 `tests/test_planner.py` 覆盖边界条件（极端分数、空候选池、铁律断言等）。

---

### IMPROVE-4 · 日志使用 print，建议引入 logging 框架
全局用 `print()` 输出，无级别、无时间戳、无文件落地，生产环境调试困难。

---

### IMPROVE-5 · app.py 过长（915行），建议按功能拆分
所有路由集中在一个文件，建议拆分为 Flask Blueprint：`routes/plan.py`、`routes/mc.py`、`routes/chat.py`。
