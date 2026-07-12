---
name: chip-arch-diagram
description: 根据 Excel 网格表 + Markdown 连接描述生成芯片架构图 SVG。用户填好模板后对 AI 说「生成架构图」即可。
---

# 芯片架构图生成器

根据 Excel 网格表 + Markdown 连接描述 → 生成架构图 SVG（单个 HTML 文件）。

## 使用方法（给用户看）

1. 复制 `input/架构图描述输入模板.md` 填好模块、连接、颜色、父框
2. 复制 `input/架构网格图模板.xlsx` 在 Excel 中画网格
3. 对 AI 说：「生成架构图」，附上 .md 和 .xlsx 路径
4. 浏览器打开 `output/架构图.html`

---

# AI 执行规范

> **核心原则：一次写对。** 下面的 Gate 是硬性检查点，任一失败必须修正后继续，不得跳过。

## 参考文档（强制查阅）

执行过程中，**每个步骤必须查阅对应 rules/ 文件**获取完整算法。以下为各步骤与参考文档的对应关系：

| 步骤 | 查阅文件 | 查阅章节 |
|------|---------|---------|
| §3 解析网格、模块分类 | `rules/通用布局规则.md` | §一 网格坐标、§二 模块分类、§三 外围模块、§四 行分配 |
| §4 解析父框 | `rules/通用布局规则.md` | §五 父框（坐标格式、子模块匹配、外框与外部模块、标签防遮挡） |
| §6 计算像素坐标 | `rules/通用布局规则.md` | §十 统一对齐（像素坐标公式）、§十一 推导流程 |
| §7 颜色与框样式 | `rules/通用布局规则.md` | §六 框样式、§八 颜色定义（映射表、渐变填充、边框色规则） |
| §8 连线计算 | `rules/通用连线规则.md` | 全文（§一~§四 Phase 1/2/3 L&Z形路由、k值分配） |
| §9 组装 HTML | `rules/通用布局规则.md` | §6.2 Z-order、§5.3 标签防遮挡、§9.1 模块标签垂直居中 |
| §10 自检 | `rules/通用检查清单.md` | 全文（A-K 全部检查项） |

> **规则**：SKILL.md 给出关键公式和 Gate 检查点。完整算法细节、边界条件、推导过程均在 `rules/` 中。遇到不确定时，以 `rules/` 为准。

---

## 0. 准备

1. `pip install openpyxl` 确保已安装
2. 在输入 `.md` 所在目录创建 `output/` 子目录
3. 写一个 Python 脚本完成全部计算，输出到 `output/架构图.html`
4. 脚本运行结束打印自检结果，全部 ✓ 才算成功

**禁止**：分步执行、写入中间 JSON 文件、多次运行 Python 逐个修 bug。一次脚本，一次运行。

### 段落提取规则（强制）

从 `.md` 中提取各段落时，regex 必须用**内容关键词**匹配标题，**禁止**用章节编号（`## 一、` `## 二、` 等）。编号随模板版本变化，关键词不变：

| 段落 | 匹配关键词 |
|------|-----------|
| 外围模块声明 | `外围模块声明` |
| 父框样式 | `父框样式` |
| 父框标注 | `父框标注` |
| 连接关系 | `连接关系` |
| 模块颜色 | `模块颜色` |
| 尺寸参数 | `尺寸参数` |
| 外框名称 | `外框名称` |

通用 regex 模板：`re.search(r'关键词.*?\n\s*```\s*\n(.*?)\n\s*```', md, re.DOTALL)`

---

## 1. 数据结构定义

脚本内部使用以下结构（Python dict / list），内存中计算，不落盘。

```python
# 模块
module = {
    "name": "MOD_A",            # Excel 中的简称
    "c0": 1, "c1": 1,          # 列范围（0-indexed）
    "r0": 0, "r1": 0,          # 行范围（0-indexed）
    "is_peri": False,           # 来自 §二 外围模块声明的显式声明
    "peri_side": None,          # 外围方向："left"/"right"/"top"/"bottom"，内部为 None
    "row_span": 1,              # 行跨度 = r1-r0+1（内部模块可 >1）
    "col_span": 1,              # 列跨度 = c1-c0+1（内部模块可 >1，Excel 同名相邻格自动合并）
    "x": 190.0, "y": 60.0,     # 像素坐标（左上角）
    "w": 120, "h": 60,         # 像素宽高
}

# 连接
conn = {
    "src": "PIO_A", "tgt": "MOD_A",   # 已映射为 Excel 简称
    "phase": 1,                        # 1/2/3
    "src_edge": "right",               # Phase3 专用
    "tgt_edge": "bottom",              # Phase3 专用
    "line_color": "black",             # 连线颜色（默认 black）
    "line_style": "",                  # 线形：""=实线，"6,4"=虚线，"3,3"=点线
    "line_label": None,                # 连线标签文字（可选）
}

# 父框
pframe = {
    "name": "GRP_A",
    "c0": 1, "c1": 2, "r0": 0, "r1": 0,  # 注意：先 c 后 r
    "children": ["MOD_A", "MOD_B"],
    "rect": {"x": 174, "y": 44, "w": 322, "h": 92},
}

# 外框
oframe = {
    "x": 170, "y": 12, "w": 1164, "h": 452,
    "name": "CHIP_TOP",
}
```

---

## 2. 参数定义

| 变量 | 默认值 | 来源 | 说明 |
|------|--------|------|------|
| LEFT_MARGIN | 20 | 用户输入（尺寸参数段） | |
| TOP_MARGIN | 60 | 用户输入（尺寸参数段） | |
| CELL_W | 120 | 用户输入（尺寸参数段） | 内部模块宽度 |
| CELL_H | 60 | 用户输入（尺寸参数段） | 内部模块高度 |
| PERIPH_W | 60 | 用户输入（尺寸参数段） | 外围模块宽度（左右） |
| COL_GAP | 50 | 用户输入（尺寸参数段） | 列间距 |
| ROW_GAP | 40 | 用户输入（尺寸参数段） | 行间距 |
| FRAME_PAD_INPUT | 12 | 用户输入（尺寸参数段） | 父框边距（最小值） |
| FRAME_PAD | max(FRAME_PAD_INPUT, 16) | 计算 | 必须 ≥16 以容纳 8px 标签 |
| OUTER_PAD_TOP_INPUT | 20 | 用户输入（尺寸参数段） | |
| OUTER_PAD_TOP | max(OUTER_PAD_TOP_INPUT, 32) | 计算 | 硬编码地板 32，不通过 FONT 公式推导 |
| OUTER_PAD_OTHER | 8 | 用户输入（尺寸参数段） | |
| FONT | 9 | 固定 | 模块标签字号 |
| FONT_S | 8 | 固定 | 父框标签字号 |
| FONT_L | 10 | 固定 | 外框标签字号 |

> 用户输入的 FRAME_PAD 和 OUTER_PAD_TOP 是**最小值**。标签空间不足时必须自动增大。最小公式：
> - FRAME_PAD ≥ FONT_S + 8 = 16
> - OUTER_PAD_TOP = max(用户输入, 32) — **硬编码 32，非公式推导。** 即使 FONT_L+FONT_S+10=28，底线仍是 32。

---

## 3. 读取并解析网格（Excel → modules dict）

> 📖 模块分类规则、外围模块定义、行分配原则详见 `rules/通用布局规则.md` §一～§四。

```python
import openpyxl
wb = openpyxl.load_workbook(xlsx_path)
ws = wb[wb.sheetnames[0]]        # 取第一个 sheet，跳过"使用说明"等
grid = {}
for r_1 in range(1, ws.max_row + 1):
    for c_1 in range(1, ws.max_column + 1):
        v = ws.cell(row=r_1, column=c_1).value
        # ⚠️ 空白单元格可能是 None 或纯空格字符串，统一处理
        s = str(v).strip() if v is not None else ""
        grid[(r_1-1, c_1-1)] = s if s else "."
```

**Gate 1 — 打印网格表**（必须打印到 stdout）：
```
       c0        c1        c2   ...
r0:  PIO_A     MOD_A     MOD_B   ...
r1:  PIO_A         .         .   ...
...
```

**模块分类**（不做推断，只读用户显式声明）：
- 从用户输入 **§二 外围模块声明** 中解析外围模块名列表 `peri_names`
- `name in peri_names` → `is_peri = True`
- 其余 → `is_peri = False`
- 同名内部模块在相邻格出现 → **自动合并**为一个多单元格模块（col_span/row_span > 1）

**外围模块交叉校验**（声明 vs 网格位置必须一致）：
- `is_peri=True` 但不在任何边缘（c=0, c=max_col, r=0, r=max_row）→ **报错**
- 位于列边缘（c=0 或 c=max_col）但 `is_peri=False` → **报错**（列边缘强制声明为外围）
- 位于行边缘（r=0 或 r=max_row）但 `is_peri=False` → **允许**（行边缘可同时容纳内部和外围模块）
- > 用户必须显式声明。脚本不靠位置推断、不靠颜色推断——只做一致性校验。外围方向由网格位置决定：单列→left/right，单行→top/bottom。

**Gate 2 — 打印模块清单**：逐个列出 name / is_peri / (r0-r1, c0-c1)。

**辅助属性**：解析后立即计算 `row_span = m["r1"] - m["r0"] + 1`、`col_span = m["c1"] - m["c0"] + 1`、`peri_side`（由网格位置决定，见上）。

### 外围模块解析

从用户输入 **§二 外围模块声明** 中提取：

```python
# ⚠️ 用内容关键词匹配，不用章节编号
peri_section = re.search(r'外围模块声明.*?\n\s*```\s*\n(.*?)\n\s*```', md, re.DOTALL)
peri_names = set()
if peri_section:
    for token in re.split(r'[,\s]+', peri_section.group(1)):
        token = token.strip()
        if token: peri_names.add(token)
```

然后在模块分类时：`m["is_peri"] = (name in peri_names)`，再执行交叉校验。

---

## 4. 解析父框（⚠️ 极易出错）

> 📖 父框坐标格式、子模块匹配、标签防遮挡详见 `rules/通用布局规则.md` §五。

### 格式：`父框名 {c范围, r范围}` —— 先列后行！

正则：`(\w+)\s*\{(\d+)-(\d+),\s*(\d+)(?:-(\d+))?\}`

| 捕获组 | 含义 | 示例 `{1-2, 0}` |
|--------|------|-----------------|
| group 1 | 父框名 | GRP_A |
| group 2 | c_start | 1 |
| group 3 | c_end | 2 |
| group 4 | r_start | 0 |
| group 5（可选）| r_end | 0（单行时=r_start） |

### 子模块匹配：用区间重叠，不用包含

```python
for name, m in modules.items():
    if (m["r0"] <= pf["r1"] and m["r1"] >= pf["r0"] and
        m["c0"] <= pf["c1"] and m["c1"] >= pf["c0"]):
        pf["children"].append(name)
```

> 模块跨行时（如 PIO_A 占 r=0-3，父框只要 r=1-2），重叠判断正确捕获，包含判断会漏掉。

**Gate 3 — 打印父框子模块**：
```
GRP_A {c=1-2, r=0} → [MOD_A, MOD_B]
GRP_B {c=4-5, r=0-2} → [MOD_D, MOD_E, MOD_F, MOD_G]
GRP_C {c=1-2, r=4} → [MOD_H, MOD_I]
```

---

## 5. 解析连接

> 📖 完整路由算法（Phase 1/2/3 详细步骤、k 值分配、交叉检测）见 `rules/通用连线规则.md` 全文。

### 输入格式

```
PIO_A → MOD_A → MOD_B → MOD_C → ...                    # 链式（→ 或 -> 均可）
【L】 MOD_H（右） → MOD_C（底）                              # L 形
```

### 步骤

1. **提取行尾属性**：`[...]` 块中解析 `颜色`/`color`、`线形`/`style`、`描述`/`label`（中英文兼容），不写用默认（黑色/实线/无标签）
2. **展开链式**：`A -> B -> C` → edges `A→B`, `B→C`，链中所有边共享行尾属性
3. **识别 L 形**：正则 `【L】\s*(\w+)\s*[（(]\s*(\S+?)\s*[）)]\s*(?:->|→)\s*(\w+)\s*[（(]\s*(\S+?)\s*[）)]`
   - ⚠️ 箭头必须兼容 `->`（ASCII）和 `→`（Unicode 全角箭头）。只写 `->` 会导致模板中的 `→` 匹配失败，L 形行被链式解析器误吃，报 `NameError`。
   - 出入边标准化：`顶/top→top`, `底/bottom→bottom`, `左/left→left`, `右/right→right`
   - > ⚠️ **陷阱1**：L 形边在构造 dict 时必须**立即**设置 `"type": "L"`。不能依赖后续的合并循环来推断类型——合并循环中无 type 的边会被错误标记为 chain，导致所有 L 形被当作"缺 L 标记"报错丢弃。
   ```python
   # 正确写法：构造时赋值
   l_edges.append({..., "type": "L"})
   ```
4. **去重**：若同一对 (src, tgt) 同时出现在链式和 L 形中，L 形覆盖链式（L 形提供了出入边信息）
5. **名称映射**：md 全称 → Excel 简称。只做确定性匹配，禁止猜测：
   - (a) 直接匹配：md_name 在 module_names 中 → 返回
   - (b) 前缀剥离 + 精确匹配：依次尝试 `PROJ_A_`, `SYS_`, `SUBSYS_` 等前缀（由调用方传入，见附录 C），剥离后的 stripped 必须在 module_names 中精确命中 → 返回。**不精确命中不返回**，继续试下一个前缀
   - (c) 以上全部失败 → **报错停止**，打印 `ERROR: 无法映射 '{md_name}' → 可用模块名: {sorted(module_names)}`。禁止子串匹配、禁止模糊匹配、禁止相似度猜测。对不上就是用户输入有问题，让用户修正。
      - > ⚠️ **陷阱3**：禁止一切猜测。`MODULE_ALPHA_FULL` 在 Excel 中写成 `MODULE_ALPHA_FUL`（少了 L）——精确匹配 + 前缀剥离全失败 → 报错。用户自己修正拼写。颜色缩写 `SUBMOD` vs Excel 全名 `SYS_SUBMOD`——对不上 → 报错。用户要么把颜色区的缩写改成全名，要么在 Excel 里用缩写。
6. **分类**：按 `e["type"]` 字段直接判断——`"L"` → Phase 3，否则按模块位置分 Phase 1/2。

| 条件 | Phase |
|------|-------|
| L 形标记 | 3 |
| 一端外围 | 1 |
| 两端内部 + (行有重叠 或 列有重叠) | 2 |
| 两端内部 + 行列均无重叠 + 无 L | **报错** |
| 两端外围 | **报错** |

> "行/列有重叠"判断用区间交叠（如模块 A 占 r0，B 占 r0-r1，有重叠→Phase 2 直连）。
> Phase 3 自动判断：异向边（右→底）→ **L 形**（2 段）；同向边（右→右）→ **Z 形**（3 段）。
> k 值分配：Phase 3 遍历所有 k 找最短路径，Phase 2 填剩余位置。

**Gate 4 — 打印连接分类**：
```
Total edges: 12
  P1 (4): PIO_A→MOD_A, MOD_G→PIO_B, ...
  P2 (5): MOD_A→MOD_B, ...
  P3 (3): MOD_H(right)→MOD_C(bottom), ...
  Self-check: P1+P2+P3 == total ✓
```

---

## 6. 计算像素坐标

> 📖 完整像素坐标公式、对齐约束、推导流程详见 `rules/通用布局规则.md` §九、§十。

### 6.1 列 X 坐标

```python
col_x = [0.0] * grid_cols
col_x[0] = LEFT_MARGIN + PERIPH_W
for c in range(1, grid_cols):
    prev_is_peri = (c - 1 == 0 or c - 1 == grid_cols - 1)
    prev_w = PERIPH_W if prev_is_peri else CELL_W
    col_x[c] = col_x[c - 1] + prev_w + COL_GAP
```

### 6.2 行 Y 坐标

```python
# 检测顶部外围 → 网格下移
has_top = any(m["is_peri"] and m.get("peri_side") == "top" for m in modules.values())
grid_top_y = TOP_MARGIN + CELL_H + ROW_GAP if has_top else TOP_MARGIN

row_y = [0.0] * grid_rows
row_y[0] = grid_top_y
for r in range(1, grid_rows):
    row_y[r] = row_y[r - 1] + CELL_H + ROW_GAP
```

### 6.3 模块像素坐标

```python
# 内部模块（支持多单元格：col_span/row_span 可 >1）
m["x"] = col_x[m["c0"]]
m["y"] = row_y[m["r0"]]
m["w"] = m["col_span"] * CELL_W + (m["col_span"] - 1) * COL_GAP
m["h"] = m["row_span"] * CELL_H + (m["row_span"] - 1) * ROW_GAP

# 外围模块 — 按 peri_side 区分形状
if m["is_peri"]:
    side = m.get("peri_side")
    if side in ("left", "right"):
        m["y"] = row_y[m["r0"]]
        m["w"] = PERIPH_W
        m["h"] = m["row_span"] * (CELL_H + ROW_GAP) - ROW_GAP
    elif side == "top":
        m["y"] = TOP_MARGIN
        m["w"] = m["col_span"] * (CELL_W + COL_GAP) - COL_GAP
        m["h"] = CELL_H
    else:  # bottom
        m["y"] = row_y[-1] + CELL_H + ROW_GAP
        m["w"] = m["col_span"] * (CELL_W + COL_GAP) - COL_GAP
        m["h"] = CELL_H
```

> 左右外围为竖条（高度公式确保相邻不重叠）；上下外围为横条（宽度公式同理）。

派生属性：`left=x, right=x+w, top=y, bottom=y+h, cx=x+w/2, cy=y+h/2`。

### 6.4 父框坐标

```python
children = pf["children"]
min_x = min(modules[c]["left"] for c in children)
min_y = min(modules[c]["top"] for c in children)
max_x = max(modules[c]["right"] for c in children)
max_y = max(modules[c]["bottom"] for c in children)
pf["rect"] = {
    "x": min_x - FRAME_PAD,
    "y": min_y - FRAME_PAD,
    "w": max_x - min_x + 2 * FRAME_PAD,
    "h": max_y - min_y + 2 * FRAME_PAD,
}
```

### 6.5 外框坐标

外框包裹所有**内部模块**（不含外围） + 所有**父框**：

```python
internals = [m for m in modules.values() if not m["is_peri"]]
omin_x = min(m["left"] for m in internals)
omin_y = min(m["top"] for m in internals)
omax_x = max(m["right"] for m in internals)
omax_y = max(m["bottom"] for m in internals)

for pf in pframes:
    if pf["rect"]:
        omin_x = min(omin_x, pf["rect"]["x"])
        omin_y = min(omin_y, pf["rect"]["y"])
        omax_x = max(omax_x, pf["rect"]["x"] + pf["rect"]["w"])
        omax_y = max(omax_y, pf["rect"]["y"] + pf["rect"]["h"])

oframe = {
    "x": omin_x - OUTER_PAD_OTHER,
    "y": omin_y - OUTER_PAD_TOP,
    "w": omax_x - omin_x + 2 * OUTER_PAD_OTHER,
    "h": omax_y - omin_y + OUTER_PAD_TOP + OUTER_PAD_OTHER,
    "name": outer_frame_name_from_input,
}
```

**Gate 5 — 打印坐标摘要**：
```
col_x: [80, 190, 360, 530, 700]
row_y: [60, 140, 220, 300]
Modules: MOD_A(190,60 120x60) | PIO_A(80,60 60x300) | ...
Outer: (170,12) 1164x452
```

---

## 7. 颜色与框样式

> 📖 完整颜色映射表、渐变填充格式、边框色选择规则详见 `rules/通用布局规则.md` §八。

### 7.0 父框与外框样式

从「父框样式」段落解析（`rules/通用布局规则.md` §六）。格式：`名称: stroke=色值, fill=色值, style=solid|dashed|dotted`。

- 外框使用 `外框`/`outer` 作为名称，映射到内部 key `__outer__`
- 默认值：外框和父框均为 `stroke=#555, fill=none, style=solid`（深灰色实线，无填充）
- style 映射：`solid`→无 dasharray, `dashed`→"6,4", `dotted`→"3,3"
- Z-order：外框在最底层，父框在外框上一层，模块在父框之上

### 7.1 颜色表（唯一真相源）

| 用户颜色名 | fill | stroke |
|-----------|------|--------|
| 灰色 | #C0C0C0 | #A0A0A0 |
| 白色 | #FFFFFF | #D0D0D0 |
| 浅蓝 | #BDD7EE | #8DB4E2 |
| 深蓝 | #4472C4 | #2F5597 |
| 浅绿 | #C5E0B4 | #A9D18E |
| 深绿 | #548235 | #375623 |
| 浅紫 | #D9C2EC | #B994D4 |
| 深紫 | #7030A0 | #4E2170 |
| 浅黄 | #FFF2CC | #D4C88C |
| 橘黄 | #FFC000 | #D4A000 |
| 浅橙 | #FCD5B4 | #F4B183 |
| 浅红 | #FF9999 | #E07878 |
| 浅粉 | #F2DCDB | #E6B8B7 |
| 浅青 | #B4D6CD | #8DBDB0 |

> ⚠️ 浅蓝 fill 是 `#BDD7EE`（非 `#ADD8E6`）。所有 14 色均配备 stroke，纯色/渐变均可使用。

### 7.2 颜色分配

从用户输入中搜索「模块颜色」标题所在的段落解析。所有 section 提取 regex 必须用**内容关键词**（如 `外围模块声明`、`连接关系`、`模块颜色`、`尺寸参数`、`外框名称`），禁止用 `## 一、` `## 二、` 等章节编号——编号随模板版本变化。**颜色行正则**：`(\S+?)\s*[：:]\s*(.+)`——`\s*` 是必须的，因为 CJK 输入习惯在全角冒号前加空格（如 `浅黄|浅蓝 ： MOD_A`）。

> ⚠️ **陷阱2**：正则写成 `(\S+?)[：:](.+)`（缺 `\s*`）会导致带空格的颜色行匹配失败，该行被当作"续行"跳过，渐变模块全部回退为默认浅蓝。Gate 不报错（颜色"看起来 OK"），但渐变丢失。

`左色|右色` 表示渐变（快切：0%/48% 左色，48%/52% 过渡，52%/100% 右色）。边框取右端的 stroke（全部 14 色均配备 stroke，不再有无 stroke 的退化情况）。

**多行处理**：颜色行匹配到新颜色名→更新 current_color；未匹配到但 current_color 非空→视为上一颜色的续行追加模块。两者都不满足→跳过。

**Token 预处理**：分割后的每个 token，先剥离中文/英文括号注释 `（...）` / `(...)` 再映射。例如 `APB（外部模块）` → `APB`。

未指定的内部模块默认浅蓝。外围模块强制灰色。

### 7.3 渐变 SVG

```xml
<linearGradient id="g_MODNAME" x1="0" y1="0" x2="1" y2="0">
  <stop offset="0%" stop-color="左色fill"/>
  <stop offset="48%" stop-color="左色fill"/>
  <stop offset="52%" stop-color="右色fill"/>
  <stop offset="100%" stop-color="右色fill"/>
</linearGradient>
```

---

## 8. 连线计算

> 📖 **执行前必须查阅 `rules/通用连线规则.md` 全文。** 以下为关键逻辑摘要和 Gate 检查点，完整算法（正交约束、不穿框验证、Phase 1/2/3 路由细节、k 值分配、交叉检测）在 rules 中。

### 8.0 边连接数统计（先于所有连线）

遍历 Phase 1+2+3 全部连接，按 `(模块名, 边)` 汇总：

```python
edge_conns = {}  # (module_name, edge) -> [conn_info, ...]

# Phase 1: 内部模块边 — 从 peri_side 取方向，不再硬编码 left/right
for conn in phase1:
    peri = modules[外围的那一端]
    int_mod = modules[内部的那一端]
    edge = peri["peri_side"]  # "left"/"right"/"top"/"bottom"
    edge_conns.setdefault((int_mod["name"], edge), []).append({"phase": 1, "conn": conn})

# Phase 2: 同行→左右边，同列→上下边
for conn in phase2:
    sm, tm = modules[conn["src"]], modules[conn["tgt"]]
    # 同行：src 在左→(right, left)，src 在右→(left, right)
    if sm["r0"] == tm["r0"]:
        if sm["c0"] < tm["c0"]: se, te = "right", "left"
        else:                    se, te = "left", "right"
    # 同列：src 在上→(bottom, top)，src 在下→(top, bottom)
    else:
        if sm["r0"] < tm["r0"]: se, te = "bottom", "top"
        else:                    se, te = "top", "bottom"
    edge_conns.setdefault((conn["src"], se), []).append({"phase": 2, "conn": conn, "k": 1})
    edge_conns.setdefault((conn["tgt"], te), []).append({"phase": 2, "conn": conn, "k": 1})

# Phase 3: 用户指定的出入边
for conn in phase3:
    edge_conns.setdefault((conn["src"], conn["src_edge"]), []).append({"phase": 3, "conn": conn})
    edge_conns.setdefault((conn["tgt"], conn["tgt_edge"]), []).append({"phase": 3, "conn": conn})
```

### 8.1 k 值分配（更近原则）

> 📖 完整分配算法见 `rules/通用连线规则.md` §4.2。

对每条 `(模块, 边)`：
1. C = 连接数, N = C + 1
2. Phase 2 固定 k = 1
3. Phase 3：遍历 k=1..N-1，对端 k=1 为参考，选距离最短的 k
4. Phase 2：按到对端模块中心距离排序，填剩余 k
5. Phase 1：同 Phase 2 逻辑，按距离排序

**Gate 6 — 打印边统计**：
```
MOD_D:bottom C=2 N=3
  P2 k=1 MOD_D->MOD_F
  P3 k=2 MOD_G->MOD_D
```

### 8.2 出入点公式

```python
def edge_point(mod, edge, k, N):
    if edge == "left":   return (mod["left"],  mod["top"] + k * mod["h"] / N)
    if edge == "right":  return (mod["right"], mod["top"] + k * mod["h"] / N)
    if edge == "top":    return (mod["left"] + k * mod["w"] / N, mod["top"])
    if edge == "bottom": return (mod["left"] + k * mod["w"] / N, mod["bottom"])
```

### 8.3 Phase 1 — 内部↔外围

> 📖 详见 `rules/通用连线规则.md` §二。

- 内部模块等分点 + 外围投影对齐
- 1 段直线 `<line>`
- 方向：数据流 src → tgt
- 连线不得穿过任何非源/目标模块内部（`rules/通用连线规则.md` §1.2）

### 8.4 Phase 2 — 直连

> 📖 详见 `rules/通用连线规则.md` §三。

- 同行：水平线。同列：垂直线。
- k=1，N_unified = max(N_src_edge, N_tgt_edge)。**两端必须用同一个 N**，否则 y 坐标（水平线）或 x 坐标（垂直线）不对齐 → 产生斜线。
- 1 段直线 `<line>`

### 8.5 Phase 3 — L 形（2段）/ Z 形（3段）

> 📖 详见 `rules/通用连线规则.md` §四（含交叉检测 §4.5）。

按源模块 (r, c) 排序处理。

**异向边 → L 形（2 段）**：
```python
N_unified = max(N_src, N_tgt)
src_pt = edge_point(src_mod, src_edge, k_src, N_unified)
tgt_pt = edge_point(tgt_mod, tgt_edge, k_tgt, N_unified)

if src_edge in ("left", "right"):
    corner = (tgt_pt[0], src_pt[1])    # 首段水平
else:
    corner = (src_pt[0], tgt_pt[1])    # 首段垂直
points = [src_pt, corner, tgt_pt]
```

**同向边 → Z 形（3 段）**：
```python
# 水平同向（右→右、左→左等）：水平→垂直→水平
mid_x = (src_pt[0] + tgt_pt[0]) / 2.0
points = [src_pt, (mid_x, src_pt[1]), (mid_x, tgt_pt[1]), tgt_pt]

# 垂直同向（底→底、顶→顶等）：垂直→水平→垂直
mid_y = (src_pt[1] + tgt_pt[1]) / 2.0
points = [src_pt, (src_pt[0], mid_y), (tgt_pt[0], mid_y), tgt_pt]
```

**k 值分配**（`allocate_k_values`）：Phase 3 逐一试 k=1..N-1，用对端 k=1 做参考，选路径最短的 k；Phase 2 按距离填剩余位置。

**交叉检测**：每条路由后，检测新路径各段与已路由段正交交叉 + 不穿框检测。冲突时报告用户。

**Gate 7 — 打印 P3 详情**：
```
MOD_E(bottom)->MOD_I(left): (1100,120)->(1100,250)->(1210,250)
MOD_M(right)->MOD_C(left): (480,315)->(505,315)->(505,95)->(530,95)
```

---

## 9. 组装 HTML（⚠️ DOM 顺序极其重要）

> 📖 标签定位公式详见 `rules/通用布局规则.md` §5.3、§9.1。

### 强制绘制顺序（SVG 画家算法——后绘制的在上层）

```
① 背景 rect（白色填充整个 viewBox）
② <defs>（渐变 + 箭头 marker）
③ 外框 rect（默认深灰实线、无填充，可通过「父框样式」配置）
④ 父框 rect（默认深灰实线、无填充，可通过「父框样式」配置）—— 仅矩形，不含标签！前面必须加 `<!-- ④ 父框 rect -->` 注释标记
⑤ 模块 rect + 模块 text（按 name 排序）—— 前面必须加 `<!-- ⑤ 模块 rect + text -->` 注释标记
⑥ 外框标签 text + 父框标签 bg rect + 父框标签 text  ← 必须在模块之后！前面必须加 `<!-- ⑥ 标签 -->` 注释标记
⑦ Phase 1 <line> → Phase 2 <line> → Phase 3 <polyline> —— 前面必须加 `<!-- ⑦ Phase 1 lines -->` 注释标记
```

> 为什么标签在模块之后：SVG 后绘制的覆盖先绘制的。标签画在模块之上才能不被遮挡。白色半透明背景 rect（`fill="white" opacity="0.9"`）覆盖模块内容，确保文字可读。

### 标签公式

**外框标签**（左上角，font-size=FONT_L=10）：
```python
lx = oframe["x"] + 10
ly = oframe["y"] + FONT_L + 2
# 无背景 rect（外框顶部 padding 足够大，没有模块遮挡）
```

**父框标签**（左上角，font-size=FONT_S=8）：
```python
lx = pf["rect"]["x"] + 4
first_top = min(modules[c]["top"] for c in pf["children"])
ly = max(pf["rect"]["y"] + FONT_S + 4, first_top - 4)
tw = len(pf["name"]) * FONT_S * 0.6 + 6  # 估算文字宽度

# 背景 rect（在 text 之前，同一位置）
bg_rect = f'<rect x="{lx-2:.1f}" y="{ly-FONT_S+1:.1f}" width="{tw:.1f}" height="{FONT_S+2:.1f}" fill="white" opacity="0.9"/>'
text_el = f'<text x="{lx:.1f}" y="{ly:.1f}" text-anchor="start" font-size="{FONT_S}" font-family="Arial, sans-serif" fill="#555">{pf["name"]}</text>'
```

**模块标签**（居中）：
```python
text_y = m["y"] + m["h"]/2 + FONT * 0.35
# 名称 > 16 字符 → 双行（在 _ 处断行）
# 双行公式详见 rules/通用布局规则.md §9.1
```

### 连线渲染

每条连线从 conn 中读取 `line_color`（默认 black）、`line_style`（默认 ""=实线）、`line_label`（默认 None）：

```python
color = conn.get("line_color", "black")
style = conn.get("line_style", "")
dash = f' stroke-dasharray="{style}"' if style else ""
# <line ... stroke="{color}" ...{dash} marker-end="url(#ar)"/>
# <polyline ... stroke="{color}" ...{dash} marker-end="url(#ar)"/>
```

标签（`line_label` 非空时）渲染在连线中点上方 8px，带白色背景 rect：
- 水平线→中点上方；垂直线→中点左侧
- P3 L 形→首段中点
- font-size=7, fill=#555, 背景 `fill="white" opacity="0.85"`

### 箭头 marker（每色一个，确保箭头颜色与线一致）

收集所有连线的 `line_color`，去重后每个颜色生成一个 marker：

```xml
<marker id="ar_black" markerWidth="8" markerHeight="6" refX="8" refY="3" orient="auto">
  <polygon points="0 0, 8 3, 0 6" fill="black"/>
</marker>
<marker id="ar_red" ... fill="red"/>
```

连线引用对应颜色的 marker：`marker-end="url(#ar_black)"`。

### viewBox

```python
all_x = [所有模块 left/right + 外框 x + 外框 right]
all_y = [所有模块 top/bottom + 外框 y + 外框 bottom]
vx, vy = min(all_x) - 20, min(all_y) - 20
vw, vh = max(all_x) - vx + 20, max(all_y) - vy + 20
viewBox = f"{vx:.0f} {vy:.0f} {vw:.0f} {vh:.0f}"
```

---

## 10. 自检（脚本末尾打印，全部 ✓ 才算通过）

> 📖 完整检查项（A-I 组共 30+ 项）见 `rules/通用检查清单.md`。以下为自检输出格式和关键检查项。

```
===== SELF-CHECK =====
G1 父框 {c,r} 顺序: 先列后行 ✓
G2 子模块重叠判断: ✓
G3 名称映射完整: ✓
H1 标签 DOM 位置: 模块之后 ✓
H2 标签白底: 3 个 ✓
H3 父框顶边距≥16: FRAME_PAD=16 ✓
H4 外框顶边距≥32: OUTER_PAD_TOP=32 ✓
I1 DOM 顺序: ③外框(at {idx_outer})<④父框(at {idx_frames})<⑤模块(at {idx_mod})<⑥标签(at {idx_label})<⑦连线(at {idx_lines}) ✓
  → 方法：html.find("<!-- ③ 外框 -->") < html.find("<!-- ④ 父框") < html.find("<!-- ⑤ 模块") < html.find("<!-- ⑥") < html.find("<!-- ⑦")
C1 正交性: 所有线段水平或垂直 ✓
C2 不穿框: ✓
C6 连线数: <line>+<polyline>={实际}=预期{22} ✓
A1 外部模块在外框外: PIO_A,PIO_B,PIO_C ✓
B1 颜色: 灰色3/渐变2/浅蓝8/深蓝2 ✓
D1 Phase1 1段直线: ✓
E1 Phase2 k=1: ✓
F1 Phase3 L形/Z形: ✓
F7 交叉检测: 0 conflicts ✓
======================
PASSED ✓
```

> 自检必须包含 `rules/通用检查清单.md` 中的 **全部 A-K 组** 检查项。上面是输出格式示例，实际检查项以清单文件为准。

---

## 11. 最终汇报

脚本成功后只输出三行：
```
Done: output/架构图.html
Modules: internal=10, peripheral=3, frames=2
Connections: P1=4, P2=5, P3=3, total=12
```

---

## 附录 A：完整 SVG 模板

```html
<!DOCTYPE html>
<html lang="zh-CN">
<head><meta charset="UTF-8">
<title>芯片架构图</title>
<style>
body{{margin:0;padding:20px;display:flex;justify-content:center;background:#f5f5f5}}
svg{{max-width:100%;height:auto}}
</style>
</head><body>
<svg xmlns="http://www.w3.org/2000/svg" viewBox="...">
<defs>
  <!-- 渐变定义 -->
  <!-- 箭头 marker -->
</defs>
<rect x="..." fill="white"/>                        <!-- ① 背景 -->
<rect ... stroke="#555" fill="none"/>                   <!-- ③ 外框（默认深灰实线，可配置） -->
<rect ... stroke="#555" fill="none"/>                   <!-- ④ 父框1（默认深灰实线，可配置） -->
<rect ... stroke="#555" fill="none"/>                   <!-- ④ 父框2 -->
<!-- ⑤ 模块 rect + text（按名称排序） -->
<!-- ⑥ 外框标签 + 父框标签 bg rect + 父框标签 text -->
<!-- ⑦ Phase1 <line> -->
<!-- ⑦ Phase2 <line> -->
<!-- ⑦ Phase3 <polyline> -->
</svg>
</body></html>
```

## 附录 B：已知实现陷阱（必读，每条都曾导致真实 Bug）

| # | 陷阱 | 错误写法 | 正确写法 | 后果 |
|---|------|---------|---------|------|
| 1 | L 形边未设 type | `l_edges.append({..., "src_edge":...})` | `l_edges.append({..., "type":"L"})` | 所有 L 形被误判为"缺 L 标记"丢弃，P3=0 |
| 2 | 颜色正则缺 `\s*` | `(\S+?)[：:](.+)` | `(\S+?)\s*[：:]\s*(.+)` | CJK 冒号前的空格导致匹配失败，渐变全部丢失 |
| 3 | 父框 `{r,c}` 顺序 | 假设 `{行,列}` | 必须 `{c, r}` 先列后行 | 父框包裹错误模块 |
| 4 | 子模块包含判断 | `frame.r0 <= m.r0 <= frame.r1` | 重叠判断 `m.r0 <= frame.r1 and m.r1 >= frame.r0` | 跨行模块被漏掉 |
| 5 | SVG 标签在模块前 | 标签 `<text>` 放在模块 `<rect>` 之前 | 标签 bg+text 放在模块 rect+text **之后** | 模块盖住标签，文字不可见 |
| 6 | 框顶边距太小 | FRAME_PAD=12, OPT=20（用户输入原值） | `max(用户值, 16)`, `max(用户值, 32)` | 标签挤在模块边框上，视觉被遮挡 |
| 7 | CSS 大括号被 `.format()` 吞噬 | `body{margin:0}` 嵌入 f-string | `body{{margin:0}}` 双写大括号 | `KeyError: 'margin'`，脚本直接崩溃 |
| 8 | Windows GBK 无法输出 ✓ | `print("✓")` | 第一行加 `sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')` | `UnicodeEncodeError`，自检无法输出 |
| 9 | 名称映射猜测 | LCS 模糊匹配 / 候选排序自动选 | 确定性匹配失败 → 报错停止，列出候选让用户修正输入 | 猜错导致级联故障（连线分类错、颜色错），且用户不知情 |
| 10 | OUTER_PAD_TOP 用公式推导 | `max(input, FONT_L+FONT_S+10)` → 28 | `max(input, 32)` 硬编码地板 | 外框顶边距不足，标签空间不够 |
| 11 | HTML 注释标记遗漏 | 父框 rect 前漏写 `<!-- ④ 父框 rect -->` | 严格按照 §9 DOM 顺序列表，每个层级前加对应注释 | 自检 I1 失败（DOM 顺序检查找不到锚点） |
| 12 | L 形正则只匹配 `->` | `【L】\s*...\s*->\s*...` | `【L】\s*...\s*(?:->|→)\s*...` | 模板中的全角 `→` 匹配失败，L 形行被链式解析器误吃，`【L】MOD_H（右）` 被当作模块名 → `NameError` |

---

## 附录 C：使用 generator.py（推荐）

> `generator.py` 封装了全部解析、计算、路由、自检逻辑和 11 个已知陷阱。优先使用。

```python
# -*- coding: utf-8 -*-
"""芯片架构图 — 最简调用。"""
from generator import ChipArchDiagram

gen = ChipArchDiagram(
    md_path=r"input/描述.md",
    xlsx_path=r"input/网格.xlsx",
    # prefixes=["PROJ_A_", "SYS_"],  # 名称映射前缀（按需）
    # params={"CELL_W": 140, "FRAME_PAD_INPUT": 16},  # 参数覆盖（按需）
)
gen.run()  # 解析→计算→路由→HTML→自检，一步完成
```

输出：`output/架构图.html` + stdout 自检报告。

### 命令行调用

```bash
python generator.py input/描述.md input/网格.xlsx
python generator.py input/描述.md input/网格.xlsx --prefixes PROJ_A_ SYS_
```

### generator.py 内部结构

| 函数/类 | 职责 |
|---------|------|
| `ChipArchDiagram` | 主类，`run()` 一键执行 |
| `generate()` | 便捷函数 |
| `parse_*()` | MD/XLSX 各段落解析 |
| `map_name()` | 确定性名称映射（禁止猜测） |
| `calculate_coords()` | 像素坐标 + 父框 + 外框 |
| `assign_colors()` | 颜色分配（外围强制灰色） |
| `classify_connections()` | Phase 1/2/3 分类 |
| `build_edge_stats()` / `allocate_k_values()` | 边统计 + 更近原则 |
| `route_phase1/2/3()` | 连线路由（含交叉检测） |
| `assemble_html()` | HTML 组装（DOM 顺序 + CSS 转义） |
