---
name: hr-interview
version: 2.1.0
description: "10 分钟对话，生成一份可直接开会用的面试打分表——含能力模型、行为题、情景题、STAR 追问和 1-5 分评分标准，最终输出本地 Excel 文件（.xlsx）。数据不出本机，不依赖任何在线服务。当用户需要设计面试题、生成面试评分表、构建岗位能力模型、做面试官培训素材时使用。"
triggers:
  - "帮我设计面试题"
  - "给我出几道面试问题"
  - "帮我做个面试评分表"
  - "我要面试一个[岗位名称]"
  - "帮我构建[岗位]的能力模型"
  - "生成 Excel 打分表"
  - "面试官打分表"
  - "STAR 追问"
  - "帮我准备面试"
  - "设计行为面试题"
  - "设计情景面试题"
metadata:
  requires:
    bins: ["python"]
  safe:
    - "所有数据仅保存在本地，不上传任何服务器"
    - "生成文件前会确认保存路径，不会静默覆盖已有文件"
    - "不会自动执行任何网络请求"
---

# HR 面试题设计工作台

## 概述

这是一个五步引导式工作流，通过对话收集信息、逐步生成内容，最终产出一份本地 Excel 打分表（.xlsx）。整个过程由 AI 主导，HR 在每步确认或修改内容后继续推进。不依赖飞书，不需要 API Key，文件直接保存到本地。

---

## 工作流

### 第一步：收集岗位信息

向用户说明流程，然后请用户提供以下信息（可以一次性提供，也可以逐项回答）：

```
请提供以下岗位信息：
1. 岗位名称（如：高级产品经理、Java 后端工程师）
2. 所在部门（如：增长事业部、技术平台团队）
3. 业务背景（这个岗位的业务场景是什么？可以粘贴现有 JD，也可以用几句话描述）
```

收集完毕后，进入第二步。

---

### 第二步：生成岗位价值分析

基于用户提供的信息，分析并输出：

**格式：**
```
【岗位价值分析】

▌ 核心问题
这个岗位主要解决：[具体描述该岗位解决的 1-2 个核心业务问题]

▌ 创造价值
该岗位为公司/团队创造的主要价值：
• [价值点 1]
• [价值点 2]
• [价值点 3（可选）]

▌ 成功画像
在这个岗位上表现优秀的人，通常具备：[1-2 句简洁描述]
```

输出后询问用户：「以上分析是否准确？有需要调整的地方吗？确认后我们进入能力建模。」

用户确认或修改后，记录最终版本，进入第三步。

---

### 第三步：构建能力模型

基于岗位价值分析，提炼能力模型，输出格式如下：

```
【能力模型】

◆ 核心能力（3-6 项，可胜任这个岗位的关键能力）
1. [能力名称]：[一句话说明这项能力在该岗位中的体现]
2. [能力名称]：[说明]
3. [能力名称]：[说明]
...

◆ 关键特质（2-3 项，价值观或行为风格层面的要求）
1. [特质名称]：[说明]
2. [特质名称]：[说明]
```

输出后询问：「能力模型是否符合您对这个岗位的判断？可以增加、删减或修改任何一项。确认后进入面试题设计。」

记录用户确认的最终能力模型，进入第四步。

---

### 第四步：生成面试题与评分标准

对每一项能力和特质，分别生成一道行为面试题 + 一道情景面试题，每道题附带详细的 1-5 分评分标准。

**单题输出格式：**

```
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【维度】[能力/特质名称]
【题型】行为面试题（过去经历）

题目：
[完整的面试问题]

评分标准：
⭐ 1 分（不达标）：[具体描述，可观察的行为表现]
⭐⭐ 2 分（基本达标）：[具体描述]
⭐⭐⭐ 3 分（达标）：[具体描述]
⭐⭐⭐⭐ 4 分（良好）：[具体描述]
⭐⭐⭐⭐⭐ 5 分（优秀）：[具体描述，往往含有量化结果或显著影响]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
```

评分标准设计原则：
- 每个分值的描述必须具体、可观察，避免"表现良好"等模糊表述
- 区分维度：从"没有经历/无法回答"→"有经历但被动"→"主动推动"→"有系统方法"→"有显著量化结果或跨组织影响"
- 1分和5分的描述要有明显的行为差异

全部题目输出后询问：「面试题和评分标准是否符合预期？可以修改任何一道题或评分标准。确认后进行 STAR 追问拆解。」

---

### 第五步：STAR 追问拆解

对每道面试题，生成 STAR 框架的追问子题：

```
【题目】[原始问题]

STAR 追问：
▸ S（情境）："当时的背景是什么？团队规模/项目阶段是怎样的？"
▸ T（任务）："您具体负责哪个部分？目标是什么？"
▸ A（行动）："您具体做了哪些事情？为什么选择这个方案？"
▸ R（结果）："最终结果怎么样？有没有可量化的数据？"
▸ 追加：[1-2 个针对该题目特点的深挖问题，如："如果重来一次，您会做什么不同的决定？"]
```

全部追问生成后，询问用户：

```
STAR 追问已全部生成。

是否现在生成 Excel 打分表？
• 输入「是」或「生成」→ 立即生成并保存到本地
• 输入「否」→ 结束，内容已在对话中，可随时复制使用
```

---

### 输出阶段：生成本地 Excel 打分表

用户确认生成后，执行以下步骤。

#### 1. 询问保存路径（可选）

询问用户：「Excel 文件保存到哪里？直接回车使用默认路径（桌面）。」

- 若用户指定路径，使用该路径
- 若用户直接回车或未填写，默认保存到桌面：
  - Windows：`%USERPROFILE%\Desktop\<岗位名称>-面试打分表-<日期>.xlsx`
  - Mac：`~/Desktop/<岗位名称>-面试打分表-<日期>.xlsx`

#### 2. 生成 Python 脚本并执行

将对话中收集到的所有数据（岗位信息、能力模型、面试题、STAR追问、评分标准）整理成结构化数据，生成一个临时 Python 脚本，使用 openpyxl 创建 Excel 文件。

**脚本结构：**

```python
# -*- coding: utf-8 -*-
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import date

# ── 样式辅助函数 ──
def hfill(h): return PatternFill("solid", fgColor=h)
def thin_border():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)
def wa(h="left", v="top"): return Alignment(wrap_text=True, horizontal=h, vertical=v)
def hdr(ws, row, cols, texts, bg="1F4E79", fg="FFFFFF"):
    for col, text in zip(cols, texts):
        c = ws.cell(row=row, column=col, value=text)
        c.fill = hfill(bg); c.font = Font(bold=True, size=10, color=fg)
        c.alignment = wa("center","center"); c.border = thin_border()
def cell(ws, row, col, val, bg=None, bold=False, h="left"):
    c = ws.cell(row=row, column=col, value=val)
    if bg: c.fill = hfill(bg)
    c.font = Font(bold=bold, size=10)
    c.alignment = wa(h, "top"); c.border = thin_border()
    return c

# ── 数据（由 AI 根据对话内容填入）──
JOB_TITLE = "<岗位名称>"
JOB_DEPT  = "<部门>"
JOB_DATE  = str(date.today())

ABILITIES = [
    ("C1", "<能力1名称>", "<能力1说明>"),
    # ...
]
TRAITS = [
    ("T1", "<特质1名称>", "<特质1说明>"),
    # ...
]

# 每道题包含：id, dim(维度), dt(核心能力/关键特质), tp(行为题/情景题),
#             q(题目), s1-s5(评分标准), sS/sT/sA/sR/sX(STAR追问)
QUESTIONS = [
    dict(id="Q01", dim="<维度>", dt="核心能力", tp="行为题",
         q="<题目>",
         s1="<1分>", s2="<2分>", s3="<3分>", s4="<4分>", s5="<5分>",
         sS="<S追问>", sT="<T追问>", sA="<A追问>", sR="<R追问>", sX="<深挖追问>"),
    # ...
]

# ── 创建工作簿（Sheet1封面 / Sheet2能力模型 / Sheet3面试题库 / Sheet4打分表）──
wb = openpyxl.Workbook()

# Sheet1: 封面
ws1 = wb.active; ws1.title = "封面"
ws1.column_dimensions["A"].width = 16; ws1.column_dimensions["B"].width = 62
ws1.merge_cells("A1:B1")
c = ws1["A1"]; c.value = "HR 面试题设计工作台"
c.fill = hfill("1F4E79"); c.font = Font(bold=True, size=16, color="FFFFFF")
c.alignment = Alignment(horizontal="center", vertical="center"); ws1.row_dimensions[1].height = 42
ws1.merge_cells("A2:B2")
c = ws1["A2"]; c.value = f"岗位：{JOB_TITLE}  |  部门：{JOB_DEPT}  |  生成日期：{JOB_DATE}"
c.fill = hfill("D6E4F0"); c.font = Font(size=11, color="1F4E79")
c.alignment = Alignment(horizontal="center", vertical="center"); ws1.row_dimensions[2].height = 24
for i, (lbl, val) in enumerate([
    ("岗位名称", JOB_TITLE), ("所在部门", JOB_DEPT), ("生成日期", JOB_DATE),
    ("能力维度数", f"{len(ABILITIES)} 项核心能力 + {len(TRAITS)} 项关键特质"),
    ("面试题总数", f"{len(QUESTIONS)} 道"),
    ("使用说明", "本表包含4个工作表：封面、能力模型、面试题库、打分表。\n面试时使用【打分表】为候选人评分，每位候选人建议单独复制一份使用。"),
], start=4):
    cell(ws1, i, 1, lbl, bg="1F4E79", bold=True, h="center")
    ws1.cell(i, 1).font = Font(bold=True, size=10, color="FFFFFF")
    cell(ws1, i, 2, val, bg="F2F2F2" if i % 2 == 0 else "FFFFFF")
    ws1.row_dimensions[i].height = 40 if i == 9 else 24

# Sheet2: 能力模型
ws2 = wb.create_sheet("能力模型")
for col, w in zip("ABCD", [12, 8, 18, 58]): ws2.column_dimensions[col].width = w
hdr(ws2, 1, range(1,5), ["类型","编号","名称","说明"]); ws2.row_dimensions[1].height = 24
row = 2
for code, name, desc in ABILITIES:
    for c_i, (val, bg, bd) in enumerate([(
        "核心能力","D6E4F0",True),("","D6E4F0",False),(name,"D6E4F0",True),(desc,"D6E4F0",False)], 1):
        cell(ws2, row, c_i, ["核心能力",code,name,desc][c_i-1], bg="D6E4F0",
             bold=(c_i in [1,3]), h="center" if c_i<=2 else "left")
    ws2.row_dimensions[row].height = 42; row += 1
for code, name, desc in TRAITS:
    for c_i in range(1,5):
        cell(ws2, row, c_i, ["关键特质",code,name,desc][c_i-1], bg="E2EFDA",
             bold=(c_i in [1,3]), h="center" if c_i<=2 else "left")
    ws2.row_dimensions[row].height = 42; row += 1
ws2.freeze_panes = "A2"

# Sheet3: 面试题库
ws3 = wb.create_sheet("面试题库")
for i, w in enumerate([6,14,8,40,28,28,28,28,34,24,24,24,24,24], 1):
    ws3.column_dimensions[get_column_letter(i)].width = w
hdr(ws3, 1, range(1,15), ["编号","考察维度","题型","面试题目",
    "S-情境追问","T-任务追问","A-行动追问","R-结果追问","深挖追问",
    "1分（不达标）","2分（基本达标）","3分（达标）","4分（良好）","5分（优秀）"])
ws3.row_dimensions[1].height = 24
for i, q in enumerate(QUESTIONS, start=2):
    bg = "D6E4F0" if q["dt"] == "核心能力" else "E2EFDA"
    for col, val in enumerate([q["id"],q["dim"],q["tp"],q["q"],
            q["sS"],q["sT"],q["sA"],q["sR"],q["sX"],
            q["s1"],q["s2"],q["s3"],q["s4"],q["s5"]], 1):
        cell(ws3, i, col, val, bg=(bg if col <= 3 else None))
    ws3.row_dimensions[i].height = 72
ws3.freeze_panes = "D2"

# Sheet4: 打分表
ws4 = wb.create_sheet("打分表")
for col, w in zip("ABCDEF", [7, 44, 14, 10, 38, 16]): ws4.column_dimensions[col].width = w
for col, lbl in enumerate(["候选人姓名：","","面试日期：","","面试官：",""], 1):
    c = ws4.cell(row=1, column=col, value=lbl)
    c.fill = hfill("F2F2F2"); c.font = Font(bold=True, size=10)
    c.alignment = Alignment(horizontal="left", vertical="center"); c.border = thin_border()
ws4.row_dimensions[1].height = 26; ws4.row_dimensions[2].height = 6
hdr(ws4, 3, range(1,7), ["编号","面试题目（简版）","考察维度","评分（1-5）","关键行为记录","备注"])
ws4.row_dimensions[3].height = 24
for i, q in enumerate(QUESTIONS, start=4):
    short_q = (q["q"][:40] + "...") if len(q["q"]) > 40 else q["q"]
    bg = "D6E4F0" if q["dt"] == "核心能力" else "E2EFDA"
    cell(ws4, i, 1, q["id"], bg=bg, h="center")
    cell(ws4, i, 2, short_q)
    cell(ws4, i, 3, q["dim"], bg=bg, h="center")
    c = ws4.cell(row=i, column=4, value="")
    c.fill = hfill("FCE4D6"); c.alignment = Alignment(horizontal="center", vertical="center"); c.border = thin_border()
    cell(ws4, i, 5, ""); cell(ws4, i, 6, "")
    ws4.row_dimensions[i].height = 44
total_row = len(QUESTIONS) + 4
for col in range(1, 7):
    c = ws4.cell(row=total_row, column=col)
    c.fill = hfill("FFF2CC"); c.font = Font(bold=True, size=10)
    c.alignment = Alignment(horizontal="center", vertical="center"); c.border = thin_border()
ws4.cell(total_row, 2).value = "总  分"
ws4.cell(total_row, 4).value = f"=SUM(D4:D{total_row-1})"
ws4.row_dimensions[total_row].height = 26
eval_row = total_row + 1
ws4.merge_cells(f"B{eval_row}:F{eval_row}")
for col in range(1, 7):
    c = ws4.cell(row=eval_row, column=col)
    c.fill = hfill("F2F2F2"); c.font = Font(bold=True, size=10)
    c.alignment = wa("left", "center"); c.border = thin_border()
ws4.cell(eval_row, 2).value = "综合评价与录用建议："
ws4.row_dimensions[eval_row].height = 54
ws4.freeze_panes = "A4"

# ── 保存 ──
output_path = r"<用户指定路径或默认桌面路径>"
wb.save(output_path)
print(f"OK: {output_path}")
```

**执行方式：**

将脚本保存为临时文件后执行：

```bash
python -X utf8 /tmp/gen_interview_sheet.py
```

Windows 下路径使用 `%TEMP%\gen_interview_sheet.py`，执行命令：

```bash
python -X utf8 "%TEMP%\gen_interview_sheet.py"
```

#### 3. 执行完成后输出

```
✅ Excel 打分表已生成！

📁 文件路径：<实际保存路径>

包含 4 个工作表：
• 封面 — 岗位信息概览
• 能力模型 — X 项核心能力 + X 项关键特质
• 面试题库 — X 道面试题（含 STAR 追问 + 1-5 分评分标准）
• 打分表 — 面试官用评分表（橙色评分列，总分自动汇总）

使用建议：每次面试新候选人前，复制一份【打分表】工作表单独使用，避免覆盖原始模板。
```

---

## 前置依赖

运行本 Skill 需要 Python 环境安装 openpyxl：

```bash
pip install openpyxl
```

---

## 失败处理

执行 Python 脚本前，先做环境检查，任何一步失败都要给出明确提示和替代方案：

### 检查 Python 是否可用

```bash
python --version
```

- ✅ 输出版本号 → 继续
- ❌ 报错 `command not found` 或 `'python' is not recognized` →
  停止执行，告知用户：
  > 「未检测到 Python 环境。请先安装 Python（https://python.org/downloads），安装完成后重新运行。如果不方便安装，我可以把面试题内容直接以 Markdown 格式输出，你复制后手动填入 Excel。」

### 检查 openpyxl 是否已安装

```bash
python -c "import openpyxl"
```

- ✅ 无报错 → 继续
- ❌ 报错 `ModuleNotFoundError` →
  自动执行 `pip install openpyxl`，若安装失败，告知用户：
  > 「openpyxl 安装失败，可能是网络或权限问题。请手动运行 `pip install openpyxl` 后重试。」

### 检查保存路径

- 若路径包含中文或空格，自动加引号
- 若同名文件已存在，先询问用户：
  > 「桌面已有同名文件，是否覆盖？输入「是」覆盖，输入「否」另存为新名称。」
- 若目标目录不存在，自动创建或提示用户

### Excel 生成失败

若脚本执行报错，将完整错误信息展示给用户，并提供降级方案：

> 「Excel 生成失败，错误信息如下：[错误内容]。
> 降级方案：以下是所有面试题的 Markdown 格式，可复制使用：[输出所有题目和评分标准]」

---

## 注意事项

- **路径中的特殊字符**：Windows 路径包含中文或空格时，用引号包裹
- **文件名冲突**：若同名文件已存在，生成前询问用户确认，不静默覆盖
- **打分表复用**：每位候选人单独复制一份打分表工作表填写，原始模板保持空白
- **脚本清理**：执行完成后可删除临时脚本文件
- **数据安全**：所有内容仅在本地处理，不上传任何平台

---

## 示例产物

`examples/` 目录包含以「高级产品经理」为样例岗位的真实生成文件：

- `examples/高级产品经理-面试打分表-样例.xlsx`：4 个工作表，7 道面试题，含完整 STAR 追问和评分标准
- `examples/gen_sample.py`：生成上述文件的脚本，可直接运行复现
