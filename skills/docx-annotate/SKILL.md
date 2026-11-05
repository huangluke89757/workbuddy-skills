---
name: docx-annotate
description: 往 Word(.docx) 注入标准 Word 批注(comments)并对问题内容 run 级标红，输出新文档。当用户要求「评审 docx / 红字标出问题 / 加批注反馈 / 新文档输出」时使用。源文档可能自带历史批注，本技能会自动避让 id 碰撞并保留源批注。
---

# docx-annotate · 带批注与标红的评审文档生成

## 触发条件
- 用户要求「评审某 docx」「把有问题的内容红字标出」「以批注形式提反馈意见」「新文档形式输出」。
- 典型输入：一份待评审的 docx 路径 + 一份「问题清单」（每个问题 = 命中子串 → 批注文本）。

## 核心约束（务必先读）
1. **开工先扫源批注 id 范围（铁律）**：源文档可能自带 Word 批注。先用脚本扫 `word/document.xml` 的 `commentRangeStart` 的 `w:id` 最大值：保留源批注则新 id 从 `max+1` 起；不保留则先剥离全部 `commentRangeStart/End/Reference` 再从 0 起。**禁止**直接覆盖 `word/comments.xml` 而不处理源标记（会 id 碰撞、错挂、孤儿）。
2. **标红用 run 级区间，不要整段染**：遍历段落 run 算累计偏移，只给与命中子串重叠的 run 设 `w:color=C00000`。
3. **避开 TOC / PAGEREF**：目录段落（`pStyle` 以 TOC 开头）和含 `PAGEREF` 的段落不参与匹配，否则会误命中。
4. **表格单元用精确值匹配**：如「备用」这类空行，必须用 `require_table=True, exact=True` 区分正文与表格，否则子串会在正文误命中、漏标表格。

## 输入
- `--src`：源 docx 绝对路径。
- `--out`：输出 docx 绝对路径（新文件，不覆盖源）。
- `--findings`：问题清单 JSON（见下）。
- `--author`（可选，默认 `"评审"`）：批注作者名。
- `--keep-source`（默认开启）：保留源文档自带批注；关闭则剥离源批注。

`findings.json` 格式：
```json
[
  {
    "tokens": ["收货员：对入库物资进行接运、验收、登记及办理入库手续的系统流程"],
    "comment": "【角色定义错误】收货员是岗位角色，不应定义为“系统流程”。应改为：负责…的人员。",
    "exact_tokens": []
  },
  {
    "tokens": ["备用"],
    "comment": "【冗余】存在多条“备用”空行，应删除或注明预留用途。",
    "exact_tokens": ["备用"]
  }
]
```
- `tokens`：用于定位与标红的子串（命中即在该段落/单元标红并挂批注）。
- `comment`：批注正文，固定格式 `【问题类别】一句话问题 + 应/建议…可执行修正`。
- `exact_tokens`：其中的 token 仅匹配表格单元精确相等值（避免子串在正文误命中）。

## 流程（六步）
1. lxml 解析 `word/document.xml`（**不要**用 python-docx 的段落对象去注入批注——对象别名会导致同 id 错挂到无关段落）。
2. 扫描已有 `commentRangeStart` 的 id 最大值 `base`（保留源则 `base=max+1`；剥离则 `base=0`）。
3. 对每个问题：`find_p` 定位段落/表格单元（表格精确值用 `require_table=True, exact=True`），run 级标红，段落首插 `commentRangeStart`、尾插 `commentRangeEnd + commentReference`。
4. 合并或新建 `word/comments.xml`：保留源则读取源并追加，否则从 0 重建；新批注 id 从 `base` 起。
5. 注入/确认 `word/_rels/document.xml.rels` 的 comments 关系与 `[Content_Types].xml` 的 comments override（源无批注时必须补）。
6. 重打包 zip（复制全部 part，替换 document.xml / comments.xml / rels / content types）。

## 批注文本写法约定（P2）
每条固定格式：`【问题类别】`一句话指出问题 + `应/建议…`给出可执行修正。
类别从 {角色定义错误 / 数据安全 / 内容错误 / 方向错误·安全隐患 / 交互反馈缺失 / 逻辑矛盾 / 校验规则存疑 / 完整性 / 命名混淆} 中选取。

## 执行脚本
`scripts/annotate_docx.py` —— 将上述流程落地为 CLI。用法见脚本内 docstring。

## 验收标准
- `python-docx` 可正常打开输出文件（无损坏）。
- 批注数 == 标记数（`commentRangeStart` 数量 == `comments.xml` 中 comment 数量），无孤儿 id。
- 标红 run 数 > 0，且标红位置精确落在命中子串上（非整段）。
- 若保留源批注：源批注全部仍可用、id 无冲突。

## 跨任务 playbook
完整五步法（含反例边界）见 `references/playbook.md`（改进手册）。
