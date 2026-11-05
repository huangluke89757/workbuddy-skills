# docx 批注注入五步法（跨任务 Playbook / 改进手册）

> 适用场景：任何需要往 `.docx` 注入 **Word 批注(comments)** 并对问题内容 **标红** 的场景（评审 / 校对 / 修订）。
> 来源：AMHS 人机交互文档评审任务，经 v2→v3→v4 三轮收敛。

## 五步法

### 1) 解析 document.xml（用 lxml，不用 python-docx 段落对象）
- 用 `zipfile` 读 `word/document.xml`，`lxml.etree.fromstring` 解析。
- **不要**用 `python-docx` 的段落对象去插入批注标记——其段落对象存在别名/缓存问题，会导致同一个 `commentRangeStart` id 错挂到无关段落（v2 实测：id=0 同时挂在「目的」标题与真实目标段）。lxml 直写最稳。

### 2) 扫源 commentRangeStart 的 id 范围，定 base
- 在 `document.xml` 与 `word/comments.xml` 中扫描所有 `w:id`。
- 保留源批注 → `base = max(existing_ids) + 1`；无既有 id → `base = 0`。
- 不保留源批注 → 先剥离全部 `commentRangeStart / commentRangeEnd / commentReference`，再从 `base = 0` 起注入。
- **禁止**直接覆盖 `word/comments.xml` 而不处理源标记：会造成 id 碰撞、批注错挂或孤儿 id（Word 中该批注不显示）。

### 3) 定位目标（避开 TOC / PAGEREF，表格精确匹配）
- 段落：全文 `substr in get_text(p)`，但跳过 `pStyle` 以 `TOC` 开头的目录段落、以及含 `PAGEREF` 的段落（否则子串会误命中目录）。
- 表格单元：当目标 token 在正文与表格都可能出现（如「备用」空行），用 `require_table=True, exact=True`（单元文本 `strip()` 后精确相等）区分，避免正文误命中、漏标表格。

### 4) 标红（run 级区间，不整段染）
- 遍历段落的 `w:r`，累加每个 run 的文本偏移量。
- 仅给与命中子串**重叠**的 run 加 `w:color=C00000`（在 `w:rPr` 内插入/修改 `w:color`）。
- 找不到子串（错位）时退化为整段标红并告警——但应先检查匹配逻辑是否命中了 TOC。

### 5) 批注标记 + 文件组装
- 段落首插 `w:commentRangeStart`，尾插 `w:commentRangeEnd` + `w:commentReference`（reference 包在 `w:r` 内）。
- `word/comments.xml`：保留源则读取源并 `SubElement` 追加新 `w:comment`；否则从 0 重建。新批注 id 从 `base` 起，`author/date/initials` 齐备。
- 补 `word/_rels/document.xml.rels` 的 comments 关系（Type 以 `/comments` 结尾，Target `comments.xml`）与 `[Content_Types].xml` 的 `/word/comments.xml` Override——**源文档无批注时必须补，否则 Word 打开报错**。
- 重打包 zip：复制全部 part，替换 `document.xml / comments.xml / rels / content types`。

## 反例边界
- **源文档无批注**：`base=-1` 实际取 0，直接从 0 起；必须新建 `comments.xml` 并补 rels + content types。
- **目标 token 正文与表格均存在**：务必用 `exact + require_table` 区分，否则 `substr in txt` 会先命中正文段落，表格单元漏标。
- **同一 token 出现在多段**：属正常（如交叉引用出现在多处），每段都会标红并挂同一 id；但若 token 也出现在目录，需 TOC 跳过逻辑兜底。

## 验收清单
- `python-docx` 能打开输出文件（无损坏）。
- `commentRangeStart` 数 == `comments.xml` 中 comment 数（无孤儿 id）。
- 标红 run 数 > 0 且精确落在命中子串。
- 保留源批注时：源批注全部仍可用、id 无冲突。
