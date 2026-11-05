# workbuddy-skills

WorkBuddy 用户级技能集合（User-level skills for WorkBuddy）。

本仓库用于沉淀、版本化、共享在 WorkBuddy 中可复用的技能。每个技能放在独立的子目录下，结构遵循 WorkBuddy 技能规范（`SKILL.md` + 可选 `scripts/` + `references/`）。

## 已收录技能

| 技能 | 路径 | 作用 |
| --- | --- | --- |
| **docx-annotate** | [`skills/docx-annotate/`](skills/docx-annotate/) | 往 Word(.docx) 注入标准 Word 批注(comments)并对问题内容 run 级标红，输出新文档。适用于「评审 docx / 红字标出问题 / 加批注反馈 / 新文档输出」场景。 |

## 本地安装

WorkBuddy 用户级技能目录为 `~/.workbuddy/skills/`。把需要的技能目录复制过去即可：

```bash
cp -r skills/docx-annotate "$HOME/.workbuddy/skills/docx-annotate"
```

## docx-annotate 快速用法

```bash
python3 skills/docx-annotate/scripts/annotate_docx.py \
  --src 待评审.docx \
  --out 评审意见.docx \
  --findings findings.json \
  --author "你的署名"
```

`findings.json` 示例：

```json
[
  {
    "tokens": ["登录密码（Hs@agv++）"],
    "comment": "【数据安全】明文密码应脱敏。",
    "exact_tokens": []
  },
  {
    "tokens": ["备用"],
    "comment": "【冗余】存在多条“备用”空行，应删除或注明预留用途。",
    "exact_tokens": ["备用"]
  }
]
```

- `tokens`：用于定位与标红的子串（命中即标红并挂批注）。
- `exact_tokens`：其中的 token 仅匹配表格单元精确相等值，避免子串在正文误命中。
- `--keep-source`（默认开）：保留源文档自带批注并自动避让 id 碰撞；关闭则剥离源批注。

> 详细约束与五步法见 [`skills/docx-annotate/SKILL.md`](skills/docx-annotate/SKILL.md) 与 [`skills/docx-annotate/references/playbook.md`](skills/docx-annotate/references/playbook.md)。

## 贡献

新增技能时，请在 `skills/<skill-name>/` 下放置 `SKILL.md`，并更新上方「已收录技能」表格。保持技能可独立运行、附带用法与验收标准。
