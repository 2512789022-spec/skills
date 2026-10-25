# Skills

个人自用的 Agent skills 合集，遵循通用的 `SKILL.md` 约定（YAML frontmatter + `references/` + `scripts/` 分层），任何支持该约定的 Agent 都能直接加载，不绑定具体客户端。

## 目录

| Skill | 说明 |
|---|---|
| [`electronic-circuit-lab-report`](electronic-circuit-lab-report/SKILL.md) | 生成、排版与修订中文《电子线路综合设计实践》等电子线路/模拟电路实验报告（.docx），并导出 PDF。固定版式：字体字号、1.5 倍行距、首行缩进、图号与图题、三线表、Word 原生公式、封面与页码 |

## 目录约定

```
skills/
└── electronic-circuit-lab-report/
    ├── SKILL.md                  # 触发条件、交付契约、格式速查、实施流程、坑位清单
    ├── references/
    │   ├── format-spec.md        # 实测出的精确版式参数（含 XML 片段）
    │   ├── content-guide.md      # 各章内容要求、数值一致性、扣分点
    │   └── build-recipes.md      # python-docx 代码骨架
    └── scripts/
        ├── inspect_reference.py  # 提取任意 docx 的版式指纹
        └── verify_report.py      # 成品格式自检（字号/缩进/行距/图号/三线表/公式）
```

## 安装

```bash
git clone https://github.com/2512789022-spec/skills.git

# 把 skill 目录放进你所使用 Agent 的 skills 目录即可
cp -r skills/electronic-circuit-lab-report "<你的 skills 目录>/"
```

触发靠 `SKILL.md` 的 `description` 字段：对话里提到「实验报告」「实验报告排版」「统一实验报告格式」时，Agent 会自动加载。

## 依赖

- Python：`python-docx`、`lxml`、`latex2mathml`
- 导出 PDF：Windows + Microsoft Office（调 Word COM 的 `ExportAsFixedFormat`）
- 公式转 OMML 需要 Office 自带的 `MML2OMML.XSL`（`C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL`）

## 用法

```bash
# 1) 有参考稿：先取版式指纹
python electronic-circuit-lab-report/scripts/inspect_reference.py 参考.docx

# 2) 生成后自检
python electronic-circuit-lab-report/scripts/verify_report.py 成品.docx
```

## License

[MIT](LICENSE)
