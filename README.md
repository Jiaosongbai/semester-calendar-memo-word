# semester-calendar-memo-word

将大学校历的 PDF、图片或纯文字说明整理为极简、可打印、可继续填写的 Word 周历记事本。

本项目是面向支持 Skills 的 ChatGPT/Codex 环境设计的 Skill。AI 负责识别和核对校历信息，随附的 Python 生成器负责根据结构化 JSON 稳定生成 `.docx`。它不是一个脱离 AI 环境、直接把 PDF 自动转换为 Word 的独立桌面程序。

## 功能

- 支持 PDF 校历、校历图片和纯文字说明。
- 根据已提供且可核实的信息生成一个学期或完整学年。
- 识别行课周、考试周、报到注册、法定节假日、校历节日标记、学校寒暑假和校园活动。
- 区分政府正式放假、校历单日标记和学校自行安排的假期。
- 支持秋学期、冬学期、小学期等特殊学期标记。
- 支持用户指定的生日和纪念日；积极日期可生成不超过 20 个汉字的安全祝福，非庆祝性纪念日只客观标注。
- 使用真正的 Word 脚注记录校园活动。
- 表格行高可随用户输入和回车自动增长。
- 输出仅为 `.docx`，暂不包含 Excel、每日视图、课程表、提醒、订阅或移动应用界面。

## 工作流程

1. 用户提供校历 PDF、图片或文字说明，并说明需要生成的范围。
2. Skill 提取并核对日期、周次、考试、假期和活动信息。
3. 信息不足时先向用户询问，不擅自补全。
4. Skill 按照 [`references/spec-schema.md`](references/spec-schema.md) 生成 JSON。
5. [`scripts/generate_notebook.py`](scripts/generate_notebook.py) 输出可编辑 Word 文档。
6. 交付前渲染并逐页检查版式、脚注和边距。

## 文件结构

```text
semester-calendar-memo-word/
├── SKILL.md
├── agents/
│   └── openai.yaml
├── references/
│   ├── spec-example.json
│   └── spec-schema.md
├── regression/
│   ├── xiaohongshu-text-spec.json
│   └── zhejiang-image-spec.json
├── scripts/
│   └── generate_notebook.py
└── tests/
    ├── run_regression.py
    └── run_final_acceptance.py
```

## 环境要求

- Python 3
- `python-docx`
- `lxml`

生成器只接收经过核对的 JSON，不负责自行识别原始 PDF 或图片。

## 使用生成器

在项目根目录运行：

```bash
python scripts/generate_notebook.py references/spec-example.json --output calendar-notebook.docx
```

输出文件名和位置可以自行修改。输入字段与分类规则见 [`references/spec-schema.md`](references/spec-schema.md)。

## 测试

运行基础回归测试：

```bash
python tests/run_regression.py
```

运行发布前的 PDF、图片、纯文字三类验收用例：

```bash
python tests/run_final_acceptance.py
```

测试生成物写入 `tests/output/`，不会进入 Git 仓库。

## 设计原则

- 只生成有可靠信息支持的范围。
- 不把校历标出的节日日期擅自扩展为正式放假区间。
- 不预测尚未正式公布的法定节假日安排。
- 不在副标题中重复“图片识别”“PDF 生成”等无必要信息。
- 个人日期只有在用户明确要求时才写入。
- 保持极简、均衡、可打印和可继续填写。

## 许可证

本项目采用 [MIT License](LICENSE)。
