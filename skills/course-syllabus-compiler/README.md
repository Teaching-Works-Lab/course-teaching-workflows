# course-syllabus-compiler Web/Core

这是一个**不依赖 Router、不自动切换模型**的完整课程大纲 Skill 包，内含编制规则、学校版式合同、空白主文档、确定性 Word renderer、Markdown/DOCX 校验器、测试和示例。

## 上传方式

ZIP 解压后的顶层目录为：

```text
course-syllabus-compiler/
```

可以直接上传这个 ZIP，或把顶层目录覆盖到仓库：

```text
course-teaching-workflows/skills/course-syllabus-compiler/
```

## 两种运行环境

### 普通 Web 对话

普通 Web 对话如果没有文件系统和 Python，只能读取 `SKILL.md`、`references/` 和 Markdown 示例来编写/审查大纲。它不能真正执行 Word renderer，也不能可靠声称已经检查 WPS 兼容性。

### DeepSeek Harness Web / Terminal / Codex

只要宿主提供文件系统和 Python，同一个模型直接运行：

```bash
pip install -r requirements.txt
python scripts/syllabus.py check --input examples/25JD31401-artificial-intelligence.md
python scripts/syllabus.py render --input examples/25JD31401-artificial-intelligence.md --output out.docx
python scripts/syllabus.py inspect --input out.docx
```

不需要 `router.yaml`、模型注册表或另一个模型。

## 包内能力

- 直接编制完整课程大纲初稿；
- 九学院正式全称校验；
- 教材 ISBN 末尾无句号、教材不附课程章节范围；
- 近5年大陆正式出版参考书优先与可核验规则；
- 目标—教学—考核闭环；
- 12×8 学校课程基本信息表；
- 理论、实践、实验、上机章节按课程实际动态生成；
- 动态考核表和评分标准；
- Word/WPS 字体、twip 列宽、批注/修订/隐藏文字检查；
- `check / render / inspect` 单入口 CLI。

## 重要说明

`examples/25JD32408-intelligent-equipment.md` 是已审阅的内容与版式示例，但任何课程代码、学分、学时和学期都必须在实际使用时重新与最新培养方案核对，不能因为示例存在就当作当前官方事实。
