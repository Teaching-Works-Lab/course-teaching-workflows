---
name: course-syllabus-compiler
description: Use when a university course syllabus must be planned, drafted, revised, checked, compiled to Word, or inspected from training-program facts, textbooks, course materials, and school formatting requirements.
---

# Course Syllabus Compiler

直接生成可逐字人工审核的**完整初稿**。上游字段写“待编制”时，本 Skill 负责完成编制；不得把内部证据状态、AI 说明、候选或建议措辞带入正式大纲。

## 环境边界

- 普通 Web 对话没有文件系统或 Python：完成 Markdown 编制/审查，但不得声称已生成或检查 DOCX。
- DeepSeek Harness Web、Codex、Terminal Agent 等具备文件系统和 Python：由当前模型直接运行 `scripts/syllabus.py`，不依赖 Router，也不自动切换模型。
- Markdown 只表达语义；Word 的列数、合并、字号、列宽和分页由确定性 renderer 与 `references/school-layout-contract.yaml` 决定。

## 编制规则

1. 从培养方案、课程矩阵和用户材料提取课程代码、学分、学时、学期、性质、学院和课程关系。硬事实冲突时先报告冲突，不静默替换。
2. 根据教材范围、课程序列和专业能力直接编制课程简介、知识/能力/素质目标、达成方法、教学单元、思政融入、预期学习成果、教学方式和课程评价。
3. 课程目标—教学单元—考核必须闭环；各自独立的权重集合均合计 100%。
4. 教材与学习资料按 `references/bibliography-contract.md` 编制；书目信息无法核验时少列，不得编造。
5. 授课学院只能使用 `references/school-format-contract.md` 中的九个正式全称。

## 必读资源

- `references/authoring-contract.md`
- `references/school-format-contract.md`
- `references/assessment-contract.md`
- `references/bibliography-contract.md`
- `references/wps-docx-contract.md`
- `references/school-layout-contract.yaml`
- `references/web-core-workflow.md`

## 构建与检查

```bash
python scripts/syllabus.py check --input syllabus.md
python scripts/syllabus.py render --input syllabus.md --output syllabus.docx
python scripts/syllabus.py inspect --input syllabus.docx
```

Word 输出必须使用 `assets/syllabus-master.docx` 的页面基座和学校 renderer。最终 DOCX 不得含批注、修订、隐藏文字、未解决占位符、模型署名或工具元数据。renderer/master 变化、WPS 显示异常或定期回归时，渲染全部页面做视觉检查。

## 示例使用边界

`examples/` 展示正文深度和版式能力。示例课程事实不能复制到新课程；每次必须重新读取最新培养方案。
