# Web/Core 工作方式

本版本不依赖外部 Router，也不自动切换模型。

- 普通 Web 对话没有文件系统和 Python 时：使用 Skill 规则生成或审查 Markdown，不能声称已经执行 Word renderer。
- DeepSeek Harness Web、Codex、Terminal Agent 等具备文件系统和 Python 时：同一个模型直接运行 `scripts/syllabus.py check/render/inspect`，不需要二次调用其他模型。
- 视觉仅用于模板/renderer 变化、WPS 显示异常或抽检；学院、学时、ISBN、权重、表格结构和 OOXML 清洁由脚本确定性检查。
