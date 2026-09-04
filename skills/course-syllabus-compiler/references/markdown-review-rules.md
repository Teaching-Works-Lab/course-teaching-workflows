# Markdown 发布检查规则

Markdown 是课程语义源，不是 Word 版式 DSL。发布前至少检查：课程基本事实、九学院白名单、学时合计、章节连续、课程目标编号、教学与考核覆盖、权重合计、教材 ISBN 末尾格式、参考书可核验性，以及终稿中不存在内部状态或 AI/工具署名。

运行：

```bash
python scripts/syllabus.py check --input <syllabus.md>
```
