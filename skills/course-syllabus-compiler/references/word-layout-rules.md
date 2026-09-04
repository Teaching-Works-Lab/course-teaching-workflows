# Word 版式与结构规则

机器版式依据见 `school-layout-contract.yaml`，兼容要求见 `wps-docx-contract.md`。

稳定结构包括：A4、学校字体层级、12×8 基本信息表、学时二级表头、课程简介内嵌、课程目标三列表、按课程学时类型动态生成教学章节、动态课程评价表、参考书目和签批。

Markdown 表格列数不得直接决定 Word 表格列数。renderer 负责将语义内容映射到学校物理布局。最终运行：

```bash
python scripts/syllabus.py inspect --input <syllabus.docx>
```
