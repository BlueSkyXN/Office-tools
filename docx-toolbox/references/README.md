# references

`docx-toolbox` 的参考脚本目录。

当前纳入：
- `docx-allinone.py`：Excel 嵌入对象处理、去水印、A3、批处理。
- `DOCX图片分离.py`：图片提取、编号标记、附图 PDF。
- `DOCX表格提取.py`：表格提取与 TXT/XLSX/PDF 导出。

说明：
- 本目录作为实现 `core` 的行为参考，不建议直接在 GUI 层调用脚本入口。
- GUI 子项目应通过统一的 `core` 适配层调用处理能力。
