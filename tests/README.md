# JXLS 迁移工具测试用例

本目录包含用于测试 JXLS 迁移工具的示例文件。

## 📁 测试文件说明

### 1. asin_top_tag.xls
**描述**: 单 Sheet 测试文件
**用途**: 测试基本的 JXLS 1.x → 2.x 迁移功能
**特点**:
- 包含 forEach 循环指令
- 简单的数据结构
- 适合验证基础迁移功能

**迁移命令**:
```bash
# 从项目根目录运行
python jxls_migration_tool.py tests/asin_top_tag.xls -f -o tests/asin_top_tag_output.xlsx --keep-extension --verbose
```

### 2. tax_contract_multi_sheets_export.xlsx
**描述**: 多 Sheet 测试文件
**用途**: 测试复杂的多工作表 Excel 文件迁移
**特点**:
- 包含 4 个工作表：SKU维度、PO维度、合同维度、发票明细
- 每个 Sheet 都有独立的 forEach 循环
- 适合验证多 Sheet 处理能力

**迁移命令**:
```bash
# 从项目根目录运行
python jxls_migration_tool.py tests/tax_contract_multi_sheets_export.xlsx -f -o tests/tax_contract_output.xlsx --keep-extension --verbose
```

## 🛠️ 迁移工具使用说明

### 基本语法
```bash
python jxls_migration_tool.py <输入文件或目录> [选项]
```

### 常用选项

| 选项 | 说明 | 示例 |
|------|------|------|
| `-f` | 迁移单个文件（而不是目录） | `-f` |
| `-o, --output` | 指定输出文件/目录路径 | `-o output/` |
| `--keep-extension` | 保持原文件后缀名 | `--keep-extension` |
| `--dry-run` | 试运行（不实际修改文件） | `--dry-run` |
| `--verbose` | 详细日志输出 | `--verbose` |
| `-h, --help` | 显示帮助信息 | `-h` |

### 完整示例

#### 示例 1: 单文件迁移
```bash
python jxls_migration_tool.py tests/asin_top_tag.xls -f -o tests/output.xlsx --keep-extension --verbose
```

#### 示例 2: 目录批量迁移
```bash
# 迁移整个目录，保持扩展名
python jxls_migration_tool.py exceltemplate_backup/ -o exceltemplate/ --keep-extension --verbose

# 迁移整个目录，转换为 .xlsx
python jxls_migration_tool.py exceltemplate_backup/ -o exceltemplate/ --verbose
```

#### 示例 3: 试运行（预览）
```bash
# 预览更改，不实际修改文件
python jxls_migration_tool.py tests/ -f -o tests/output/ --dry-run --verbose
```

## ✅ 验证迁移结果

### 查看输出
迁移成功后，工具会显示：
```
✅ 迁移成功: 输出文件路径
🔧 发现 X 个命令，转换 Y 个
```

### 手动验证
1. **打开输出文件** - 使用 Excel 或 WPS 打开生成的 `.xlsx` 文件
2. **检查注释** - 查看 A1 和 A2 单元格是否包含正确的 JXLS 2.x 注释
3. **验证表达式** - 确认 `<jx:...>` 标签已转换为 `${...}` 表达式

### 自动化验证脚本
可以使用 Python 脚本验证结果：

```python
from openpyxl import load_workbook

# 打开输出文件
wb = load_workbook('tests/output.xlsx')
ws = wb.active

# 检查 A1 注释
if ws['A1'].comment:
    print(f"A1 注释: {ws['A1'].comment.text}")

# 检查 A2 注释
if ws['A2'].comment:
    print(f"A2 注释: {ws['A2'].comment.text}")
```

## 🔍 常见问题

### Q: 如果迁移失败怎么办？
A: 使用 `--verbose` 选项查看详细错误信息，或检查日志文件。

### Q: 注释位置不对怎么办？
A: v3.4.1 已修复智能注释位置问题。注释现在会正确地出现在数据行的第一个有数据的单元格中。

### Q: 支持哪些文件格式？
A: 支持 `.xls` 和 `.xlsx` 格式。工具会自动检测真实格式，即使扩展名与实际格式不匹配。

### Q: 能否批量迁移？
A: 可以。直接指定目录路径，工具会自动处理目录下的所有 Excel 文件。

## 📋 测试检查清单

- [ ] 迁移成功无错误
- [ ] JXLS 指令正确转换
- [ ] A1 单元格有 jx:area 注释
- [ ] A2 单元格有 jx:each 注释
- [ ] 表达式格式正确 `${...}`
- [ ] 原有格式保留
- [ ] 合并单元格保留
- [ ] 列宽行高保留

## 📞 技术支持

如果遇到问题，请：
1. 使用 `--verbose` 选项获取详细日志
2. 查看 `jxls_migration.log` 文件
3. 检查 JXLS 迁移工具的 CHANGELOG.md 了解最新修复

---

**当前版本**: v3.4.1 (Smart Comment Position)
**更新日期**: 2025-11-07
