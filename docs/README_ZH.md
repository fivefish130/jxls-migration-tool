# JXLS 迁移工具

[![版本](https://img.shields.io/badge/version-3.0-blue.svg)](https://github.com/fivefish130/jxls-migration-tool)
[![Python](https://img.shields.io/badge/python-3.6+-green.svg)](https://www.python.org/downloads/)
[![许可证](https://img.shields.io/badge/license-MIT-yellow.svg)](LICENSE)

一个强大的、生产就绪的自动化工具，用于将 JXLS 1.x Excel 模板迁移到 JXLS 2.14.0。

## ✨ 特性

- **完整的 JXLS 指令支持**: 迁移所有 JXLS 指令（forEach、if、out、area、multiSheet）
- **智能文件格式检测**: 自动检测 Excel 文件格式，不依赖扩展名
- **格式保留**: 保持所有单元格样式、列宽、行高、合并单元格
- **Windows Terminal 优化**: 自动检测并配置 UTF-8 支持
- **详细报告**: 生成 Markdown、JSON 和 DEBUG 日志
- **生产就绪**: 全面的错误处理和日志记录

## 📦 安装

### 要求

- **Python**: 3.6 或更高版本（推荐 Python 3.12）
- **依赖**: xlrd==2.0.1, openpyxl>=3.0.0

### 快速安装

```bash
pip install xlrd==2.0.1 openpyxl
```

## 🚀 快速开始

### 迁移整个目录

```bash
# 保持原始文件扩展名（推荐）
python jxls_migration_tool.py input_directory --keep-extension

# 指定输出目录
python jxls_migration_tool.py input_directory -o output_directory --keep-extension

# 试运行（预览更改）
python jxls_migration_tool.py input_directory --dry-run --verbose
```

### 迁移单个文件

```bash
python jxls_migration_tool.py input.xls -f output.xlsx

# 保持扩展名
python jxls_migration_tool.py input.xls -f output.xls --keep-extension
```

## 📋 JXLS 指令转换

| JXLS 1.x | JXLS 2.14.0 | 描述 |
|----------|-------------|------|
| `<jx:forEach items="..." var="...">` | `jx:each(items="..." var="..." lastCell="...")` | 基于注释，自动删除标签行 |
| `<jx:if test="...">` | `jx:if(condition="..." lastCell="...")` | test → condition，基于注释 |
| `<jx:out select="..."/>` | `${...}` | 直接表达式替换 |
| `<jx:area lastCell="...">` | `jx:area(lastCell="...")` | 保留现有或自动生成 |
| `<jx:multiSheet data="...">` | `jx:multiSheet(data="...")` | 完整多工作表支持 |

## 📊 实际应用结果

**迁移统计（923 个 Excel 文件）**:
- ✅ 成功迁移: 50 个 JXLS 模板
- ⏭️  跳过: 873 个文件（无 JXLS 指令）
- ❌ 失败: 0 个文件
- 🎯 成功率: **100%**

**模块分布**:
- Module A: 7 个模板
- Module B: 1 个模板
- Module C: 42 个模板

**指令统计**:
- 发现 JXLS 指令总数: 106
- 成功转换: 106
- 主要类型: forEach (50), area (50), if (6)

## 📚 文档

- **[使用指南](USAGE_ZH.md)** - 详细使用说明
- **[API 文档](API.md)** - 程序化 API 参考
- **[示例](examples/)** - 代码示例和用例
- **[更新日志](CHANGELOG.md)** - 版本历史和变更

## 🛠️ 命令行选项

```
用法: jxls_migration_tool.py [-h] [-o 输出] [-f] [--keep-extension]
                             [--dry-run] [-v] [--verbose]
                             输入路径

位置参数:
  输入路径              输入文件或目录路径

可选参数:
  -h, --help              显示帮助信息并退出
  -o 输出, --output 输出  输出目录（默认：与输入相同）
  -f 输出文件, --file 输出文件
                          输出文件（单文件迁移）
  --keep-extension        保持原始文件扩展名
  --dry-run               预览更改而不修改文件
  -v, --verbose           启用详细日志
```

## 🎯 迁移示例

### 迁移前 (JXLS 1.x)
```
行 1: <jx:area lastCell="E5">
行 2: 采购单列表
行 3: <jx:forEach items="datas" var="item">
行 4: ${item.sku} | ${item.qty} | ${item.price}
行 5: </jx:forEach>
行 6: 总计: ${total}
```

### 迁移后 (JXLS 2.x)
```
行 1: (空白)
       └─ [注释] jx:area(lastCell="E4")
行 2: 采购单列表
行 3: ${item.sku} | ${item.qty} | ${item.price}
       └─ [注释] jx:each(items="datas" var="item" lastCell="C3")
行 4: 总计: ${total}
```

**关键变更**:
- ✅ 删除了标签行（第 3 行和第 5 行）
- ✅ 在第 1 行添加了 jx:area 注释
- ✅ 在第 3 行添加了 jx:each 注释
- ✅ 自动计算了 lastCell="C3"（C 列，第 3 行）

## 🔧 智能文件格式检测

工具通过读取文件头自动检测实际 Excel 文件格式，而不依赖扩展名：

| 文件头 | 格式 | 描述 |
|--------|------|------|
| `D0 CF 11 E0 A1 B1 1A E1` | XLS | OLE2/复合文档 |
| `PK` (50 4B) | XLSX | ZIP 格式 |

这允许：
- 实际是 .xlsx 格式但扩展名为 .xls 的文件被正确处理
- 实际是 .xls 格式但扩展名为 .xlsx 的文件被正确处理
- 自动选择正确的处理器（xlrd vs openpyxl）

## 💡 Windows Terminal 优化

自动检测并优化 Windows Terminal 环境：

```python
# 自动检测这些环境变量
WT_SESSION, WT_PROFILE_ID

# Windows Terminal: 使用原生 UTF-8（无需配置）
# 传统 cmd/PowerShell: 自动设置 chcp 65001
```

## 📄 输出报告

迁移后，生成三种类型的报告：

1. **migration_report.md** - 人类可读的 Markdown 报告
2. **migration_report.json** - 机器可读的 JSON 报告
3. **jxls_migration.log** - 完整的 DEBUG 日志

## ⚠️ 已知限制

1. **varStatus** - JXLS 2.x 不再支持 varStatus（必须在 Java 代码中手动实现）
2. **加密文件** - 无法处理密码保护的 Excel 文件
3. **损坏文件** - 无法处理损坏的 Excel 文件
4. **特殊格式** - 极少数特殊格式可能需要手动调整

## 🆘 故障排除

### Q: 提示缺少 xlrd 库
```bash
pip install xlrd openpyxl
```

### Q: 文件扩展名是 .xls 但实际是 .xlsx 格式
这是正常的！工具会自动检测和处理。使用 `--keep-extension` 保持原始扩展名。

### Q: lastCell 参数不正确
手动打开 .xlsx 文件，右键单击注释，修改 lastCell 参数（通常是表格的最后一列和最后一行）。

### Q: 某些模板迁移失败
检查日志文件以获取具体错误。常见原因：
- 文件损坏
- 密码保护
- 特殊 JXLS 指令
- 复杂的嵌套结构

### Q: 如何处理加密的 Excel 文件
工具无法处理加密文件。请先解密。

### Q: 未检测到 Windows Terminal
检查环境变量 `WT_SESSION` 或 `WT_PROFILE_ID` 是否存在。

## 📝 迁移最佳实践

1. **备份原始文件** - 建议使用 `_backup` 后缀备份目录
2. **先试运行** - 使用 `--dry-run` 预览更改
3. **测试关键模板** - 迁移后验证业务功能
4. **查看报告** - 检查 migration_report.md 了解详情
5. **验证导出** - 确保所有导出功能正常工作

## 🤝 贡献

欢迎贡献！请随时提交 Pull Request。

### 开发设置
```bash
# 克隆仓库
git clone https://github.com/fivefish130/jxls-migration-tool.git
cd jxls-migration-tool

# 安装开发依赖
pip install -r requirements-dev.txt

# 运行测试
python -m pytest tests/
```

## 📄 许可证

MIT 许可证 - 详见 [LICENSE](LICENSE) 文件。

## 👨‍💻 作者

**fivefish**
- 版本: 3.1.0
- 日期: 2025-11-07

## 🙏 致谢

- [JXLS 项目](https://jxls.sourceforge.net/) - 优秀的 Java Excel 模板库
- [xlrd 库](https://github.com/python-excel/xlrd) - Python .xls 文件读取库
- [openpyxl 库](https://openpyxl.readthedocs.io/) - Python .xlsx 文件读写库

## 📚 相关项目

- [JXLS 官方文档](https://jxls.sourceforge.net/)
- [POI 项目](https://poi.apache.org/) - Apache POI - Java Microsoft 文档 API

---

**快乐迁移！** 🎉
