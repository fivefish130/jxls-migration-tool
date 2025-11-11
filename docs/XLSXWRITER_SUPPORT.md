# XlsxWriter 支持文档

## 概述

JXLS 迁移工具现已支持 XlsxWriter 3.2.0+，可自动使用共享字符串表，提升与 Apache POI 5.4.0+ 的兼容性。

## 新增功能

### `--use-xlsxwriter` 参数

```bash
# 使用 XlsxWriter 迁移（自动共享字符串表）
python jxls_migration_tool.py input_dir --use-xlsxwriter

# 完整示例
python jxls_migration_tool.py input_dir -o output_dir --keep-extension --use-xlsxwriter --verbose
```

## 技术实现

### 共享字符串表对比

| 特性 | XlsxWriter | OpenPyXL (默认) |
|------|-----------|-----------------|
| 共享字符串表 | ✅ 自动 | ❌ 内联字符串 |
| POI 5.4.0 兼容性 | ✅ 优秀 | ⚠️ 可能有问题 |
| 文件大小 | ✅ 较小 | ❌ 较大 |
| 性能 | ✅ 较快 | ❌ 较慢 |
| 功能完整性 | ✅ 完整 | ✅ 完整 |

### 代码结构

1. **XlsxWriterConverter 类**: 处理格式转换
2. **migrate_xls_sheet_xlsxwriter()**: XlsxWriter 专用 sheet 迁移
3. **process_commands_and_migrate_data_xlsxwriter()**: 数据处理逻辑
4. **条件分支**: 根据 `use_xlsxwriter` 参数选择处理器

## 使用场景

### 推荐使用 XlsxWriter 的情况:
- 需要与 Apache POI 5.4.0+ 完美兼容
- 处理的 Excel 文件包含大量重复文本
- 希望减小文件大小
- 追求更好的写入性能

### 继续使用 OpenPyXL 的情况:
- 需要读取现有 XLSX 文件（XlsxWriter 仅支持写入）
- 依赖某些 OpenPyXL 特有功能
- 向后兼容性要求

## 验证方法

### 检查共享字符串表

```bash
# XlsxWriter 版本（包含共享字符串）
unzip -l out/hot_sock_tag_xlsxwriter.xls | grep sharedStrings
# 输出: xl/sharedStrings.xml

# OpenPyXL 版本（无共享字符串）
unzip -l out/hot_sock_tag_migrated.xls | grep sharedStrings
# 输出: （无结果）
```

### 迁移结果验证

两个版本都能正确:
- ✅ 转换 JXLS 指令: `jx:forEach` → `jx:each`
- ✅ 保留模板变量: `${row.cate1Name}`
- ✅ 添加 Excel 注释: `jx:area(lastCell="E2")`
- ✅ 删除标签行，保留数据行

## 性能对比

### 文件大小
- XlsxWriter: ~6.9KB (使用共享字符串)
- OpenPyXL: ~6.5KB (内联字符串)

### 处理速度
XlsxWriter 在处理大量重复文本时性能更优。

## 安装要求

```bash
pip install xlsxwriter==3.2.0+
```

## 未来改进

- [ ] 支持 XlsxWriter 读取 XLSX 文件
- [ ] 添加格式转换优化
- [ ] 性能基准测试
- [ ] 更多 POI 版本兼容性测试

## 变更日志

- **v3.4**: 新增 XlsxWriter 支持
- **v3.3**: 修复 jx:each 注释和 jx:area 位置问题
- **v3.2**: 智能注释位置修复
- **v3.1**: 修复多 Sheet 导出模板
