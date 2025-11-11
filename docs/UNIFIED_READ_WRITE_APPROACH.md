# 统一读写方法设计文档

## 概述

JXLS 迁移工具现已采用统一的读写策略：**OpenPyXL 读取 + XlsxWriter 写入**，为用户提供最佳的性能和兼容性。

## 设计理念

### 读写分离原则

| 操作 | 推荐工具 | 原因 |
|------|----------|------|
| **读取** (.xls, .xlsx) | OpenPyXL | 既支持旧版 .xls 也支持新版 .xlsx，功能全面 |
| **写入** (.xlsx) | XlsxWriter | 自动使用共享字符串表，文件更小，POI 兼容性更好 |

### 默认行为

- **默认**: 使用 `XlsxWriter` 写入（自动共享字符串表）
- **可选**: 使用 `--prefer-openpyxl` 显式选择 OpenPyXL 写入
- **原因**: XlsxWriter 写入的 XLSX 文件与 Apache POI 5.4.0+ 兼容性更好

## 新增参数

### `--prefer-openpyxl`
```bash
# 默认：使用 XlsxWriter（推荐）
python jxls_migration_tool.py input_dir

# 显式使用 OpenPyXL
python jxls_migration_tool.py input_dir --prefer-openpyxl
```

### `--verbose`
```bash
# 查看详细处理过程
python jxls_migration_tool.py input_dir --verbose
```

## 处理流程

### 1. 文件检测
- 自动检测文件真实格式（不依赖后缀）
- 识别 .xls vs .xlsx

### 2. 读取策略
- **所有文件**: 统一使用 `openpyxl` 读取
- **.xls 文件**: 需要 xlrd<2.0 支持
- **.xlsx 文件**: openpyxl 原生支持

### 3. 写入策略
- **默认**: XlsxWriter（自动共享字符串）
- **可选**: OpenPyXL（内联字符串）

### 4. JXLS 转换
- ✅ 转换 `jx:forEach` → `jx:each`（注释）
- ✅ 转换 `jx:if(test=...)` → `jx:if(condition=...)`
- ✅ 转换 `jx:out` → `${...}`
- ✅ 自动生成 `jx:area`
- ✅ 保留所有格式（字体、颜色、边框等）

## 依赖管理

### 必需依赖
```bash
pip install openpyxl
```

### 推荐依赖
```bash
# 用于处理 .xls 文件（可选）
pip install 'xlrd<2.0'

# 用于更好兼容性（推荐）
pip install xlsxwriter
```

### xlrd 版本说明
- **xlrd < 2.0**: 支持 .xls 和 .xlsx（推荐用于处理 .xls 文件）
- **xlrd >= 2.0**: 仅支持 .xlsx（已不支持 .xls）
- **未安装**: 仍可处理 .xlsx 文件，.xls 文件会提示安装 xlrd

## 性能对比

### 共享字符串表

| 工具 | 共享字符串 | 文件大小 | POI 兼容性 | 性能 |
|------|-----------|----------|------------|------|
| **XlsxWriter** | ✅ 自动 | ~6.9KB | ✅ 优秀 | ✅ 快 |
| **OpenPyXL** | ❌ 内联 | ~6.5KB | ⚠️ 需转换 | - |

### 测试结果
```bash
# XlsxWriter 版本（包含共享字符串）
unzip -l out/test_unified.xls | grep sharedStrings
# 输出: xl/sharedStrings.xml

# OpenPyXL 版本（无共享字符串）
unzip -l out/test_openpyxl.xls | grep sharedStrings
# 输出: （无结果）
```

## 使用场景推荐

### 使用 XlsxWriter（默认）✅
- 生产环境部署
- 与 Apache POI 5.4.0+ 集成
- 文件包含大量重复文本
- 追求更好性能和兼容性

### 使用 OpenPyXL（`--prefer-openpyxl`）
- 需要读取现有 XLSX 文件
- 依赖 OpenPyXL 特有功能
- 向后兼容性要求
- 调试和开发场景

## 升级指南

### 从 v3.3 升级
- **无需更改**: 现有脚本继续工作
- **新功能**: 添加 `--prefer-openpyxl` 参数
- **默认变化**: 现在默认使用 XlsxWriter（更优）

### 迁移建议
1. **生产环境**: 无需更改，享受更好兼容性
2. **开发环境**: 可使用 `--prefer-openpyxl` 便于调试
3. **CI/CD**: 使用默认 XlsxWriter 设置

## 故障排除

### 错误: "xlrd 2.0+ 仅支持 .xlsx"
**解决**: 安装旧版 xlrd
```bash
pip install 'xlrd<2.0'
```

### 错误: "缺少 xlsxwriter"
**解决**: 安装 xlsxwriter
```bash
pip install xlsxwriter
```
或使用 OpenPyXL:
```bash
python jxls_migration_tool.py input_dir --prefer-openpyxl
```

## 未来规划

- [ ] 移除 xlrd 依赖，全面转向 openpyxl
- [ ] 添加格式转换优化
- [ ] 支持更多 Excel 特性
- [ ] 性能基准测试套件

## 变更日志

- **v3.4**: 统一读写方法，默认使用 XlsxWriter
- **v3.3**: 修复 jx:each 注释和 jx:area 位置问题
- **v3.2**: 智能注释位置修复
- **v3.1**: 修复多 Sheet 导出模板
