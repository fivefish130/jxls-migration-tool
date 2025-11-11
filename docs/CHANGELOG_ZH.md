# 更新日志

本文件记录 JXLS 迁移工具的所有重要更改。

格式基于 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.0.0/)，
项目版本遵循 [语义化版本](https://semver.org/lang/zh-CN/)。

## [Unreleased]

## [3.4.4] - 2025-11-11

### 移除
- **移除OpenPyXL写入模式** - 基于用户反馈，OpenPyXL模式不稳定
  - 删除 `--prefer-openpyxl` 命令行参数
  - 强制使用 XlsxWriter 作为唯一写入引擎
  - 如果未安装 xlsxwriter，工具会报错退出
  - 提供更清晰的错误消息指导安装

- **简化依赖管理**
  - xlsxwriter 改为必需依赖（而不是推荐依赖）
  - openpyxl 仍为必需依赖（用于读取）
  - xlrd 保持为可选依赖（仅用于 .xls 文件）

### 改进
- **强制XlsxWriter模式**
  - 所有输出文件自动使用共享字符串表
  - 确保与 Apache POI 5.4.0+ 的最佳兼容性
  - 文件更小，性能更好
  - 消除双模式导致的混淆

### 文档更新
- 更新 USAGE.md 移除 `--prefer-openpyxl` 参考
- 更新依赖安装说明
- 明确 xlsxwriter 为必需依赖
- 移除 OpenPyXL 写入模式相关文档

### 兼容性
- **Apache POI 5.4.0+**: 完全兼容 ✅
- **JXLS 2.14.0**: 完全兼容 ✅
- **所有迁移场景**: 使用 XlsxWriter ✅

### 迁移指南
- **生产环境**: 无需更改，XlsxWriter 提供更好稳定性
- **开发环境**: 安装 xlsxwriter (pip install xlsxwriter)
- **CI/CD**: 添加 xlsxwriter 到依赖列表

## [3.4.3] - 2025-11-11

### 新增
- **统一读写方法** - 采用 OpenPyXL 读取 + XlsxWriter 写入的策略
  - 读取: 统一使用 OpenPyXL（支持 .xls 和 .xlsx）
  - 写入: 默认使用 XlsxWriter（自动共享字符串表）
  - 添加 `--prefer-openpyxl` 参数用于特殊场景
  - 更好的 Apache POI 5.4.0+ 兼容性

- **XlsxWriter 支持** - 新增 XlsxWriter 3.2.0+ 作为写入引擎
  - 自动使用共享字符串表（xl/sharedStrings.xml）
  - 文件更小，性能更好
  - POI 兼容性显著提升
  - 新增 XlsxWriterConverter 类处理格式转换

### 改进
- **默认行为优化**
  - 默认使用 XlsxWriter（更优性能）
  - 移除 `--use-xlsxwriter` 参数
  - 新增 `--prefer-openpyxl` 参数（反向逻辑）
  - 更好的开箱即用体验

- **依赖管理**
  - xlrd 改为可选依赖（仅用于 .xls 文件）
  - xlrd 2.0+ 不再支持 .xls，自动提示安装 'xlrd<2.0'
  - xlsxwriter 作为推荐依赖（更好的兼容性）

### 技术细节
- 读写分离: OpenPyXL(读) + XlsxWriter(写)
- 共享字符串: XlsxWriter 自动 / OpenPyXL 内联
- 格式转换: 新增 XlsxWriterConverter 类
- 智能回退: xlsxwriter 不可用时自动使用 openpyxl

### 测试验证
- ✅ XlsxWriter 模式: 包含 xl/sharedStrings.xml
- ✅ OpenPyXL 模式: 内联字符串
- ✅ JXLS 转换: forEach→each, if(test→condition), out→${}
- ✅ 模板变量: ${row.cate1Name} 等完整保留

### 迁移指南
- **生产环境**: 无需更改，享受更好兼容性
- **开发环境**: 可使用 `--prefer-openpyxl` 便于调试
- **CI/CD**: 使用默认 XlsxWriter 设置

## [3.4.2] - 2025-11-11

### 修复
- **修复 keep-extension 参数行为**
  - 问题：.xls文件迁移后保持.xls后缀，但Jxls 2.14.0要求.xlsx格式文件
  - 解决：.xls文件迁移后保持.xls文件名，但文件内容转换为.xlsx格式
  - 这样既保持文件名不变（避免后端代码修改），又确保Jxls 2.14.0可正常读取
  - 目录迁移：.xls → .xls文件名（.xlsx内容），.xlsx → .xlsx文件名（.xlsx内容）

- **修复富文本格式导致的兼容性问题**
  - 问题：迁移后文件包含富文本格式，导致Jxls 2.14.0读取失败
  - 解决：简化格式复制逻辑
    - 不再复制字体名称和大小（避免中文字体兼容性问题）
    - 使用标准Calibri字体替代
    - 仅复制粗体/斜体等基本样式
    - 仅复制纯色填充和边框
  - 效果：避免富文本兼容性问题，确保Jxls 2.14.0可正常读取

### 改进
- **优化帮助信息**
  - 更新 --keep-extension 参数说明，明确文件格式行为
  - 添加详细的注释说明代码逻辑

### 测试验证
- 文件头检测：504b（ZIP格式，.xlsx本质）✅
- 文件类型检测：Microsoft Excel 2007+ ✅
- openpyxl可正常打开 ✅
- 迁移命令成功转换 ✅

### 推荐用法
```bash
# 使用 keep-extension 避免修改后端代码
python jxls_migration_tool.py templates/ -o migrated/ --keep-extension

# 结果：
# - 文件名保持不变（.xls还是.xls）
# - 文件内容是.xlsx（Jxls 2.14.0兼容）
# - 无需修改后端代码
```

### 兼容性
- **Jxls 2.14.0**: 完全兼容 ✅
- **Jxls 2.x**: 完全兼容 ✅
- **Jxls 1.x**: 自动迁移 ✅

## [3.4.1] - 2025-11-07

### 修复
- **智能注释位置计算** - 使用 `find_first_data_column` 方法找到数据行的第一个有数据的单元格
  - 精确位置计算：基于原模板的实际数据结构，而不是固定列位置
  - 详细日志：添加了注释位置的详细日志信息
  - 错误处理：如果没有找到有数据的列，回退到命令所在的列
  - 现在注释会正确地出现在数据行的第一个有数据的单元格中

### 改进
- **增强日志记录**
  - 添加注释位置的详细日志
  - 跟踪数据列查找过程
  - 记录回退机制使用情况

## [3.4.0] - 2025-11-07

### 修复
- **修复 jx:each 注释不生成问题**
  - 增强指令检测逻辑，使用更宽松的匹配模式
  - 添加详细的调试日志来跟踪命令处理过程
  - 确保 forEach 命令被正确识别和处理
  - 修复 `process_commands_and_migrate_data` 和 `process_commands_xlsx` 中的处理逻辑

- **修复 jx:area 位置错误问题**
  - 自动生成的 jx:area 现在正确添加到 A1 单元格
  - 修复注释位置计算逻辑
  - 添加专门的日志来确认 jx:area 注释位置
  - 在 XLS 和 XLSX 处理中都统一修复

### 改进
- **增强日志记录**
  - 添加更详细的调试信息
  - 跟踪命令处理全流程
  - 确认注释添加位置

### 版本更新
- 版本号更新至 v3.4
- 横幅和文档更新为"修复版 v3.4"

## [3.3.1] - 2025-11-07

### 修复
- **修复健壮迁移回退逻辑** - 解决 `migrate_xls_file` 内部处理异常导致回退机制失效的问题
- 移除依赖异常的 try-except 包装，改为检查第一次尝试的返回结果
- 如果第一次尝试失败，自动进行第二次尝试
- 保留完整的尝试记录日志

### 测试
- ✅ 成功处理 `.xls` 扩展名但实际是 `.xlsx` 格式的文件
- 自动回退机制正常工作

## [3.3.0] - 2025-11-07

### 重大改进
- **统一健壮API** - 将 `robust_migrate_file()` 重命名为 `migrate_file()`，简化外部接口
- 删除旧的 `migrate_file()` 方法（无回退机制）
- 现在 `migrate_file()` 统一提供健壮迁移能力

### 特性
- 所有迁移操作都使用健壮版本
- 自动格式检测
- 双重处理器回退
- 详细错误日志
- 尝试记录

### API 兼容性
- ✅ 外部调用无需修改（仍使用 `migrate_file`）
- ✅ 内部实现更健壮
- ✅ 单文件和目录迁移统一使用相同逻辑

### 代码变化
- 净减少: 44行代码
- 新增: 15行
- 删除: 59行

## [3.2.0] - 2025-11-07

### 新增
- **健壮迁移方案** - 新增 `robust_migrate_file()` 方法，支持自动格式检测和回退
- 新增 `safe_detect_excel_format()` 函数，安全检测文件格式
  - 自动检测 Excel 文件真实格式
  - 带详细日志记录
  - 出错时自动回退到扩展名判断

### 特性
- **双重处理器回退机制**
  - 第一次尝试：根据检测的格式选择处理器
  - 如果失败，自动回退到另一种处理器
- **详细尝试记录**
  - 返回结果包含 `attempts` 数组
  - 记录所有尝试和错误信息
- **增强日志记录**
  - 显示尝试记录（如果有多次尝试）
  - 详细的处理器切换日志

### 修复的问题
- ✅ 文件实际格式与扩展名不匹配（.xls 后缀但实际是 .xlsx）
- ✅ xlrd 无法处理 .xlsx 文件导致的错误
- ✅ 格式检测错误导致迁移失败

### 代码变化
- +116行 (新增)
- -8行 (删除)

## [3.1.0] - 2025-11-07

### 新增
- **增强错误处理** - 为 ExcelFormatConverter 添加完整的错误处理机制

### 修复
- **修复格式转换错误** - 解决 `'Format' object has no attribute 'font_index'` 错误
- 使用 `hasattr()` 和 `getattr()` 安全访问属性
- 为缺失属性提供默认值（字体: Calibri, 大小: 11）
- 新增 `copy_cell_format()` 静态方法统一处理格式复制

### 增强
- 所有格式转换方法（字体、填充、边框、对齐）使用安全的属性访问
- 详细错误日志记录用于调试
- 添加默认回退值策略

### 技术改进
- 字体名称默认值: 'Calibri'
- 字体大小默认值: 11
- 智能属性检查避免 AttributeError
- 防御性编程实践

### 兼容性
- ✅ xlrd 1.x 和 2.x 版本
- ✅ 各种 Excel 文件格式（.xls, .xlsx）
- ✅ 包含不完整格式信息的文件
- ✅ 包含自定义样式的文件

### 代码变化
- +190行 (新增)
- -87行 (删除)

## [3.0.0] - 2025-11-07

### 新增
- **完整迁移工具** - JXLS 1.x → 2.14.0 自动化迁移

### 核心特性
- **指令转换**
  - forEach → each
  - if(test → condition
  - out → ${}
  - area 自动生成
  - multiSheet 支持

- **格式保留**
  - 样式、列宽、行高
  - 合并单元格
  - 背景色
  - 字体、边框、对齐

- **智能识别**
  - 基于文件头检测真实格式
  - 不依赖文件后缀名

- **终端优化**
  - Windows Terminal 自动 UTF-8 检测与配置
  - 现代终端识别

- **报告生成**
  - Markdown 报告
  - JSON 数据
  - DEBUG 日志

### 工具头部
- 增强的 Unicode 支持
- 现代终端检测
- 改进的错误处理

## [Unreleased] - 早期版本

详细的早期版本更改未记录在此文件中。

---

## 版本说明

### 版本号格式
- 主版本号：不兼容的 API 更改
- 次版本号：向下兼容的功能性新增
- 修订号：向下兼容的问题修正

### 变更类型
- `新增` - 新功能
- `修改` - 对现有功能的变更
- `弃用` - 即将删除的功能
- `移除` - 已删除的功能
- `修复` - 任何问题修复
- `安全` - 安全相关修复

### 链接
- [Unreleased]: https://github.com/fivefish130/jxls-migration-tool/compare/v3.3.1...HEAD
- [3.3.1]: https://github.com/fivefish130/jxls-migration-tool/compare/v3.3.0...v3.3.1
- [3.3.0]: https://github.com/fivefish130/jxls-migration-tool/compare/v3.2.0...v3.3.0
- [3.2.0]: https://github.com/fivefish130/jxls-migration-tool/compare/v3.1.0...v3.2.0
- [3.1.0]: https://github.com/fivefish130/jxls-migration-tool/compare/v3.0.0...v3.1.0
- [3.0.0]: https://github.com/fivefish130/jxls-migration-tool/releases/tag/v3.0.0
