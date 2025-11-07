# JXLS 迁移工具 - 使用指南

本指南提供了如何有效使用 JXLS 迁移工具的详细说明。

## 目录

- [安装](#安装)
- [基本用法](#基本用法)
- [命令行选项](#命令行选项)
- [迁移模式](#迁移模式)
- [文件格式处理](#文件格式处理)
- [报告生成](#报告生成)
- [错误处理](#错误处理)
- [高级场景](#高级场景)
- [最佳实践](#最佳实践)
- [常见问题](#常见问题)

## 安装

### 系统要求

- **Python**: 3.6 或更高版本（推荐 Python 3.12）
- **操作系统**: Windows 10/11, Linux, macOS
- **内存**: 最小 512MB RAM（大文件推荐 1GB+）
- **磁盘空间**: Excel 文件大小的 2 倍（用于处理）

### 安装依赖

#### 使用 pip
```bash
pip install xlrd==2.0.1 openpyxl
```

#### 使用 conda
```bash
conda install xlrd=2.0.1 openpyxl
```

#### 验证安装
```bash
python -c "import xlrd, openpyxl; print('依赖安装成功')"
```

### Python 环境设置

#### 使用 virtualenv
```bash
# 创建虚拟环境
python -m venv jxls-migration-env

# 在 Windows 上激活
jxls-migration-env\Scripts\activate

# 在 Linux/macOS 上激活
source jxls-migration-env/bin/activate

# 安装依赖
pip install xlrd==2.0.1 openpyxl
```

#### 使用 conda
```bash
# 创建 conda 环境
conda create -n jxls-migration python=3.12
conda activate jxls-migration
pip install xlrd==2.0.1 openpyxl
```

## 基本用法

### 单文件迁移

#### 转换单个文件
```bash
# 将 input.xls 转换为 output.xlsx
python jxls_migration_tool.py input.xls -f output.xlsx

# 转换并保持扩展名
python jxls_migration_tool.py input.xls -f output.xls --keep-extension
```

#### 处理流程：
1. 读取输入文件
2. 检测 JXLS 指令
3. 转换为 JXLS 2.x 格式
4. 保留所有格式
5. 写入输出文件

### 目录迁移

#### 迁移目录中的所有 Excel 文件
```bash
# 基本目录迁移
python jxls_migration_tool.py /path/to/excel/templates

# 迁移到指定输出目录
python jxls_migration_tool.py /path/to/excel/templates -o /path/to/output

# 保持原始文件扩展名
python jxls_migration_tool.py /path/to/excel/templates -o /path/to/output --keep-extension

# 详细日志
python jxls_migration_tool.py /path/to/excel/templates --keep-extension --verbose
```

#### 处理流程：
1. 扫描目录中的 Excel 文件（.xls, .xlsx）
2. 检测每个文件中的 JXLS 指令
3. 处理包含 JXLS 指令的文件
4. 跳过不含 JXLS 指令的文件
5. 生成迁移报告
6. 创建日志

### 试运行模式

#### 预览更改而不修改文件
```bash
# 试运行 + 详细输出
python jxls_migration_tool.py /path/to/excel/templates --dry-run --verbose

# 试运行 + 指定输出目录
python jxls_migration_tool.py /path/to/excel/templates -o /path/to/output --dry-run
```

**你会得到：**
- 将要处理的文件列表
- 每个文件中发现的 JXLS 指令数量
- 估计的转换详情
- 无实际文件修改

## 命令行选项

### 完整选项参考

```bash
python jxls_migration_tool.py [选项] 输入路径
```

#### 位置参数

- `输入路径` (必需)
  - 类型: 字符串
  - 描述: 输入文件或目录路径
  - 示例:
    - `input.xls` (单个文件)
    - `input_directory` (包含 Excel 文件的目录)

#### 可选参数

##### `-h, --help`
- 描述: 显示帮助信息并退出
- 示例: `python jxls_migration_tool.py --help`

##### `-o 输出目录, --output 输出目录`
- 描述: 迁移文件的输出目录
- 默认: 与输入目录相同
- 类型: 字符串
- 示例: `python jxls_migration_tool.py input_dir -o output_dir`

##### `-f 输出文件, --file 输出文件`
- 描述: 输出文件（仅限单文件迁移）
- 默认: 不适用
- 类型: 字符串
- 示例: `python jxls_migration_tool.py input.xls -f output.xlsx`

##### `--keep-extension`
- 描述: 保持原始文件扩展名 (.xls → .xls, .xlsx → .xlsx)
- 默认: 将所有 .xls 转换为 .xlsx
- 类型: 布尔标志（不需要值）
- 示例: `python jxls_migration_tool.py input_dir --keep-extension`

##### `--dry-run`
- 描述: 预览更改而不修改文件
- 默认: False
- 类型: 布尔标志
- 示例: `python jxls_migration_tool.py input_dir --dry-run`

##### `-v, --verbose`
- 描述: 启用详细日志
- 默认: False
- 类型: 布尔标志
- 示例: `python jxls_migration_tool.py input_dir --verbose`

### 常用组合

```bash
# 1. 标准迁移 + 扩展名保持
python jxls_migration_tool.py input_dir -o output_dir --keep-extension

# 2. 迁移前预览
python jxls_migration_tool.py input_dir --dry-run --verbose

# 3. 就地迁移（覆盖原文件）
python jxls_migration_tool.py input_dir --keep-extension

# 4. 转换为 .xlsx（不保持扩展名）
python jxls_migration_tool.py input_dir -o output_dir

# 5. 单文件 + 指定输出名称
python jxls_migration_tool.py input.xls -f specific_output.xlsx
```

## 迁移模式

### 模式 1: 标准转换 (.xls → .xlsx)

**目的**: 将所有 .xls 文件转换为 .xlsx 格式

**命令**:
```bash
python jxls_migration_tool.py input_dir -o output_dir
```

**行为**:
- 所有 .xls 文件 → .xlsx
- 所有 .xlsx 文件保持 .xlsx
- JXLS 指令转换
- 格式保留

**适用场景**:
- 标准迁移
- 想要所有文件都是现代 .xlsx 格式
- 不关心保持原始扩展名

### 模式 2: 扩展名保持

**目的**: 保持原始文件扩展名

**命令**:
```bash
python jxls_migration_tool.py input_dir --keep-extension
```

**行为**:
- .xls 文件保持 .xls
- .xlsx 文件保持 .xlsx
- JXLS 指令转换
- 格式保留
- 智能格式检测（处理不匹配的扩展名）

**适用场景**:
- 当文件扩展名是工作流的一部分
- 需要保持原始文件类型
- 想要最小化更改

### 模式 3: 就地迁移

**目的**: 在同一目录中迁移文件

**命令**:
```bash
python jxls_migration_tool.py input_dir --keep-extension
```

**行为**:
- 文件在原位置覆盖
- 原文件被替换
- 建议使用前先备份

**适用场景**:
- 有备份时
- 乐于覆盖文件时
- 测试环境

### 模式 4: 预览模式

**目的**: 看到将要迁移的内容而不做更改

**命令**:
```bash
python jxls_migration_tool.py input_dir --dry-run --verbose
```

**行为**:
- 扫描文件
- 报告发现的 JXLS 指令
- 显示将要转换的内容
- 不修改文件

**适用场景**:
- 首次使用工具
- 迁移前审计
- 故障排除
- 学习工具工作方式

## 文件格式处理

### 智能格式检测

工具使用文件头分析来检测实际格式，而不仅仅是扩展名：

#### 检测逻辑

```python
# 读取文件前 8 字节
with open(filepath, 'rb') as f:
    header = f.read(8)

# 检查文件头
if header.startswith(b'\xd0\xcf\x11\xe0'):
    # OLE2 格式 = .xls
    actual_format = 'xls'
elif header.startswith(b'PK'):
    # ZIP 格式 = .xlsx
    actual_format = 'xlsx'
```

#### 处理格式不匹配

**场景 1**: 文件扩展名是 .xls 但实际是 .xlsx
- 正确检测
- 作为 .xlsx 处理
- 输出遵守 `--keep-extension` 标志

**场景 2**: 文件扩展名是 .xlsx 但实际是 .xls
- 正确检测
- 作为 .xls 处理
- 输出遵守 `--keep-extension` 标志

**日志输出**:
```
⚠️  文件扩展名为 .xls 但实际格式是 .xlsx
   检测格式: xlsx
   处理为: xlsx
```

### 支持的格式

| 格式 | 扩展名 | 检测 | 处理器 |
|------|--------|------|--------|
| Excel 97-2003 | .xls | OLE2 文件头 (D0CF11E0) | xlrd |
| Excel 2007+ | .xlsx | ZIP 文件头 (PK) | openpyxl |

### 不支持的格式

- **.xlsb** (Excel 二进制工作簿)
- **.xlam** (Excel 加载项)
- **.xlsm** (Excel 宏启用工作簿) - *可以读取但不保留宏*
- **密码保护文件** - *无法读取*
- **损坏文件** - *会失败*

## 报告生成

### 自动报告生成

迁移后会生成三种类型的报告：

#### 1. Markdown 报告 (migration_report.md)

**位置**: 输出目录
**用途**: 人类可读摘要
**内容**:
- 带版本和日期的标头
- 统计（总数、成功、失败、成功率）
- 已处理文件列表
- 成功转换列表
- 失败转换列表（带错误详情）
- 主要更改摘要
- 后续步骤建议

**示例**:
```markdown
# JXLS 迁移报告

## 统计
- 文件总数: 50
- 成功迁移: 50
- 失败: 0
- 成功率: 100%

## 成功迁移的文件
- purchase_order_export.xls
  - 发现 JXLS 命令: 2
  - 转换命令: 2
  - 工作表: 采购单

## 失败文件
(无)
```

#### 2. JSON 报告 (migration_report.json)

**位置**: 输出目录
**用途**: 机器可读详细数据
**内容**:
- 时间戳
- 完整统计
- 每个文件的详细记录
- 文件前后信息
- 转换详情
- 错误信息（如有）

**示例**:
```json
{
  "时间戳": "2025-11-07 10:00:00",
  "统计": {
    "文件总数": 50,
    "成功": 50,
    "失败": 0,
    "成功率": 100.0
  },
  "文件": [
    {
      "文件名": "purchase_order_export.xls",
      "状态": "成功",
      "发现命令": 2,
      "转换命令": 2,
      "工作表": ["采购单"]
    }
  ]
}
```

#### 3. 调试日志 (jxls_migration.log)

**位置**: 输出目录
**用途**: 完整执行日志
**内容**:
- 带时间戳的日志条目
- DEBUG 级别详情
- 错误堆栈跟踪
- 文件处理详情
- 指令检测和转换日志

**示例**:
```
2025-11-07 10:00:00 - INFO - 开始迁移目录: input_dir
2025-11-07 10:00:01 - DEBUG - 处理文件: purchase_order_export.xls
2025-11-07 10:00:01 - DEBUG - 检测格式: xlsx
2025-11-07 10:00:01 - DEBUG - 发现 2 个 JXLS 命令
2025-11-07 10:00:02 - INFO - 成功迁移: purchase_order_export.xls
```

### 自定义报告位置

报告总是生成在输出目录中。要指定不同的位置：

```bash
# 报告将在 output_dir 中
python jxls_migration_tool.py input_dir -o output_dir --keep-extension --verbose
```

## 错误处理

### 常见错误和解决方案

#### 错误: "找不到 xlrd 库"

**解决方案**:
```bash
pip install xlrd==2.0.1
```

#### 错误: "权限被拒绝"

**原因**: 对输出目录没有写权限

**解决方案**:
- 检查目录权限
- 使用适当用户权限运行
- 指定不同的输出目录

#### 错误: "文件已加密"

**原因**: Excel 文件受密码保护

**解决方案**:
- 手动解密文件
- 移除密码保护
- 重新运行迁移

#### 错误: "文件已损坏"

**原因**: Excel 文件损坏或无效

**解决方案**:
- 验证文件可以在 Excel 中打开
- 从备份恢复
- 必要时重新创建文件

#### 错误: "lastCell 参数不正确"

**原因**: 复杂嵌套结构可能导致 lastCell 不正确

**解决方案**:
- 手动审查迁移后的文件
- 调整 Excel 注释中的 lastCell
- 报告为工具改进问题

### 错误恢复

即使个别文件失败，工具也会继续处理：

```bash
# 带有失败的示例输出
✅ 成功迁移: 48 个文件
❌ 失败: 2 个文件
📊 总计: 50 个文件
🎯 成功率: 96.00%
```

检查报告以查看哪些文件失败以及原因。

## 高级场景

### 大文件迁移

对于包含许多大文件的目录：

```bash
# 使用详细日志处理
python jxls_migration_tool.py large_directory --keep-extension --verbose

# 先试运行估计时间
python jxls_migration_tool.py large_directory --dry-run
```

**性能提示**:
- 确保有足够磁盘空间（文件大小的 2 倍）
- 使用 SSD 存储以获得更好性能
- 关闭其他应用程序以释放内存

### 混合文件类型

包含各种 Excel 文件的目录：

```bash
python jxls_migration_tool.py mixed_directory --keep-extension --verbose
```

工具将：
- 处理含 JXLS 指令的 .xls/.xlsx 文件
- 跳过不含 JXLS 指令的文件
- 跳过非 Excel 文件
- 生成综合报告

### 自动化脚本集成

```bash
#!/bin/bash
# migrate.sh - 批量迁移脚本

输入目录=$1
输出目录=$2

if [ -z "$输入目录" ] || [ -z "$输出目录" ]; then
    echo "用法: $0 <input_dir> <output_dir>"
    exit 1
fi

# 创建输出目录
mkdir -p "$输出目录"

# 运行迁移
python jxls_migration_tool.py "$输入目录" -o "$输出目录" --keep-extension --verbose

# 检查结果
if [ $? -eq 0 ]; then
    echo "迁移成功完成"
    echo "报告位置: $输出目录"
else
    echo "迁移失败。检查日志。"
    exit 1
fi
```

### Git 集成

```bash
# 添加到 .gitignore
migration_report.*
jxls_migration.log

# 预提交钩子示例
#!/bin/bash
# 对更改的文件运行试运行
git diff --name-only | grep '\.xls$|\.xlsx$' | while read file; do
    echo "检查 $file 的 JXLS 1.x 语法..."
    if grep -q '<jx:' "$file"; then
        echo "错误: $file 包含 JXLS 1.x 语法。运行迁移工具。"
        exit 1
    fi
done
```

## 最佳实践

### 1. 总是先备份

```bash
# 创建备份
cp -r templates templates_backup_$(date +%Y%m%d)

# 运行迁移
python jxls_migration_tool.py templates --keep-extension
```

### 2. 先使用试运行

```bash
# 总是预览更改
python jxls_migration_tool.py templates --dry-run --verbose

# 查看输出
cat migration_report.md
```

### 3. 测试关键模板

迁移后，测试关键模板以确保它们工作：
- 运行你的应用程序
- 测试导出功能
- 验证输出符合预期
- 检查是否有运行时错误

### 4. 查看报告

总是检查生成的报告：
- `migration_report.md` - 人类摘要
- `migration_report.json` - 详细数据
- `jxls_migration.log` - 调试信息

### 5. 版本控制

```bash
# 提交迁移后的文件
git add -A
git commit -m "迁移 JXLS 模板到 2.14.0"

# 标记版本
git tag -a v2.14.0 -m "JXLS 2.14.0 迁移"
```

### 6. 记录你的迁移

```markdown
# 迁移记录

日期: 2025-11-07
版本: 3.0
迁移文件: 50
成功率: 100%

已测试文件:
- [ ] purchase_order_export.xls
- [ ] inventory_report.xlsx
...

发现的问题:
- (列出任何问题及其解决方案)
```

### 7. 逐步推出

```bash
# 阶段 1: 测试环境
python jxls_migration_tool.py test_templates -o test_output --keep-extension

# 阶段 2: 预发布环境
python jxls_migration_tool.py staging_templates -o staging_output --keep-extension

# 阶段 3: 生产环境
python jxls_migration_tool.py prod_templates -o prod_output --keep-extension
```

## 常见问题

### Q: 迁移需要多长时间？

**A**: 通常每个文件 < 1 秒。对于 50 个文件，预期 < 30 秒。大文件可能需要更长时间。

### Q: 我的公式会保留吗？

**A**: 是的，单元格公式（${...} 表达式）会被保留。工具只修改 JXLS 指令。

### Q: 模板中的图片会怎样？

**A**: 图片会被保留。工具保留所有非文本内容。

### Q: 可以从 2.x 迁移回 1.x 吗？

**A**: 不，迁移是单向的。保留原文件备份。

### Q: 工具支持中文/日文/韩文字符吗？

**A**: 是的，完全支持 Unicode。Windows Terminal 被自动检测和配置。

### Q: 如果我有自定义 JXLS 指令怎么办？

**A**: 工具支持标准 JXLS 指令（forEach, if, out, area, multiSheet）。自定义指令需要手动迁移。

### Q: 如何验证迁移成功？

**A**:
1. 检查迁移报告
2. 在 Excel 中打开迁移后的文件
3. 验证存在 JXLS 注释
4. 测试模板渲染
5. 运行应用程序测试

### Q: 可以在 CI/CD 流水线中使用吗？

**A**: 可以！GitHub Actions 示例：

```yaml
- name: 迁移 JXLS 模板
  run: |
    python jxls_migration_tool.py templates --keep-extension
    # 检查失败
    if grep -q "失败:" migration_report.md; then
      echo "迁移失败!"
      exit 1
    fi
```

### Q: 报告在哪里？

**A**: 报告在输出目录中：
- `migration_report.md` (Markdown 摘要)
- `migration_report.json` (详细 JSON)
- `jxls_migration.log` (调试日志)

### Q: 如何报告问题？

**A**:
1. 启用 `--verbose` 标志
2. 包含日志文件 (`jxls_migration.log`)
3. 包含迁移报告
4. 如可能提供示例文件
5. 报告工具版本和 Python 版本

---

**需要更多帮助？** 查看 [API 文档](API.md) 或访问 [项目仓库](https://github.com/fivefish130/jxls-migration-tool)。
