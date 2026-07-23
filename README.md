# 📊 表格脱敏工具

批量对 Excel 表格中的敏感字段（员工号、姓名、组织等）进行脱敏处理。

> 使用 [uv](https://docs.astral.sh/uv/) 管理项目，无需手动创建虚拟环境。

## 快速上手

```bash
# 1. 克隆
git clone https://github.com/zzhcng/table-desensitizer.git
cd table-desensitizer

# 2. 安装依赖（自动创建 .venv）
uv sync

# 3. 生成测试数据试试
uv run python generate_test_data.py

# 4. 运行脱敏
uv run python mask_tables.py
```

## 如何使用你自己的表格

### 方法一：替换文件（推荐）

1. 清空 `input/` 目录，放入你的 `.xlsx` 文件
2. 编辑 `config.yaml`，将 `column` 改为你表格的实际列名
3. 运行脱敏：`uv run python mask_tables.py`
4. 脱敏后的文件在 `output/` 目录

### 方法二：指定目录

```bash
uv run python mask_tables.py --input-dir ./你的表格目录 --output-dir ./脱敏结果
```

### 预览模式

先看看哪些列会被匹配、脱敏后会是什么效果：

```bash
uv run python mask_tables.py --dry-run
```

## 配置文件

编辑 `config.yaml`，核心是 `rules` 部分：

```yaml
rules:
  - column: "员工号"              # 表格中的列名（支持模糊匹配）
    strategy: "partial_mask"     # 脱敏策略
    options:
      prefix_len: 1              # 保留前 N 位
      suffix_len: 3              # 保留后 N 位
  - column: "姓名"
    strategy: "fake_chinese_name"
  - column: "组织"
    strategy: "generic_org"
```

## 脱敏策略

| 策略 | 说明 | 示例 |
|------|------|------|
| `partial_mask` | 部分掩码，保留头尾 | `E2024001` → `E****001` |
| `fake_chinese_name` | 伪中文姓名（确定性） | `张三丰` → `苗宇昊` |
| `generic_org` | 通用组织名称 | `研发中心/后端组` → `技术研发-A/运营服务-B` |
| `hash` | 哈希替换 | `zhangsan@c.com` → `a1b2c3d4e5` |
| `fake_employee_id` | 伪员工号（保留格式） | `E2024001` → `X8392047` |

## 命令行选项

```
uv run python mask_tables.py [选项]

  --config PATH      配置文件路径（默认: config.yaml）
  --input-dir DIR    输入目录（覆盖配置文件）
  --output-dir DIR   输出目录（覆盖配置文件）
  --suffix SUFFIX    输出文件后缀（默认: _masked）
  --dry-run          预览模式，仅展示将要处理的文件
```

## 项目结构

```
table-desensitizer/
├── mask_tables.py         # 主程序
├── config.yaml            # 配置文件（编辑它！）
├── generate_test_data.py  # 测试数据生成
├── pyproject.toml         # 项目元数据 + 依赖（uv 管理）
├── README.md              # 本文件
├── input/                 # 放你的原始表格
└── output/                # 脱敏结果输出到这里
```

## 注意事项

1. ⚠ **请先备份原始数据**，输出到 `output/` 目录，不覆盖原文件
2. **确定性映射** — 同一原始值跨文件脱敏结果一致，不影响表间关联
3. **公式保留** — Excel 公式中的引用值替换，公式本身不动
4. **仅改指定列** — 未配置的列原样保留
5. **列名模糊匹配** — `姓名` 自动匹配 `员工姓名`，`组织` 匹配 `所属组织`
