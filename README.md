# 📊 表格脱敏工具

批量对 Excel 表格中的敏感字段（员工号、姓名、组织等）进行脱敏处理。

## 快速上手

```bash
# 1. 安装依赖
cd /home/ubuntu/table-desensitizer
source /home/ubuntu/.venv/bin/activate

# 2. 生成测试数据
python generate_test_data.py

# 3. 运行脱敏
python mask_tables.py
```

## 如何使用你自己的表格

### 方法一：替换文件（推荐）

1. 清空 `input/` 目录，放入你的 `.xlsx` 文件
2. 编辑 `config.yaml`，将 `column` 改为你表格的实际列名
3. 运行 `python mask_tables.py`
4. 脱敏后的文件在 `output/` 目录

### 方法二：指定目录

```bash
python mask_tables.py --input-dir ./你的表格目录 --output-dir ./脱敏结果
```

## 配置文件说明

编辑 `config.yaml`，核心是 `rules` 部分：

```yaml
rules:
  - column: "员工号"        # 表格中的列名（支持模糊匹配）
    strategy: "partial_mask" # 脱敏策略
    options:
      prefix_len: 1         # 保留前 N 位
      suffix_len: 3         # 保留后 N 位
```

### 支持的脱敏策略

| 策略 | 说明 | 示例 |
|------|------|------|
| `partial_mask` | 部分掩码，保留头尾 | `E2024001` → `E****001` |
| `fake_chinese_name` | 伪中文姓名（确定性） | `张三丰` → `赵明华` |
| `generic_org` | 通用组织名称 | `研发中心/后端组` → `技术研发-A/运营服务-B` |
| `hash` | 哈希脱敏 | `zhangsan@c.com` → `a1b2c3d4e5` |
| `fake_employee_id` | 伪员工号（保留格式特征） | `E2024001` → `X8392047` |

### 确定性映射

**同一原始值在整个运行中始终映射为同一脱敏值**。例如：
- 所有表格中工号 `E2024001` 都被脱敏为 `E****001`
- 这在关联不同表格时非常重要，不会丢失数据间的联系

## 命令行选项

```
python mask_tables.py --help

选项：
  --config PATH      配置文件路径（默认: config.yaml）
  --input-dir DIR    输入目录（覆盖配置文件）
  --output-dir DIR   输出目录（覆盖配置文件）
  --suffix SUFFIX    输出文件后缀（默认: _masked）
  --dry-run          预览模式，仅展示将要处理的文件
```

### 预览模式

先看看哪些列会被匹配、脱敏后会是什么效果：

```bash
python mask_tables.py --dry-run
```

## 项目结构

```
table-desensitizer/
├── mask_tables.py         # 主程序
├── config.yaml            # 配置文件（编辑它！）
├── generate_test_data.py  # 测试数据生成脚本
├── requirements.txt       # 依赖列表
├── README.md              # 本文件
├── input/                 # 放你的原始表格
│   ├── 员工信息表.xlsx
│   ├── 绩效评估表.xlsx
│   └── 薪酬统计表.xlsx
└── output/                # 脱敏结果输出到这里
    ├── 员工信息表_masked.xlsx
    ├── 绩效评估表_masked.xlsx
    └── 薪酬统计表_masked.xlsx
```

## 注意事项

1. ⚠ **请先备份原始数据**，脱敏后的文件保存在 `output/` 目录，不会覆盖原文件
2. 公式会被脱敏后的值替换（公式结果仍在）
3. 同一原始值重复出现时，脱敏结果一致（确定性映射）
4. 如果列名没匹配到，程序会提示跳过，不会破坏表格
