# Joint Stiffness & Damping 回归计算工具

基于阻抗模型的关节刚度/阻尼回归计算，自动读取 CSV 配对文件，输出格式化 XLSX 结果。

**模型**: τ = K · (θ_ref − θ_sta) + D · (θ̇_ref − θ̇_sta)

## 依赖

只需安装 [uv](https://docs.astral.sh/uv/) (Python ≥ 3.10)：

```bash
pip install uv       # 或用官网安装脚本
```

numpy / openpyxl 由 uv 自动管理，无需手动安装。

## 一键运行

### 方式 A：零安装管道运行（无需克隆，最简）

```bash
# Windows PowerShell
irm https://raw.githubusercontent.com/SalicylicAcid/stiffness_damping_computer/main/compute.py | uv run -

# Linux / macOS / WSL
curl -sL https://raw.githubusercontent.com/SalicylicAcid/stiffness_damping_computer/main/compute.py | uv run -
```

↓ 等价于这种简化写法（uv ≥ 0.4）：

```bash
uv run --script https://raw.githubusercontent.com/SalicylicAcid/stiffness_damping_computer/main/compute.py
```

运行前确保当前目录下有 `dof_ref_*.csv` / `dof_sta_*.csv` 配对文件。
结果文件 `stiffness_damping_*.xlsx` 输出到当前目录。

### 方式 B：本地运行

```bash
git clone https://github.com/SalicylicAcid/stiffness_damping_computer.git
cd stiffness_damping_computer
uv run compute.py
```

`uv run compute.py` 会自动创建 `.venv` 并安装依赖。首次安装依赖，后续直接执行。

### 高级选项

```bash
# 指定输入目录和输出目录
uv run compute.py -C ./data -o ./results
```

## 输入文件

脚本自动扫描当前目录下的 CSV 配对：

| 文件 | 必需 | 说明 |
|------|------|------|
| `dof_ref_<suffix>.csv` | ✅ | 参考轨迹 (joint_pos / joint_vel) |
| `dof_sta_<suffix>.csv` | ✅ | 跟随状态 (joint_pos / joint_vel / joint_tor) |

命名规则：`dof_ref_` 与 `dof_sta_` 共用时间后缀，脚本自动匹配合并。

### CSV 列说明

| 列 | 含义 |
|----|------|
| `timestamp` | 时间戳，用于对齐 |
| `index` | 序列号 |
| `joint_pos{i}` | 第 i 个关节位置 (rad) |
| `joint_vel{i}` | 第 i 个关节速度 (rad/s) |
| `joint_tor{i}` | 第 i 个关节力矩 (Nm) — 仅 sta 文件 |

## 输出

| 文件 | 说明 |
|------|------|
| `stiffness_damping_<suffix>.xlsx` | 每个配对生成一个 XLSX |

每份 Excel 包含两个 Sheet：

- **Summary** — 整体统计 (均值/标准差/范围) + 指标说明
- **Data** — 各关节 K/D/R²/RMSE 逐关节结果，按关节分组排版（7-7 / 6-6 并排对比）

### 字体颜色说明

| 颜色 | 含义 |
|------|------|
| 深绿 | 正常 / 较好 |
| 棕色 | 一般 / 偏弱 |
| 深红 | 异常 / 较差 |
| 灰色 | 无效数据 (力矩恒为零) |

### 异常判定依据

| 指标 | 正常 | 偏弱 | 异常 |
|------|------|------|------|
| R² | ≥ 0.8 | 0.6 ~ 0.8 | < 0.6 |
| RMSE | ≤ P75 | P75 ~ P90 | > P90 |
| K / D | \|z\| ≤ 1 | 1 < \|z\| ≤ 2 | \|z\| > 2 |

## 关节分组

| 关节数 | 分组 |
|--------|------|
| 29 | 7-7-3-6-6 |
| 27 | 7-7-1-6-6 |
| 其它 | 自动适配 |

A-B 组 (7-7) 和 D-E 组 (6-6) 并排对比。

## 系统兼容性

| 系统 | 状态 |
|------|------|
| Windows 10+ | ✅ 已验证 |
| Ubuntu 22.04+ | ✅ |
| macOS | ✅ |

## 项目结构

```
.
├── compute.py                     # 主脚本
├── pyproject.toml                 # 项目配置 & 依赖声明
├── .gitignore
├── README.md
├── dof_ref_*.csv                  # 输入: 参考轨迹 (不纳入版本控制)
├── dof_sta_*.csv                  # 输入: 跟随状态 (不纳入版本控制)
└── stiffness_damping_*.xlsx       # 输出: 结果文件 (不纳入版本控制)
```
