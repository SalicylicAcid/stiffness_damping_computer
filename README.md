# Joint Stiffness & Damping 回归计算工具

基于阻抗模型的关节刚度/阻尼回归计算，自动读取机器人会话日志，输出格式化 XLSX 结果。

**模型**:

$$\tau = K(\theta_{\text{ref}} - \theta_{\text{sta}}) + D(\dot{\theta}_{\text{ref}} - \dot{\theta}_{\text{sta}})$$

其中 $\tau$ 为关节力矩，$\theta$ 为关节位置，$\dot{\theta}$ 为关节速度，$K$ 为刚度系数，$D$ 为阻尼系数。

## 依赖

只需安装 [uv](https://docs.astral.sh/uv/) (Python ≥ 3.10)：

```bash
pip install uv       # 或用官网安装脚本
```

numpy / openpyxl 由 uv 自动管理，无需手动安装。

## 一键运行

### 方式 A：零安装远程运行（无需克隆，最简）

脚本本身声明了全部依赖（numpy / openpyxl），uv 自动读取安装。

**⭐ 推荐（全平台，支持传参，uv ≥ 0.4）**

```bash
# 在日志目录下执行即处理当前目录
uv run --script https://cdn.jsdelivr.net/gh/SalicylicAcid/stiffness_damping_computer@main/compute.py

# 也可显式指定输入/输出目录
uv run --script https://cdn.jsdelivr.net/gh/SalicylicAcid/stiffness_damping_computer@main/compute.py -C ./logs -o ./results
```

**或先下载再运行（便于反复使用）**

```bash
# Windows PowerShell
Invoke-WebRequest -Uri https://cdn.jsdelivr.net/gh/SalicylicAcid/stiffness_damping_computer@main/compute.py -OutFile compute.py
uv run compute.py

# Linux / macOS / WSL
curl -sLO https://cdn.jsdelivr.net/gh/SalicylicAcid/stiffness_damping_computer@main/compute.py
uv run compute.py
```

**管道运行（仅 Linux / macOS / WSL）**

```bash
curl -sL https://cdn.jsdelivr.net/gh/SalicylicAcid/stiffness_damping_computer@main/compute.py | uv run -
```

> 管道方式不支持向脚本传参（请先 `cd` 到目标目录）；且 Windows 上 uv 会把脚本作为命令行参数
> 传给 python，脚本超过约 32KB 时会报错，Windows 请使用上面的 `uv run --script` 或下载方式。

> **备用源**：若 CDN 不可用，可将 `cdn.jsdelivr.net/gh/...@main/` 改为 `raw.githubusercontent.com/.../main/`。

运行前将工作目录切换到会话日志目录（或日志根目录），结果 XLSX 输出到当前目录。

### 方式 B：本地运行

```bash
git clone https://github.com/SalicylicAcid/stiffness_damping_computer.git
cd stiffness_damping_computer
uv run compute.py -C <会话目录或父目录>
```

`uv run compute.py` 会自动创建 `.venv` 并安装依赖。首次安装依赖，后续直接执行。

### 高级选项

```bash
# 指定输入目录和输出目录
uv run compute.py -C ./logs -o ./results

# 关闭分析窗口裁剪（调试对比用，结果通常会严重失真）
uv run compute.py -C ./20261014_094915 --no-trim
```

## 输入：新版会话日志

每次运行生成一个会话目录（如 `20261014_094915/`）：

```
<session>/
├── recorder/
│   ├── <session>_joint_commands.csv   # 参考轨迹（关节位置）
│   └── <session>_joint_states.csv     # 跟随状态（位置 / 速度 / 力矩）
├── ros/                               # 忽略
└── zl_robot.log                       # 忽略
```

脚本扫描 `-C` 指定的目录：若该目录本身是会话目录（含 `recorder/`）则只处理它；否则处理其下所有含 `recorder/` 的子目录（每个会话输出一份 XLSX）。

### joint_commands.csv 列说明

| 列 | 含义 |
|----|------|
| `recorder_time_ms` | 记录时间戳 (ms)，用于与状态对齐 |
| `stage` / `type` | 命令阶段（如 `prepare_interp` / `control_output`），仅作参考信息展示 |
| `<joint>_pos` | 关节参考位置 (rad)，关节名如 `arm_l_1`、`leg_r_3` |

### joint_states.csv 列说明

| 列 | 含义 |
|----|------|
| `recorder_time_ms` | 记录时间戳 (ms) |
| `<joint>_pos` | 关节实际位置 (rad) |
| `<joint>_vel` | 关节实际速度 (rad/s) |
| `<joint>_tor` | 关节实际力矩 (Nm) |

`---` 为分块分隔列；err / merr / temp / mos / imu 等其余列不参与计算。

### 参考速度

命令日志只有位置，没有速度列，参考速度由位置序列数值差分（中心差分）得到。

### 分析窗口

默认只使用**机器人实际运行段**（参考轨迹正在运动的采样）：

- 命令生效前，控制器输出的是默认值（机器人并未执行该命令，例如启动阶段命令被拒绝），
  计入会使回归完全失真；
- 长时间静态保持段（参考速度 ≈ 0）不提供 $K$/$D$ 辨识信息，一并剔除；
- 速度阈值按「会话主运动速度」自适应（p99 的 5%），并对速度做 50ms 平滑以抑制
  命令跳变产生的单点尖峰；
- 状态日志开头的全零初始化段（驱动上线前）会被自动剔除；
- 用 `--no-trim` 可关闭筛选，使用全部重叠数据（对比排查用）。

> **关于负刚度**：稳定位置控制器不可能出现 $K<0$。若日志中只有缓慢的准备/归位插值
> （准静态、单向运动），力矩主要随位姿（重力/前馈）变化，此时回归得到的 $K$ 反映的是
> 负载随位姿的变化，而不是控制器刚度——腿部等承载关节常出现负值。工具会在 Summary
> 中给出预警。**这类日志无法分离刚度与负载**，请使用包含机器人真实运行
> （往复/振荡运动，如行走段）的日志。

## 关节分组

按身体部位自动分组（31 关节 → 2-3-7-7-6-6）：

| 部位 | 关节数 | 显示方式 |
|------|--------|----------|
| head | 2 | 独立成块 |
| waist | 3 | 独立成块 |
| arm_l / arm_r | 7 / 7 | 并排对比 |
| leg_l / leg_r | 6 / 6 | 并排对比 |

关节数量或命名变化时自动适配。

## 输出

| 文件 | 说明 |
|------|------|
| `stiffness_damping_<session>.xlsx` | 每个会话生成一个 XLSX |

每份 Excel 包含两个 Sheet：

- **Summary** — 分析窗口 / 整体统计 (均值/标准差/范围) + 指标说明
- **<session>** — 各关节 $K$ / $D$ / $R^2$ / RMSE 逐关节结果，按身体部位分组排版

### 字体颜色说明

| 颜色 | 含义 |
|------|------|
| 深绿 | 正常 / 较好 |
| 棕色 | 一般 / 偏弱 |
| 深红 | 异常 / 较差 |
| 灰色 | 无效数据 (力矩恒为零 / 无激励) |

### 异常判定依据

| 指标 | 正常 | 偏弱 | 异常 |
|------|------|------|------|
| $R^2$ | $\ge 0.8$ | $0.6 \sim 0.8$ | $< 0.6$ |
| RMSE | $\le P_{75}$ | $P_{75} \sim P_{90}$ | $> P_{90}$ |
| $K$ / $D$ | $\lvert z \rvert \le 1$ | $1 < \lvert z \rvert \le 2$ | $\lvert z \rvert > 2$ |

## 旧版日志

旧版 `dof_ref_*.csv` / `dof_sta_*.csv` 配对格式请使用
[`legacy` 分支](https://github.com/SalicylicAcid/stiffness_damping_computer/tree/legacy)。

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
└── stiffness_damping_*.xlsx       # 输出: 结果文件 (不纳入版本控制)
```
