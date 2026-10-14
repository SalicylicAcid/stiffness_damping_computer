#!/usr/bin/env -S uv run --script
# -*- coding: utf-8 -*-
# /// script
# dependencies = [
#   "numpy>=1.21",
#   "openpyxl>=3.0",
# ]
# ///
"""
Joint Stiffness & Damping 回归计算 (新版日志) + 格式化 XLSX 输出
==============================================================
基于阻抗模型: tau = K·(theta_ref − theta_sta) + D·(theta_dot_ref − theta_dot_sta)

新版日志: 每次运行生成一个会话目录 (如 20261014_094915/):
  <session>/
  ├── recorder/
  │   ├── <session>_joint_commands.csv   # 参考轨迹 (仅关节位置, 速度由位置差分得到)
  │   └── <session>_joint_states.csv     # 跟随状态 (位置 / 速度 / 力矩)
  ├── ros/                               # 忽略
  └── zl_robot.log                       # 忽略

用法:
  uv run compute.py                          # 扫描当前目录下的会话
  uv run compute.py -C <会话目录或父目录> -o <输出目录>
  uv run compute.py --no-trim                # 不裁剪分析窗口 (调试对比用)

分析窗口:
  默认只使用「参考轨迹实际运动」的采样段 (机器人实际运行段):
    - 命令生效前记录的是控制器默认输出 (机器人并未执行该命令), 计入会使回归完全失真;
    - 长时间静态保持段 (参考速度≈0) 不提供 K/D 辨识信息, 一并剔除;
    - 速度阈值按会话主运动速度自适应 (p99 的 5%), 并做 50ms 平滑抑制命令跳变尖峰;
    - 状态日志开头的全零初始化段同样会被自动剔除;
    - 用 --no-trim 可关闭筛选, 使用全部重叠数据。

  注意: 若日志中只有缓慢的准备/归位插值 (准静态单向), 力矩随位姿 (重力/前馈) 变化,
  回归得到的 K 反映负载而非控制器刚度 (会出现 K<0, 工具会给出预警); 请使用包含
  机器人真实运行 (往复/振荡运动) 的日志段。

关节分组 (按身体部位, 31 关节 → 2-3-7-7-6-6):
  head(2) / waist(3) 独立成块; arm_l↔arm_r (7-7)、leg_l↔leg_r (6-6) 并排对比。
"""

import argparse, csv, glob, os, re, sys
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ============================================================
# 0. 样式定义
# ============================================================
HEADER_FILL   = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
HEADER_FONT   = Font(name="Microsoft YaHei", bold=True, color="FFFFFF", size=11)
TITLE_FONT    = Font(name="Microsoft YaHei", bold=True, size=14, color="1F4E79")
GROUP_FONT    = Font(name="Microsoft YaHei", bold=True, size=11, color="2E75B6")
NORMAL_FONT   = Font(name="Consolas", size=10)
WARN_FONT     = Font(name="Consolas", size=10, color="999999")
THIN_BORDER   = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)
CENTER_ALIGN  = Alignment(horizontal="center", vertical="center")
LEFT_ALIGN    = Alignment(horizontal="left", vertical="center")
LIGHT_FILL    = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
DARK_FILL     = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
SEP_FILL      = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")

# 异常程度颜色 — 用字体颜色标注, 不改底纹
GOOD_COLOR    = "006100"   # 深绿: 正常 / 较好
WEAK_COLOR    = "9C6500"   # 棕色: 一般 / 偏弱
BAD_COLOR     = "C00000"   # 深红: 异常 / 较差

# 参考轨迹「开始运动」判定阈值 (rad, 相对命令起始值的累计偏差)
MOTION_EPS = 1e-5
# 参考轨迹「正在运动」判定: 速度平滑窗 (s) / 相对阈值 (占会话主运动速度 p99 的比例) / 绝对下限 (rad/s)
MOTION_SMOOTH_S = 0.05
MOTION_SPEED_FRAC = 0.05
MOTION_SPEED_EPS = 2e-3
# 运动分段合并间隔 (s): 间隔小于该值的两段运动视为同一段 (如步态摆动过零)
MOTION_GAP_S = 0.5


def get_quality_font(value, metric, valid_scores):
    """
    返回异常标注用的 Font 对象 (只改颜色，不改底纹)。
    metric: "R2" / "RMSE" / "K" / "D"
    """
    if valid_scores is None:
        return None
    if metric == "R2":
        if value < 0.6:   return Font(name="Consolas", size=10, color=BAD_COLOR, bold=True)
        if value < 0.8:   return Font(name="Consolas", size=10, color=WEAK_COLOR, bold=True)
        return Font(name="Consolas", size=10, color=GOOD_COLOR)
    elif metric == "RMSE":
        arr = valid_scores.get("rmse_vals")
        if arr is None or len(arr) == 0:
            return None
        p75 = np.percentile(arr, 75)
        p90 = np.percentile(arr, 90)
        if value > p90:   return Font(name="Consolas", size=10, color=BAD_COLOR, bold=True)
        if value > p75:   return Font(name="Consolas", size=10, color=WEAK_COLOR, bold=True)
        return Font(name="Consolas", size=10, color=GOOD_COLOR)
    elif metric == "K":
        arr = valid_scores.get("k_vals")
        if arr is None or len(arr) == 0: return None
        mean, std = np.mean(arr), np.std(arr)
        if std < 1e-12: return None
        z = abs(value - mean) / std
        if z > 2.0:     return Font(name="Consolas", size=10, color=BAD_COLOR, bold=True)
        if z > 1.0:     return Font(name="Consolas", size=10, color=WEAK_COLOR, bold=True)
        return Font(name="Consolas", size=10, color=GOOD_COLOR)
    elif metric == "D":
        arr = valid_scores.get("d_vals")
        if arr is None or len(arr) == 0: return None
        mean, std = np.mean(arr), np.std(arr)
        if std < 1e-12: return None
        z = abs(value - mean) / std
        if z > 2.0:     return Font(name="Consolas", size=10, color=BAD_COLOR, bold=True)
        if z > 1.0:     return Font(name="Consolas", size=10, color=WEAK_COLOR, bold=True)
        return Font(name="Consolas", size=10, color=GOOD_COLOR)
    return None


def style_header_row(ws, row, col_start, col_end):
    for c in range(col_start, col_end + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER_ALIGN
        cell.border = THIN_BORDER


def style_data_cell(ws, row, col, value, fmt="0.0000", is_warn=False, fill=None, font=None):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = font if font is not None else (WARN_FONT if is_warn else NORMAL_FONT)
    cell.alignment = CENTER_ALIGN
    cell.border = THIN_BORDER
    if fill is not None:
        cell.fill = fill
    if isinstance(value, (int, float)) and not is_warn and fmt:
        cell.number_format = fmt
    return cell


def apply_banding(ws, start_row, end_row, col_start, col_end):
    """每个子表格独立: 第1行浅色, 第2行深色, 交替重复"""
    for r in range(start_row, end_row + 1):
        fill = DARK_FILL if (r - start_row) % 2 == 1 else LIGHT_FILL
        for c in range(col_start, col_end + 1):
            ws.cell(row=r, column=c).fill = fill


# ============================================================
# 1. 会话发现
# ============================================================
def find_session_dirs(root):
    """
    返回会话目录列表:
      - root 本身含 recorder/          → [root]
      - root 直接是 recorder 目录      → [root]
      - 否则扫描一级子目录中含 recorder/ 的目录
    """
    if os.path.isdir(os.path.join(root, "recorder")):
        return [root]
    if glob.glob(os.path.join(root, "*_joint_commands.csv")):
        return [root]
    out = []
    for name in sorted(os.listdir(root)):
        d = os.path.join(root, name)
        if os.path.isdir(os.path.join(d, "recorder")):
            out.append(d)
    return out


def session_label(path):
    """会话标签: 目录名; 若直接指向 recorder 目录则取父目录名。"""
    p = os.path.normpath(path)
    label = os.path.basename(p)
    if label == "recorder":
        parent = os.path.basename(os.path.dirname(p))
        if parent:
            label = parent
    return label


def find_csv_pair(session_dir):
    """定位会话的 (joint_commands.csv, joint_states.csv); 优先匹配会话目录名前缀。"""
    label = session_label(session_dir)
    rec = os.path.join(session_dir, "recorder")
    if not os.path.isdir(rec):
        rec = session_dir

    cmds = sorted(glob.glob(os.path.join(rec, "*_joint_commands.csv")))
    stas = sorted(glob.glob(os.path.join(rec, "*_joint_states.csv")))
    if not cmds or not stas:
        return None

    def prefer(files):
        for f in files:
            if os.path.basename(f).startswith(label):
                return f
        return files[0]

    return prefer(cmds), prefer(stas)


# ============================================================
# 2. CSV 读取
# ============================================================
def load_table(path):
    """读取 CSV: 自动跳过空行 / 列数不符 / 含 null 字节的坏行。"""
    header, rows = None, []
    with open(path, "r", encoding="utf-8", errors="replace", newline="") as f:
        for row in csv.reader(f):
            row = [c.replace("\0", "").strip() for c in row]
            if header is None:
                header = row
                continue
            if not row or len(row) != len(header) or all(c == "" for c in row):
                continue
            rows.append(row)
    if header is None:
        raise ValueError(f"空文件: {path}")
    return header, rows


def col_index(header):
    idx = {}
    for i, name in enumerate(header):
        idx.setdefault(name, i)   # 重复列名 (states 中的 '---') 取首次出现
    return idx


def to_float(s):
    if s == "" or s.lower() in ("nan", "none"):
        return np.nan
    try:
        return float(s)
    except ValueError:
        return np.nan


# ============================================================
# 3. 关节解析与分组
# ============================================================
PART_ORDER = ["head", "waist", "arm_l", "arm_r", "leg_l", "leg_r"]


def parse_joints(header):
    """从 commands 表头提取关节名 (xxx_pos 列), 按身体部位 + 序号规范排序。"""
    names = [n[:-4] for n in header if n.endswith("_pos")]

    def key(j):
        m = re.match(r"^(.*)_(\d+)$", j)
        base, num = (m.group(1), int(m.group(2))) if m else (j, 0)
        gi = PART_ORDER.index(base) if base in PART_ORDER else len(PART_ORDER)
        return (gi, base, num)

    return sorted(names, key=key)


def group_joints(joints):
    """
    按身体部位分组 → (groups, ordered, pairs, blocks)
      groups:  {部位: [关节名, ...]}
      ordered: 显示顺序的部位列表
      pairs:   [(左, 右), ...] 并排对比的部位对
      blocks:  [("pair", 左, 右) | ("single", 部位, None), ...] 按显示顺序
    """
    groups = {}
    for j in joints:               # joints 已按规范顺序排列
        base = re.sub(r"_\d+$", "", j)
        groups.setdefault(base, []).append(j)

    ordered = [g for g in PART_ORDER if g in groups] + \
              [g for g in sorted(groups) if g not in PART_ORDER]

    pairs = []
    if "arm_l" in groups and "arm_r" in groups:
        pairs.append(("arm_l", "arm_r"))
    if "leg_l" in groups and "leg_r" in groups:
        pairs.append(("leg_l", "leg_r"))

    # 按身体部位顺序生成显示块: ("pair", 左, 右) 或 ("single", 部位)
    left_of = {gl: gr for gl, gr in pairs}
    right_of = {gr for _, gr in pairs}
    blocks = []
    for g in ordered:
        if g in left_of:
            blocks.append(("pair", g, left_of[g]))
        elif g in right_of:
            continue
        else:
            blocks.append(("single", g, None))
    return groups, ordered, pairs, blocks


# ============================================================
# 4. 数据处理 (对齐 + 回归)
# ============================================================
def load_and_compute(cmd_path, sta_path, label, trim=True, t_from=None, t_to=None):
    h_cmd, rows_cmd = load_table(cmd_path)
    h_sta, rows_sta = load_table(sta_path)
    ci, si = col_index(h_cmd), col_index(h_sta)

    if not rows_cmd or not rows_sta:
        raise ValueError("CSV 无有效数据行")

    t_key_c = "recorder_time_ms" if "recorder_time_ms" in ci else "msg_stamp_ms"
    t_key_s = "recorder_time_ms" if "recorder_time_ms" in si else "msg_stamp_ms"

    all_joints = parse_joints(h_cmd)
    joints = [j for j in all_joints
              if f"{j}_pos" in si and f"{j}_vel" in si and f"{j}_tor" in si]
    if not joints:
        raise ValueError("commands / states 之间没有可匹配的关节列")

    t_cmd = np.array([to_float(r[ci[t_key_c]]) for r in rows_cmd])
    t_sta = np.array([to_float(r[si[t_key_s]]) for r in rows_sta])

    P_ref = np.array([[to_float(r[ci[f"{j}_pos"]]) for j in joints] for r in rows_cmd])
    P_sta = np.array([[to_float(r[si[f"{j}_pos"]]) for j in joints] for r in rows_sta])
    V_sta = np.array([[to_float(r[si[f"{j}_vel"]]) for j in joints] for r in rows_sta])
    T_sta = np.array([[to_float(r[si[f"{j}_tor"]]) for j in joints] for r in rows_sta])

    # --- 参考轨迹裁剪: 从首次运动开始 (之前的命令未被执行) ---
    i0 = 0
    motion_detected = True
    if trim and len(P_ref) > 1:
        dev = np.abs(P_ref - P_ref[0])
        dev = np.where(np.isfinite(dev), dev, 0.0).max(axis=1)
        nz = np.flatnonzero(dev > MOTION_EPS)
        if len(nz):
            i0 = int(nz[0])
        else:
            motion_detected = False

    t_cmd_u, P_ref_u = t_cmd[i0:], P_ref[i0:]
    # 去掉时间戳重复/倒退的采样, 保证差分有效
    keep = np.concatenate([[True], np.diff(t_cmd_u) > 0])
    t_cmd_u, P_ref_u = t_cmd_u[keep], P_ref_u[keep]

    # --- 参考速度: 先截断再差分, 避免截断处的伪跳变尖峰 ---
    V_ref_u = np.gradient(P_ref_u, t_cmd_u / 1000.0, axis=0)

    # --- 状态侧窗口: 跳过开头的全零初始化段 ---
    i_s0 = 0
    if len(P_sta) > 1:
        all_zero = (np.nan_to_num(np.abs(P_sta)) == 0).all(axis=1) & \
                   (np.nan_to_num(np.abs(V_sta)) == 0).all(axis=1) & \
                   (np.nan_to_num(np.abs(T_sta)) == 0).all(axis=1)
        nz = np.flatnonzero(~all_zero)
        if len(nz):
            i_s0 = int(nz[0])

    t_start, t_end = float(t_cmd_u[0]), float(t_cmd_u[-1])
    mask = (t_sta >= t_start) & (t_sta <= t_end)
    mask[:i_s0] = False

    # --- 用户指定的时间窗 (相对日志起始, 秒) ---
    if t_from is not None:
        mask &= t_sta >= float(t_cmd[0]) + t_from * 1000.0
    if t_to is not None:
        mask &= t_sta <= float(t_cmd[0]) + t_to * 1000.0

    # --- 只保留「参考轨迹正在运动」的采样 (机器人实际运行段) ---
    n_seg = 1
    seg_list = []
    win_start_s = (t_start - float(t_cmd[0])) / 1000.0
    win_end_s = (t_end - float(t_cmd[0])) / 1000.0
    win_dur_s = (t_end - t_start) / 1000.0
    if trim and motion_detected:
        speed = np.abs(V_ref_u).max(axis=1)          # 每采样点最大关节参考速度
        dt_med = float(np.median(np.diff(t_cmd_u))) / 1000.0
        k_sm = max(1, int(round(MOTION_SMOOTH_S / dt_med))) if dt_med > 0 else 1
        if k_sm > 1:                                  # 平滑, 抑制命令跳变产生的单点尖峰
            speed = np.convolve(speed, np.ones(k_sm) / k_sm, mode="same")
        thr = max(MOTION_SPEED_EPS, MOTION_SPEED_FRAC * float(np.percentile(speed, 99)))
        moving = speed > thr
        if moving.any():
            mask &= np.interp(t_sta, t_cmd_u, moving.astype(float)) > 0.5
            idx_mv = np.flatnonzero(moving)
            gap_idx = int(round(MOTION_GAP_S / dt_med)) if dt_med > 0 else 1
            cut = np.diff(idx_mv) > max(gap_idx, 1)
            starts = np.concatenate([[idx_mv[0]], idx_mv[1:][cut]])
            ends = np.concatenate([idx_mv[1:][cut], [idx_mv[-1]]])
            n_seg = len(starts)
            win_start_s = (t_cmd_u[starts[0]] - float(t_cmd[0])) / 1000.0
            win_end_s = (t_cmd_u[ends[-1]] - float(t_cmd[0])) / 1000.0
            win_dur_s = float(sum((t_cmd_u[e] - t_cmd_u[s]) / 1000.0
                                  for s, e in zip(starts, ends)))
            seg_list = [((t_cmd_u[s] - float(t_cmd[0])) / 1000.0,
                         (t_cmd_u[e] - float(t_cmd[0])) / 1000.0)
                        for s, e in zip(starts, ends)]

    n_sta_used = int(mask.sum())
    if n_sta_used < 10:
        raise ValueError(f"参考窗口内状态采样过少 ({n_sta_used} 行)")

    t_win = t_sta[mask]
    nj = len(joints)
    Pref_i = np.column_stack([np.interp(t_win, t_cmd_u, P_ref_u[:, k]) for k in range(nj)])
    Vref_i = np.column_stack([np.interp(t_win, t_cmd_u, V_ref_u[:, k]) for k in range(nj)])
    Psta_w, Vsta_w, Tsta_w = P_sta[mask], V_sta[mask], T_sta[mask]

    # --- 逐关节回归 ---
    results = []
    for k, j in enumerate(joints):
        dpos = Pref_i[:, k] - Psta_w[:, k]
        dvel = Vref_i[:, k] - Vsta_w[:, k]
        tor = Tsta_w[:, k]
        ok = np.isfinite(dpos) & np.isfinite(dvel) & np.isfinite(tor)

        dp_std = float(np.std(dpos[ok])) if ok.any() else np.nan
        dv_std = float(np.std(dvel[ok])) if ok.any() else np.nan
        tq_std = float(np.std(tor[ok])) if ok.any() else np.nan

        res = {"joint": j, "K": np.nan, "D": np.nan, "R2": np.nan, "RMSE": np.nan,
               "MAE": np.nan, "ExplVar": np.nan, "dp_std": dp_std, "dv_std": dv_std,
               "tq_std": tq_std, "valid": False}

        if int(ok.sum()) >= 10 and np.isfinite(tq_std) and tq_std >= 1e-12:
            X = np.column_stack([dpos[ok], dvel[ok]])
            y = tor[ok]
            try:
                beta, _, _, _ = np.linalg.lstsq(X, y, rcond=None)
                y_pred = X @ beta
                resid = y - y_pred
                ss_res = float(np.sum(resid ** 2))
                ss_tot = float(np.sum((y - np.mean(y)) ** 2))
                var_y = float(np.var(y))
                res.update({
                    "K": float(beta[0]), "D": float(beta[1]),
                    "R2": 1.0 - ss_res / ss_tot if ss_tot > 1e-15 else np.nan,
                    "RMSE": float(np.sqrt(ss_res / len(y))),
                    "MAE": float(np.mean(np.abs(resid))),
                    "ExplVar": 1.0 - float(np.var(resid)) / var_y if var_y > 1e-15 else np.nan,
                    "valid": True,
                })
            except np.linalg.LinAlgError:
                pass
        results.append(res)

    # --- 元信息 ---
    stages = None
    if "stage" in ci and "type" in ci:
        seen = {}
        for r in rows_cmd[i0:]:
            key = f"{r[ci['stage']]}/{r[ci['type']]}"
            seen[key] = seen.get(key, 0) + 1
        stages = "  ".join(f"{k}×{v}" for k, v in seen.items())

    meta = {
        "label": label,
        "n_cmd_total": len(rows_cmd), "n_cmd_used": int(len(t_cmd_u)),
        "n_sta_total": len(rows_sta), "n_sta_used": n_sta_used,
        "win_start_s": win_start_s,
        "win_end_s": win_end_s,
        "win_dur_s": win_dur_s,
        "n_seg": n_seg,
        "seg_list": seg_list,
        "trim_s": (t_start - float(t_cmd[0])) / 1000.0,
        "t_from": t_from, "t_to": t_to,
        "stages": stages,
        "motion_detected": motion_detected,
        "dropped_joints": [j for j in all_joints if j not in joints],
    }
    return results, joints, meta


# ============================================================
# 5. 分组辅助
# ============================================================
def build_scores(valid_results):
    if not valid_results:
        return None
    return {
        "r2_vals":   np.array([r["R2"] for r in valid_results if not np.isnan(r["R2"])]),
        "rmse_vals": np.array([r["RMSE"] for r in valid_results if not np.isnan(r["RMSE"])]),
        "k_vals":    np.array([r["K"] for r in valid_results]),
        "d_vals":    np.array([r["D"] for r in valid_results]),
    }


def fmt_metric(val, valid, nan_text="—"):
    """非空数值 → float, 否则返回 nan_text"""
    if valid and not np.isnan(val):
        return float(val)
    return nan_text


def window_text(meta):
    """分析窗口说明文本 (数据表 / Summary / 控制台共用)"""
    txt = (f"分析窗口: 实际运行段 +{meta['win_start_s']:.2f}s ~ +{meta['win_end_s']:.2f}s"
           f" (运动 {meta['win_dur_s']:.1f}s / {meta['n_seg']} 段)")
    if meta.get("t_from") is not None or meta.get("t_to") is not None:
        a = f"{meta['t_from']:.2f}s" if meta.get("t_from") is not None else "起始"
        b = f"{meta['t_to']:.2f}s" if meta.get("t_to") is not None else "结束"
        txt = f"指定区间 +{a} ~ +{b}  |  " + txt
    elif meta["trim_s"] > 0.01:
        txt += f"  |  已剔除前 {meta['trim_s']:.2f}s 命令未生效段"
    if meta.get("stages"):
        txt += f"  |  Stages: {meta['stages']}"
    return txt


def check_load_dominated(results):
    """
    负刚度预警: 稳定位置控制器不可能出现 K<0。
    若拟合质量好 (R²>0.8) 但 K<0, 说明力矩与位置误差符号相反,
    力矩由负载/工况主导, K 反映负载而非控制器刚度。
    返回负刚度且拟合良好的关节名列表。
    """
    out = []
    for r in results:
        if r["valid"] and r["K"] < 0 and not np.isnan(r["R2"]) and r["R2"] > 0.8:
            out.append(r["joint"])
    return out


# 指标说明文本
METRICS_LEGEND = [
    ("R² (决定系数)", "Coefficient of Determination",
     "衡量回归模型对观测值的拟合程度。取值范围 [0, 1]，越接近 1 表示模型解释力越强。"
     "R² = 1 - SS_res / SS_tot"),
    ("RMSE (均方根误差)", "Root Mean Squared Error",
     "预测值与真实值偏差的平方和与样本数比值的平方根。与扭矩同单位 (Nm)，越小越好。"
     "RMSE = sqrt(Σ(y_true - y_pred)² / n)"),
    ("MAE (平均绝对误差)", "Mean Absolute Error",
     "预测值与真实值偏差绝对值的平均值。与扭矩同单位 (Nm)，对异常值不如 RMSE 敏感。"
     "MAE = Σ|y_true - y_pred| / n"),
    ("ExplVar (可解释方差)", "Explained Variance Score",
     "衡量模型捕捉数据变异程度的能力。取值范围 (-∞, 1]，越接近 1 越好。"
     "ExplVar = 1 - Var(y_true - y_pred) / Var(y_true)"),
    ("K (刚度)", "Stiffness",
     "位置误差到力矩的映射系数。单位: Nm/rad。正值表示弹性恢复力。"),
    ("D (阻尼)", "Damping",
     "速度误差到力矩的映射系数。单位: Nm·s/rad。正值表示粘性阻尼。"),
]


# ============================================================
# 6. 写入数据 Sheet（每个会话一张，只含数据表）
# ============================================================
def write_joint_cells(ws, row, col0, name, r, valid_scores, band_fill):
    """写一个关节的 5 列: Joint / K / D / R² / RMSE"""
    kf = get_quality_font(r["K"], "K", valid_scores) if r["valid"] else None
    df = get_quality_font(r["D"], "D", valid_scores) if r["valid"] else None
    rf = get_quality_font(r["R2"], "R2", valid_scores) if r["valid"] else None
    mf = get_quality_font(r["RMSE"], "RMSE", valid_scores) if r["valid"] else None
    style_data_cell(ws, row, col0, name, fmt="", fill=band_fill, font=NORMAL_FONT)
    style_data_cell(ws, row, col0 + 1, fmt_metric(r["K"], r["valid"]),
                    is_warn=not r["valid"], fill=band_fill, font=kf)
    style_data_cell(ws, row, col0 + 2, fmt_metric(r["D"], r["valid"]),
                    is_warn=not r["valid"], fill=band_fill, font=df)
    style_data_cell(ws, row, col0 + 3, fmt_metric(r["R2"], r["valid"]),
                    fmt="0.0000" if r["valid"] else "", fill=band_fill, font=rf)
    style_data_cell(ws, row, col0 + 4, fmt_metric(r["RMSE"], r["valid"]),
                    fmt="0.0000" if r["valid"] else "",
                    is_warn=not r["valid"], fill=band_fill, font=mf)


def write_data_sheet(ws, results, joints, meta, valid_scores):
    groups, ordered, pairs, blocks = group_joints(joints)
    rmap = {r["joint"]: r for r in results}

    # 列宽预设 (数据表 B~M 统一 13)
    ws.column_dimensions["A"].width = 2
    for col_letter in ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]:
        ws.column_dimensions[col_letter].width = 13

    row = 1

    # -------- 标题 --------
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=13)
    title_cell = ws.cell(row=row, column=2,
                         value=f"Joint Stiffness & Damping   [{meta['label']}]")
    title_cell.font = TITLE_FONT
    title_cell.alignment = LEFT_ALIGN
    row += 1

    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=13)
    info_cell = ws.cell(row=row, column=2,
                        value=f"Joints: {len(joints)}  |  分组: "
                              f"{'-'.join(str(len(groups[g])) for g in ordered)}"
                              f"  |  Cmd: {meta['n_cmd_used']}/{meta['n_cmd_total']} 行"
                              f"  |  Sta: {meta['n_sta_used']}/{meta['n_sta_total']} 行")
    info_cell.font = Font(name="Microsoft YaHei", size=9, color="666666")
    row += 1

    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=13)
    ws.cell(row=row, column=2, value=window_text(meta)).font = Font(
        name="Microsoft YaHei", size=9, color="666666")
    row += 2

    # -------- 按身体部位顺序写块 (head → waist → arm_l↔arm_r → leg_l↔leg_r) --------
    for kind, gl, gr in blocks:
        if kind == "pair":
            ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=13)
            ws.cell(row=row, column=2,
                    value=f"{gl} ↔ {gr} 并排对比   "
                          f"({len(groups[gl])}-{len(groups[gr])})").font = GROUP_FONT
            row += 1

            sb_headers = ["Joint", "K", "D", "R²", "RMSE", "", "Joint", "K", "D", "R²", "RMSE"]
            for ci_, h in enumerate(sb_headers):
                if h:
                    ws.cell(row=row, column=2 + ci_, value=h)
            style_header_row(ws, row, 2, 12)
            ws.cell(row=row, column=7).fill = SEP_FILL
            row += 1

            n_rows = max(len(groups[gl]), len(groups[gr]))
            for i in range(n_rows):
                band_fill = DARK_FILL if i % 2 == 1 else LIGHT_FILL
                if i < len(groups[gl]):
                    write_joint_cells(ws, row, 2, groups[gl][i], rmap[groups[gl][i]],
                                      valid_scores, band_fill)
                if i < len(groups[gr]):
                    write_joint_cells(ws, row, 8, groups[gr][i], rmap[groups[gr][i]],
                                      valid_scores, band_fill)
                ws.cell(row=row, column=7).border = THIN_BORDER
                ws.cell(row=row, column=7).fill = SEP_FILL
                row += 1
            row += 1
        else:
            g = gl
            ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=7)
            ws.cell(row=row, column=2,
                    value=f"{g}  ({len(groups[g])} joints)").font = GROUP_FONT
            row += 1

            for ci_, h in enumerate(["Joint", "K", "D", "R²", "RMSE"]):
                ws.cell(row=row, column=2 + ci_, value=h)
            style_header_row(ws, row, 2, 6)
            row += 1

            for i, jname in enumerate(groups[g]):
                band_fill = DARK_FILL if i % 2 == 1 else LIGHT_FILL
                write_joint_cells(ws, row, 2, jname, rmap[jname], valid_scores, band_fill)
                row += 1
            row += 1

    ws.freeze_panes = "A1"


# ============================================================
# 7. 写入汇总 Sheet（首表：Summary + Legend）
# ============================================================
def write_summary_sheet(ws, meta, results, joints):
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 26
    ws.column_dimensions["D"].width = 38
    ws.column_dimensions["E"].width = 26

    groups, ordered, _, _ = group_joints(joints)
    valid = [r for r in results if r["valid"]]
    invalid = [r for r in results if not r["valid"]]

    row = 1

    # -------- 标题 --------
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=5)
    ws.cell(row=row, column=2,
            value="Joint Stiffness & Damping — Summary & Legend").font = TITLE_FONT
    row += 2

    # -------- Summary --------
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=5)
    ws.cell(row=row, column=2,
            value=f"Summary   [{meta['label']}]   Joints: {len(joints)}  "
                  f"分组: {'-'.join(str(len(groups[g])) for g in ordered)}  "
                  f"Cmd: {meta['n_cmd_used']}/{meta['n_cmd_total']}  "
                  f"Sta: {meta['n_sta_used']}/{meta['n_sta_total']}").font = GROUP_FONT
    row += 1

    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=5)
    ws.cell(row=row, column=2, value=window_text(meta)).font = Font(
        name="Microsoft YaHei", size=9, color="666666")
    row += 1

    stats_headers = ["Metric", "Mean ± Std", "Range (Min / Max)", "说明"]
    for ci_, h in enumerate(stats_headers):
        ws.cell(row=row, column=2 + ci_, value=h)
    style_header_row(ws, row, 2, 5)
    row += 1

    if valid:
        Ks    = np.array([r["K"] for r in valid])
        Ds    = np.array([r["D"] for r in valid])
        R2s   = np.array([r["R2"] for r in valid if not np.isnan(r["R2"])])
        RMSEs = np.array([r["RMSE"] for r in valid if not np.isnan(r["RMSE"])])
        MAEs  = np.array([r["MAE"] for r in valid if not np.isnan(r["MAE"])])

        invalid_txt = ("无激励(τ≡0)关节: " + ", ".join(r["joint"] for r in invalid)) \
            if invalid else "无"

        stat_rows = [
            ("Valid Joints",
             f"{len(valid)} / {len(joints)}",
             "",
             invalid_txt),
            ("Stiffness K (Nm/rad)",
             f"{Ks.mean():.4f} ± {Ks.std():.4f}",
             f"[{Ks.min():.4f}, {Ks.max():.4f}]",
             "位置误差→力矩映射系数"),
            ("Damping D (Nm·s/rad)",
             f"{Ds.mean():.4f} ± {Ds.std():.4f}",
             f"[{Ds.min():.4f}, {Ds.max():.4f}]",
             "速度误差→力矩映射系数"),
        ]
        if len(R2s) > 0:
            stat_rows.append(("R² (决定系数)",
                              f"{R2s.mean():.4f} ± {R2s.std():.4f}",
                              f"[{R2s.min():.4f}, {R2s.max():.4f}]",
                              "1=完美拟合, 0=无解释力"))
        if len(RMSEs) > 0:
            stat_rows.append(("RMSE (均方根误差, Nm)",
                              f"{RMSEs.mean():.4f} ± {RMSEs.std():.4f}",
                              f"[{RMSEs.min():.4f}, {RMSEs.max():.4f}]",
                              "越小越好, 对异常值敏感"))
        if len(MAEs) > 0:
            stat_rows.append(("MAE (平均绝对误差, Nm)",
                              f"{MAEs.mean():.4f} ± {MAEs.std():.4f}",
                              f"[{MAEs.min():.4f}, {MAEs.max():.4f}]",
                              "越小越好, 对异常值不敏感"))

        stat_start = row
        for sr in stat_rows:
            for ci_, v in enumerate(sr):
                cell = ws.cell(row=row, column=2 + ci_, value=v)
                cell.font = Font(name="Consolas", bold=(ci_ == 0), size=10)
                cell.alignment = CENTER_ALIGN if ci_ > 0 else LEFT_ALIGN
                cell.border = THIN_BORDER
            ws.row_dimensions[row].height = 22
            row += 1
        apply_banding(ws, stat_start, row - 1, 2, 5)

        neg = check_load_dominated(results)
        if neg:
            ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=5)
            cell = ws.cell(row=row, column=2,
                           value=f"⚠ {len(neg)} 个关节 K<0 且 R²>0.8 "
                                 f"({', '.join(neg)}) — 力矩与位置误差符号相反, "
                                 f"该段力矩由负载/工况主导, K 不代表控制器刚度")
            cell.font = Font(name="Microsoft YaHei", size=9, color=BAD_COLOR)
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            ws.row_dimensions[row].height = 28
            row += 1
    else:
        ws.cell(row=row, column=2, value="No valid joints").font = WARN_FONT
        row += 1

    row += 1

    # -------- 指标说明 Legend --------
    row += 1
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
    ws.cell(row=row, column=2, value="Metrics Legend").font = GROUP_FONT
    row += 1

    legend_headers = ["指标", "英文名", "含义 & 公式"]
    for ci_, h in enumerate(legend_headers):
        ws.cell(row=row, column=2 + ci_, value=h)
    style_header_row(ws, row, 2, 4)
    row += 1

    LEGEND_FILL_LIGHT = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    LEGEND_FILL_DARK  = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    for li, (name_cn, name_en, desc) in enumerate(METRICS_LEGEND):
        fill = LEGEND_FILL_DARK if li % 2 == 1 else LEGEND_FILL_LIGHT
        vals = [name_cn, name_en, desc]
        for ci_, v in enumerate(vals):
            cell = ws.cell(row=row, column=2 + ci_, value=v)
            cell.font = Font(name="Microsoft YaHei", bold=(ci_ == 0), size=9,
                             color="333333")
            cell.alignment = LEFT_ALIGN if ci_ == 2 else CENTER_ALIGN
            cell.border = THIN_BORDER
            cell.fill = fill
        desc_cell = ws.cell(row=row, column=4)
        desc_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.row_dimensions[row].height = 50
        row += 1

    # 字体颜色说明
    row += 1
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
    ws.cell(row=row, column=2, value="字体颜色说明").font = Font(
        name="Microsoft YaHei", bold=True, size=9, color="333333")
    row += 1
    color_legend = [
        ("深绿 006100", GOOD_COLOR,  "正常 / 较好"),
        ("棕色 9C6500", WEAK_COLOR,  "一般 / 偏弱 (R²<0.8, z-score 1~2, 或 RMSE>P75)"),
        ("深红 C00000", BAD_COLOR,   "异常 / 较差 (R²<0.6, z-score>2, 或 RMSE>P90)"),
        ("灰色 999999", "999999",    "无效数据 (力矩恒为零 / 无激励)"),
    ]
    for cl_name, cl_hex, cl_desc in color_legend:
        cell_b = ws.cell(row=row, column=2, value=cl_name)
        cell_b.font = Font(name="Consolas", size=10, color=cl_hex, bold=True)
        cell_b.alignment = CENTER_ALIGN
        cell_b.border = THIN_BORDER

        ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=4)
        cell_c = ws.cell(row=row, column=3, value=cl_desc)
        cell_c.font = Font(name="Microsoft YaHei", size=9, color="666666")
        cell_c.alignment = LEFT_ALIGN
        cell_c.border = THIN_BORDER
        ws.row_dimensions[row].height = 18
        row += 1

    ws.freeze_panes = "A1"


def safe_sheet_name(name):
    name = re.sub(r"[\[\]:*?/\\]", "_", name)
    return name[:31] or "Data"


# ============================================================
# 8. Main
# ============================================================
def parse_time_arg(s):
    """解析时间参数: '33.8' / '+33.8s' → 秒 (相对日志起始)"""
    if s is None:
        return None
    txt = s.strip().lower().rstrip("s").lstrip("+")
    try:
        return float(txt)
    except ValueError:
        raise argparse.ArgumentTypeError(
            f"无法解析时间 '{s}' (示例: 33.8 或 +33.8s)")


def main():
    parser = argparse.ArgumentParser(
        description="Joint Stiffness & Damping 回归计算 (新版会话日志)",
        epilog="示例: uv run compute.py -C ./20261014_094915 -o ./results",
    )
    parser.add_argument("-C", "--cwd", default=".",
                        help="会话目录 / 含会话目录的父目录 (默认: 当前目录)")
    parser.add_argument("-o", "--output-dir", default=None,
                        help="XLSX 输出目录 (默认: 与输入相同)")
    parser.add_argument("--no-trim", action="store_true",
                        help="不筛选分析窗口, 使用全部重叠数据 "
                             "(默认只用参考轨迹实际运动段)")
    parser.add_argument("--start", type=parse_time_arg, default=None,
                        help="分析起始时刻 (相对日志起始的秒数, 如 33.8 或 +33.8s)")
    parser.add_argument("--end", type=parse_time_arg, default=None,
                        help="分析结束时刻 (相对日志起始的秒数)")
    args = parser.parse_args()

    work_dir = os.path.abspath(args.cwd)
    out_dir = os.path.abspath(args.output_dir) if args.output_dir else work_dir

    sessions = find_session_dirs(work_dir)
    if not sessions:
        print(f"ERROR: 在 '{work_dir}' 中未找到会话目录 "
              f"(需包含 recorder/*_joint_commands.csv 与 *_joint_states.csv)")
        sys.exit(1)

    n_ok = 0
    for sess in sessions:
        label = session_label(sess)
        pair = find_csv_pair(sess)
        if pair is None:
            print(f"\nSKIP: {label} — recorder 目录缺少 CSV 配对")
            continue
        cmd_path, sta_path = pair

        print(f"\nProcessing: {label}")
        print(f"  Commands: {os.path.basename(cmd_path)}")
        print(f"  States:   {os.path.basename(sta_path)}")

        try:
            results, joints, meta = load_and_compute(
                cmd_path, sta_path, label, trim=not args.no_trim,
                t_from=args.start, t_to=args.end)
        except Exception as e:
            print(f"  ERROR: {e}")
            continue

        groups, ordered, _, _ = group_joints(joints)
        valid = [r for r in results if r["valid"]]
        print(f"  Joints: {len(joints)}  Groups: "
              f"{'-'.join(str(len(groups[g])) for g in ordered)}  Valid: {len(valid)}")
        print(f"  Window: 运行段 +{meta['win_start_s']:.2f}s ~ +{meta['win_end_s']:.2f}s "
              f"(运动 {meta['win_dur_s']:.1f}s / {meta['n_seg']} 段, "
              f"{meta['n_sta_used']} samples)")
        if meta.get("stages"):
            print(f"  Stages: {meta['stages']}")
        if not meta.get("motion_detected", True):
            print("  WARN: 未检测到参考轨迹运动, 回归结果可能不可用")
        neg = check_load_dominated(results)
        if neg:
            print(f"  WARN: {len(neg)} 个关节 K<0 且拟合良好 ({', '.join(neg[:6])}"
                  f"{' ...' if len(neg) > 6 else ''})")
            print("        → 力矩与位置误差符号相反 (稳定位置控制器不可能如此),")
            print("          该段力矩由负载/工况主导, K 不代表控制器刚度")

        cross_scores = build_scores(valid)

        wb = Workbook()
        wb.remove(wb.active)

        ws_summary = wb.create_sheet(title="Summary")
        write_summary_sheet(ws_summary, meta, results, joints)

        ws = wb.create_sheet(title=safe_sheet_name(label))
        write_data_sheet(ws, results, joints, meta, cross_scores)

        os.makedirs(out_dir, exist_ok=True)
        if args.start is not None or args.end is not None:
            a = f"{args.start:g}" if args.start is not None else "0"
            b = f"{args.end:g}" if args.end is not None else "end"
            out_name = f"stiffness_damping_{label}_{a}-{b}s.xlsx"
        else:
            out_name = f"stiffness_damping_{label}.xlsx"
        output_path = os.path.join(out_dir, out_name)
        wb.save(output_path)
        print(f"  Saved: {os.path.basename(output_path)}")

        # 控制台速览
        print(f"  {'joint':<9} {'K':>10} {'D':>10} {'R2':>7}")
        for r in results:
            if r["valid"]:
                print(f"  {r['joint']:<9} {r['K']:>10.3f} {r['D']:>10.3f} {r['R2']:>7.3f}")
            else:
                print(f"  {r['joint']:<9} {'—':>10} {'—':>10} {'—':>7}")
        n_ok += 1

    print(f"\nDone. {n_ok} file(s) generated.")


if __name__ == "__main__":
    main()
