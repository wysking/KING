import pandas as pd
import numpy as np

# 读取 API优化.xlsx 文件
df = pd.read_excel("问题2_淬火排产结果_API优化.xlsx", engine="openpyxl")

# ================== 1. 提取炉批16的所有物料 ==================
batch16_mask = df["炉批号"] == 16
batch16_df = df[batch16_mask].copy()
other_df = df[~batch16_mask].copy()

# ================== 2. 定义硬铝系合炉规则 ==================
# 厚度区间及允许最大厚度差（硬铝系：7050,7075,2A12,2024等）
hard_rules = [
    (0, 20, None),      # <20mm 不合炉（必须单独）
    (20, 51, 5),        # 20-50mm 允许差≤5mm
    (50, 101, 10),      # 50-100mm 允许差≤10mm
    (100, 151, 40),     # 100-150mm 允许差≤40mm
    (150, 261, 60)      # 150-260mm 允许差≤60mm
]

def get_thickness_interval(thickness):
    for low, high, gap in hard_rules:
        if low <= thickness < high:
            return low, high, gap
    return None, None, None

def can_merge_in_batch(m1, m2):
    """同炉批内两块物料是否可以合并（同温度，同合金大类，厚度差≤允许差）"""
    if m1["板温度设定值"] != m2["板温度设定值"]:
        return False
    if m1["合金大类"] != m2["合金大类"]:
        return False
    # 必须单独一炉的不能合并（但炉批16中所有物料的“必须单独一炉”都是0或1？实际上原始数据中<20mm的应为1）
    # 这里简单检查厚度区间
    _, _, gap1 = get_thickness_interval(m1["入口厚度(mm)"])
    _, _, gap2 = get_thickness_interval(m2["入口厚度(mm)"])
    if gap1 is None or gap2 is None:
        return False
    max_gap = min(gap1, gap2)
    return abs(m1["入口厚度(mm)"] - m2["入口厚度(mm)"]) <= max_gap

# ================== 3. 对炉批16内的物料重新分组 ==================
materials = batch16_df.to_dict('records')
# 按温度、合金大类分组（炉批16内都是335°C，硬铝系）
# 进一步按厚度所在区间分组
interval_groups = {}
for m in materials:
    low, high, _ = get_thickness_interval(m["入口厚度(mm)"])
    if low is None:
        # 异常情况，单独成批
        key = (m["入口厚度(mm)"],)
    else:
        key = (low, high)
    interval_groups.setdefault(key, []).append(m)

# 对每个厚度区间内的物料，按厚度排序后贪心合并（满足厚度差）
new_batches = []
for interval, group in interval_groups.items():
    if interval[1] - interval[0] <= 10:  # 窄区间（如<20），必须单独一炉
        for m in group:
            new_batches.append([m])
    else:
        # 按厚度排序
        group.sort(key=lambda x: x["入口厚度(mm)"])
        batches = []
        for m in group:
            placed = False
            for batch in batches:
                if all(can_merge_in_batch(m, b) for b in batch):
                    # 检查新batch内厚度差是否仍≤允许差（取当前batch最大最小厚度）
                    batch_thicks = [b["入口厚度(mm)"] for b in batch] + [m["入口厚度(mm)"]]
                    min_t, max_t = min(batch_thicks), max(batch_thicks)
                    # 获取允许的最大差（取两个区间允许差的最小值）
                    _, _, gap_rule = get_thickness_interval(min_t)
                    if gap_rule is not None and (max_t - min_t) <= gap_rule:
                        batch.append(m)
                        placed = True
                        break
            if not placed:
                batches.append([m])
        new_batches.extend(batches)

# ================== 4. 为拆分后的炉批分配新炉批号 ==================
# 获取当前最大炉批号（从other_df中）
max_batch = other_df["炉批号"].max()
if pd.isna(max_batch):
    max_batch = 0
new_batch_id = max_batch + 1

new_batch_records = []
for batch in new_batches:
    batch_id = new_batch_id
    new_batch_id += 1
    for m in batch:
        m["炉批号"] = batch_id
        new_batch_records.append(m)

new_batch_df = pd.DataFrame(new_batch_records)

# ================== 5. 合并其他炉批和新的炉批 ==================
df_fixed = pd.concat([other_df, new_batch_df], ignore_index=True)

# ================== 6. 重新排序（按温度升序，同温内按炉批中位数厚度降序） ==================
# 计算每个炉批的厚度中位数和温度
batch_stats = df_fixed.groupby("炉批号").agg(
    温度=("板温度设定值", "first"),
    厚度中位数=("入口厚度(mm)", "median")
).reset_index()
# 按温度升序，同温内按厚度中位数降序
batch_stats = batch_stats.sort_values(["温度", "厚度中位数"], ascending=[True, False])
ordered_batches = batch_stats["炉批号"].tolist()

# 重新分配全局顺序
df_fixed["全局顺序"] = -1
global_seq = 1
for batch_id in ordered_batches:
    indices = df_fixed[df_fixed["炉批号"] == batch_id].index
    df_fixed.loc[indices, "全局顺序"] = range(global_seq, global_seq + len(indices))
    global_seq += len(indices)

df_fixed = df_fixed.sort_values("全局顺序").reset_index(drop=True)

# ================== 7. 重新生成淬火工艺说明 ==================
def get_quench_param(batch_df):
    if batch_df.empty:
        return ""
    thicknesses = batch_df["入口厚度(mm)"].tolist()
    if any(t >= 50 for t in thicknesses):
        ref_thick = max(thicknesses)
        param_type = "最厚厚度"
    else:
        ref_thick = np.median(thicknesses)
        param_type = "中间厚度"
    return f"淬火工艺采用厚度{ref_thick:.1f}mm对应的参数（{param_type}）"

df_fixed["淬火工艺"] = ""
for batch_id, group in df_fixed.groupby("炉批号"):
    quench_str = get_quench_param(group)
    df_fixed.loc[group.index, "淬火工艺"] = quench_str

# ================== 8. 保存结果 ==================
output_file = "问题2_淬火排产结果_修正版.xlsx"
df_fixed.to_excel(output_file, index=False, engine="openpyxl")
print(f"修正后的排产结果已保存至 {output_file}")
print(f"原有炉批16被拆分为 {len(new_batches)} 个新炉批，总炉批数：{df_fixed['炉批号'].nunique()}")