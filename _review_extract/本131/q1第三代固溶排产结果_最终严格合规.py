import pandas as pd
import numpy as np
from openai import OpenAI
import time
import sys
import os

# ================= 配置 =================
API_KEY = os.getenv("DEEPSEEK_API_KEY", "")
client = OpenAI(api_key=API_KEY, base_url="https://api.deepseek.com/v1")

INPUT_FILE = "清洗完成_排产数据.xlsx"
OUTPUT_FILE = "问题1_固溶排产结果_最终严格合规.xlsx"

# 合炉厚度差允许（mm）
THICK_ALLOW = {
    ("硬铝系", "<20"): 0, ("硬铝系", "20-50"): 5,
    ("硬铝系", "50-100"): 10, ("硬铝系", "100-150"): 40,
    ("硬铝系", "150-260"): 60,
    ("铝镁硅系", "<30"): 0, ("铝镁硅系", "30-50"): 5,
    ("铝镁硅系", "50-100"): 10, ("铝镁硅系", "100-150"): 40,
    ("铝镁硅系", "150-260"): 60,
}

# ================= 辅助函数 =================
def thickness_diff_pct(t1, t2):
    if t1 == t2:
        return 0
    thick = max(t1, t2)
    return abs(t1 - t2) / thick * 100

def width_ok(w1, w2):
    if w1 >= w2:
        return (w1 - w2) <= 800
    else:
        return (w2 - w1) <= 500

def can_adjacent(row1, row2):
    if thickness_diff_pct(row1["入口厚度(mm)"], row2["入口厚度(mm)"]) > 30:
        return False
    if not width_ok(row1["入口宽度(mm)"], row2["入口宽度(mm)"]):
        return False
    return True

def call_api_sort(plate_list):
    """调用DeepSeek API排序，失败则降级为本地贪心排序"""
    if len(plate_list) <= 1:
        return plate_list

    print(f"      → 调用API排序，大小={len(plate_list)}...", end="", flush=True)
    plates_text = ""
    for i, p in enumerate(plate_list):
        plates_text += (f"板材{i}: 合金{p['合金牌号']}, 厚度={p['入口厚度(mm)']:.2f}mm, "
                        f"宽度={p['入口宽度(mm)']}mm, 紧急={p['紧急标记']}, 交期={p['交货日期']}\n")
    prompt = f"""排产专家。将以下{len(plate_list)}块板材排成顺序，要求：
1. 相邻：厚度差≤30%，宽度差：宽→窄≤800，窄→宽≤500。
2. 优先紧急高、交期早。
输出数字列表，如 [2,0,3,1]。只输出列表。

板材：
{plates_text}
"""
    try:
        resp = client.chat.completions.create(
            model="deepseek-chat",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.2,
            max_tokens=500,
            timeout=30
        )
        result = resp.choices[0].message.content.strip()
        if result.startswith('['):
            result = result[1:-1]
        order = [int(x.strip()) for x in result.split(',')]
        if len(order) != len(plate_list):
            raise ValueError
        sorted_plates = [plate_list[i] for i in order]
        for i in range(len(sorted_plates)-1):
            if not can_adjacent(sorted_plates[i], sorted_plates[i+1]):
                print(" 不满足约束，回退本地排序")
                return sorted(plate_list, key=lambda x: (x["入口厚度(mm)"], x["入口宽度(mm)"]))
        print(" 成功")
        return sorted_plates
    except Exception as e:
        print(f" 失败({e})，使用本地排序")
        return sorted(plate_list, key=lambda x: (x["入口厚度(mm)"], x["入口宽度(mm)"]))

# ================= 数据准备 =================
df = pd.read_excel(INPUT_FILE)
print(f"读取数据：{len(df)} 条")

# 确保排序优先级字段
if "优先级分数" not in df.columns:
    df["优先级分数"] = 0
df["排序分"] = df["紧急标记"] * 10000 + df["拖期标记"] * 1000 - pd.to_datetime(df["交货日期"]).apply(lambda x: x.toordinal())

temperatures = sorted(df["板温度设定值"].unique())
print(f"温度值: {temperatures}")

# ================= 核心分组：严格隔离 must_alone =================
must_alone_batches = []   # 每个元素是 (温度, DataFrame)
combine_batches = []      # 每个元素是 (温度, DataFrame)

for temp in temperatures:
    df_temp = df[df["板温度设定值"] == temp].copy()
    # 必须单独一炉的板材
    alone = df_temp[df_temp["必须单独一炉"] == 1]
    for _, row in alone.iterrows():
        must_alone_batches.append((temp, pd.DataFrame([row])))
    # 可合炉的板材
    can_combine = df_temp[df_temp["必须单独一炉"] == 0]
    if can_combine.empty:
        continue
    # 按合金大类 + 厚度组_合炉 分组
    for (alloy, thick_interval), group in can_combine.groupby(["合金大类", "厚度组_合炉"]):
        allow = THICK_ALLOW.get((alloy, thick_interval), 0)
        if allow == 0:
            # 理论不会出现，安全处理
            for _, row in group.iterrows():
                combine_batches.append((temp, pd.DataFrame([row])))
            continue
        # 按厚度排序，拆分厚度极差≤allow的小组
        group_sorted = group.sort_values("入口厚度(mm)")
        sub_groups = []
        current = []
        min_t = None
        for _, row in group_sorted.iterrows():
            if not current:
                current.append(row)
                min_t = row["入口厚度(mm)"]
            else:
                if row["入口厚度(mm)"] - min_t <= allow:
                    current.append(row)
                else:
                    sub_groups.append(pd.DataFrame(current))
                    current = [row]
                    min_t = row["入口厚度(mm)"]
        if current:
            sub_groups.append(pd.DataFrame(current))
        for sub in sub_groups:
            if len(sub) <= 1:
                combine_batches.append((temp, sub))
            else:
                plates = sub.to_dict('records')
                sorted_plates = call_api_sort(plates)
                sorted_df = pd.DataFrame(sorted_plates)
                combine_batches.append((temp, sorted_df))
                time.sleep(0.2)

# ================= 合并所有炉批，但保持独立 =================
all_batches = must_alone_batches + combine_batches
print(f"共生成 {len(all_batches)} 个炉批（其中单独炉批 {len(must_alone_batches)} 个）")

# ================= 炉批间排序 =================
# 为每个炉批计算温度、最大优先级
batch_info = []
for temp, batch_df in all_batches:
    max_pri = batch_df["排序分"].max()
    batch_info.append((temp, max_pri, batch_df))

# 按温度升序，同温内按优先级降序
batch_info.sort(key=lambda x: (x[0], -x[1]))

# 构建最终序列（不合并不同炉批）
final_batches = []
for temp, pri, batch_df in batch_info:
    final_batches.append(batch_df)

result_df = pd.concat(final_batches, ignore_index=True)
result_df.insert(0, "全局序号", range(1, len(result_df)+1))

# ================= 重新划分炉批号 =================
# 注意：我们已经保证了每个 batch 内部满足约束，且必须单独一炉的 batch 只有1行
# 因此直接按 original batch 顺序分配炉批号即可
batch_ids = []
current_batch = 1
for i, (temp, pri, batch_df) in enumerate(batch_info):
    batch_ids.extend([current_batch] * len(batch_df))
    current_batch += 1
result_df["炉批号"] = batch_ids

# ================= 固溶工艺描述 =================
sol_process = []
for batch in result_df["炉批号"].unique():
    mask = result_df["炉批号"] == batch
    max_thick = result_df.loc[mask, "入口厚度(mm)"].max()
    desc = f"固溶工艺采用厚度{max_thick:.1f}mm对应的参数"
    sol_process.extend([desc] * mask.sum())
result_df["固溶工艺"] = sol_process

# ================= 保存结果 =================
result_df.to_excel(OUTPUT_FILE, index=False)
print(f"\n✅ 排产完成！结果保存至 {OUTPUT_FILE}")
print(f"总板材数: {len(result_df)}")
print(f"总炉批数: {result_df['炉批号'].nunique()}")

# ================= 严格自检（违规即报错） =================
print("\n🔍 开始自检...")
violations = 0

for batch in result_df["炉批号"].unique():
    batch_df = result_df[result_df["炉批号"] == batch]
    # 1. 检查必须单独一炉的板材是否独立
    if any(batch_df["必须单独一炉"] == 1) and len(batch_df) > 1:
        print(f"  ❌ 炉批{batch}包含{len(batch_df)}块板材，但含有必须单独一炉的物料")
        violations += 1
    # 2. 检查同批内是否混合不同厚度区间
    for alloy, sub in batch_df.groupby("合金大类"):
        intervals = sub["厚度组_合炉"].unique()
        if len(intervals) > 1:
            print(f"  ❌ 炉批{batch}中合金{alloy}混合厚度区间 {intervals}")
            violations += 1
    # 3. 检查厚度极差是否超过允许值
    if len(batch_df) > 1:
        alloy = batch_df["合金大类"].iloc[0]
        thick_interval = batch_df["厚度组_合炉"].iloc[0]
        allow = THICK_ALLOW.get((alloy, thick_interval), 0)
        if allow > 0:
            thick_range = batch_df["入口厚度(mm)"].max() - batch_df["入口厚度(mm)"].min()
            if thick_range > allow:
                print(f"  ❌ 炉批{batch}厚度极差{thick_range:.1f}mm > 允许{allow}mm")
                violations += 1
    # 4. 检查相邻板材约束（每个炉批内部）
    if len(batch_df) > 1:
        for i in range(len(batch_df)-1):
            row1 = batch_df.iloc[i]
            row2 = batch_df.iloc[i+1]
            if not can_adjacent(row1, row2):
                print(f"  ❌ 炉批{batch}中相邻板材 {row1['入口材料号']} → {row2['入口材料号']} 不满足约束")
                violations += 1

if violations == 0:
    print("  ✅ 所有检查通过，排产结果完全合规！")
else:
    print(f"  ⚠️ 发现 {violations} 处违规，请检查代码逻辑。")
    sys.exit(1)

print("\n🎉 程序成功结束，输出文件已生成。")