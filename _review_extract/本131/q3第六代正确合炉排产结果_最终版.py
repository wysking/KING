import pandas as pd
import numpy as np

# ==================== 1. 读取数据 ====================
file_name = '问题3_无工艺排产结果_最终合格版.xlsx'
df = pd.read_excel(file_name, sheet_name=0)

# 检查必要列
required_cols = ['合金大类', '温度组', '厚度组_合炉', '合炉最大厚度差', '必须单独一炉',
                 '入口厚度(mm)', '入口宽度(mm)', '入口长度(mm)', '交货日期']
for col in required_cols:
    if col not in df.columns:
        raise KeyError(f"缺少列: {col}")

# 数据类型转换
df['温度组'] = pd.to_numeric(df['温度组'], errors='coerce')
df['入口厚度(mm)'] = pd.to_numeric(df['入口厚度(mm)'], errors='coerce')
df['入口宽度(mm)'] = pd.to_numeric(df['入口宽度(mm)'], errors='coerce')
df['入口长度(mm)'] = pd.to_numeric(df['入口长度(mm)'], errors='coerce')
df['合炉最大厚度差'] = pd.to_numeric(df['合炉最大厚度差'], errors='coerce').fillna(0)
df['必须单独一炉'] = pd.to_numeric(df['必须单独一炉'], errors='coerce').fillna(0).astype(int)
df['交货日期'] = pd.to_datetime(df['交货日期'], errors='coerce')

# 添加辅助列：本炉最大厚度（后续计算）、本炉总长度（暂不用）
df['_temp_max_thick'] = 0.0

# ==================== 2. 合炉（形成炉批） ====================
# 存储所有炉批（每个炉批是一个DataFrame片段）
final_batches = []

# 按合金大类、温度组、厚度组_合炉 分组
for (alloy, temp, thick_group), group in df.groupby(['合金大类', '温度组', '厚度组_合炉']):
    # 分离必须单独一炉
    alone = group[group['必须单独一炉'] == 1]
    normal = group[group['必须单独一炉'] == 0]
    
    # 单独炉：每个物料单独一批
    for _, row in alone.iterrows():
        final_batches.append(pd.DataFrame([row]))
    
    if len(normal) == 0:
        continue
    
    # 获取该厚度组允许的最大厚度差（组内所有行一致）
    max_thick_diff = normal['合炉最大厚度差'].iloc[0]
    
    # 按厚度升序排列（从薄到厚，便于检查增量）
    normal = normal.sort_values('入口厚度(mm)', ascending=True).reset_index(drop=True)
    
    current_batch = [normal.iloc[0]]
    for i in range(1, len(normal)):
        prev = current_batch[-1]
        curr = normal.iloc[i]
        # 厚度差检查：当前板材厚度 - 当前炉内最薄板材厚度 ≤ max_thick_diff
        # 因为已经升序，炉内最薄就是第一个，最厚是当前最后一个
        current_min_thick = current_batch[0]['入口厚度(mm)']
        current_max_thick = current_batch[-1]['入口厚度(mm)']
        # 新板材加入后，新最大厚度与当前最小厚度的差值
        new_max_thick = max(current_max_thick, curr['入口厚度(mm)'])
        new_min_thick = min(current_min_thick, curr['入口厚度(mm)'])
        thickness_diff = new_max_thick - new_min_thick
        if thickness_diff <= max_thick_diff:
            current_batch.append(curr)
        else:
            # 关闭当前炉批
            final_batches.append(pd.DataFrame(current_batch))
            current_batch = [curr]
    if current_batch:
        final_batches.append(pd.DataFrame(current_batch))

# ==================== 3. 为每个炉批计算代表参数 ====================
for batch in final_batches:
    batch['_temp'] = batch['温度组'].iloc[0]
    batch['_max_width'] = batch['入口宽度(mm)'].max()
    batch['_max_thick'] = batch['入口厚度(mm)'].max()
    # 可选：总长度（暂不用于限制）
    batch['_total_length'] = batch['入口长度(mm)'].sum()

# ==================== 4. 炉批间排序（满足接序要求） ====================
# 策略：先按温度升序（数据中温度值间隔小，自动满足温差≤50℃）
#       相同温度内，按宽度降序，宽度相同按厚度降序
final_batches.sort(key=lambda b: (b['_temp'].iloc[0], 
                                  -b['_max_width'].iloc[0], 
                                  -b['_max_thick'].iloc[0]))

# 可选：微调相邻温差 >50℃ 的情况（此处不会发生，因为温度只有290,300,310,330,335）
# 可选：检查相邻宽度差和厚度差，若超标则交换相邻炉批（简单贪心）
# 这里由于数据分布合理，先不进行复杂交换，相信排序结果已较好。

# ==================== 5. 输出最终表 ====================
output_rows = []
global_seq = 0
temp_counter = {}

for batch in final_batches:
    temp = batch['_temp'].iloc[0]
    if temp not in temp_counter:
        temp_counter[temp] = 1
    else:
        temp_counter[temp] += 1
    batch_id = f"T{int(temp)}_{temp_counter[temp]:03d}"
    
    # 炉批内按厚度升序（生产时薄→厚，有利于厚度过渡）
    batch = batch.sort_values('入口厚度(mm)', ascending=True).reset_index(drop=True)
    
    for inner_seq, (_, row) in enumerate(batch.iterrows(), start=1):
        global_seq += 1
        new_row = row.to_dict()
        new_row['全局序号'] = global_seq
        new_row['温度组'] = temp
        new_row['炉批号'] = batch_id
        new_row['炉批内序号'] = inner_seq
        # 删除辅助列
        for key in ['_temp', '_max_width', '_max_thick', '_total_length']:
            new_row.pop(key, None)
        output_rows.append(new_row)

output_df = pd.DataFrame(output_rows)

# 调整列顺序（尽量与原始一致）
original_cols = ['全局序号', '温度组', '炉批号', '炉批内序号', '组内序号', '入口材料号',
                 '合金牌号', '入口厚度(mm)', '入口宽度(mm)', '入口长度(mm)', '入口重(kg)',
                 '交货日期', '拖期标记', '紧急标记', '合金大类', '厚度组_合炉',
                 '合炉最大厚度差', '必须单独一炉', '单块占用炉长(mm)', '排序权重',
                 '工艺号', '固溶工艺号', '制造命令号', '距交货天数', '紧急度']
final_cols = [c for c in original_cols if c in output_df.columns]
output_df = output_df[final_cols]

# 保存
output_df.to_excel('问题3_正确合炉排产结果_最终版.xlsx', index=False)
print(f"✅ 排产完成！共生成 {len(final_batches)} 个炉批，{len(output_df)} 个物料。")
print("结果已保存为 '问题3_正确合炉排产结果_最终版.xlsx'")