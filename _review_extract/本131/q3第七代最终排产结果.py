import pandas as pd

# ==================== 1. 读取数据 ====================
file_name = '问题3_无工艺排产结果_最终合格版.xlsx'
df = pd.read_excel(file_name, sheet_name=0)

# 类型转换
df['温度组'] = pd.to_numeric(df['温度组'], errors='coerce')
df['入口厚度(mm)'] = pd.to_numeric(df['入口厚度(mm)'], errors='coerce')
df['入口宽度(mm)'] = pd.to_numeric(df['入口宽度(mm)'], errors='coerce')
df['合炉最大厚度差'] = pd.to_numeric(df['合炉最大厚度差'], errors='coerce').fillna(0)
df['必须单独一炉'] = pd.to_numeric(df['必须单独一炉'], errors='coerce').fillna(0).astype(int)

# ==================== 2. 合炉（严格按厚度组和允许差） ====================
def combine_batches(df):
    batches = []
    for (alloy, temp, thick_group), group in df.groupby(['合金大类', '温度组', '厚度组_合炉']):
        alone = group[group['必须单独一炉'] == 1]
        normal = group[group['必须单独一炉'] == 0]
        for _, row in alone.iterrows():
            batches.append(pd.DataFrame([row]))
        if len(normal) == 0:
            continue
        max_diff = normal['合炉最大厚度差'].iloc[0]
        # 按厚度升序，便于检查厚度差
        normal = normal.sort_values('入口厚度(mm)', ascending=True).reset_index(drop=True)
        cur = [normal.iloc[0]]
        for i in range(1, len(normal)):
            min_t = cur[0]['入口厚度(mm)']
            max_t = cur[-1]['入口厚度(mm)']
            new_t = normal.iloc[i]['入口厚度(mm)']
            if max(max_t, new_t) - min(min_t, new_t) <= max_diff:
                cur.append(normal.iloc[i])
            else:
                batches.append(pd.DataFrame(cur))
                cur = [normal.iloc[i]]
        if cur:
            batches.append(pd.DataFrame(cur))
    return batches

batches = combine_batches(df)
print(f"合炉完成，共 {len(batches)} 个炉批")

# 为每个炉批计算代表参数
for b in batches:
    b['_temp'] = b['温度组'].iloc[0]
    b['_max_width'] = b['入口宽度(mm)'].max()
    b['_max_thick'] = b['入口厚度(mm)'].max()

# ==================== 3. 过渡约束函数 ====================
def thickness_diff_ok(h1, h2):
    thick, thin = max(h1, h2), min(h1, h2)
    return (thick - thin) / thick <= 0.3

def width_diff_ok(w1, w2):
    if w1 >= w2:
        return (w1 - w2) <= 800
    else:
        return (w2 - w1) <= 500

def temp_diff_ok(t1, t2):
    if t2 >= t1:
        return (t2 - t1) <= 50
    else:
        return (t1 - t2) <= 20

# ==================== 4. 排序（温度升序，组内厚度降序） ====================
# 按温度分组
temp_groups = {}
for b in batches:
    t = b['_temp'].iloc[0]
    temp_groups.setdefault(t, []).append(b)

# 每个温度组内按厚度降序排列
for t in temp_groups:
    temp_groups[t].sort(key=lambda b: -b['_max_thick'].iloc[0])

# 构建最终序列
sequence = []
for t in sorted(temp_groups.keys()):
    sequence.extend(temp_groups[t])

# ==================== 5. 计算违规情况 ====================
violations = {'temp': 0, 'width': 0, 'thickness': 0}
for i in range(len(sequence)-1):
    b1 = sequence[i]
    b2 = sequence[i+1]
    t1 = b1['_temp'].iloc[0]
    t2 = b2['_temp'].iloc[0]
    if not temp_diff_ok(t1, t2):
        violations['temp'] += 1
    w1 = b1['_max_width'].iloc[0]
    w2 = b2['_max_width'].iloc[0]
    if not width_diff_ok(w1, w2):
        violations['width'] += 1
    h1 = b1['_max_thick'].iloc[0]
    h2 = b2['_max_thick'].iloc[0]
    if not thickness_diff_ok(h1, h2):
        violations['thickness'] += 1

objective = {
    'total_batches': len(sequence),
    'temp_violations': violations['temp'],
    'width_violations': violations['width'],
    'thickness_violations': violations['thickness'],
    'total_violations': sum(violations.values())
}
print("目标函数值：", objective)

# ==================== 6. 输出最终排产表 ====================
output_rows = []
global_seq = 0
temp_counter = {}
for batch in sequence:
    temp = batch['_temp'].iloc[0]
    temp_counter[temp] = temp_counter.get(temp, 0) + 1
    batch_id = f"T{int(temp)}_{temp_counter[temp]:03d}"
    # 炉批内按厚度升序（从薄到厚生产）
    batch = batch.sort_values('入口厚度(mm)', ascending=True).reset_index(drop=True)
    for inner_seq, (_, row) in enumerate(batch.iterrows(), start=1):
        global_seq += 1
        new_row = row.to_dict()
        new_row['全局序号'] = global_seq
        new_row['温度组'] = temp
        new_row['炉批号'] = batch_id
        new_row['炉批内序号'] = inner_seq
        # 删除辅助列
        for k in ['_temp', '_max_width', '_max_thick']:
            new_row.pop(k, None)
        output_rows.append(new_row)

output_df = pd.DataFrame(output_rows)
output_df.to_excel('问题3_最终排产结果.xlsx', index=False)
print("排产表已保存：问题3_最终排产结果.xlsx")

# ==================== 7. 生成验证表 ====================
verification = []
for i in range(len(sequence)-1):
    b1 = sequence[i]
    b2 = sequence[i+1]
    t1 = b1['_temp'].iloc[0]
    t2 = b2['_temp'].iloc[0]
    w1 = b1['_max_width'].iloc[0]
    w2 = b2['_max_width'].iloc[0]
    h1 = b1['_max_thick'].iloc[0]
    h2 = b2['_max_thick'].iloc[0]
    verification.append({
        '前炉批号': f"T{int(t1)}_{i}",
        '后炉批号': f"T{int(t2)}_{i+1}",
        '温度差': abs(t1-t2),
        '温度合规': temp_diff_ok(t1, t2),
        '宽度差': abs(w1-w2),
        '宽度合规': width_diff_ok(w1, w2),
        '厚度差': abs(h1-h2),
        '厚度比': round((max(h1,h2)-min(h1,h2))/max(h1,h2), 4),
        '厚度合规': thickness_diff_ok(h1, h2)
    })
verif_df = pd.DataFrame(verification)
verif_df.to_excel('验证表_相邻炉批过渡情况.xlsx', index=False)
print("验证表已保存：验证表_相邻炉批过渡情况.xlsx")

print("排产完成！")