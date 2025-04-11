import pandas as pd

# 读取 Excel 文件
file_path = "2. Unemployed_Skäggetorp_Linköping auto.xlsx"
sheet_name = "AL53"
df = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=3)  # 跳过前 3 行

# 清理数据：删除完全为空的行和列
df = df.dropna(how='all').dropna(axis=1, how='all')

# 重命名列（根据原始结构）
df.columns = [
    "Year", "Area",
    "Men_Foreign-born_1-9_months", "Men_Foreign-born_10-12_months", "Men_Foreign-born_Not_unemployed",
    "Men_Born_in_Sweden_1-9_months", "Men_Born_in_Sweden_10-12_months", "Men_Born_in_Sweden_Not_unemployed",
    "Men_Total_1-9_months", "Men_Total_10-12_months", "Men_Total_Not_unemployed",
    "Women_Foreign-born_1-9_months", "Women_Foreign-born_10-12_months", "Women_Foreign-born_Not_unemployed",
    "Women_Born_in_Sweden_1-9_months", "Women_Born_in_Sweden_10-12_months", "Women_Born_in_Sweden_Not_unemployed",
    "Women_Total_1-9_months", "Women_Total_10-12_months", "Women_Total_Not_unemployed",
    "Both_genders_Foreign-born_1-9_months", "Both_genders_Foreign-born_10-12_months",
    "Both_genders_Foreign-born_Not_unemployed",
    "Both_genders_Born_in_Sweden_1-9_months", "Both_genders_Born_in_Sweden_10-12_months",
    "Both_genders_Born_in_Sweden_Not_unemployed",
    "Both_genders_Total_1-9_months", "Both_genders_Total_10-12_months", "Both_genders_Total_Not_unemployed"
]

# 转换为长格式（Tableau 友好格式）
melted_df = pd.melt(
    df,
    id_vars=["Year", "Area"],
    var_name="Category_Metric",
    value_name="Value"
)


# 修复拆分逻辑：确保拆分正确
def split_category_metric(category_metric):
    parts = category_metric.split("_")
    gender = parts[0]

    # 处理 Population_Group 和 Metric
    if "Foreign-born" in category_metric:
        population_group = "Foreign_born"
    elif "Born_in_Sweden" in category_metric:
        population_group = "Born_in_Sweden"
    elif "Total" in category_metric:
        population_group = "Total"
    else:
        population_group = "Unknown"

    # 拼接 Metric
    if "1-9" in category_metric:
        metric = "1_9_months"
    elif "10-12" in category_metric:
        metric = "10_12_months"
    elif "Not_unemployed" in category_metric:
        metric = "Not_unemployed"
    else:
        metric = "Unknown"

    return gender, population_group, metric


# 应用拆分函数
split_data = melted_df["Category_Metric"].apply(split_category_metric)
split_df = pd.DataFrame(split_data.tolist(), columns=["Gender", "Population_Group", "Metric"])

# 合并拆分后的字段到原数据框
melted_df = pd.concat([melted_df, split_df], axis=1)

# 删除临时列并重新排序
final_df = melted_df[["Year", "Area", "Gender", "Population_Group", "Metric", "Value"]]

# 保存为 CSV 文件（UTF-8 编码，兼容瑞典语字符）
output_path = "unemployment_data_tableau_ready_fixed.csv"
final_df.to_csv(output_path, index=False, encoding="utf-8-sig")

print(f"数据已转换并保存到: {output_path}")
print("关键字段说明:")
print("- Gender: Men / Women / Both_genders")
print("- Population_Group: Foreign_born / Born_in_Sweden / Total")
print("- Metric: 1_9_months / 10_12_months / Not_unemployed")
