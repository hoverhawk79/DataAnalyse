import pandas as pd

# 读取Excel文件
# file_path = "temp auto.xlsx"
file_path = r"C:\Users\User\Desktop\DevOps\Worldline\DataAnalytics\temp auto.xlsx"
df = pd.read_excel(file_path, sheet_name="Both genders", skiprows=2)  # 跳过前两行标题

# 清理数据：删除完全为空的行和列
df = df.dropna(how='all').dropna(axis=1, how='all')

# 重命名列（根据原始结构）
df.columns = [
    "Year", "Area", 
    "Foreign-born_1-9_months", "Foreign-born_10-12_months", "Foreign-born_Not_unemployed",
    "Born_in_Sweden_1-9_months", "Born_in_Sweden_10-12_months", "Born_in_Sweden_Not_unemployed",
    "Total_1-9_months", "Total_10-12_months", "Total_Not_unemployed"
]

# 转换为长格式（Tableau友好格式）
melted_df = pd.melt(
    df,
    id_vars=["Year", "Area"],
    var_name="Category_Metric",
    value_name="Value"
)

# 拆分复合列（例如将"Foreign-born_1-9_months"拆分为两列）
melted_df[['Population_Group', 'Metric']] = melted_df['Category_Metric'].str.split('_', n=1, expand=True)

# 删除临时列并重新排序
final_df = melted_df[['Year', 'Area', 'Population_Group', 'Metric', 'Value']]

# 保存为新的CSV文件（Tableau可直接读取）
# output_path = "tableau_ready_data.csv"
output_path = r"C:\Users\User\Desktop\DevOps\Worldline\DataAnalytics\tableau_ready_data.csv"
# final_df.to_csv(output_path, index=False)
final_df.to_csv(output_path, index=False, encoding='utf-8-sig')  # 使用 utf-8-sig 解决Excel乱码

print(f"数据已转换并保存到: {output_path}")