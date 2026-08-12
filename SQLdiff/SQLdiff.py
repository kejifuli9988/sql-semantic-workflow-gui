import pandas as pd
from pathlib import Path


# ==========================
# 文件
file1 = "慢SQL表_2026_6_30.xls"
file2 = "慢SQL表_2026_7_1.xls"

sheet1 = 0
sheet2 = 0

sql_col = "SQL语句"
# ==========================


def clean_col_name(x):
    """清洗列名"""
    return str(x).replace("\n", "").replace("\r", "").replace(" ", "").strip()


def read_excel_auto(file, sheet=0):
    """自动读取xls/xlsx"""

    suffix = Path(file).suffix.lower()

    if suffix == ".xlsx":
        return pd.read_excel(
            file,
            sheet_name=sheet,
            engine="openpyxl",
            header=None
        )

    elif suffix == ".xls":

        # 先尝试真正xls
        try:
            return pd.read_excel(
                file,
                sheet_name=sheet,
                engine="xlrd",
                header=None
            )

        # 如果失败，则按HTML读取
        except Exception:
            tables = pd.read_html(file, header=None)
            return tables[sheet]

    else:
        raise Exception(f"不支持文件类型：{suffix}")


def fix_header(df, target_col="SQL语句"):
    """自动寻找真正表头"""

    target = clean_col_name(target_col)

    for i in range(min(30, len(df))):

        row = [clean_col_name(x) for x in df.iloc[i].tolist()]

        if target in row:

            new_df = df.iloc[i + 1:].copy()
            new_df.columns = row
            new_df.reset_index(drop=True, inplace=True)

            return new_df

    raise Exception(f"没有找到表头【{target_col}】")


# ------------------------------
# 读取
df1 = read_excel_auto(file1, sheet1)
df2 = read_excel_auto(file2, sheet2)
# print("===== df1 前5行 =====")
# print(df1.head())

# print("===== df2 前5行 =====")
# print(df2.head())
# 自动识别表头
df1 = fix_header(df1, sql_col)
df2 = fix_header(df2, sql_col)

# 去掉SQL为空
df1 = df1[df1[sql_col].notna()].copy()
df2 = df2[df2[sql_col].notna()].copy()

# SQL标准化
df1[sql_col] = (
    df1[sql_col]
    .astype(str)
    .str.strip()
)

df2[sql_col] = (
    df2[sql_col]
    .astype(str)
    .str.strip()
)

# 去重
df1 = df1.drop_duplicates(subset=[sql_col])
df2 = df2.drop_duplicates(subset=[sql_col])

# 集合比较
sql1 = set(df1[sql_col])
sql2 = set(df2[sql_col])

only1 = sql1 - sql2
only2 = sql2 - sql1
both = sql1 & sql2

# 四个结果
only_table1 = df1[df1[sql_col].isin(only1)]
only_table2 = df2[df2[sql_col].isin(only2)]

both_table1 = df1[df1[sql_col].isin(both)]
both_table2 = df2[df2[sql_col].isin(both)]

name1 = Path(file1).stem
name2 = Path(file2).stem
# 导出
output = f"SQL比较结果_{name1}_VS_{name2}.xlsx"



with pd.ExcelWriter(output, engine="openpyxl") as writer:
    only_table1.to_excel(writer, sheet_name=f"仅{name1}", index=False)
    only_table2.to_excel(writer, sheet_name=f"仅{name2}", index=False)
    both_table1.to_excel(writer, sheet_name=f"共有({name1})", index=False)
    both_table2.to_excel(writer, sheet_name=f"共有({name2})", index=False)

print("====================================")
print("完成！")
print(f"表1总数：{len(df1)}")
print(f"表2总数：{len(df2)}")
print(f"仅表1：{len(only_table1)}")
print(f"仅表2：{len(only_table2)}")
print(f"共有：{len(both)}")
print("输出文件：SQL比较结果.xlsx")
print("====================================")