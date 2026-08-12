def main(params):
    import os
    import re
    import zipfile
    import tempfile
    from pathlib import Path

    import pandas as pd

    files = params.get("files", [])

    if not files:
        return {
            "success": False,
            "message": "请上传一个zip压缩包",
        }

    upload_file = files[0] if isinstance(files, list) else files

    zip_path = (
        upload_file.get("path")
        or upload_file.get("url")
        or upload_file.get("fileUrl")
        or upload_file.get("localPath")
    )

    zip_name = (
        upload_file.get("name")
        or upload_file.get("fileName")
        or "上传文件.zip"
    )

    if not zip_path:
        return {
            "success": False,
            "message": f"没有获取到上传文件路径，当前文件对象：{upload_file}",
        }

    sql_col = "SQL语句"

    def clean_col_name(x):
        return str(x).replace("\n", "").replace("\r", "").replace(" ", "").strip()

    def clean_sql(x):
        if pd.isna(x):
            return ""
        s = str(x).strip()
        s = re.sub(r"\s+", " ", s)
        return s

    def safe_sheet_name(name):
        name = str(name)
        for ch in ['\\', '/', '*', '?', ':', '[', ']']:
            name = name.replace(ch, "_")
        return name[:31]

    def read_excel_auto(file, sheet=0):
        suffix = Path(file).suffix.lower()

        if suffix == ".xlsx":
            return pd.read_excel(file, sheet_name=sheet, engine="openpyxl", header=None)

        elif suffix == ".xls":
            try:
                return pd.read_excel(file, sheet_name=sheet, engine="xlrd", header=None)
            except Exception:
                tables = pd.read_html(file, header=None)
                return tables[sheet]

        else:
            raise Exception(f"不支持的文件类型：{suffix}")

    def fix_header(df, target_col="SQL语句"):
        target = clean_col_name(target_col)

        for i in range(min(30, len(df))):
            row = [clean_col_name(x) for x in df.iloc[i].tolist()]

            if target in row:
                new_df = df.iloc[i + 1:].copy()
                new_df.columns = row
                new_df.reset_index(drop=True, inplace=True)
                return new_df

        raise Exception(f"没有找到表头【{target_col}】")

    def df_to_html_sheet(sheet_name, df):
        return f"""
        <h2>{sheet_name}</h2>
        {df.to_html(index=False, border=1, escape=False)}
        <br/>
        """

    work_dir = tempfile.mkdtemp(prefix="sql_diff_")
    extract_dir = os.path.join(work_dir, "extract")
    os.makedirs(extract_dir, exist_ok=True)

    try:
        with zipfile.ZipFile(zip_path, "r") as zip_ref:
            zip_ref.extractall(extract_dir)

        excel_files = []

        for root, dirs, file_names in os.walk(extract_dir):
            for file_name in file_names:
                if file_name.startswith("~$"):
                    continue

                suffix = Path(file_name).suffix.lower()
                if suffix in [".xls", ".xlsx"]:
                    excel_files.append(os.path.join(root, file_name))

        if len(excel_files) != 2:
            return {
                "success": False,
                "message": f"压缩包里必须有且只有2个Excel文件，当前找到{len(excel_files)}个：{[Path(x).name for x in excel_files]}",
            }

        file1, file2 = excel_files[0], excel_files[1]

        name1 = Path(file1).stem
        name2 = Path(file2).stem

        df1_raw = read_excel_auto(file1)
        df2_raw = read_excel_auto(file2)

        df1 = fix_header(df1_raw, sql_col)
        df2 = fix_header(df2_raw, sql_col)

        df1 = df1[df1[sql_col].notna()].copy()
        df2 = df2[df2[sql_col].notna()].copy()

        compare_col = "__SQL_COMPARE_KEY__"

        df1[compare_col] = df1[sql_col].apply(clean_sql)
        df2[compare_col] = df2[sql_col].apply(clean_sql)

        df1 = df1[df1[compare_col] != ""].copy()
        df2 = df2[df2[compare_col] != ""].copy()

        df1 = df1.drop_duplicates(subset=[compare_col])
        df2 = df2.drop_duplicates(subset=[compare_col])

        sql1 = set(df1[compare_col])
        sql2 = set(df2[compare_col])

        only1 = sql1 - sql2
        only2 = sql2 - sql1
        both = sql1 & sql2

        only_table1 = df1[df1[compare_col].isin(only1)].copy()
        only_table2 = df2[df2[compare_col].isin(only2)].copy()
        both_table1 = df1[df1[compare_col].isin(both)].copy()
        both_table2 = df2[df2[compare_col].isin(both)].copy()

        for d in [only_table1, only_table2, both_table1, both_table2]:
            if compare_col in d.columns:
                d.drop(columns=[compare_col], inplace=True)

        sheet1_name = safe_sheet_name(f"仅{name1}")
        sheet2_name = safe_sheet_name(f"仅{name2}")
        sheet3_name = safe_sheet_name(f"共有_{name1}")
        sheet4_name = safe_sheet_name(f"共有_{name2}")

        html_body = ""
        html_body += df_to_html_sheet(sheet1_name, only_table1)
        html_body += df_to_html_sheet(sheet2_name, only_table2)
        html_body += df_to_html_sheet(sheet3_name, both_table1)
        html_body += df_to_html_sheet(sheet4_name, both_table2)

        html_output = f"""
        <html>
        <head>
            <meta charset="utf-8">
        </head>
        <body>
            {html_body}
        </body>
        </html>
        """

        output_file_name = f"SQL比较结果_{name1}_VS_{name2}.xls"

        return {
            "success": True,
            "message": "SQL比较完成",
            "htmlOutput": html_output,
            "fileName": output_file_name,
            "file1Name": name1,
            "file2Name": name2,
            "file1Count": int(len(df1)),
            "file2Count": int(len(df2)),
            "onlyFile1Count": int(len(only_table1)),
            "onlyFile2Count": int(len(only_table2)),
            "bothCount": int(len(both)),
        }

    except Exception as e:
        return {
            "success": False,
            "message": f"处理失败：{str(e)}",
        }