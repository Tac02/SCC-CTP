import pandas as pd
import warnings


def process_cadet_file(input_file, output_file, master_module_file="module_list.csv"):
    warnings.filterwarnings("ignore", message="Could not infer format")

    # -----------------------------
    # CONFIG
    # -----------------------------
    rank_order = [
        "Junior Cadet",
        "Junior Cadet First Class",
        "Able Junior Cadet",
        "Leading Junior Cadet",
        "New Entry Cadet",
        "Cadet",
        "Cadet 1st Class",
        "Ordinary Cadet",
        "Able Cadet",
        "Leading Cadet",
        "Petty Officer Cadet",
    ]

    syllabus_order = [
        "SCC CTP1 - New Entry to Cadet",
        "SCC CTP2 - Cadet to Cadet 1st Class",
        "SCC CTP3 - Cadet 1st Class to Ordinary Cadet",
        "SCC CTP4 - Ordinary Cadet to Able Cadet",
        "SCC CTP5 - Able Cadet to Leading Cadet",
        "SCC CTP6 - Leading Cadet to Petty Officer Cadet",
    ]

    syllabus_sheet_names = {
        "SCC CTP1 - New Entry to Cadet": "New Entry to Cadet",
        "SCC CTP2 - Cadet to Cadet 1st Class": "Cadet to Cadet First Class",
        "SCC CTP3 - Cadet 1st Class to Ordinary Cadet": "Cadet First Class to Ordinary Cadet",
        "SCC CTP4 - Ordinary Cadet to Able Cadet": "Ordinary Cadet to Able Cadet",
        "SCC CTP5 - Able Cadet to Leading Cadet": "Able Cadet to Leading Cadet",
        "SCC CTP6 - Leading Cadet to Petty Officer Cadet": "Leading Cadet to PO Cadet",
    }

    rank_filter = {
        "SCC CTP1 - New Entry to Cadet": ["New Entry Cadet"],
        "SCC CTP2 - Cadet to Cadet 1st Class": ["New Entry Cadet", "Cadet"],
        "SCC CTP3 - Cadet 1st Class to Ordinary Cadet": ["Cadet", "Cadet 1st Class"],
        "SCC CTP4 - Ordinary Cadet to Able Cadet": ["Cadet 1st Class", "Ordinary Cadet"],
        "SCC CTP5 - Able Cadet to Leading Cadet": ["Ordinary Cadet", "Able Cadet"],
        "SCC CTP6 - Leading Cadet to Petty Officer Cadet": ["Leading Cadet", "Petty Officer Cadet"],
    }

    # -----------------------------
    # LOAD & CLEAN CADET FILE
    # -----------------------------
    df = pd.read_csv(input_file) if input_file.lower().endswith(".csv") else pd.read_excel(input_file)

    df["Rank"] = pd.Categorical(df["Rank"], categories=rank_order, ordered=True)
    df["Date Achieved"] = pd.to_datetime(df["Date Achieved"], dayfirst=True, errors="coerce")

    df["Module"] = df["Module"].astype(str).str.strip()
    split_cols = df["Module"].str.split("-", n=1, expand=True)
    df["Module Code"] = split_cols[0].str.strip()
    df["Module Description"] = split_cols[1].str.strip() if split_cols.shape[1] > 1 else ""

    # Build description lookup
    module_descriptions = (
        df[["Module Code", "Module Description"]]
        .dropna()
        .drop_duplicates(subset=["Module Code"])
        .set_index("Module Code")["Module Description"]
        .to_dict()
    )

    # -----------------------------
    # LOAD MASTER MODULE LIST
    # -----------------------------
    master_modules = pd.read_csv(master_module_file)
    master_modules["Syllabus"] = master_modules["Syllabus"].astype(str).str.strip()
    master_modules["Module Code"] = master_modules["Module Code"].astype(str).str.strip()

    # -----------------------------
    # MERGE MODULES (MASTER + INPUT)
    # -----------------------------
    all_modules = pd.concat(
        [
            master_modules[["Syllabus", "Module Code"]],
            df[["Syllabus", "Module Code"]],
        ],
        ignore_index=True,
    ).drop_duplicates(subset=["Module Code"])

    # Sort modules by syllabus order then alphabetically
    def syllabus_sort_key(syl):
        return syllabus_order.index(syl) if syl in syllabus_order else len(syllabus_order)

    all_modules = all_modules.sort_values(
        by=["Syllabus", "Module Code"],
        key=lambda col: col.map(
            lambda x: syllabus_sort_key(x) if col.name == "Syllabus" else x
        ),
    )

    # Add descriptions
    all_modules["Module Description"] = all_modules["Module Code"].map(module_descriptions).fillna("")

    module_syllabus_map = all_modules.set_index("Module Code")["Syllabus"].to_dict()

    # -----------------------------
    # MASTER PIVOT (ONE ROW PER CADET)
    # -----------------------------
    pivot_df = df.pivot_table(
        index=["PNumber", "Rank", "Name"],
        columns="Module Code",
        values="Date Achieved",
        aggfunc="first",
        observed=False,
    ).reset_index()

    pivot_df = pivot_df.sort_values(["Rank", "Name"])

    # Sort module columns
    module_cols = [c for c in pivot_df.columns if c not in ["PNumber", "Rank", "Name"]]

    def module_sort_key(code):
        syllabus = module_syllabus_map.get(code, None)
        return (syllabus_order.index(syllabus), code) if syllabus in syllabus_order else (len(syllabus_order), code)

    module_cols_sorted = sorted(module_cols, key=module_sort_key)

    # Force include missing modules
    for code in all_modules["Module Code"]:
        if code not in pivot_df.columns:
            pivot_df[code] = pd.NaT
            module_cols_sorted.append(code)

    module_cols_sorted = sorted(set(module_cols_sorted), key=module_sort_key)

    pivot_df = pivot_df[["PNumber", "Rank", "Name"] + module_cols_sorted]

    # -----------------------------
    # SPLIT JUNIOR / SENIOR
    # -----------------------------
    is_junior = pivot_df["Rank"].astype(str).str.contains("junior", case=False, na=False)
    master_junior = pivot_df[is_junior].copy()
    master_main = pivot_df[~is_junior].copy()

    # -----------------------------
    # MODULE LIST SHEET
    # -----------------------------
    module_list = (
        all_modules[["Syllabus", "Module Code", "Module Description"]]
        .drop_duplicates()
        .sort_values(["Syllabus", "Module Code"])
    )

    # -----------------------------
    # WRITE TO EXCEL
    # -----------------------------
    with pd.ExcelWriter(output_file, engine="xlsxwriter", datetime_format="dd/mm/yy") as writer:
        workbook = writer.book

        fmt_green = workbook.add_format({"bg_color": "#C6EFCE"})
        fmt_red = workbook.add_format({"bg_color": "#FFC7CE"})
        fmt_header = workbook.add_format({"bold": True})

        def format_sheet(ws, df_sheet):
            ws.freeze_panes(1, 3)
            ws.set_row(0, None, fmt_header)
            last_row = len(df_sheet)
            last_col = len(df_sheet.columns) - 1
            if last_col >= 3:
                ws.conditional_format(1, 3, last_row, last_col, {"type": "blanks", "format": fmt_red})
                ws.conditional_format(1, 3, last_row, last_col, {"type": "no_blanks", "format": fmt_green})
            for col_idx, col_name in enumerate(df_sheet.columns):
                width = max(10, min(40, df_sheet[col_name].astype(str).str.len().max() + 2))
                ws.set_column(col_idx, col_idx, width)

        # MASTER SHEETS
        master_main.to_excel(writer, sheet_name="Master", index=False)
        format_sheet(writer.sheets["Master"], master_main)

        master_junior.to_excel(writer, sheet_name="Master_Junior", index=False)
        format_sheet(writer.sheets["Master_Junior"], master_junior)

        # SYLLABUS SHEETS
        for syllabus in syllabus_order:
            syllabus_modules = list(all_modules.loc[all_modules["Syllabus"] == syllabus, "Module Code"])
            syllabus_modules = [m for m in module_cols_sorted if m in syllabus_modules]

            sheet_df = master_main.copy()
            sheet_df = sheet_df[["PNumber", "Rank", "Name"] + syllabus_modules]

            allowed_ranks = rank_filter.get(syllabus, rank_order)
            sheet_df = sheet_df[sheet_df["Rank"].isin(allowed_ranks)]

            if sheet_df.empty:
                continue

            sheet_name = syllabus_sheet_names[syllabus].replace(" ", "_")[:31]
            sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
            format_sheet(writer.sheets[sheet_name], sheet_df)

        # MODULE LIST
        module_list.to_excel(writer, sheet_name="Module_List", index=False)
        writer.sheets["Module_List"].set_row(0, None, fmt_header)

        # DASHBOARD
        dashboard = workbook.add_worksheet("Dashboard")
        dashboard.write("A1", "Metric", fmt_header)
        dashboard.write("B1", "Value", fmt_header)

        dashboard.write("A2", "Total Cadets")
        dashboard.write("B2", pivot_df["PNumber"].nunique())

        dashboard.write("A3", "Junior Cadets")
        dashboard.write("B3", master_junior["PNumber"].nunique())

        dashboard.write("A4", "Senior Cadets")
        dashboard.write("B4", master_main["PNumber"].nunique())

        dashboard.write("A6", "Cadets per Rank", fmt_header)
        rank_counts = pivot_df.groupby("Rank")["PNumber"].nunique().reindex(rank_order).fillna(0)
        row = 7
        for rank, count in rank_counts.items():
            dashboard.write(row, 0, rank)
            dashboard.write(row, 1, int(count))
            row += 1

        dashboard.write(f"A{row+1}", "Approx Modules Completed (Senior)", fmt_header)
        dashboard.write(f"B{row+1}", int(master_main[module_cols_sorted].notna().sum().sum()))

    print(f"Processed file saved as: {output_file}")



process_cadet_file("downloads/Cadet Qualifications Report.csv", "CTP_Tracker.xlsx")
