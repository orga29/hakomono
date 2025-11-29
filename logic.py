import pandas as pd
import openpyxl
from datetime import datetime
import io
import re

def load_source_data(file_path_or_buffer):
    """
    Loads the source Excel file and extracts necessary information.
    """
    # Read the file
    # Row 6 (index 5): Customer Code (e.g. 254-4, 303, 900...)
    # Row 7 (index 6): Delivery Type (配送種別) - e.g. "ﾗﾐ", "直売所"
    # Row 8 (index 7): Customer Name (得意先名)
    # Row 9 (index 8): Start of data
    
    df_raw = pd.read_excel(file_path_or_buffer, header=None)
    
    customer_codes = df_raw.iloc[5, 3:] # Col D onwards
    delivery_types = df_raw.iloc[6, 3:] # Col D onwards
    customer_names = df_raw.iloc[7, 3:] # Col D onwards
    
    # Create a mapping of column index to metadata
    col_mapping = []
    for idx, (code, dtype, cname) in enumerate(zip(customer_codes, delivery_types, customer_names)):
        real_col_idx = idx + 3 # Offset by 3 (A, B, C)
        
        # Check exclusion rule: Code >= 900
        is_excluded = False
        code_str = str(code).strip()
        match = re.match(r'^(\d+)', code_str)
        if match:
            num = int(match.group(1))
            if num >= 900:
                is_excluded = True
        
        if not is_excluded:
            col_mapping.append({
                "col_idx": real_col_idx,
                "customer_code": code_str,
                "delivery_type": str(dtype).strip(),
                "customer_name": str(cname).strip(),
                "sort_key": _get_sort_key(code_str)
            })
        
    # Read data part
    data_df = df_raw.iloc[8:].copy()
    
    # Rename first 3 columns
    data_df = data_df.rename(columns={0: "product_code", 1: "product_name", 2: "box_type"})
    
    # Filter by "箱" in Column C (index 2)
    filtered_df = data_df[data_df.iloc[:, 2] == "箱"].copy()
    
    # Sort by Product Code (A column)
    filtered_df["product_code"] = filtered_df["product_code"].astype(str)
    filtered_df = filtered_df.sort_values("product_code")
    
    return filtered_df, col_mapping

def _get_sort_key(code_str):
    """
    Helper to create a sort key from customer code (e.g. '254-4' -> (254, 4))
    """
    parts = re.split(r'\D+', code_str)
    return tuple(int(p) for p in parts if p)

def process_data(filtered_df, col_mapping):
    """
    Splits the data into two DataFrames based on delivery type.
    """
    # Base columns
    base_cols = filtered_df.iloc[:, 0:2].copy()
    base_cols.columns = ["商品コード", "商品名"]
    base_cols["箱/こ/不"] = "箱"
    
    # --- Split Logic ---
    koda_mapping = []
    yamato_mapping = []
    
    for m in col_mapping:
        if m["delivery_type"] == "ﾗﾐ(こだ)":
            koda_mapping.append(m)
        else:
            yamato_mapping.append(m)
            
    # --- Sort Logic ---
    koda_mapping.sort(key=lambda x: x["sort_key"])
    
    priority_list = ["ﾗﾐ", "ﾗﾐ(ｻﾈｯﾄ行)", "桜花便", "社内便1", "社内便2", "社内便3"]
    
    def yamato_sort_key(m):
        dtype = m["delivery_type"]
        try:
            p_idx = priority_list.index(dtype)
        except ValueError:
            p_idx = 999
        return (p_idx, m["sort_key"])
        
    yamato_mapping.sort(key=yamato_sort_key)
    
    # --- Construct DataFrames ---
    def build_df_and_headers(mapping, include_delivery_types=False):
        cols = []
        headers = []
        delivery_types = []
        for m in mapping:
            col_data = filtered_df.iloc[:, m["col_idx"]]
            col_data.name = m["customer_name"]
            cols.append(col_data)
            headers.append(m["customer_name"])
            if include_delivery_types:
                delivery_types.append(m["delivery_type"])
            
        if cols:
            df = pd.concat([base_cols] + cols, axis=1)
        else:
            df = base_cols.copy()
        
        if include_delivery_types:
            return df, headers, delivery_types
        return df, headers

    df_lami_koda, lami_koda_headers = build_df_and_headers(koda_mapping)
    df_lami_yamato, lami_yamato_headers, yamato_delivery_types = build_df_and_headers(yamato_mapping, include_delivery_types=True)
        
    return df_lami_koda, df_lami_yamato, lami_koda_headers, lami_yamato_headers, yamato_delivery_types

def write_to_template(template_path, df_lami_koda, df_lami_yamato, koda_headers, yamato_headers, yamato_delivery_types, output_filename):
    """
    Writes the processed data to the Excel template.
    """
    wb = openpyxl.load_workbook(template_path, keep_vba=True)
    
    if "ラミ（こだ）" in wb.sheetnames:
        ws_koda = wb["ラミ（こだ）"]
        _write_koda_sheet(ws_koda, df_lami_koda, koda_headers)
        
    if "ラミヤマトその他" in wb.sheetnames:
        ws_yamato = wb["ラミヤマトその他"]
        _write_yamato_sheet(ws_yamato, df_lami_yamato, yamato_headers, yamato_delivery_types)
        
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def _write_koda_sheet(ws, df, customer_names):
    """
    Writing logic for Lami (Koda) sheet - 2 row header structure.
    """
    header_row = None
    footer_row = None
    
    for r in range(1, 100):
        val_a = ws.cell(row=r, column=1).value
        val_b = ws.cell(row=r, column=2).value
        
        # Check if this is a header/footer row by looking for exact "商品コード" in A
        # and "商品名" in B (not product data)
        is_header_footer = False
        if val_a and val_b:
            # Must be exact match for header strings, not product codes
            if str(val_a).strip() == "商品コード" and "商品" in str(val_b):
                is_header_footer = True
            
        if is_header_footer:
            if header_row is None:
                header_row = r
            else:
                footer_row = r
                break
    
    if header_row is None: header_row = 2
    if footer_row is None: footer_row = 15
    
    # Find limit column
    limit_col = ws.max_column
    check_row = header_row + 1
    if check_row >= footer_row: check_row = header_row
    
    for c in range(4, ws.max_column + 1):
        cell = ws.cell(row=check_row, column=c)
        if cell.data_type == 'f' or (isinstance(cell.value, str) and cell.value.startswith('=')):
            limit_col = c
            break
            
    # Clear Data Rows
    if footer_row > header_row + 1:
        for r in range(header_row + 1, footer_row):
            for c in range(4, limit_col):
                ws.cell(row=r, column=c).value = None
            
    # Clear Header/Footer Customer Names
    for c in range(4, limit_col):
        ws.cell(row=header_row, column=c).value = None
        ws.cell(row=footer_row, column=c).value = None
        
    # Write Customer Headers
    for idx, name in enumerate(customer_names):
        col_idx = 4 + idx
        if col_idx < limit_col:
            ws.cell(row=header_row, column=col_idx, value=name)
            ws.cell(row=footer_row, column=col_idx, value=name)
            
    # Write Product Data
    num_data_rows = len(df)
    available_rows = footer_row - header_row - 1
    
    if num_data_rows > available_rows:
        rows_to_insert = num_data_rows - available_rows
        ws.insert_rows(footer_row, amount=rows_to_insert)
        footer_row += rows_to_insert
        
    for r_idx, row in enumerate(df.itertuples(index=False), 1):
        current_row = header_row + r_idx
        
        ws.cell(row=current_row, column=1, value=row[0])
        ws.cell(row=current_row, column=2, value=row[1])
        ws.cell(row=current_row, column=3, value=row[2])
        
        for c_idx, val in enumerate(row[3:], 0):
            target_col = 4 + c_idx
            if target_col < limit_col:
                ws.cell(row=current_row, column=target_col, value=val)

def _write_yamato_sheet(ws, df, customer_names, delivery_types):
    """
    Writing logic for Lami Yamato sheet - 3 row header structure.
    Row 2: Delivery Types
    Row 3: Customer Names
    """
    header_row = None
    footer_row = None
    
    for r in range(1, 100):
        val_a = ws.cell(row=r, column=1).value
        val_b = ws.cell(row=r, column=2).value
        
        # Check if this is a header/footer row
        is_header_footer = False
        if val_a and val_b:
            if str(val_a).strip() == "商品コード" and "商品" in str(val_b):
                is_header_footer = True
            
        if is_header_footer:
            if header_row is None:
                header_row = r
            else:
                footer_row = r
                break
    
    if header_row is None: header_row = 3
    if footer_row is None: footer_row = 16
    
    delivery_type_row = header_row - 1
    
    # Find limit column
    limit_col = ws.max_column
    check_row = header_row + 1
    if check_row >= footer_row: check_row = header_row
    
    for c in range(4, ws.max_column + 1):
        cell = ws.cell(row=check_row, column=c)
        if cell.data_type == 'f' or (isinstance(cell.value, str) and cell.value.startswith('=')):
            limit_col = c
            break
    
    # Clear data rows
    if footer_row > header_row + 1:
        for r in range(header_row + 1, footer_row):
            for c in range(4, limit_col):
                ws.cell(row=r, column=c).value = None
    
    # Clear Delivery Type Row
    for c in range(4, limit_col):
        ws.cell(row=delivery_type_row, column=c).value = None
        
    # Clear Header/Footer Customer Names
    for c in range(4, limit_col):
        ws.cell(row=header_row, column=c).value = None
        ws.cell(row=footer_row, column=c).value = None
    
    # Write Delivery Types
    for idx, dtype in enumerate(delivery_types):
        col_idx = 4 + idx
        if col_idx < limit_col:
            ws.cell(row=delivery_type_row, column=col_idx, value=dtype)
    
    # Write Customer Names
    for idx, name in enumerate(customer_names):
        col_idx = 4 + idx
        if col_idx < limit_col:
            ws.cell(row=header_row, column=col_idx, value=name)
            ws.cell(row=footer_row, column=col_idx, value=name)
    
    # Write Product Data
    num_data_rows = len(df)
    available_rows = footer_row - header_row - 1
    
    if num_data_rows > available_rows:
        rows_to_insert = num_data_rows - available_rows
        ws.insert_rows(footer_row, amount=rows_to_insert)
        footer_row += rows_to_insert
    
    for r_idx, row in enumerate(df.itertuples(index=False), 1):
        current_row = header_row + r_idx
        
        ws.cell(row=current_row, column=1, value=row[0])
        ws.cell(row=current_row, column=2, value=row[1])
        ws.cell(row=current_row, column=3, value=row[2])
        
        for c_idx, val in enumerate(row[3:], 0):
            target_col = 4 + c_idx
            if target_col < limit_col:
                ws.cell(row=current_row, column=target_col, value=val)
