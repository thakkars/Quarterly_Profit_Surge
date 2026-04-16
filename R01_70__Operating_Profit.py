def append_rows_to_raw_excel(rows, sheet_name):
    """Append row-dicts to the quarter sheet with dynamic transformations."""
    if not rows:
        return

    # Create new DataFrame and apply our transformation
    df_new = transform_and_order_data(pd.DataFrame(rows))

    if os.path.exists(RAW_FILE_PATH):
        try:
            # Read existing data so we don't lose previously saved companies
            wb = load_workbook(RAW_FILE_PATH)
            if sheet_name in wb.sheetnames:
                df_existing = pd.read_excel(RAW_FILE_PATH, sheet_name=sheet_name)
                # Combine old and new, and re-run transformation to sync column structures perfectly
                df_combined = pd.concat([df_existing, df_new], ignore_index=True)
                df_final = transform_and_order_data(df_combined)
            else:
                df_final = df_new
            wb.close()
            
            # Use Pandas ExcelWriter to overlay the perfectly structured data
            with pd.ExcelWriter(RAW_FILE_PATH, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df_final.to_excel(writer, sheet_name=sheet_name, index=False)
                
        except Exception as e:
            print(f"⚠️  Error appending to existing Excel: {e}")
            return
    else:
        # First time file creation
        with pd.ExcelWriter(RAW_FILE_PATH, engine='openpyxl') as writer:
            df_new.to_excel(writer, sheet_name=sheet_name, index=False)

    # Re-apply styling formatting
    wb = load_workbook(RAW_FILE_PATH)
    ws = wb[sheet_name]
    format_worksheet(ws)
    wb.save(RAW_FILE_PATH)
    
    print(f"  ↳ Saved and formatted {len(rows)} rows to sheet '{sheet_name}'.")
