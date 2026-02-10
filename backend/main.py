# from fastapi import FastAPI, File, UploadFile, Form
# from fastapi.middleware.cors import CORSMiddleware
# from fastapi.responses import JSONResponse
# import pandas as pd
# import io
# import base64
# from typing import Optional

# app = FastAPI(title="TAXI Insurance Policy Processor")

# # Enable CORS
# app.add_middleware(
#     CORSMiddleware,
#     allow_origins=["https://digit-excel-taxi.vercel.app"],
#     allow_credentials=True,
#     allow_methods=["*"],
#     allow_headers=["*"],
# )

# # ------------------- FORMULA DATA -------------------
# FORMULA_DATA = [
#     {"LOB": "TW", "SEGMENT": "1+5", "PO": "90% of Payin", "REMARKS": "NIL"},
#     {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "PO": "-2%", "REMARKS": "Payin Below 20%"},
#     {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
#     {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
#     {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "PO": "-5%", "REMARKS": "Payin Above 50%"},
#     {"LOB": "TW", "SEGMENT": "TW TP", "PO": "-2%", "REMARKS": "Payin Below 20%"},
#     {"LOB": "TW", "SEGMENT": "TW TP", "PO": "-3%", "REMARKS": "Payin Above 20%"},
#     {"LOB": "PVT CAR", "SEGMENT": "PVT CAR COMP + SAOD", "PO": "90% of Payin", "REMARKS": "NIL"},
#     {"LOB": "PVT CAR", "SEGMENT": "PVT CAR TP", "PO": "-2%", "REMARKS": "Payin Below 20%"},
#     {"LOB": "PVT CAR", "SEGMENT": "PVT CAR TP", "PO": "-3%", "REMARKS": "Payin Above 20%"},
#     {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-2%", "REMARKS": "Payin Below 20%"},
#     {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
#     {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
#     {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-5%", "REMARKS": "Payin Above 50%"},
#     {"LOB": "BUS", "SEGMENT": "SCHOOL BUS", "PO": "Less 2% of Payin", "REMARKS": "NIL"},
#     {"LOB": "BUS", "SEGMENT": "STAFF BUS", "PO": "88% of Payin", "REMARKS": "NIL"},
#     {"LOB": "TAXI", "SEGMENT": "TAXI", "PO": "-2%", "REMARKS": "Payin Below 20%"},
#     {"LOB": "TAXI", "SEGMENT": "TAXI", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
#     {"LOB": "TAXI", "SEGMENT": "TAXI", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
#     {"LOB": "TAXI", "SEGMENT": "TAXI", "PO": "-5%", "REMARKS": "Payin Above 50%"},
#     {"LOB": "MISD", "SEGMENT": "Misd, Tractor", "PO": "88% of Payin", "REMARKS": "NIL"}
# ]

# # ------------------- STATE MAPPING -------------------
# STATE_MAPPING = {
#     "DELHI": "DELHI", "MUMBAI": "MAHARASHTRA", "PUNE": "MAHARASHTRA", "GOA": "GOA",
#     "KOLKATA": "WEST BENGAL", "HYDERABAD": "TELANGANA", "AHMEDABAD": "GUJARAT",
#     "SURAT": "GUJARAT", "JAIPUR": "RAJASTHAN", "LUCKNOW": "UTTAR PRADESH",
#     "PATNA": "BIHAR", "RANCHI": "JHARKHAND", "BHUVANESHWAR": "ODISHA",
#     "SRINAGAR": "JAMMU AND KASHMIR", "DEHRADUN": "UTTARAKHAND", "HARIDWAR": "UTTARAKHAND",
#     "HIMACHAL PRADESH": "HIMACHAL PRADESH", "ANDAMAN": "ANDAMAN AND NICOBAR ISLANDS",
#     "BANGALORE": "KARNATAKA", "JHARKHAND": "JHARKHAND", "BIHAR": "BIHAR",
#     "WEST BENGAL": "WEST BENGAL", "NORTH BENGAL": "WEST BENGAL", "ORISSA": "ODISHA",
#     "GOOD GJ": "GUJARAT", "BAD GJ": "GUJARAT", "ROM1": "REST OF MAHARASHTRA",
#     "ROM2": "REST OF MAHARASHTRA", "GOOD VIZAG": "ANDHRA PRADESH", "GOOD TN": "TAMIL NADU",
#     "KERALA": "KERALA", "GOOD MP": "MADHYA PRADESH", "GOOD CG": "CHHATTISGARH",
#     "GOOD RJ": "RAJASTHAN", "BAD RJ": "RAJASTHAN", "GOOD UP": "UTTAR PRADESH",
#     "BAD UP": "UTTAR PRADESH", "GOOD UK": "UTTARAKHAND", "BAD UK": "UTTARAKHAND",
#     "PUNJAB": "PUNJAB", "JAMMU": "JAMMU AND KASHMIR", "ASSAM": "ASSAM",
#     "NE EX ASSAM": "NORTH EAST", "GOOD NL": "NAGALAND", "GOOD KA": "KARNATAKA",
#     "BAD KA": "KARNATAKA", "HR REF": "HARYANA", "DEHRADUN, HARIDWAR": "UTTARAKHAND",
#     "HIMACHAL": "HIMACHAL PRADESH"
# }

# # ------------------- PAYOUT LOGIC -------------------
# def get_payin_category(payin: float):
#     if payin <= 20: return "Payin Below 20%"
#     elif payin <= 30: return "Payin 21% to 30%"
#     elif payin <= 50: return "Payin 31% to 50%"
#     else: return "Payin Above 50%"

# def safe_float(value):
#     if pd.isna(value): return None
#     val_str = str(value).strip().upper().replace('%', '')
#     if val_str in ["D", "NA", "", "NAN", "NONE"]: return None
#     try:
#         num = float(val_str)
#         return num * 100 if 0 < num < 1 else num
#     except: return None

# def get_formula_from_data(lob: str, segment: str, policy_type: str, payin: float):
#     segment_key = segment.upper()
#     if lob == "TW":
#         segment_key = "TW TP" if policy_type == "TP" else "TW SAOD + COMP"
#     elif lob == "PVT CAR":
#         segment_key = "PVT CAR TP" if policy_type == "TP" else "PVT CAR COMP + SAOD"
#     elif lob in ["TAXI", "CV", "BUS", "MISD"]:
#         segment_key = segment.upper()

#     payin_category = get_payin_category(payin)
#     matching_rule = None
    
#     for rule in FORMULA_DATA:
#         if rule["LOB"] == lob and rule["SEGMENT"] == segment_key:
#             if rule["REMARKS"] == payin_category or rule["REMARKS"] == "NIL":
#                 matching_rule = rule
#                 break
    
#     if not matching_rule and payin > 20:
#         for rule in FORMULA_DATA:
#             if rule["LOB"] == lob and rule["SEGMENT"] == segment_key and "Above 20%" in rule["REMARKS"]:
#                 matching_rule = rule
#                 break

#     if not matching_rule:
#         deduction = 2 if payin <= 20 else 3 if payin <= 30 else 4 if payin <= 50 else 5
#         return f"-{deduction}%", round(payin - deduction, 2)

#     formula = matching_rule["PO"]
#     if "% of Payin" in formula:
#         perc = float(formula.split("%")[0].strip())
#         return formula, round(payin * perc / 100, 2)
#     elif formula.startswith("-") and "%" in formula:
#         ded = float(formula.replace("-", "").replace("%", ""))
#         return formula, round(payin - ded, 2)
#     elif "Less" in formula and "%" in formula:
#         ded = float(formula.split()[1].replace("%", ""))
#         return formula, round(payin - ded, 2)
#     else:
#         return "-2%", round(payin - 2, 2)

# def calculate_payout_with_formula(lob: str, segment: str, policy_type: str, payin: float):
#     if payin == 0:
#         return 0, "0% (No Payin)", "Payin is 0"
#     formula, payout = get_formula_from_data(lob, segment, policy_type, payin)
#     return payout, formula, f"Match: LOB={lob}, Segment={segment}, Policy={policy_type}, {get_payin_category(payin)}"

# # ------------------- PATTERN DETECTION -------------------
# def detect_sheet_pattern(df):
#     """
#     Detects which processing pattern the sheet follows:
#     - 'electric': Electric vehicle sheet (simpler structure)
#     - 'regular': Regular taxi sheet with multiple columns (6+ combinations)
#     - 'compact': Compact sheet with just 2 CD2 columns
#     """
#     # Check first few rows for patterns
#     first_10_rows = df.head(10).astype(str).apply(lambda x: ' '.join(x), axis=1).str.upper()
#     all_text = ' '.join(first_10_rows)
    
#     # Electric sheet detection
#     if 'ELECTRIC' in all_text or 'EV' in all_text:
#         # Check if it has simple structure (fewer columns)
#         if df.shape[1] <= 10:
#             return 'electric'
    
#     # Check for CD2 columns
#     cd2_count = 0
#     for col_idx in range(df.shape[1]):
#         for r in range(min(10, df.shape[0])):
#             cell = str(df.iloc[r, col_idx]).strip().upper()
#             if "CD2" in cell:
#                 cd2_count += 1
#                 break
    
#     # Compact pattern: typically has exactly 2 CD2 columns
#     if cd2_count == 2:
#         # Check if it has fewer detail columns (compact format)
#         if df.shape[1] <= 8:
#             return 'compact'
    
#     # Regular pattern: has many columns for different combinations
#     if df.shape[1] >= 12:
#         return 'regular'
    
#     # Default to regular if uncertain
#     return 'regular'

# # ------------------- PROCESSORS -------------------
# def process_electric_sheet(df, override_enabled, override_lob, override_segment, override_policy_type):
#     """Process electric vehicle taxi sheets"""
#     records = []
#     for _, row in df.iterrows():
#         if pd.isna(row.iloc[0]) or str(row.iloc[0]).strip() == "":
#             continue
        
#         city_cluster = str(row.iloc[0]).strip()
#         rto_remarks = str(row.iloc[1]).strip() if len(row) > 1 else ""
#         fuel = str(row.iloc[2]).strip() if len(row) > 2 else "Electric"
#         make = str(row.iloc[3]).strip() if len(row) > 3 else "All"
#         seating = str(row.iloc[4]).strip() if len(row) > 4 else "5"
        
#         state = next((v for k, v in STATE_MAPPING.items() if k.upper() in city_cluster.upper()), "UNKNOWN")
#         cvod_cd2 = safe_float(row.iloc[6]) if len(row) > 6 else None
#         cvtp_cd2 = safe_float(row.iloc[7]) if len(row) > 7 else None
        
#         segment_desc = f"Taxi {fuel} {make} {rto_remarks} Seating:{seating}".strip()
        
#         lob_final = override_lob if override_enabled and override_lob else "TAXI"
#         segment_final = override_segment if override_enabled and override_segment else "TAXI"
        
#         if cvod_cd2 is not None:
#             policy_type = override_policy_type or "Comp"
#             payout, formula, rule_exp = calculate_payout_with_formula(lob_final, segment_final, policy_type, cvod_cd2)
#             records.append({
#                 "State": state.upper(),
#                 "Location/Cluster": city_cluster,
#                 "Original Segment": segment_desc,
#                 "Mapped Segment": segment_final,
#                 "LOB": lob_final,
#                 "Policy Type": policy_type,
#                 "Payin (CD2)": f"{cvod_cd2:.2f}%",
#                 "Payin Category": get_payin_category(cvod_cd2),
#                 "Calculated Payout": f"{payout:.2f}%",
#                 "Formula Used": formula,
#                 "Rule Explanation": rule_exp
#             })
        
#         if cvtp_cd2 is not None:
#             payout, formula, rule_exp = calculate_payout_with_formula(lob_final, segment_final, "TP", cvtp_cd2)
#             records.append({
#                 "State": state.upper(),
#                 "Location/Cluster": city_cluster,
#                 "Original Segment": segment_desc,
#                 "Mapped Segment": segment_final,
#                 "LOB": lob_final,
#                 "Policy Type": "TP",
#                 "Payin (CD2)": f"{cvtp_cd2:.2f}%",
#                 "Payin Category": get_payin_category(cvtp_cd2),
#                 "Calculated Payout": f"{payout:.2f}%",
#                 "Formula Used": formula,
#                 "Rule Explanation": rule_exp
#             })
#     return records

# def process_regular_sheet(df, override_enabled, override_lob, override_segment, override_policy_type):
#     """Process regular taxi sheets with multiple column combinations"""
#     records = []
#     prev_location = ""
    
#     for idx, row in df.iterrows():
#         if idx < 5:  # Skip header rows
#             continue
        
#         location = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else prev_location
#         if location: prev_location = location
        
#         fuel = str(row.iloc[1]).strip() if len(row) > 1 else ""
#         make = str(row.iloc[2]).strip() if len(row) > 2 else ""
#         remarks = str(row.iloc[3]).strip() if len(row) > 3 else ""
#         seating = str(row.iloc[4]).strip() if len(row) > 4 else ""
        
#         state = next((v for k, v in STATE_MAPPING.items() if k.upper() in location.upper()), "UNKNOWN")
#         cols = row.values
        
#         # Column combinations for different policy types
#         combinations = [
#             ("Without Add On Cover", "<=1000 CC", "Comp", 5, 6),
#             ("Without Add On Cover", ">1000 CC",  "Comp", 7, 8),
#             ("With Add On Cover",    "<=1000 CC", "Comp", 9, 10),
#             ("With Add On Cover",    ">1000 CC",  "Comp", 11, 12),
#             ("", "<=1000 CC", "TP", None, 13),
#             ("", ">1000 CC",  "TP", None, 14),
#         ]
        
#         for addon, cc, ptype, _, cd2_idx in combinations:
#             if cd2_idx >= len(cols): continue
#             payin = safe_float(cols[cd2_idx])
#             if payin is None: continue
            
#             segment_desc = f"Taxi {fuel} {make} {remarks}".strip()
#             if seating: segment_desc += f" Seating:{seating}"
#             if cc: segment_desc += f" {cc}"
#             if addon: segment_desc += f" {addon}"
            
#             lob_final = override_lob if override_enabled and override_lob else "TAXI"
#             segment_final = override_segment if override_enabled and override_segment else "TAXI"
#             policy_type_final = override_policy_type or ptype
            
#             payout, formula, rule_exp = calculate_payout_with_formula(lob_final, segment_final, policy_type_final, payin)
#             records.append({
#                 "State": state.upper(),
#                 "Location/Cluster": location,
#                 "Original Segment": segment_desc.strip(),
#                 "Mapped Segment": segment_final,
#                 "LOB": lob_final,
#                 "Policy Type": policy_type_final,
#                 "Payin (CD2)": f"{payin:.2f}%",
#                 "Payin Category": get_payin_category(payin),
#                 "Calculated Payout": f"{payout:.2f}%",
#                 "Formula Used": formula,
#                 "Rule Explanation": rule_exp
#             })
#     return records

# def process_compact_sheet(df, override_enabled, override_lob, override_segment, override_policy_type):
#     """Process compact taxi sheets with dynamic CD2 column detection"""
#     records = []
#     cluster = ""
    
#     comp_cd2_col = None
#     satp_cd2_col = None
    
#     # Scan all columns for headers containing "CD2"
#     cd2_candidates = []
#     for col_idx in range(df.shape[1]):
#         for r in range(min(10, df.shape[0])):
#             cell = str(df.iloc[r, col_idx]).strip().upper()
#             if "CD2" in cell:
#                 cd2_candidates.append((col_idx, r))
#                 break
    
#     if len(cd2_candidates) < 2:
#         return records
    
#     # Assign Comp and SATP based on nearby cells
#     for col_idx, row_idx in cd2_candidates:
#         group = ""
#         # Check left/right in same row
#         for dcol in [-1, 1]:
#             if 0 <= col_idx + dcol < df.shape[1]:
#                 nearby = str(df.iloc[row_idx, col_idx + dcol]).strip().upper()
#                 if nearby:
#                     group = nearby
#                     break
#         # Check above
#         if not group:
#             for dr in range(1, 5):
#                 if row_idx - dr >= 0:
#                     above = str(df.iloc[row_idx - dr, col_idx]).strip().upper()
#                     if above:
#                         group = above
#                         break
#         # Check diagonal (left columns above)
#         if not group:
#             for dcol in [-1, -2]:
#                 if 0 <= col_idx + dcol < df.shape[1]:
#                     for dr in range(1, 5):
#                         if row_idx - dr >= 0:
#                             above_left = str(df.iloc[row_idx - dr, col_idx + dcol]).strip().upper()
#                             if above_left:
#                                 group = above_left
#                                 break
        
#         if "COMP" in group or "OD" in group:
#             comp_cd2_col = col_idx
#         elif "SATP" in group or "TP" in group:
#             satp_cd2_col = col_idx
    
#     if comp_cd2_col is None or satp_cd2_col is None:
#         return records
    
#     # Find data start (first row with numeric in Comp column)
#     data_start = 0
#     for r in range(df.shape[0]):
#         if safe_float(df.iloc[r, comp_cd2_col]) is not None:
#             data_start = r
#             break
    
#     for r in range(data_start, df.shape[0]):
#         row = df.iloc[r]
        
#         if pd.notna(row.iloc[0]) and str(row.iloc[0]).strip():
#             cluster = str(row.iloc[0]).strip()
        
#         if df.shape[1] < 2 or pd.isna(row.iloc[1]) or str(row.iloc[1]).strip() == "":
#             continue
        
#         segment = str(row.iloc[1]).strip()
#         make = str(row.iloc[2]).strip() if df.shape[1] > 2 and pd.notna(row.iloc[2]) else ""
#         remarks = str(row.iloc[-1]).strip() if pd.notna(row.iloc[-1]) else ""
        
#         segment_desc = f"Taxi {segment} {make}".strip()
#         if remarks: segment_desc += f" {remarks}"
        
#         state = next((v for k, v in STATE_MAPPING.items() if k.upper() in cluster.upper()), "UNKNOWN")
        
#         lob_final = override_lob if override_enabled and override_lob else "TAXI"
#         segment_final = override_segment if override_enabled and override_segment else "TAXI"
        
#         comp_payin = safe_float(row.iloc[comp_cd2_col])
#         if comp_payin is not None:
#             policy_type = override_policy_type or "Comp"
#             payout, formula, rule_exp = calculate_payout_with_formula(lob_final, segment_final, policy_type, comp_payin)
#             records.append({
#                 "State": state.upper(),
#                 "Location/Cluster": cluster,
#                 "Original Segment": segment_desc,
#                 "Mapped Segment": segment_final,
#                 "LOB": lob_final,
#                 "Policy Type": policy_type,
#                 "Payin (CD2)": f"{comp_payin:.2f}%",
#                 "Payin Category": get_payin_category(comp_payin),
#                 "Calculated Payout": f"{payout:.2f}%",
#                 "Formula Used": formula,
#                 "Rule Explanation": rule_exp
#             })
        
#         tp_payin = safe_float(row.iloc[satp_cd2_col])
#         if tp_payin is not None:
#             payout, formula, rule_exp = calculate_payout_with_formula(lob_final, segment_final, "TP", tp_payin)
#             records.append({
#                 "State": state.upper(),
#                 "Location/Cluster": cluster,
#                 "Original Segment": segment_desc,
#                 "Mapped Segment": segment_final,
#                 "LOB": lob_final,
#                 "Policy Type": "TP",
#                 "Payin (CD2)": f"{tp_payin:.2f}%",
#                 "Payin Category": get_payin_category(tp_payin),
#                 "Calculated Payout": f"{payout:.2f}%",
#                 "Formula Used": formula,
#                 "Rule Explanation": rule_exp
#             })
    
#     return records

# # ------------------- INTELLIGENT DISPATCHER -------------------
# def intelligent_dispatcher(df, sheet_name, override_enabled, override_lob, override_segment, override_policy_type):
#     """
#     Intelligent dispatcher that detects the Excel pattern and calls the appropriate processor
#     """
#     pattern = detect_sheet_pattern(df)
    
#     print(f"   [DISPATCHER] Detected pattern: {pattern.upper()}")
    
#     if pattern == 'electric':
#         processor_name = "Electric Vehicle Processor"
#         records = process_electric_sheet(df, override_enabled, override_lob, override_segment, override_policy_type)
#     elif pattern == 'compact':
#         processor_name = "Compact Sheet Processor"
#         records = process_compact_sheet(df, override_enabled, override_lob, override_segment, override_policy_type)
#     else:  # regular
#         processor_name = "Regular Sheet Processor"
#         records = process_regular_sheet(df, override_enabled, override_lob, override_segment, override_policy_type)
    
#     print(f"   [DISPATCHER] Used: {processor_name}")
#     print(f"   [DISPATCHER] Extracted {len(records)} records")
    
#     return records, processor_name, pattern

# # ------------------- API ENDPOINTS -------------------
# @app.get("/")
# async def root():
#     return {
#         "message": "TAXI Insurance Policy Processing API",
#         "endpoints": {
#             "/taxi": "Process TAXI insurance data",
#             "/get-sheets": "Get worksheet names from Excel file"
#         },
#         "features": [
#             "Intelligent pattern detection",
#             "Electric vehicle sheet processing",
#             "Regular taxi sheet processing",
#             "Compact sheet processing",
#             "Manual worksheet selection",
#             "Automatic processing for single-sheet files"
#         ]
#     }

# @app.post("/get-sheets")
# async def get_sheets(file: UploadFile = File(...)):
#     """
#     Get list of all worksheet names from Excel file
#     Returns sheet names for manual selection
#     """
#     try:
#         contents = await file.read()
#         xls = pd.ExcelFile(io.BytesIO(contents))
#         sheet_names = xls.sheet_names
        
#         return JSONResponse(content={
#             "success": True,
#             "sheets": sheet_names,
#             "total_sheets": len(sheet_names),
#             "message": f"Found {len(sheet_names)} worksheet(s)"
#         })
#     except Exception as e:
#         return JSONResponse(
#             status_code=400,
#             content={
#                 "success": False,
#                 "error": f"Error reading Excel file: {str(e)}",
#                 "sheets": [],
#                 "total_sheets": 0
#             }
#         )

# @app.post("/taxi")
# async def process_taxi(
#     file: UploadFile = File(...),
#     company_name: str = Form("Digit"),
#     sheet_name: Optional[str] = Form(None),
#     override_segment: Optional[str] = Form(None),
#     override_policy_type: Optional[str] = Form(None)
# ):
#     """
#     Process TAXI insurance data with intelligent pattern detection
    
#     Logic:
#     1. If sheet_name is provided -> Process only that specific sheet
#     2. If sheet_name is None:
#        - If file has only 1 worksheet -> Process it directly
#        - If file has multiple worksheets -> Return error asking for sheet selection
#     """
#     try:
#         # Read the uploaded Excel file
#         contents = await file.read()
        
#         try:
#             xls = pd.ExcelFile(io.BytesIO(contents))
#             sheet_names = xls.sheet_names
            
#             print(f"[INFO] File has {len(sheet_names)} worksheet(s): {sheet_names}")
            
#             # ==================== WORKSHEET SELECTION LOGIC ====================
#             sheets_to_process = []
            
#             if sheet_name:
#                 # User specified a sheet name - validate and process only that sheet
#                 if sheet_name not in sheet_names:
#                     return JSONResponse(
#                         status_code=400,
#                         content={
#                             "success": False,
#                             "error": f"Worksheet '{sheet_name}' not found in the file.",
#                             "available_sheets": sheet_names,
#                             "message": f"Available worksheets: {', '.join(sheet_names)}"
#                         }
#                     )
#                 sheets_to_process = [sheet_name]
#                 print(f"[INFO] Processing user-selected worksheet: {sheet_name}")
                
#             else:
#                 # No sheet name provided - check number of worksheets
#                 if len(sheet_names) == 1:
#                     # Only one worksheet - process it directly
#                     sheets_to_process = sheet_names
#                     print(f"[INFO] Single worksheet detected - processing directly: {sheet_names[0]}")
                    
#                 else:
#                     # Multiple worksheets - require manual selection
#                     return JSONResponse(
#                         status_code=400,
#                         content={
#                             "success": False,
#                             "error": "Multiple worksheets found. Please select a worksheet to process.",
#                             "available_sheets": sheet_names,
#                             "total_sheets": len(sheet_names),
#                             "message": f"This file contains {len(sheet_names)} worksheets. Please select one to process.",
#                             "require_sheet_selection": True
#                         }
#                     )
            
#             # ==================== PROCESS SELECTED SHEET(S) ====================
#             all_records = []
#             processors_used = []
#             patterns_detected = []
            
#             for sheet in sheets_to_process:
#                 print(f"\n[PROCESSING] Sheet: {sheet}")
#                 df = pd.read_excel(io.BytesIO(contents), sheet_name=sheet, header=None)
                
#                 # Use intelligent dispatcher
#                 records, processor_name, pattern = intelligent_dispatcher(
#                     df, 
#                     sheet, 
#                     override_enabled=False,  # Not using overrides for TAXI
#                     override_lob="TAXI",      # Always TAXI
#                     override_segment=override_segment,
#                     override_policy_type=override_policy_type
#                 )
                
#                 all_records.extend(records)
#                 processors_used.append(f"{sheet}: {processor_name} ({pattern})")
#                 patterns_detected.append(pattern)
            
#             # ==================== VALIDATION ====================
#             if not all_records:
#                 return JSONResponse(
#                     status_code=400,
#                     content={
#                         "success": False,
#                         "error": "No processable data found in the uploaded file",
#                         "message": "The selected worksheet does not contain valid TAXI data or the format is not recognized."
#                     }
#                 )
            
#             # ==================== GENERATE OUTPUT ====================
#             # Create result DataFrame
#             result_df = pd.DataFrame(all_records)
            
#             # Calculate statistics
#             payin_values = []
#             for record in all_records:
#                 payin_str = record.get("Payin (CD2)", "0%")
#                 try:
#                     payin_values.append(float(payin_str.replace("%", "")))
#                 except:
#                     pass
            
#             avg_payin = round(sum(payin_values) / len(payin_values), 2) if payin_values else 0
#             unique_segments = len(result_df["Mapped Segment"].unique()) if "Mapped Segment" in result_df.columns else 0
            
#             # Generate Excel file
#             output = io.BytesIO()
#             with pd.ExcelWriter(output, engine='openpyxl') as writer:
#                 result_df.to_excel(writer, index=False, sheet_name='Processed Data')
#             output.seek(0)
#             excel_base64 = base64.b64encode(output.read()).decode()
            
#             # Generate CSV
#             csv_output = result_df.to_csv(index=False)
            
#             # Formula summary
#             formula_summary = {}
#             for record in all_records:
#                 formula = record.get("Formula Used", "Unknown")
#                 formula_summary[formula] = formula_summary.get(formula, 0) + 1
            
#             # Pattern summary
#             pattern_summary = {}
#             for pattern in patterns_detected:
#                 pattern_summary[pattern] = pattern_summary.get(pattern, 0) + 1
            
#             # ==================== RETURN RESPONSE ====================
#             return JSONResponse(content={
#                 "success": True,
#                 "company_name": company_name,
#                 "lob": "TAXI",
#                 "sheet_processed": sheets_to_process[0] if sheets_to_process else "Unknown",
#                 "total_sheets_in_file": len(sheet_names),
#                 "processors_used": processors_used,
#                 "patterns_detected": pattern_summary,
#                 "total_records": len(all_records),
#                 "avg_payin": avg_payin,
#                 "unique_segments": unique_segments,
#                 "calculated_data": all_records,
#                 "formula_data": FORMULA_DATA,
#                 "formula_summary": formula_summary,
#                 "excel_data": excel_base64,
#                 "csv_data": csv_output,
#                 "extracted_text": f"Processed {len(all_records)} TAXI records from worksheet '{sheets_to_process[0]}' using intelligent pattern detection",
#                 "parsed_data": all_records[:5]  # First 5 records for preview
#             })
            
#         except Exception as e:
#             import traceback
#             traceback.print_exc()
#             return JSONResponse(
#                 status_code=400,
#                 content={
#                     "success": False,
#                     "error": f"Error processing Excel file: {str(e)}",
#                     "message": "There was an error reading or processing the Excel file. Please check the file format."
#                 }
#             )
    
#     except Exception as e:
#         import traceback
#         traceback.print_exc()
#         return JSONResponse(
#             status_code=500,
#             content={
#                 "success": False,
#                 "error": f"Server error: {str(e)}",
#                 "message": "An internal server error occurred. Please try again."
#             }
#         )

# if __name__ == "__main__":
#     import uvicorn
#     print("\n" + "="*80)
#     print(" "*25 + "TAXI INSURANCE PROCESSOR API")
#     print("="*80)
#     print("\nStarting server on http://localhost:8000")
#     print("\nFeatures:")
#     print("  ✓ Manual worksheet selection for multi-sheet files")
#     print("  ✓ Automatic processing for single-sheet files")
#     print("  ✓ Intelligent pattern detection (Electric/Regular/Compact)")
#     print("  ✓ Real-time payout calculation")
#     print("\nEndpoints:")
#     print("  • GET  /          - API information")
#     print("  • POST /get-sheets - Get worksheet names")
#     print("  • POST /taxi       - Process TAXI data")
#     print("="*80 + "\n")
    
#     uvicorn.run(app, host="0.0.0.0", port=8000)





from fastapi import FastAPI, File, UploadFile, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
import pandas as pd
import io
import base64
from typing import Optional

app = FastAPI(title="TAXI Insurance Policy Processor")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://digit-excel-taxi.vercel.app"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ------------------- FORMULA DATA -------------------
FORMULA_DATA = [
    {"LOB": "TW", "SEGMENT": "1+5", "PO": "90% of Payin", "REMARKS": "NIL"},
    {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "PO": "-5%", "REMARKS": "Payin Above 50%"},
    {"LOB": "TW", "SEGMENT": "TW TP", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "TW", "SEGMENT": "TW TP", "PO": "-3%", "REMARKS": "Payin Above 20%"},
    {"LOB": "PVT CAR", "SEGMENT": "PVT CAR COMP + SAOD", "PO": "90% of Payin", "REMARKS": "NIL"},
    {"LOB": "PVT CAR", "SEGMENT": "PVT CAR TP", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "PVT CAR", "SEGMENT": "PVT CAR TP", "PO": "-3%", "REMARKS": "Payin Above 20%"},
    {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-5%", "REMARKS": "Payin Above 50%"},
    {"LOB": "BUS", "SEGMENT": "SCHOOL BUS", "PO": "Less 2% of Payin", "REMARKS": "NIL"},
    {"LOB": "BUS", "SEGMENT": "STAFF BUS", "PO": "88% of Payin", "REMARKS": "NIL"},
    {"LOB": "TAXI", "SEGMENT": "TAXI", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "TAXI", "SEGMENT": "TAXI", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "TAXI", "SEGMENT": "TAXI", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "TAXI", "SEGMENT": "TAXI", "PO": "-5%", "REMARKS": "Payin Above 50%"},
    {"LOB": "MISD", "SEGMENT": "Misd, Tractor", "PO": "88% of Payin", "REMARKS": "NIL"}
]

# ------------------- STATE MAPPING -------------------
STATE_MAPPING = {
    "DELHI": "DELHI", "MUMBAI": "MAHARASHTRA", "PUNE": "MAHARASHTRA", "GOA": "GOA",
    "KOLKATA": "WEST BENGAL", "HYDERABAD": "TELANGANA", "AHMEDABAD": "GUJARAT",
    "SURAT": "GUJARAT", "JAIPUR": "RAJASTHAN", "LUCKNOW": "UTTAR PRADESH",
    "PATNA": "BIHAR", "RANCHI": "JHARKHAND", "BHUVANESHWAR": "ODISHA",
    "SRINAGAR": "JAMMU AND KASHMIR", "DEHRADUN": "UTTARAKHAND", "HARIDWAR": "UTTARAKHAND",
    "HIMACHAL PRADESH": "HIMACHAL PRADESH", "ANDAMAN": "ANDAMAN AND NICOBAR ISLANDS",
    "BANGALORE": "KARNATAKA", "JHARKHAND": "JHARKHAND", "BIHAR": "BIHAR",
    "WEST BENGAL": "WEST BENGAL", "NORTH BENGAL": "WEST BENGAL", "ORISSA": "ODISHA",
    "GOOD GJ": "GUJARAT", "BAD GJ": "GUJARAT", "ROM1": "REST OF MAHARASHTRA",
    "ROM2": "REST OF MAHARASHTRA", "GOOD VIZAG": "ANDHRA PRADESH", "GOOD TN": "TAMIL NADU",
    "KERALA": "KERALA", "GOOD MP": "MADHYA PRADESH", "GOOD CG": "CHHATTISGARH",
    "GOOD RJ": "RAJASTHAN", "BAD RJ": "RAJASTHAN", "GOOD UP": "UTTAR PRADESH",
    "BAD UP": "UTTAR PRADESH", "GOOD UK": "UTTARAKHAND", "BAD UK": "UTTARAKHAND",
    "PUNJAB": "PUNJAB", "JAMMU": "JAMMU AND KASHMIR", "ASSAM": "ASSAM",
    "NE EX ASSAM": "NORTH EAST", "GOOD NL": "NAGALAND", "GOOD KA": "KARNATAKA",
    "BAD KA": "KARNATAKA", "HR REF": "HARYANA", "DEHRADUN, HARIDWAR": "UTTARAKHAND",
    "HIMACHAL": "HIMACHAL PRADESH"
}

# ------------------- HELPER: SAFE CELL TO STRING -------------------
def cell_to_str(val) -> str:
    """
    THE FIX: Safely converts ANY cell value (float, int, NaN, None, str) to string.
    This prevents 'expected str instance, float found' errors in join() calls.
    """
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except (TypeError, ValueError):
        pass
    return str(val).strip()

# ------------------- PAYOUT LOGIC -------------------
def get_payin_category(payin: float):
    if payin <= 20: return "Payin Below 20%"
    elif payin <= 30: return "Payin 21% to 30%"
    elif payin <= 50: return "Payin 31% to 50%"
    else: return "Payin Above 50%"

def safe_float(value):
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    val_str = str(value).strip().upper().replace('%', '')
    if val_str in ["D", "NA", "", "NAN", "NONE", "DECLINE"]:
        return None
    try:
        num = float(val_str)
        if num < 0:
            return None
        return num * 100 if 0 < num < 1 else num
    except:
        return None

def get_formula_from_data(lob: str, segment: str, policy_type: str, payin: float):
    segment_key = segment.upper()
    if lob == "TW":
        segment_key = "TW TP" if policy_type == "TP" else "TW SAOD + COMP"
    elif lob == "PVT CAR":
        segment_key = "PVT CAR TP" if policy_type == "TP" else "PVT CAR COMP + SAOD"
    elif lob in ["TAXI", "CV", "BUS", "MISD"]:
        segment_key = segment.upper()

    payin_category = get_payin_category(payin)
    matching_rule = None

    for rule in FORMULA_DATA:
        if rule["LOB"] == lob and rule["SEGMENT"] == segment_key:
            if rule["REMARKS"] == payin_category or rule["REMARKS"] == "NIL":
                matching_rule = rule
                break

    if not matching_rule and payin > 20:
        for rule in FORMULA_DATA:
            if rule["LOB"] == lob and rule["SEGMENT"] == segment_key and "Above 20%" in rule["REMARKS"]:
                matching_rule = rule
                break

    if not matching_rule:
        deduction = 2 if payin <= 20 else 3 if payin <= 30 else 4 if payin <= 50 else 5
        return f"-{deduction}%", round(payin - deduction, 2)

    formula = matching_rule["PO"]
    if "% of Payin" in formula:
        perc = float(formula.split("%")[0].strip())
        return formula, round(payin * perc / 100, 2)
    elif formula.startswith("-") and "%" in formula:
        ded = float(formula.replace("-", "").replace("%", ""))
        return formula, round(payin - ded, 2)
    elif "Less" in formula and "%" in formula:
        ded = float(formula.split()[1].replace("%", ""))
        return formula, round(payin - ded, 2)
    else:
        return "-2%", round(payin - 2, 2)

def calculate_payout_with_formula(lob: str, segment: str, policy_type: str, payin: float):
    if payin == 0:
        return 0, "0% (No Payin)", "Payin is 0"
    formula, payout = get_formula_from_data(lob, segment, policy_type, payin)
    return payout, formula, f"Match: LOB={lob}, Segment={segment}, Policy={policy_type}, {get_payin_category(payin)}"

# ------------------- PATTERN DETECTION (FIXED) -------------------
def detect_sheet_pattern(df):
    """
    Detects sheet pattern. Patterns:
      - 'cluster_segment' : Cluster/Segment/Make cols + COMP & SATP group headers (NEW)
      - 'electric'        : Electric vehicle sheet
      - 'compact'         : Compact 2-CD2-column sheet
      - 'regular'         : Regular multi-column sheet (12+ cols)
    """
    sample_rows = []
    for i in range(min(10, df.shape[0])):
        row_str = " ".join(cell_to_str(val) for val in df.iloc[i])
        sample_rows.append(row_str.upper())
    all_text = " ".join(sample_rows)

    # NEW PATTERN: Cluster/Segment/Make + COMP + SATP group headers
    if ("COMP" in all_text and "SATP" in all_text
            and "CLUSTER" in all_text and "SEGMENT" in all_text):
        return "cluster_segment"

    # Electric sheet detection
    if "ELECTRIC" in all_text or "EV" in all_text:
        if df.shape[1] <= 10:
            return "electric"

    # Count CD2 columns
    cd2_count = 0
    for col_idx in range(df.shape[1]):
        for r in range(min(10, df.shape[0])):
            cell = cell_to_str(df.iloc[r, col_idx]).upper()
            if "CD2" in cell:
                cd2_count += 1
                break

    if cd2_count == 2 and df.shape[1] <= 8:
        return "compact"

    if df.shape[1] >= 12:
        return "regular"

    return "regular"

# ------------------- PROCESSORS -------------------
def process_electric_sheet(df, override_enabled, override_lob, override_segment, override_policy_type):
    """Process electric vehicle taxi sheets - FIXED with cell_to_str"""
    records = []

    # Find the actual data start row (skip empty rows and header rows)
    data_start = 0
    for i in range(min(10, df.shape[0])):
        cell = cell_to_str(df.iloc[i, 0]).upper()
        # Skip empty rows and header rows
        if cell and cell not in ["CITY/CLUSTER", "RTO LOCATION", ""] and "CITY" not in cell:
            data_start = i
            break

    for row_idx in range(data_start, df.shape[0]):
        row = df.iloc[row_idx]

        city_cluster = cell_to_str(row.iloc[0])
        if not city_cluster:
            continue

        # Skip header-like rows
        if city_cluster.upper() in ["CITY/CLUSTER", "RTO LOCATION", "CLUSTER"]:
            continue

        rto_remarks   = cell_to_str(row.iloc[1]) if len(row) > 1 else ""
        fuel          = cell_to_str(row.iloc[2]) if len(row) > 2 else "Electric"
        make          = cell_to_str(row.iloc[3]) if len(row) > 3 else "All"
        seating       = cell_to_str(row.iloc[4]) if len(row) > 4 else "5"

        # CD1 is column F (index 5) — IGNORED
        # CVOD CD2 is column G (index 6)
        # CVTP CD2 is column H (index 7)
        cvod_cd2 = safe_float(row.iloc[6]) if len(row) > 6 else None
        cvtp_cd2 = safe_float(row.iloc[7]) if len(row) > 7 else None

        state = next(
            (v for k, v in STATE_MAPPING.items() if k.upper() in city_cluster.upper()),
            "UNKNOWN"
        )

        segment_desc = f"Taxi {fuel} {make}".strip()
        if rto_remarks:
            segment_desc += f" {rto_remarks}"
        if seating:
            segment_desc += f" Seating:{seating}"

        lob_final     = override_lob     if override_enabled and override_lob     else "TAXI"
        segment_final = override_segment if override_enabled and override_segment else "TAXI"

        if cvod_cd2 is not None:
            policy_type = override_policy_type or "Comp"
            payout, formula, rule_exp = calculate_payout_with_formula(
                lob_final, segment_final, policy_type, cvod_cd2
            )
            records.append({
                "State": state.upper(),
                "Location/Cluster": city_cluster,
                "Original Segment": segment_desc,
                "Mapped Segment": segment_final,
                "LOB": lob_final,
                "Policy Type": policy_type,
                "Payin (CD2)": f"{cvod_cd2:.2f}%",
                "Payin Category": get_payin_category(cvod_cd2),
                "Calculated Payout": f"{payout:.2f}%",
                "Formula Used": formula,
                "Rule Explanation": rule_exp
            })

        if cvtp_cd2 is not None:
            payout, formula, rule_exp = calculate_payout_with_formula(
                lob_final, segment_final, "TP", cvtp_cd2
            )
            records.append({
                "State": state.upper(),
                "Location/Cluster": city_cluster,
                "Original Segment": segment_desc,
                "Mapped Segment": segment_final,
                "LOB": lob_final,
                "Policy Type": "TP",
                "Payin (CD2)": f"{cvtp_cd2:.2f}%",
                "Payin Category": get_payin_category(cvtp_cd2),
                "Calculated Payout": f"{payout:.2f}%",
                "Formula Used": formula,
                "Rule Explanation": rule_exp
            })

    return records


def process_regular_sheet(df, override_enabled, override_lob, override_segment, override_policy_type):
    """
    Process regular taxi sheets (Petrol/CNG/Diesel) - FIXED with cell_to_str.
    Structure: Row 1=empty, Row 2=CVOD header, Row 3=Without/With Add On, 
               Row 4=CC ranges, Row 5=CD1/CD2_N headers, Row 6+ = data
    """
    records = []
    prev_location = ""

    # Find the actual data start row dynamically
    data_start = 6  # Default based on your sheet structure (rows 1-5 are headers)
    for i in range(min(15, df.shape[0])):
        cell = cell_to_str(df.iloc[i, 0]).upper()
        # First row where column A has a real location name (not a header)
        if cell and cell not in ["RTO LOCATION", "CLUSTER", "CITY/CLUSTER", ""] \
                and "RTO" not in cell and "LOCATION" not in cell \
                and "CD" not in cell and "COVER" not in cell:
            data_start = i
            break

    # Dynamically find CD2_N columns and their policy types
    # Scan rows 0-5 for column metadata
    comp_cd2_cols = []   # (col_index, label)
    tp_cd2_cols   = []   # (col_index, label)

    for col_idx in range(df.shape[1]):
        cd2_found   = False
        is_tp       = False
        is_comp     = False
        cc_label    = ""
        addon_label = ""

        for row_idx in range(min(6, df.shape[0])):
            cell = cell_to_str(df.iloc[row_idx, col_idx]).upper()

            if "CD2" in cell or "CD2_N" in cell:
                cd2_found = True

            # Look at cells above this column to determine policy type
            if "SATP" in cell or ("TP" in cell and "COMP" not in cell):
                is_tp = True
            if "CVOD" in cell or "COMP" in cell or "OD" in cell:
                is_comp = True
            if "<=1000" in cell or ">1000" in cell:
                cc_label = cell
            if "WITHOUT" in cell:
                addon_label = "Without Add On"
            elif "WITH" in cell and "WITHOUT" not in cell:
                addon_label = "With Add On"

        if cd2_found:
            # Also scan the entire column header area for TP/Comp context
            col_context = " ".join(
                cell_to_str(df.iloc[r, col_idx]).upper()
                for r in range(min(6, df.shape[0]))
            )
            # Also scan nearby columns for group headers (merged cells appear as NaN)
            nearby_context = ""
            for dcol in range(-3, 1):
                if 0 <= col_idx + dcol < df.shape[1]:
                    for row_idx in range(min(6, df.shape[0])):
                        nearby_context += " " + cell_to_str(
                            df.iloc[row_idx, col_idx + dcol]
                        ).upper()

            if "SATP" in nearby_context or ("CVTP" in nearby_context):
                tp_cd2_cols.append((col_idx, cc_label or ""))
            else:
                comp_cd2_cols.append((col_idx, cc_label or ""))

    print(f"   [REGULAR] Comp CD2 cols: {comp_cd2_cols}")
    print(f"   [REGULAR] TP   CD2 cols: {tp_cd2_cols}")
    print(f"   [REGULAR] Data starts at row: {data_start}")

    # Fallback to hardcoded indices if detection failed (matches your screenshot)
    # Cols: F=CD1(ignored), G=CD2_N(Comp,<=1000), H=CD1(ignored), I=CD2_N(Comp,>1000),
    #       J=CD1(ignored), K=CD2_N(Comp,<=1000,with addon), ... TP cols near end
    if not comp_cd2_cols and not tp_cd2_cols:
        comp_cd2_cols = [(6, "<=1000 CC"), (8, ">1000 CC"), (10, "<=1000 CC With AddOn"), (12, ">1000 CC With AddOn")]
        tp_cd2_cols   = [(14, "<=1000 CC"), (16, ">1000 CC")]

    for row_idx in range(data_start, df.shape[0]):
        row = df.iloc[row_idx]

        location = cell_to_str(row.iloc[0])
        if location:
            prev_location = location
        else:
            location = prev_location

        if not location:
            continue

        fuel    = cell_to_str(row.iloc[1]) if len(row) > 1 else ""
        make    = cell_to_str(row.iloc[2]) if len(row) > 2 else ""
        remarks = cell_to_str(row.iloc[3]) if len(row) > 3 else ""
        seating = cell_to_str(row.iloc[4]) if len(row) > 4 else ""

        # Skip rows that are clearly not data
        if fuel.upper() in ["FUEL", ""] and make.upper() in ["MAKE", ""]:
            continue

        state = next(
            (v for k, v in STATE_MAPPING.items() if k.upper() in location.upper()),
            "UNKNOWN"
        )

        lob_final     = override_lob     if override_enabled and override_lob     else "TAXI"
        segment_final = override_segment if override_enabled and override_segment else "TAXI"

        # Process Comp columns (CD2 only, skip CD1)
        for col_idx, cc_label in comp_cd2_cols:
            if col_idx >= len(row):
                continue
            payin = safe_float(row.iloc[col_idx])
            if payin is None:
                continue

            segment_desc = f"Taxi {fuel} {make}".strip()
            if remarks: segment_desc += f" {remarks}"
            if seating: segment_desc += f" Seating:{seating}"
            if cc_label: segment_desc += f" {cc_label}"

            policy_type = override_policy_type or "Comp"
            payout, formula, rule_exp = calculate_payout_with_formula(
                lob_final, segment_final, policy_type, payin
            )
            records.append({
                "State": state.upper(),
                "Location/Cluster": location,
                "Original Segment": segment_desc.strip(),
                "Mapped Segment": segment_final,
                "LOB": lob_final,
                "Policy Type": policy_type,
                "Payin (CD2)": f"{payin:.2f}%",
                "Payin Category": get_payin_category(payin),
                "Calculated Payout": f"{payout:.2f}%",
                "Formula Used": formula,
                "Rule Explanation": rule_exp
            })

        # Process TP columns (CD2 only, skip CD1)
        for col_idx, cc_label in tp_cd2_cols:
            if col_idx >= len(row):
                continue
            payin = safe_float(row.iloc[col_idx])
            if payin is None:
                continue

            segment_desc = f"Taxi {fuel} {make}".strip()
            if remarks: segment_desc += f" {remarks}"
            if seating: segment_desc += f" Seating:{seating}"
            if cc_label: segment_desc += f" {cc_label}"

            policy_type = override_policy_type or "TP"
            payout, formula, rule_exp = calculate_payout_with_formula(
                lob_final, segment_final, policy_type, payin
            )
            records.append({
                "State": state.upper(),
                "Location/Cluster": location,
                "Original Segment": segment_desc.strip(),
                "Mapped Segment": segment_final,
                "LOB": lob_final,
                "Policy Type": policy_type,
                "Payin (CD2)": f"{payin:.2f}%",
                "Payin Category": get_payin_category(payin),
                "Calculated Payout": f"{payout:.2f}%",
                "Formula Used": formula,
                "Rule Explanation": rule_exp
            })

    return records


def process_compact_sheet(df, override_enabled, override_lob, override_segment, override_policy_type):
    """Process compact taxi sheets - FIXED with cell_to_str"""
    records = []
    cluster = ""

    comp_cd2_col = None
    satp_cd2_col = None

    # Scan all columns for headers containing "CD2" — FIXED using cell_to_str
    cd2_candidates = []
    for col_idx in range(df.shape[1]):
        for r in range(min(10, df.shape[0])):
            cell = cell_to_str(df.iloc[r, col_idx]).upper()
            if "CD2" in cell:
                cd2_candidates.append((col_idx, r))
                break

    if len(cd2_candidates) < 2:
        return records

    for col_idx, row_idx in cd2_candidates:
        group = ""
        for dcol in [-1, 1]:
            if 0 <= col_idx + dcol < df.shape[1]:
                nearby = cell_to_str(df.iloc[row_idx, col_idx + dcol]).upper()
                if nearby:
                    group = nearby
                    break
        if not group:
            for dr in range(1, 5):
                if row_idx - dr >= 0:
                    above = cell_to_str(df.iloc[row_idx - dr, col_idx]).upper()
                    if above:
                        group = above
                        break
        if not group:
            for dcol in [-1, -2]:
                if 0 <= col_idx + dcol < df.shape[1]:
                    for dr in range(1, 5):
                        if row_idx - dr >= 0:
                            above_left = cell_to_str(df.iloc[row_idx - dr, col_idx + dcol]).upper()
                            if above_left:
                                group = above_left
                                break

        if "COMP" in group or "OD" in group or "CVOD" in group:
            comp_cd2_col = col_idx
        elif "SATP" in group or "TP" in group or "CVTP" in group:
            satp_cd2_col = col_idx

    if comp_cd2_col is None or satp_cd2_col is None:
        return records

    data_start = 0
    for r in range(df.shape[0]):
        if safe_float(df.iloc[r, comp_cd2_col]) is not None:
            data_start = r
            break

    for r in range(data_start, df.shape[0]):
        row = df.iloc[r]

        first_cell = cell_to_str(row.iloc[0])
        if first_cell:
            cluster = first_cell

        second_cell = cell_to_str(row.iloc[1]) if df.shape[1] > 1 else ""
        if not second_cell:
            continue

        segment = second_cell
        make    = cell_to_str(row.iloc[2]) if df.shape[1] > 2 else ""
        remarks = cell_to_str(row.iloc[-1]) if df.shape[1] > 3 else ""

        segment_desc = f"Taxi {segment} {make}".strip()
        if remarks:
            segment_desc += f" {remarks}"

        state = next(
            (v for k, v in STATE_MAPPING.items() if k.upper() in cluster.upper()),
            "UNKNOWN"
        )

        lob_final     = override_lob     if override_enabled and override_lob     else "TAXI"
        segment_final = override_segment if override_enabled and override_segment else "TAXI"

        comp_payin = safe_float(row.iloc[comp_cd2_col])
        if comp_payin is not None:
            policy_type = override_policy_type or "Comp"
            payout, formula, rule_exp = calculate_payout_with_formula(
                lob_final, segment_final, policy_type, comp_payin
            )
            records.append({
                "State": state.upper(),
                "Location/Cluster": cluster,
                "Original Segment": segment_desc,
                "Mapped Segment": segment_final,
                "LOB": lob_final,
                "Policy Type": policy_type,
                "Payin (CD2)": f"{comp_payin:.2f}%",
                "Payin Category": get_payin_category(comp_payin),
                "Calculated Payout": f"{payout:.2f}%",
                "Formula Used": formula,
                "Rule Explanation": rule_exp
            })

        tp_payin = safe_float(row.iloc[satp_cd2_col])
        if tp_payin is not None:
            payout, formula, rule_exp = calculate_payout_with_formula(
                lob_final, segment_final, "TP", tp_payin
            )
            records.append({
                "State": state.upper(),
                "Location/Cluster": cluster,
                "Original Segment": segment_desc,
                "Mapped Segment": segment_final,
                "LOB": lob_final,
                "Policy Type": "TP",
                "Payin (CD2)": f"{tp_payin:.2f}%",
                "Payin Category": get_payin_category(tp_payin),
                "Calculated Payout": f"{payout:.2f}%",
                "Formula Used": formula,
                "Rule Explanation": rule_exp
            })

    return records



# ------------------- NEW PATTERN: CLUSTER/SEGMENT/MAKE + COMP & SATP -------------------
def process_cluster_segment_sheet(df, override_enabled, override_lob, override_segment, override_policy_type):
    """
    Processes sheets structured as:
      Row 0: (empty) | (empty) | (empty) | COMP (merged) | (empty) | SATP | (empty)
      Row 1: Cluster | Segment | Make    | CD1           | CD2     | CD2  | Remarks
      Row 2+: data   (Cluster blank = carry forward)

    Rules:
      - CD1 columns are IGNORED entirely
      - COMP CD2  -> Policy Type = Comp
      - SATP CD2  -> Policy Type = TP
      - Make value goes into Remarks in output
    """
    records = []

    # Step 1: find GROUP HEADER row (has both COMP and SATP)
    group_header_row = None
    for i in range(min(10, df.shape[0])):
        row_text = " ".join(cell_to_str(v) for v in df.iloc[i]).upper()
        if "COMP" in row_text and "SATP" in row_text:
            group_header_row = i
            break

    if group_header_row is None:
        print("   [CLUSTER_SEG] Could not find COMP/SATP header row")
        return records

    # Step 2: find COLUMN HEADER row (Cluster, Segment, CD1, CD2 ...)
    col_header_row = group_header_row + 1
    for i in range(group_header_row + 1, min(group_header_row + 4, df.shape[0])):
        row_text = " ".join(cell_to_str(v) for v in df.iloc[i]).upper()
        if "CLUSTER" in row_text or "SEGMENT" in row_text or "CD2" in row_text:
            col_header_row = i
            break

    # Step 3: map group headers -> COMP / SATP for each column index
    group_map = {}
    current_group = None
    for col_idx in range(df.shape[1]):
        val = cell_to_str(df.iloc[group_header_row, col_idx]).upper()
        if "COMP" in val:
            current_group = "COMP"
        elif "SATP" in val or ("TP" in val and "COMP" not in val):
            current_group = "SATP"
        if current_group:
            group_map[col_idx] = current_group

    print(f"   [CLUSTER_SEG] Group map: {group_map}")

    # Step 4: identify exact column roles from column header row
    comp_cd2_col = None
    satp_cd2_col = None
    cluster_col  = 0
    segment_col  = 1
    make_col     = 2
    remarks_col  = None

    for col_idx in range(df.shape[1]):
        header = cell_to_str(df.iloc[col_header_row, col_idx]).upper()
        group  = group_map.get(col_idx, "")

        if "CLUSTER" in header:
            cluster_col = col_idx
        elif "SEGMENT" in header:
            segment_col = col_idx
        elif "MAKE" in header:
            make_col = col_idx
        elif "REMARK" in header:
            remarks_col = col_idx
        elif "CD2" in header and "CD1" not in header:
            # Assign to correct policy group — ignore CD1 completely
            if group == "COMP" and comp_cd2_col is None:
                comp_cd2_col = col_idx
            elif group == "SATP" and satp_cd2_col is None:
                satp_cd2_col = col_idx

    # Fallback column positions (from screenshot: A=0,B=1,C=2,D=CD1,E=CD2-Comp,F=CD2-SATP,G=Remarks)
    if comp_cd2_col is None:
        comp_cd2_col = 4
    if satp_cd2_col is None:
        satp_cd2_col = 5
    if remarks_col is None:
        remarks_col = df.shape[1] - 1

    print(f"   [CLUSTER_SEG] cluster={cluster_col}, segment={segment_col}, "
          f"make={make_col}, comp_cd2={comp_cd2_col}, satp_cd2={satp_cd2_col}, remarks={remarks_col}")

    # Step 5: process data rows
    data_start      = col_header_row + 1
    current_cluster = ""

    lob_final              = override_lob      if override_enabled and override_lob      else "TAXI"
    segment_final_override = override_segment  if override_enabled and override_segment  else None

    for row_idx in range(data_start, df.shape[0]):
        row = df.iloc[row_idx]

        # Carry-forward cluster name
        cluster_val = cell_to_str(row.iloc[cluster_col])
        if cluster_val:
            current_cluster = cluster_val
        if not current_cluster:
            continue

        segment_val = cell_to_str(row.iloc[segment_col])
        if not segment_val:
            continue        # skip rows with no segment info

        make_val    = cell_to_str(row.iloc[make_col])
        remarks_val = cell_to_str(row.iloc[remarks_col]) if remarks_col < len(row) else ""

        # Make goes into Remarks field in output
        output_remarks = " | ".join(filter(None, [make_val, remarks_val]))

        state = next(
            (v for k, v in STATE_MAPPING.items() if k.upper() in current_cluster.upper()),
            "UNKNOWN"
        )

        mapped_seg = segment_final_override or segment_val

        # COMP CD2
        comp_payin = safe_float(row.iloc[comp_cd2_col]) if comp_cd2_col < len(row) else None
        if comp_payin is not None:
            policy_type = override_policy_type or "Comp"
            payout, formula, rule_exp = calculate_payout_with_formula(
                lob_final, mapped_seg, policy_type, comp_payin
            )
            records.append({
                "State":             state.upper(),
                "Location/Cluster":  current_cluster,
                "Original Segment":  segment_val,
                "Mapped Segment":    mapped_seg,
                "LOB":               lob_final,
                "Policy Type":       policy_type,
                "Status":            "STP",
                "Payin (CD2)":       f"{comp_payin:.2f}%",
                "Payin Category":    get_payin_category(comp_payin),
                "Calculated Payout": f"{payout:.2f}%",
                "Formula Used":      formula,
                "Rule Explanation":  rule_exp,
                "Remarks":           output_remarks,
            })

        # SATP = TP CD2
        satp_payin = safe_float(row.iloc[satp_cd2_col]) if satp_cd2_col < len(row) else None
        if satp_payin is not None:
            policy_type = override_policy_type or "TP"
            payout, formula, rule_exp = calculate_payout_with_formula(
                lob_final, mapped_seg, policy_type, satp_payin
            )
            records.append({
                "State":             state.upper(),
                "Location/Cluster":  current_cluster,
                "Original Segment":  segment_val,
                "Mapped Segment":    mapped_seg,
                "LOB":               lob_final,
                "Policy Type":       policy_type,
                "Status":            "STP",
                "Payin (CD2)":       f"{satp_payin:.2f}%",
                "Payin Category":    get_payin_category(satp_payin),
                "Calculated Payout": f"{payout:.2f}%",
                "Formula Used":      formula,
                "Rule Explanation":  rule_exp,
                "Remarks":           output_remarks,
            })

    return records

# ------------------- INTELLIGENT DISPATCHER -------------------
def intelligent_dispatcher(df, sheet_name, override_enabled, override_lob, override_segment, override_policy_type):
    pattern = detect_sheet_pattern(df)
    print(f"   [DISPATCHER] Detected pattern: {pattern.upper()}")

    if pattern == "cluster_segment":
        processor_name = "Cluster/Segment/Make Processor (NEW)"
        records = process_cluster_segment_sheet(df, override_enabled, override_lob, override_segment, override_policy_type)
    elif pattern == "electric":
        processor_name = "Electric Vehicle Processor"
        records = process_electric_sheet(df, override_enabled, override_lob, override_segment, override_policy_type)
    elif pattern == "compact":
        processor_name = "Compact Sheet Processor"
        records = process_compact_sheet(df, override_enabled, override_lob, override_segment, override_policy_type)
    else:
        processor_name = "Regular Sheet Processor"
        records = process_regular_sheet(df, override_enabled, override_lob, override_segment, override_policy_type)

    print(f"   [DISPATCHER] Used: {processor_name}")
    print(f"   [DISPATCHER] Extracted {len(records)} records")
    return records, processor_name, pattern


# ------------------- API ENDPOINTS -------------------
@app.get("/")
async def root():
    return {
        "message": "TAXI Insurance Policy Processing API",
        "version": "2.0.0 (Fixed)",
        "fix": "Resolved 'expected str instance, float found' error",
        "endpoints": {
            "/taxi": "Process TAXI insurance data",
            "/get-sheets": "Get worksheet names from Excel file"
        }
    }


@app.post("/get-sheets")
async def get_sheets(file: UploadFile = File(...)):
    try:
        contents = await file.read()
        xls = pd.ExcelFile(io.BytesIO(contents))
        sheet_names = xls.sheet_names
        return JSONResponse(content={
            "success": True,
            "sheets": sheet_names,
            "total_sheets": len(sheet_names),
            "message": f"Found {len(sheet_names)} worksheet(s)"
        })
    except Exception as e:
        return JSONResponse(status_code=400, content={
            "success": False,
            "error": f"Error reading Excel file: {str(e)}",
            "sheets": [],
            "total_sheets": 0
        })


@app.post("/taxi")
async def process_taxi(
    file: UploadFile = File(...),
    company_name: str = Form("Digit"),
    sheet_name: Optional[str] = Form(None),
    override_segment: Optional[str] = Form(None),
    override_policy_type: Optional[str] = Form(None)
):
    try:
        contents = await file.read()

        try:
            xls = pd.ExcelFile(io.BytesIO(contents))
            sheet_names = xls.sheet_names
            print(f"[INFO] File has {len(sheet_names)} worksheet(s): {sheet_names}")

            sheets_to_process = []

            if sheet_name:
                # User specified a sheet — validate and process only that one
                if sheet_name not in sheet_names:
                    return JSONResponse(status_code=400, content={
                        "success": False,
                        "error": f"Worksheet '{sheet_name}' not found in the file.",
                        "available_sheets": sheet_names
                    })
                sheets_to_process = [sheet_name]
                print(f"[INFO] Processing user-selected worksheet: {sheet_name}")
            else:
                # No sheet specified — process ALL sheets regardless of count
                sheets_to_process = sheet_names
                print(f"[INFO] No sheet specified — processing ALL {len(sheet_names)} worksheet(s): {sheet_names}")

            all_records = []
            processors_used = []
            patterns_detected = []

            for sheet in sheets_to_process:
                print(f"\n[PROCESSING] Sheet: {sheet}")
                df = pd.read_excel(io.BytesIO(contents), sheet_name=sheet, header=None)

                records, processor_name, pattern = intelligent_dispatcher(
                    df,
                    sheet,
                    override_enabled=False,
                    override_lob="TAXI",
                    override_segment=override_segment,
                    override_policy_type=override_policy_type
                )

                all_records.extend(records)
                processors_used.append(f"{sheet}: {processor_name} ({pattern})")
                patterns_detected.append(pattern)

            if not all_records:
                return JSONResponse(status_code=400, content={
                    "success": False,
                    "error": "No processable data found in the uploaded file",
                    "message": "The selected worksheet does not contain valid TAXI data."
                })

            result_df = pd.DataFrame(all_records)

            payin_values = []
            for record in all_records:
                try:
                    payin_values.append(float(record.get("Payin (CD2)", "0%").replace("%", "")))
                except:
                    pass

            avg_payin = round(sum(payin_values) / len(payin_values), 2) if payin_values else 0

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                result_df.to_excel(writer, index=False, sheet_name="Processed Data")
            output.seek(0)
            excel_base64 = base64.b64encode(output.read()).decode()

            csv_output = result_df.to_csv(index=False)

            formula_summary = {}
            for record in all_records:
                formula = record.get("Formula Used", "Unknown")
                formula_summary[formula] = formula_summary.get(formula, 0) + 1

            pattern_summary = {}
            for p in patterns_detected:
                pattern_summary[p] = pattern_summary.get(p, 0) + 1

            return JSONResponse(content={
                "success": True,
                "company_name": company_name,
                "lob": "TAXI",
                "sheet_processed": sheets_to_process[0] if len(sheets_to_process) == 1 else f"All {len(sheets_to_process)} sheets",
                "sheets_processed": sheets_to_process,
                "total_sheets_in_file": len(sheet_names),
                "processors_used": processors_used,
                "patterns_detected": pattern_summary,
                "total_records": len(all_records),
                "avg_payin": avg_payin,
                "unique_segments": len(result_df["Mapped Segment"].unique()) if "Mapped Segment" in result_df.columns else 0,
                "calculated_data": all_records,
                "formula_data": FORMULA_DATA,
                "formula_summary": formula_summary,
                "excel_data": excel_base64,
                "csv_data": csv_output,
                "extracted_text": f"Processed {len(all_records)} TAXI records from {len(sheets_to_process)} sheet(s): {', '.join(sheets_to_process)}",
                "parsed_data": all_records[:5]
            })

        except Exception as e:
            import traceback
            traceback.print_exc()
            return JSONResponse(status_code=400, content={
                "success": False,
                "error": f"Error processing Excel file: {str(e)}"
            })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return JSONResponse(status_code=500, content={
            "success": False,
            "error": f"Server error: {str(e)}"
        })


if __name__ == "__main__":
    import uvicorn
    print("\n" + "=" * 70)
    print("TAXI INSURANCE PROCESSOR API - v2.0.0 (Fixed)")
    print("Fix: 'expected str instance, float found' resolved")
    print("=" * 70 + "\n")
    uvicorn.run(app, host="0.0.0.0", port=8000)
