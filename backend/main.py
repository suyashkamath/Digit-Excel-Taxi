from fastapi import FastAPI, File, UploadFile, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
import pandas as pd
import io
import base64
from typing import Optional

app = FastAPI(title="TAXI Insurance Policy Processor")

# Enable CORS
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

# ------------------- PAYOUT LOGIC -------------------
def get_payin_category(payin: float):
    if payin <= 20: return "Payin Below 20%"
    elif payin <= 30: return "Payin 21% to 30%"
    elif payin <= 50: return "Payin 31% to 50%"
    else: return "Payin Above 50%"

def safe_float(value):
    if pd.isna(value): return None
    val_str = str(value).strip().upper().replace('%', '')
    if val_str in ["D", "NA", "", "NAN", "NONE"]: return None
    try:
        num = float(val_str)
        return num * 100 if 0 < num < 1 else num
    except: return None

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

# ------------------- PATTERN DETECTION -------------------
def detect_sheet_pattern(df):
    """
    Detects which processing pattern the sheet follows:
    - 'electric': Electric vehicle sheet (simpler structure)
    - 'regular': Regular taxi sheet with multiple columns (6+ combinations)
    - 'compact': Compact sheet with just 2 CD2 columns
    """
    # Check first few rows for patterns
    first_10_rows = df.head(10).astype(str).apply(lambda x: ' '.join(x), axis=1).str.upper()
    all_text = ' '.join(first_10_rows)
    
    # Electric sheet detection
    if 'ELECTRIC' in all_text or 'EV' in all_text:
        # Check if it has simple structure (fewer columns)
        if df.shape[1] <= 10:
            return 'electric'
    
    # Check for CD2 columns
    cd2_count = 0
    for col_idx in range(df.shape[1]):
        for r in range(min(10, df.shape[0])):
            cell = str(df.iloc[r, col_idx]).strip().upper()
            if "CD2" in cell:
                cd2_count += 1
                break
    
    # Compact pattern: typically has exactly 2 CD2 columns
    if cd2_count == 2:
        # Check if it has fewer detail columns (compact format)
        if df.shape[1] <= 8:
            return 'compact'
    
    # Regular pattern: has many columns for different combinations
    if df.shape[1] >= 12:
        return 'regular'
    
    # Default to regular if uncertain
    return 'regular'

# ------------------- PROCESSORS -------------------
def process_electric_sheet(df, override_enabled, override_lob, override_segment, override_policy_type):
    """Process electric vehicle taxi sheets"""
    records = []
    for _, row in df.iterrows():
        if pd.isna(row.iloc[0]) or str(row.iloc[0]).strip() == "":
            continue
        
        city_cluster = str(row.iloc[0]).strip()
        rto_remarks = str(row.iloc[1]).strip() if len(row) > 1 else ""
        fuel = str(row.iloc[2]).strip() if len(row) > 2 else "Electric"
        make = str(row.iloc[3]).strip() if len(row) > 3 else "All"
        seating = str(row.iloc[4]).strip() if len(row) > 4 else "5"
        
        state = next((v for k, v in STATE_MAPPING.items() if k.upper() in city_cluster.upper()), "UNKNOWN")
        cvod_cd2 = safe_float(row.iloc[6]) if len(row) > 6 else None
        cvtp_cd2 = safe_float(row.iloc[7]) if len(row) > 7 else None
        
        segment_desc = f"Taxi {fuel} {make} {rto_remarks} Seating:{seating}".strip()
        
        lob_final = override_lob if override_enabled and override_lob else "TAXI"
        segment_final = override_segment if override_enabled and override_segment else "TAXI"
        
        if cvod_cd2 is not None:
            policy_type = override_policy_type or "Comp"
            payout, formula, rule_exp = calculate_payout_with_formula(lob_final, segment_final, policy_type, cvod_cd2)
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
            payout, formula, rule_exp = calculate_payout_with_formula(lob_final, segment_final, "TP", cvtp_cd2)
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
    """Process regular taxi sheets with multiple column combinations"""
    records = []
    prev_location = ""
    
    for idx, row in df.iterrows():
        if idx < 5:  # Skip header rows
            continue
        
        location = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else prev_location
        if location: prev_location = location
        
        fuel = str(row.iloc[1]).strip() if len(row) > 1 else ""
        make = str(row.iloc[2]).strip() if len(row) > 2 else ""
        remarks = str(row.iloc[3]).strip() if len(row) > 3 else ""
        seating = str(row.iloc[4]).strip() if len(row) > 4 else ""
        
        state = next((v for k, v in STATE_MAPPING.items() if k.upper() in location.upper()), "UNKNOWN")
        cols = row.values
        
        # Column combinations for different policy types
        combinations = [
            ("Without Add On Cover", "<=1000 CC", "Comp", 5, 6),
            ("Without Add On Cover", ">1000 CC",  "Comp", 7, 8),
            ("With Add On Cover",    "<=1000 CC", "Comp", 9, 10),
            ("With Add On Cover",    ">1000 CC",  "Comp", 11, 12),
            ("", "<=1000 CC", "TP", None, 13),
            ("", ">1000 CC",  "TP", None, 14),
        ]
        
        for addon, cc, ptype, _, cd2_idx in combinations:
            if cd2_idx >= len(cols): continue
            payin = safe_float(cols[cd2_idx])
            if payin is None: continue
            
            segment_desc = f"Taxi {fuel} {make} {remarks}".strip()
            if seating: segment_desc += f" Seating:{seating}"
            if cc: segment_desc += f" {cc}"
            if addon: segment_desc += f" {addon}"
            
            lob_final = override_lob if override_enabled and override_lob else "TAXI"
            segment_final = override_segment if override_enabled and override_segment else "TAXI"
            policy_type_final = override_policy_type or ptype
            
            payout, formula, rule_exp = calculate_payout_with_formula(lob_final, segment_final, policy_type_final, payin)
            records.append({
                "State": state.upper(),
                "Location/Cluster": location,
                "Original Segment": segment_desc.strip(),
                "Mapped Segment": segment_final,
                "LOB": lob_final,
                "Policy Type": policy_type_final,
                "Payin (CD2)": f"{payin:.2f}%",
                "Payin Category": get_payin_category(payin),
                "Calculated Payout": f"{payout:.2f}%",
                "Formula Used": formula,
                "Rule Explanation": rule_exp
            })
    return records

def process_compact_sheet(df, override_enabled, override_lob, override_segment, override_policy_type):
    """Process compact taxi sheets with dynamic CD2 column detection"""
    records = []
    cluster = ""
    
    comp_cd2_col = None
    satp_cd2_col = None
    
    # Scan all columns for headers containing "CD2"
    cd2_candidates = []
    for col_idx in range(df.shape[1]):
        for r in range(min(10, df.shape[0])):
            cell = str(df.iloc[r, col_idx]).strip().upper()
            if "CD2" in cell:
                cd2_candidates.append((col_idx, r))
                break
    
    if len(cd2_candidates) < 2:
        return records
    
    # Assign Comp and SATP based on nearby cells
    for col_idx, row_idx in cd2_candidates:
        group = ""
        # Check left/right in same row
        for dcol in [-1, 1]:
            if 0 <= col_idx + dcol < df.shape[1]:
                nearby = str(df.iloc[row_idx, col_idx + dcol]).strip().upper()
                if nearby:
                    group = nearby
                    break
        # Check above
        if not group:
            for dr in range(1, 5):
                if row_idx - dr >= 0:
                    above = str(df.iloc[row_idx - dr, col_idx]).strip().upper()
                    if above:
                        group = above
                        break
        # Check diagonal (left columns above)
        if not group:
            for dcol in [-1, -2]:
                if 0 <= col_idx + dcol < df.shape[1]:
                    for dr in range(1, 5):
                        if row_idx - dr >= 0:
                            above_left = str(df.iloc[row_idx - dr, col_idx + dcol]).strip().upper()
                            if above_left:
                                group = above_left
                                break
        
        if "COMP" in group or "OD" in group:
            comp_cd2_col = col_idx
        elif "SATP" in group or "TP" in group:
            satp_cd2_col = col_idx
    
    if comp_cd2_col is None or satp_cd2_col is None:
        return records
    
    # Find data start (first row with numeric in Comp column)
    data_start = 0
    for r in range(df.shape[0]):
        if safe_float(df.iloc[r, comp_cd2_col]) is not None:
            data_start = r
            break
    
    for r in range(data_start, df.shape[0]):
        row = df.iloc[r]
        
        if pd.notna(row.iloc[0]) and str(row.iloc[0]).strip():
            cluster = str(row.iloc[0]).strip()
        
        if df.shape[1] < 2 or pd.isna(row.iloc[1]) or str(row.iloc[1]).strip() == "":
            continue
        
        segment = str(row.iloc[1]).strip()
        make = str(row.iloc[2]).strip() if df.shape[1] > 2 and pd.notna(row.iloc[2]) else ""
        remarks = str(row.iloc[-1]).strip() if pd.notna(row.iloc[-1]) else ""
        
        segment_desc = f"Taxi {segment} {make}".strip()
        if remarks: segment_desc += f" {remarks}"
        
        state = next((v for k, v in STATE_MAPPING.items() if k.upper() in cluster.upper()), "UNKNOWN")
        
        lob_final = override_lob if override_enabled and override_lob else "TAXI"
        segment_final = override_segment if override_enabled and override_segment else "TAXI"
        
        comp_payin = safe_float(row.iloc[comp_cd2_col])
        if comp_payin is not None:
            policy_type = override_policy_type or "Comp"
            payout, formula, rule_exp = calculate_payout_with_formula(lob_final, segment_final, policy_type, comp_payin)
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
            payout, formula, rule_exp = calculate_payout_with_formula(lob_final, segment_final, "TP", tp_payin)
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

# ------------------- INTELLIGENT DISPATCHER -------------------
def intelligent_dispatcher(df, sheet_name, override_enabled, override_lob, override_segment, override_policy_type):
    """
    Intelligent dispatcher that detects the Excel pattern and calls the appropriate processor
    """
    pattern = detect_sheet_pattern(df)
    
    print(f"   [DISPATCHER] Detected pattern: {pattern.upper()}")
    
    if pattern == 'electric':
        processor_name = "Electric Vehicle Processor"
        records = process_electric_sheet(df, override_enabled, override_lob, override_segment, override_policy_type)
    elif pattern == 'compact':
        processor_name = "Compact Sheet Processor"
        records = process_compact_sheet(df, override_enabled, override_lob, override_segment, override_policy_type)
    else:  # regular
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
        "endpoints": {
            "/taxi": "Process TAXI insurance data",
            "/get-sheets": "Get worksheet names from Excel file"
        },
        "features": [
            "Intelligent pattern detection",
            "Electric vehicle sheet processing",
            "Regular taxi sheet processing",
            "Compact sheet processing",
            "Manual worksheet selection",
            "Automatic processing for single-sheet files"
        ]
    }

@app.post("/get-sheets")
async def get_sheets(file: UploadFile = File(...)):
    """
    Get list of all worksheet names from Excel file
    Returns sheet names for manual selection
    """
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
        return JSONResponse(
            status_code=400,
            content={
                "success": False,
                "error": f"Error reading Excel file: {str(e)}",
                "sheets": [],
                "total_sheets": 0
            }
        )

@app.post("/taxi")
async def process_taxi(
    file: UploadFile = File(...),
    company_name: str = Form("Digit"),
    sheet_name: Optional[str] = Form(None),
    override_segment: Optional[str] = Form(None),
    override_policy_type: Optional[str] = Form(None)
):
    """
    Process TAXI insurance data with intelligent pattern detection
    
    Logic:
    1. If sheet_name is provided -> Process only that specific sheet
    2. If sheet_name is None:
       - If file has only 1 worksheet -> Process it directly
       - If file has multiple worksheets -> Return error asking for sheet selection
    """
    try:
        # Read the uploaded Excel file
        contents = await file.read()
        
        try:
            xls = pd.ExcelFile(io.BytesIO(contents))
            sheet_names = xls.sheet_names
            
            print(f"[INFO] File has {len(sheet_names)} worksheet(s): {sheet_names}")
            
            # ==================== WORKSHEET SELECTION LOGIC ====================
            sheets_to_process = []
            
            if sheet_name:
                # User specified a sheet name - validate and process only that sheet
                if sheet_name not in sheet_names:
                    return JSONResponse(
                        status_code=400,
                        content={
                            "success": False,
                            "error": f"Worksheet '{sheet_name}' not found in the file.",
                            "available_sheets": sheet_names,
                            "message": f"Available worksheets: {', '.join(sheet_names)}"
                        }
                    )
                sheets_to_process = [sheet_name]
                print(f"[INFO] Processing user-selected worksheet: {sheet_name}")
                
            else:
                # No sheet name provided - check number of worksheets
                if len(sheet_names) == 1:
                    # Only one worksheet - process it directly
                    sheets_to_process = sheet_names
                    print(f"[INFO] Single worksheet detected - processing directly: {sheet_names[0]}")
                    
                else:
                    # Multiple worksheets - require manual selection
                    return JSONResponse(
                        status_code=400,
                        content={
                            "success": False,
                            "error": "Multiple worksheets found. Please select a worksheet to process.",
                            "available_sheets": sheet_names,
                            "total_sheets": len(sheet_names),
                            "message": f"This file contains {len(sheet_names)} worksheets. Please select one to process.",
                            "require_sheet_selection": True
                        }
                    )
            
            # ==================== PROCESS SELECTED SHEET(S) ====================
            all_records = []
            processors_used = []
            patterns_detected = []
            
            for sheet in sheets_to_process:
                print(f"\n[PROCESSING] Sheet: {sheet}")
                df = pd.read_excel(io.BytesIO(contents), sheet_name=sheet, header=None)
                
                # Use intelligent dispatcher
                records, processor_name, pattern = intelligent_dispatcher(
                    df, 
                    sheet, 
                    override_enabled=False,  # Not using overrides for TAXI
                    override_lob="TAXI",      # Always TAXI
                    override_segment=override_segment,
                    override_policy_type=override_policy_type
                )
                
                all_records.extend(records)
                processors_used.append(f"{sheet}: {processor_name} ({pattern})")
                patterns_detected.append(pattern)
            
            # ==================== VALIDATION ====================
            if not all_records:
                return JSONResponse(
                    status_code=400,
                    content={
                        "success": False,
                        "error": "No processable data found in the uploaded file",
                        "message": "The selected worksheet does not contain valid TAXI data or the format is not recognized."
                    }
                )
            
            # ==================== GENERATE OUTPUT ====================
            # Create result DataFrame
            result_df = pd.DataFrame(all_records)
            
            # Calculate statistics
            payin_values = []
            for record in all_records:
                payin_str = record.get("Payin (CD2)", "0%")
                try:
                    payin_values.append(float(payin_str.replace("%", "")))
                except:
                    pass
            
            avg_payin = round(sum(payin_values) / len(payin_values), 2) if payin_values else 0
            unique_segments = len(result_df["Mapped Segment"].unique()) if "Mapped Segment" in result_df.columns else 0
            
            # Generate Excel file
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                result_df.to_excel(writer, index=False, sheet_name='Processed Data')
            output.seek(0)
            excel_base64 = base64.b64encode(output.read()).decode()
            
            # Generate CSV
            csv_output = result_df.to_csv(index=False)
            
            # Formula summary
            formula_summary = {}
            for record in all_records:
                formula = record.get("Formula Used", "Unknown")
                formula_summary[formula] = formula_summary.get(formula, 0) + 1
            
            # Pattern summary
            pattern_summary = {}
            for pattern in patterns_detected:
                pattern_summary[pattern] = pattern_summary.get(pattern, 0) + 1
            
            # ==================== RETURN RESPONSE ====================
            return JSONResponse(content={
                "success": True,
                "company_name": company_name,
                "lob": "TAXI",
                "sheet_processed": sheets_to_process[0] if sheets_to_process else "Unknown",
                "total_sheets_in_file": len(sheet_names),
                "processors_used": processors_used,
                "patterns_detected": pattern_summary,
                "total_records": len(all_records),
                "avg_payin": avg_payin,
                "unique_segments": unique_segments,
                "calculated_data": all_records,
                "formula_data": FORMULA_DATA,
                "formula_summary": formula_summary,
                "excel_data": excel_base64,
                "csv_data": csv_output,
                "extracted_text": f"Processed {len(all_records)} TAXI records from worksheet '{sheets_to_process[0]}' using intelligent pattern detection",
                "parsed_data": all_records[:5]  # First 5 records for preview
            })
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            return JSONResponse(
                status_code=400,
                content={
                    "success": False,
                    "error": f"Error processing Excel file: {str(e)}",
                    "message": "There was an error reading or processing the Excel file. Please check the file format."
                }
            )
    
    except Exception as e:
        import traceback
        traceback.print_exc()
        return JSONResponse(
            status_code=500,
            content={
                "success": False,
                "error": f"Server error: {str(e)}",
                "message": "An internal server error occurred. Please try again."
            }
        )

if __name__ == "__main__":
    import uvicorn
    print("\n" + "="*80)
    print(" "*25 + "TAXI INSURANCE PROCESSOR API")
    print("="*80)
    print("\nStarting server on http://localhost:8000")
    print("\nFeatures:")
    print("  ✓ Manual worksheet selection for multi-sheet files")
    print("  ✓ Automatic processing for single-sheet files")
    print("  ✓ Intelligent pattern detection (Electric/Regular/Compact)")
    print("  ✓ Real-time payout calculation")
    print("\nEndpoints:")
    print("  • GET  /          - API information")
    print("  • POST /get-sheets - Get worksheet names")
    print("  • POST /taxi       - Process TAXI data")
    print("="*80 + "\n")
    
    uvicorn.run(app, host="0.0.0.0", port=8000)
