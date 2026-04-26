import gspread
from google.oauth2.service_account import Credentials
import re
import time

# --- Config ---
SPREADSHEET_ID = "1Xa8l50qpEtt8a4TMq_VMwCKNYpi2UbpEZ2P4lNz_SYo"
CREDENTIALS_FILE = "credentials.json"
MASTER_SHEET_NAME = "Ingredients Master list"
MEALS = ["Breakfast", "Pre-Workout", "Post-Workout", "Lunch", "Snack", "Dinner", "All Day"]

# --- Auth ---
scopes = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
client = gspread.authorize(creds)

# --- Connect ---
spreadsheet = client.open_by_key(SPREADSHEET_ID)
master_sheet = spreadsheet.worksheet(MASTER_SHEET_NAME)

# --- Add Calories to master sheet if missing ---
def ensure_calories_in_master(sheet):
    headers = sheet.row_values(1)
    if "Calories" not in headers:
        print("Adding Calories column to master sheet...")
        col = len(headers) + 1
        updates = [["Calories"]]
        data = sheet.get_all_records()
        for row in data:
            calories = round((row["Protein (g)"] * 4) + (row["Carbs (g)"] * 4) + (row["Fat (g)"] * 9), 1)
            updates.append([calories])
        col_letter = chr(64 + col)
        sheet.update(f"{col_letter}1", updates)
        print(f"Calories added for {len(data)} ingredients.")
    else:
        print("Calories column already exists in master sheet.")

# --- Create meal plan tab ---
def create_meal_plan(plan_name, rows_per_meal=8):
    existing = [ws.title for ws in spreadsheet.worksheets()]
    if plan_name in existing:
        print(f"Tab '{plan_name}' already exists. Deleting and recreating...")
        spreadsheet.del_worksheet(spreadsheet.worksheet(plan_name))

    ws = spreadsheet.add_worksheet(title=plan_name, rows=300, cols=10)
    ingredients = master_sheet.get_all_records()
    ingredient_names = [row["Ingredient"] for row in ingredients]

    # Build all cell data in memory first
    all_values = {}  # {(row, col): value}

    # Headers
    headers = ["Meal", "Ingredient", "Reference", "Protein (g)", "Fat (g)", "Carbs (g)", "Multiplier", "Quantity", "Calories"]
    for ci, h in enumerate(headers, 1):
        all_values[(1, ci)] = h

    current_row = 2
    subtotal_rows = []

    for meal in MEALS:
        # Meal header
        all_values[(current_row, 1)] = meal
        current_row += 1

        ing_start = current_row
        ing_end = current_row + rows_per_meal - 1

        # Ingredient rows - formulas
        master = f"'{MASTER_SHEET_NAME}'!$A:$F"
        for row_num in range(ing_start, ing_end + 1):
            b = f"B{row_num}"
            g = f"G{row_num}"
            all_values[(row_num, 3)] = f'=IF({b}="","",VLOOKUP({b},{master},5,FALSE))'
            all_values[(row_num, 4)] = f'=IF(OR({b}="",{g}=""),"",ROUND(VLOOKUP({b},{master},2,FALSE)*{g},1))'
            all_values[(row_num, 5)] = f'=IF(OR({b}="",{g}=""),"",ROUND(VLOOKUP({b},{master},3,FALSE)*{g},1))'
            all_values[(row_num, 6)] = f'=IF(OR({b}="",{g}=""),"",ROUND(VLOOKUP({b},{master},4,FALSE)*{g},1))'
            all_values[(row_num, 9)] = f'=IF(OR({b}="",{g}=""),"",ROUND(D{row_num}*4+F{row_num}*4+E{row_num}*9,1))'
            all_values[(row_num, 8)] = f'=IF(OR({b}="",{g}=""),"",{g}&" x "&VLOOKUP({b},{master},5,FALSE))'

        current_row = ing_end + 1

        # Subtotal row
        subtotal_row = current_row
        subtotal_rows.append((meal, subtotal_row, ing_start, ing_end))
        all_values[(subtotal_row, 1)] = f"{meal} Total"
        for col, col_letter in [(4, "D"), (5, "E"), (6, "F"), (9, "I")]:
            all_values[(subtotal_row, col)] = f"=SUM({col_letter}{ing_start}:{col_letter}{ing_end})"

        current_row += 2  # blank row after subtotal

    # Grand total row
    grand_total_row = current_row
    all_values[(grand_total_row, 1)] = "TOTAL"
    all_values[(grand_total_row, 4)] = f"=SUM({','.join([f'D{r}' for _, r, _, _ in subtotal_rows])})"
    all_values[(grand_total_row, 5)] = f"=SUM({','.join([f'E{r}' for _, r, _, _ in subtotal_rows])})"
    all_values[(grand_total_row, 6)] = f"=SUM({','.join([f'F{r}' for _, r, _, _ in subtotal_rows])})"
    all_values[(grand_total_row, 9)] = f"=SUM({','.join([f'I{r}' for _, r, _, _ in subtotal_rows])})"

    # Macros in calories row
    macro_cal_row = grand_total_row + 1
    all_values[(macro_cal_row, 1)] = "Macros in Calories"
    all_values[(macro_cal_row, 4)] = f"=D{grand_total_row}*4"
    all_values[(macro_cal_row, 5)] = f"=E{grand_total_row}*9"
    all_values[(macro_cal_row, 6)] = f"=F{grand_total_row}*4"

    # Convert to batch update format
    max_row = max(r for r, c in all_values)
    max_col = max(c for r, c in all_values)
    grid = [[""] * max_col for _ in range(max_row)]
    for (r, c), v in all_values.items():
        grid[r - 1][c - 1] = v

    print("Writing meal plan structure...")
    ws.update("A1", grid, value_input_option="USER_ENTERED")
    print("Structure written.")

    # Dropdowns via batch_update
    print("Adding ingredient dropdowns...")
    requests = []
    for _, subtotal_row, ing_start, ing_end in subtotal_rows:
        requests.append({
            "setDataValidation": {
                "range": {
                    "sheetId": ws.id,
                    "startRowIndex": ing_start - 1,
                    "endRowIndex": ing_end,
                    "startColumnIndex": 1,
                    "endColumnIndex": 2,
                },
                "rule": {
                    "condition": {
                        "type": "ONE_OF_LIST",
                        "values": [{"userEnteredValue": name} for name in ingredient_names],
                    },
                    "showCustomUi": True,
                    "strict": False,
                },
            }
        })
    spreadsheet.batch_update({"requests": requests})
    print("Dropdowns added.")

    # Set fixed column widths
    col_widths = [120, 160, 120, 90, 90, 90, 90, 120, 90]  # A through I
    width_requests = []
    for i, width in enumerate(col_widths):
        width_requests.append({
            "updateDimensionProperties": {
                "range": {
                    "sheetId": ws.id,
                    "dimension": "COLUMNS",
                    "startIndex": i,
                    "endIndex": i + 1
                },
                "properties": {"pixelSize": width},
                "fields": "pixelSize"
            }
        })
    spreadsheet.batch_update({"requests": width_requests})
    print("Column widths set.")

    print(f"\nMeal plan '{plan_name}' created successfully.")
    print(f"Open your sheet, go to the '{plan_name}' tab, pick ingredients from dropdowns and enter multipliers.")

# --- Main ---
ensure_calories_in_master(master_sheet)
plan_name = input("Enter meal plan name (this becomes the tab name): ")
create_meal_plan(plan_name)